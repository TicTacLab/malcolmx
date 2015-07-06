(ns malcolmx.core
  (:require [clojure.java.io :as io]
            [clojure.tools.logging :as log]
            [malcolmx.math :as math])
  (:import [org.apache.poi.ss.usermodel WorkbookFactory Workbook Sheet Cell Row FormulaEvaluator]
           [java.util List]
           [org.apache.poi.ss.util CellReference]))

(defn sheet-header [^Sheet sheet]
  (let [row (.getRow sheet 0)]
    (mapv #(.getStringCellValue ^Cell %) row)))

(defn set-row! [^Row row header row-data]
  (mapv (fn [^Cell cell header-name]
          (when-let [value (get row-data header-name)]
            (.setCellValue cell value)))
        row header))

(defn make-evaluator [^Workbook workbook]
  (math/register-udf-funs! workbook)
  (-> workbook
      (.getCreationHelper)
      (.createFormulaEvaluator)))

(defn error-code [code]
  (condp = code
    15 "VALUE!"
    7  "DIV/0"
    (str "formula error:" code )))

(defn formula-cell-value
  "eval excel formulas"
  [^FormulaEvaluator evaluator ^Cell cell ]
  (try
    (let [cell-value (.evaluate evaluator cell)]
      (condp = (.getCellType cell-value)
        Cell/CELL_TYPE_NUMERIC (.getNumberValue cell-value)
        Cell/CELL_TYPE_STRING (.getStringValue cell-value)
        Cell/CELL_TYPE_BOOLEAN (.getBooleanValue cell-value)
        Cell/CELL_TYPE_ERROR (error-code (.getErrorValue cell-value))
        (log/errorf "Undefined cell type: %" (.getCellType cell-value))))
    (catch Exception e
      (log/errorf e "Can't evaluate\nCell: '%s!%s'\nFormula: '%s'\n"
                  (.getSheetName (.getSheet cell))
                  (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell)))
                  (.getCellFormula cell)))))

(defn cell-value [^FormulaEvaluator evaluator ^Cell cell]
  (condp = (.getCellType cell)
    Cell/CELL_TYPE_FORMULA (formula-cell-value evaluator cell)
    Cell/CELL_TYPE_NUMERIC (.getNumericCellValue cell)
    Cell/CELL_TYPE_STRING (.getStringCellValue cell)
    Cell/CELL_TYPE_BOOLEAN (.getBooleanCellValue cell)
    Cell/CELL_TYPE_ERROR (error-code (.getErrorCellValue cell))
    Cell/CELL_TYPE_BLANK nil
    (log/errorf "Undefined cell type: %" (.getCellType cell))))


(defmacro with-timer-when [profile? & body]
  `(if ~profile?
     (let [start# (System/nanoTime)
           ret# (do ~@body)
           time-call#
           (/
             (double
               (- (System/nanoTime) start#))
             1000000.0)]
       (assoc ret# "timer" (int time-call#)))
     (do ~@body)))

;; PUBLIC API

(defn parse
  "Parses coerceable to the InputStream input
   (see clojure.java.io/input-stream) to the Workbook."
  [input]
  (with-open [in (io/input-stream input)]
    (WorkbookFactory/create in)))

(defn update-sheet!
  "Update sheet with sheet-name at workbook using sheet-data.
   It updates corresponding rows with sheet-data using id-column
   as index. It updates only those rows and cells that are present
   at sheet-data, otherwise left them untouched."
  [^Workbook workbook sheet-name sheet-data & {id-column :by}]
  (let [sheet (.getSheet workbook sheet-name)
        header (sheet-header sheet)
        id-index (.indexOf ^List header id-column)
        evaluator (make-evaluator workbook)]
    (loop [index 1
           data (sort-by #(get % id-column) sheet-data)]
      (when (seq data)
        (let [row (.getRow sheet index)
              row-data (first data)]
          (if (= (cell-value evaluator (.getCell row id-index))
                 (get row-data id-column))
            (do
              (set-row! row header row-data)
              (recur (inc index) (rest data)))
            (recur (inc index) data)))))))

(defn get-sheet [^Workbook workbook sheet-name & {profile? :profile?}]
  (let [sheet (.getSheet workbook sheet-name)
        header (sheet-header sheet)
        evaluator (make-evaluator workbook)]
    (map (fn [row]
           (with-timer-when profile?
             (zipmap header (map (partial cell-value evaluator)
                                 row))))
         (rest sheet))))