(ns malcolmx.core
  (:require [clojure.java.io :as io]
            [clojure.tools.logging :as log])
  (:import [org.apache.poi.ss.usermodel WorkbookFactory Workbook Sheet Cell Row FormulaEvaluator]
           [java.util List]))

(defn sheet-header [^Sheet sheet]
  (let [row (.getRow sheet 0)]
    (mapv #(.getStringCellValue ^Cell %) row)))

(defn set-row! [^Row row header row-data]
  (mapv (fn [^Cell cell header-name]
          (when-let [value (get row-data header-name)]
            (.setCellValue cell value)))
        row header))

(defn formula-eval
  "eval excel formulas"
  [^Cell cell ^FormulaEvaluator evaluator]
  (try
    (.evaluate evaluator cell)
    (catch Exception e
      (log/errorf e "Cant eval formula '%s'\nSheet: %s\nRow: %d\nCell: %d)" (.getCellFormula cell)
                  (.getSheetName (.getSheet cell))
                  (.getRowIndex cell)
                  (.getColumnIndex cell)))))

(defn error-code [code]
  (condp = code
    15 "VALUE!"
    7  "DIV/0"
    (str "formula error:" code )))

(defn cell-value [^FormulaEvaluator evaluator ^Cell cell]
  (condp = (.getCellType cell)
    Cell/CELL_TYPE_FORMULA    (cell-value (formula-eval cell evaluator) evaluator)
    Cell/CELL_TYPE_NUMERIC    (.getNumericCellValue cell)
    Cell/CELL_TYPE_STRING     (.getStringCellValue cell)
    Cell/CELL_TYPE_BOOLEAN    (.getBooleanCellValue cell)
    Cell/CELL_TYPE_ERROR      (error-code (.getErrorCellValue cell))
    Cell/CELL_TYPE_BLANK      nil
    (log/errorf "Undef cell type: %" (.getCellType cell))))

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
        evaluator (-> workbook
                      (.getCreationHelper)
                      (.createFormulaEvaluator))]
    (loop [index 1
           data sheet-data]
      (when (seq data)
        (let [row (.getRow sheet index)
              row-data (first data)]
          (if (= (cell-value evaluator (.getCell row id-index))
                 (get row-data id-column))
            (do
              (set-row! row header row-data)
              (recur (inc index) (rest data)))
            (recur (inc index) data)))))))

(defn get-sheet [^Workbook workbook sheet-name]
  (let [sheet (.getSheet workbook sheet-name)
        header (sheet-header sheet)
        evaluator (-> workbook
                      (.getCreationHelper)
                      (.createFormulaEvaluator))]
    (map (fn [row]
           (zipmap header (map (partial cell-value evaluator)
                               row)))
         (rest sheet))))