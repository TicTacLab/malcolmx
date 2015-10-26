(ns malcolmx.core
  (:require [clojure.java.io :as io]
            [clojure.tools.logging :as log]
            [malcolmx.math :as math]
            [clojure.core.typed :refer [ann non-nil-return] :as t])
  (:import [org.apache.poi.ss.usermodel WorkbookFactory Workbook Sheet Cell Row FormulaEvaluator FormulaError]
           [java.util List]
           [org.apache.poi.ss.util CellReference]
           [clojure.java.io IOFactory]
           [java.io InputStream BufferedInputStream]
           [clojure.lang Seqable Fn]
           (org.apache.poi POIXMLDocument)
           (org.apache.poi.poifs.filesystem POIFSFileSystem)
           (java.lang.reflect Array)))

(defmacro nil-return [& body]
  `(do (t/tc-ignore ~@body)
      nil))

(t/defalias ColumnName String)
(t/defalias CellValue (t/U Number String Boolean nil))
(t/defalias RowData (t/Map ColumnName CellValue))
(t/defalias SheetData (Seqable RowData))
(t/defalias Header (t/Vec ColumnName))
(t/defalias SheetName String)
(t/defalias SheetNames [SheetName])


(defprotocol CellProtocol
  (set-value [this] "set cell value"))

(extend-protocol CellProtocol
  Long
  (set-value [this] (double this))
  Float
  (set-value [this] (double this))
  Double
  (set-value [this] this)
  String
  (set-value [this] this)
  nil
  (set-value [_] ""))

(ann ^:no-check get-cells (Fn [Row -> (Seqable Cell)]
                              [Row Number -> (Seqable (t/U Cell nil))]))
(defn get-cells
  ([^Row row]
   (seq row))
  ([^Row row cell-count]
   (map #(.getCell row %) (range cell-count))))

(ann ^:no-check get-rows [Sheet -> (Seqable Row)])
(defn get-rows [sheet]
  (seq sheet))

(ann ^:no-check index-of [Header ColumnName -> Integer])
(defn index-of [header column-name]
  (.indexOf ^List header column-name))

(ann ^:no-check set-cell-value [Cell CellValue -> nil])
(defn set-cell-value [^Cell cell value]
  (.setCellValue cell (set-value value)))

(ann sheet-header [Sheet -> Header])
(defn sheet-header [^Sheet sheet]
  (if-let [row (.getRow sheet 0)]
    (mapv (t/fn [c :- Cell] :- String
            (.getStringCellValue ^Cell c))
          (get-cells row))
    []))

(ann set-row! [Row Header RowData -> nil])
(defn set-row! [^Row row header row-data]
  (mapv (t/fn [^Cell cell :- (t/U Cell nil) header-name :- ColumnName] :- nil
          (when cell
            (when-let [value (get row-data header-name)]
              (set-cell-value cell value))))
        (get-cells row (count header)) header)
  nil)

(ann make-evaluator [Workbook -> FormulaEvaluator])
(defn make-evaluator [^Workbook workbook]
  (math/register-udf-funs! workbook)
  (-> workbook
      (.getCreationHelper)
      (.createFormulaEvaluator)))

(ann error-code [Byte -> String])
(defn error-code [code]
  (.getString (FormulaError/forInt ^byte code)))

(ann cell-value [FormulaEvaluator Cell -> CellValue])
(defn cell-value [^FormulaEvaluator evaluator ^Cell cell]
  (try
    (when cell
      (when-let [cell-value (.evaluate evaluator cell)]
        (condp = (.getCellType cell-value)
          Cell/CELL_TYPE_NUMERIC (.getNumberValue cell-value)
          Cell/CELL_TYPE_STRING (.getStringValue cell-value)
          Cell/CELL_TYPE_BOOLEAN (.getBooleanValue cell-value)
          Cell/CELL_TYPE_ERROR (error-code (.getErrorValue cell-value))
          Cell/CELL_TYPE_BLANK nil                          ;; this clause, actually, unreacable, because eveluator returns nil as cell-value for this cell type. Here just for readablility
          (nil-return
            (log/errorf "Undefined cell type: %" (.getCellType cell-value))))))
    (catch Exception e
      (nil-return
        (log/errorf e "Can't evaluate\nCell: '%s!%s'\nValue: '%s'\n"
                    (.getSheetName (.getSheet cell))
                    (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell)))
                    (str cell))))))


(defmacro with-timer-when [profile? & body]
  `(if ~profile?
     (let [start# (System/nanoTime)
           ret# (do ~@body)
           time-call#
           (/
             (double
               (- (System/nanoTime) start#))
             1000000.0)]
       ;; assoc-in instead of simply assoc, becouse of http://dev.clojure.org/jira/browse/CTYP-161
       (assoc ret# "timer" (num time-call#)))
     (do ~@body)))

;; PUBLIC API

(ann clojure.java.io/input-stream [IOFactory -> InputStream])


(ann parse [IOFactory -> Workbook])
(defn parse
  "Parses coerceable to the InputStream input
   (see clojure.java.io/input-stream) to the Workbook."
  [input]
  (with-open [in (io/input-stream input)]
    (WorkbookFactory/create in)))

(ann ^:no-check sort-by-id-column [ColumnName SheetData -> SheetData])
(defn sort-by-id-column [id-column sheet-data]
  (sort-by #(get % id-column) sheet-data))

(ann make-row-data [Header (Seqable CellValue) -> RowData])
(defn make-row-data [header cell-values]
  (zipmap header cell-values))


(ann update-sheet! [Workbook SheetName SheetData & :optional {:by ColumnName} -> Workbook])
(defn ^Workbook update-sheet!
  "Update sheet with sheet-name at workbook using sheet-data.
   It updates corresponding rows with sheet-data using id-column
   as index. It updates only those rows and cells that are present
   at sheet-data, otherwise left them untouched."
  [^Workbook workbook sheet-name sheet-data & {id-column :by}]
  (assert id-column "You should always specify id-column using :by key: (update-sheet! w name data :by \"id\")")
  (if-let [sheet (.getSheet workbook sheet-name)]
    (let [header (sheet-header sheet)
          id-index (index-of header id-column)
          evaluator (make-evaluator workbook)]
      (t/loop [index :- Long 1
               data :- SheetData (sort-by-id-column id-column sheet-data)]
        (when-first [row-data data]
          (if-let [row (.getRow sheet index)]
            (if-let [id-value (.getCell row id-index)]
              (if (= (cell-value evaluator id-value)
                     (get row-data id-column))
                (do
                  (set-row! row header row-data)
                  (recur (inc index) (rest data)))
                (recur (inc index) data))
              (recur (inc index) data))
            (recur (inc index) data))))
      workbook)
    (throw (ex-info "Sheet does not exists" {:sheet-name sheet-name
                                             :workbook   workbook}))))
(ann get-sheet [Workbook SheetName & :optional {:profile? Boolean} -> SheetData])
(defn get-sheet [^Workbook workbook sheet-name & {profile? :profile?}]
  (if-let [sheet (.getSheet workbook sheet-name)]
    (let [header (sheet-header sheet)
          evaluator (make-evaluator workbook)]
      (->> sheet
           (get-rows)
           (rest)
           (map (t/fn [row :- Row] :- RowData
                  (with-timer-when profile?
                    (make-row-data header
                                   (map (t/fn [^Cell cell :- (t/U Cell nil)] :- CellValue
                                          (when cell (cell-value evaluator cell)))
                                        (get-cells row (count header)))))))
           (remove empty?)))
    (throw (ex-info "Sheet does not exists" {:sheet-name sheet-name
                                             :workbook   workbook}))))

(ann get-sheet-header [Workbook SheetName -> Header])
(defn get-sheet-header [^Workbook workbook sheet-name]
  (if-let [sheet (.getSheet workbook sheet-name)]
    (sheet-header sheet)
    (throw (ex-info "Sheet does not exists" {:sheet-name sheet-name
                                             :workbook   workbook}))))

(ann excel-file? [(Array Byte) String -> Boolean])
(defn excel-file? [^BufferedInputStream is type]
  (case type
    "xls" (POIFSFileSystem/hasPOIFSHeader is)
    "xlsx" (POIXMLDocument/hasOOXMLHeader is)
    :else false))

(ann get-sheets-names [Workbook -> SheetNames])
(defn get-sheets-names [^Workbook workbook]
  (->> (range (.getNumberOfSheets workbook))
       (map #(.getSheetAt workbook ^int %))
       (map #(.getSheetName ^Sheet %))))


(defn count-rows [sheet]
  (.getPhysicalNumberOfRows sheet))
;(.getLastRowNum sheet))

(defn get-cell [workbook sheet-name row-index column-index]
  (some-> workbook
          (.getSheet sheet-name)
          (.getRow row-index)
          (.getCell column-index)))

(defn get-cell-value [workbook sheet-name row-index column-index]
  (when-let [cell (get-cell workbook sheet-name row-index column-index)]
    (cell-value (make-evaluator workbook) cell)))

(defn create-row!
  ;; Safe return clean of existent row
  [sheet row-index]
  (if-let [row (.getRow sheet row-index)]
    row
    (.createRow sheet row-index)))

(defn create-cell! [sheet row-index column-index]
  (let [row (create-row! sheet row-index)]
    (when-let [cell (.getCell row column-index)]
      (.removeCell row cell))
    (.createCell row column-index)))

(defn set-cells!
  "set cell value by row and column indexes
  example of new-sheet-data:
  [[0 0 \"a\"][1 1 \"b\"] - Set A1 -> a B2 -> b "
  [^Workbook workbook sheet-name new-sheet-data]
  (doseq [cell-data new-sheet-data]
    (if-let [sheet (.getSheet workbook sheet-name)]
      (let [[row-id coll-id cell-value] cell-data
            cell (create-cell! sheet row-id coll-id)]
        (set-cell-value cell cell-value))
      (throw (ex-info "Sheet does not exists" {:sheet-name sheet-name
                                               :workbook   workbook})))))
(defn index-offset
  ([data] (range (count data)))
  ([data offset]
   (range offset (+ offset (count data)))))

(defn add-address
  ([data] (add-address data 0))
  ([data row-offset]
   (into [] (mapcat (fn [row row-index]
                     (map (fn [cell cell-index]
                             [row-index cell-index cell])
                           row (index-offset row)))
                   data
                   (index-offset data row-offset)))))

(defn append-rows!
  [^Workbook workbook sheet-name new-sheet-data]
  (let [sheet (.getSheet workbook sheet-name)
        row-offset (count-rows sheet)
        new-sheet-data-with-addresses (add-address new-sheet-data row-offset)]
    (set-cells! workbook sheet-name new-sheet-data-with-addresses)))

(defn create-rows!
  ([^Workbook workbook sheet-name]
   (create-rows! workbook sheet-name 0))
  ([^Workbook workbook sheet-name row-offset]
   (let [sheet (.getSheet workbook sheet-name)
         last-row (inc (count-rows sheet))]
     (doseq [row-number (range row-offset last-row)]
       ;; rows must be overwritten
       (.createRow sheet row-number)))))

(defn remove-rows!([^Workbook workbook sheet-name]
   (remove-rows! workbook sheet-name 0))
  ([^Workbook workbook sheet-name row-offset]
   (let [sheet (.getSheet workbook sheet-name)
         last-row (inc (count-rows sheet))]
     (doseq [row-number (range row-offset last-row)]
       (when-let [row (.getRow sheet row-number)]
         (.removeRow sheet row))))))
