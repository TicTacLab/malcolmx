(ns malcolmx.math
  (:use [clojure.tools.trace])
  (:import [org.apache.commons.math3.distribution NormalDistribution
                                                  PascalDistribution
                                                  PoissonDistribution
                                                  BinomialDistribution]
           [org.apache.poi.ss.formula.functions FreeRefFunction Fixed4ArgFunction Function Fixed3ArgFunction AggregateFunction]
           [org.apache.poi.ss.formula.eval NumberEval StringEval OperandResolver ErrorEval RefEvalBase BoolEval ValueEval]
           [org.apache.poi.ss.formula.udf AggregatingUDFFinder DefaultUDFFinder]
           [org.apache.poi.ss.formula WorkbookEvaluator]
           [org.apache.poi.ss.usermodel Workbook]))

(defn probability? [^double probability_s]
  (< 0.0 probability_s 1.0))

(defn get-eval-value [value-eval]
  (condp instance? value-eval
    ErrorEval {:error (.getErrorCode ^ErrorEval value-eval)}
    NumberEval (.getNumberValue ^NumberEval value-eval)
    BoolEval (.getBooleanValue ^BoolEval value-eval)
    RefEvalBase (get-eval-value (.getInnerValueEval ^RefEvalBase value-eval
                                                    (.getFirstSheetIndex ^RefEvalBase value-eval)))))

(defn extract-operand [args n row col]
  (-> (nth args n)
      (OperandResolver/getSingleValue  row col)
      (get-eval-value)))

(defn register-fun! [^String nm ^Function fun]
  (WorkbookEvaluator/registerFunction nm fun))

(defn udf-fun [name fun]
  (AggregatingUDFFinder.
    (into-array [(DefaultUDFFinder.
                   (into-array [name])
                   (into-array [fun]))])))


;;; Distributions

;; Normal Distribution
(defn normal-distribution [x mean standard-dev cumulative]
  (let [norm (NormalDistribution. mean standard-dev)]
    (if (true? cumulative)
      (.cumulativeProbability norm x)
      (.density norm x))))

;;; NORMDIST
(def normal-distribution-fun
  (proxy [Fixed4ArgFunction] []
    (evaluate [col-index  row-index x mean standard-dev cumulative]
      (let [cumulative-value (get-eval-value cumulative)
            standard-dev-value (get-eval-value standard-dev)
            x-value (get-eval-value x)
            mean-value (get-eval-value mean)]
        (if (every? #(or (integer? %) (float? %)) [standard-dev-value x-value mean-value])
          (NumberEval. ^double (normal-distribution x-value mean-value standard-dev-value cumulative-value))
          (StringEval. (str [x-value mean-value standard-dev-value])))))))

;;; NORM.DIST
(def normal-distribution-free-fun
  (reify FreeRefFunction
    (evaluate [this args ec]
      (let [row  (.getRowIndex ec)
            col  (.getColumnIndex ec)
            x (extract-operand args 0 row col)
            mean (extract-operand args 1 row col)
            standard-dev (extract-operand args 2 row col)
            cumulative (extract-operand args 3 row col)]
        (if (every? #(or (integer? %) (float? %)) [standard-dev x mean])
          (NumberEval. ^double (normal-distribution x mean standard-dev cumulative))
          (StringEval. "VALUE!"))))))

(def normdist-udf (udf-fun "_xlfn.NORM.DIST" normal-distribution-free-fun))


;;; Pascal Distribution

(defn pascal-distribution [number-f number-s probability-s cumulative]
  (let [pasc (PascalDistribution. number-s probability-s)]
    (if (true? cumulative)
      (.cumulativeProbability pasc number-f)
      (.probability pasc number-f))))

;;; NEGBINOM.DIST
(def pascal-distribution-free-fun
  (reify FreeRefFunction
    (evaluate [this args ec]
      (let [row  (.getRowIndex ec)
            col  (.getColumnIndex ec)
            number-f (extract-operand args 0 row col)
            number-s (extract-operand args 1 row col)
            probability-s (extract-operand args 2 row col)
            cumulative (extract-operand args 3 row col)]
        (if (every? #(or (integer? %) (float? %)) [number-f number-s probability-s])
          (NumberEval. ^double (pascal-distribution number-f number-s probability-s cumulative))
          (StringEval. "VALUE!"))))))

(def pascal-udf (udf-fun "_xlfn.NEGBINOM.DIST" pascal-distribution-free-fun))

;;; Binomial Distribution

(defn binomial-distribution [number-s trials probability-s cumulative]
  (let [bin (BinomialDistribution. trials probability-s)]
    (if (true? cumulative)
      (.cumulativeProbability bin number-s)
      (.probability bin number-s))))

;;; BINOMIAL
(def binomial-distribution-free-fun
  (reify FreeRefFunction
    (evaluate [this args ec]
      (let [row  (.getRowIndex ec)
            col  (.getColumnIndex ec)
            number-s (extract-operand args 0 row col)
            trials (extract-operand args 1 row col)
            probability-s (extract-operand args 2 row col)
            cumulative (extract-operand args 3 row col)]
        (if (every? #(or (integer? %) (float? %)) [number-s probability-s trials])
          (NumberEval. ^double (binomial-distribution number-s trials probability-s cumulative))
          (StringEval. "VALUE!"))))))

(def binomial-udf (udf-fun "_xlfn.BINOM.DIST" binomial-distribution-free-fun))


;;; Binomial Inverse Distribution

(defn binom-inv [^long trials ^double probability-s ^double alpha]
  (if (and (> trials 0)
           (probability? probability-s)
           (probability? alpha))
    (->> (for [x (range trials)
               :let [p (binomial-distribution x trials probability-s true)]
               :when (>= p alpha)]
           [p x])
         sort
         first
         second)
    'error))

;;; BINOM.INV
(def binom-inv-free
  (reify FreeRefFunction
    (evaluate [this args ec]
      (let [row  (.getRowIndex ec)
            col  (.getColumnIndex ec)
            trials (extract-operand args 0 row col)
            probability-s (extract-operand args 1 row col)
            alpha (extract-operand args 2 row col)]
        (if (every? #(or (integer? %) (float? %)) [trials probability-s alpha])
          (NumberEval. (double (binom-inv trials probability-s alpha)))  ;;; DOUBLE!!!!
          (StringEval. "VALUE!"))))))

(def binom-inv-udf (udf-fun "_xlfn.BINOM.INV" binom-inv-free))

;;; CRITBINOM
(def critbinom-fun
  (proxy [Fixed3ArgFunction] []
    (evaluate [col-index  row-index trials probability-s alpha]
      (let [trials-value (get-eval-value trials)
            probability-s-value (get-eval-value  probability-s)
            alpha-value (get-eval-value alpha)]
        (if (every? #(or (integer? %) (float? %)) [trials-value probability-s-value alpha-value])
          (NumberEval. (double (binom-inv trials-value probability-s-value alpha-value)))
          (StringEval. "VALUE!"))))))

;;; SUM
(def sum-free
  (proxy [AggregateFunction] []
    (evaluate [args]
      (areduce ^doubles args i acc 0.0
               (unchecked-add acc (aget ^doubles args i))))
    (getMaxNumOperands [] 255)))

(defn register-funs! []
  (register-fun! "NORMDIST" normal-distribution-fun)
  (register-fun! "CRITBINOM" critbinom-fun)
  (register-fun! "SUM" sum-free))

(register-funs!)

(defn register-udf-funs! [^Workbook workbook]
  (.addToolPack workbook normdist-udf)
  (.addToolPack workbook pascal-udf)
  (.addToolPack workbook binomial-udf)
  (.addToolPack workbook binom-inv-udf))

