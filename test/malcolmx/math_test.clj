(ns malcolmx.math-test
  (:require [clojure.test :refer :all]
            [malcolmx.math :as math]))

(defn roughly= [x y]
  (< (Math/abs ^double (- x y))
     0.0001))

(deftest cdf-test
  (testing "NORMALDISTS"
    (is (roughly= 0.36944134 (math/normal-distribution 10 20 30 true)))))

(deftest pdf-test
  (testing "NORMALDIST"
    (is (roughly= 0.012579441 (math/normal-distribution 10 20 30 false)))))

(deftest pascal-cdf-test
  (testing "Pascal CDF DIST"
    (is (roughly= 0.974383745 (math/pascal-distribution 10 20 0.8 true)))))

(deftest pascal-pdf-test
  (testing "Pascal PDF DIST"
    (is (roughly= 0.023647262 (math/pascal-distribution 10 20 0.8 false)))))

(deftest poisson-test
  (are [res x mean cummulative] (= res (math/poisson-distribution x mean cummulative))
    1                   1 0 true
    0.3678794411714609  0 1 true))

(deftest binom-inv-test
  (testing "BINOM.INV/ CRITBINOM"
    (is (= 75 (math/binom-inv 100 0.75 0.5)))
    (is (= 75 (math/binom-inv 100 0.75 0.5)))))


