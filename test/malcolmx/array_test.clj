(ns malcolmx.array-test
  (:use malcolmx.core
        clojure.pprint
        clojure.test))

(def w (parse "/home/serzh/AutoCalc_Soccer_EventLog.xlsx"))

(deftest array-test
  (is (= [{"Value" 4.0}
          {"Value" 2.0}
          {"Value" 2.0}
          {"Value" 0.0}
          {"Value" 1.0}
          {"Value" 2.0}
          {"Value" 0.0}
          {"Value" 1.0}
          {"Value" 1.0}
          {"Value" 0.0}
          {"Value" 1.0}
          {"Value" 1.0}
          {"Value" 0.0}
          {"Value" 0.0}
          {"Value" 0.0}
          {"Value" 0.0}
          {"Value" 0.0}
          {"Value" false}
          {"Value" false}
          {"Value" 0.0}
          {"Value" true}
          {"Value" true}
          {"Value" 0.0}
          {"Value" false}
          {"Value" true}
          {"Value" 0.0}
          {"Value" "Team1"}
          {"Value" 52.0}]
         (get-sheet w "TEST"))))