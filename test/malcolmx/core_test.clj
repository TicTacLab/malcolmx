 (ns malcolmx.core-test
  (:require [clojure.test :refer :all]
            [malcolmx.core :refer :all]))

(deftest super-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        result [{"code" "P_SCORE_A_PART_1",
                       "value" 1.0,
                       "name" "Score A Part 1",
                       "type" "Parameter",
                       "id" 1.0}
                      {"code" "P_SCORE_B_PART_1",
                       "value" 2.0,
                       "name" "Score B Part 1",
                       "type" "Parameter",
                       "id" 2.0}
                      {"code" "P_SCORE_A_PART_2",
                       "value" 3.0,
                       "name" "Score A Part 2",
                       "type" "Parameter",
                       "id" 3.0}
                      {"code" "P_SCORE_B_PART_2",
                       "value" 4.0,
                       "name" "Score B Part 2",
                       "type" "Parameter",
                       "id" 4.0}
                      {"code" "P_SCORE_A_PART_3",
                       "value" 5.0,
                       "name" "Score A Part 3",
                       "type" "Parameter",
                       "id" 5.0}
                      {"code" "P_SCORE_B_PART_3",
                       "value" 6.0,
                       "name" "Score B Part 3",
                       "type" "Parameter",
                       "id" 6.0}
                      {"code" "P_TIME",
                       "value" 7.0,
                       "name" "Time",
                       "type" "Parameter",
                       "id" 7.0}
                      {"code" "P_HANDICAP",
                       "value" 8.0,
                       "name" "Handicap",
                       "type" "Parameter",
                       "id" 8.0}
                      {"code" "P_TOTAL",
                       "value" 9.0,
                       "name" "Total",
                       "type" "Parameter",
                       "id" 9.0}
                      {"code" "P_QUIRRELL_COEF",
                       "value" 10.0,
                       "name" "Quirrell coef",
                       "type" "Parameter",
                       "id" 10.0}
                      {"code" "P_PROB_A",
                       "value" 11.0,
                       "name" "Prob A",
                       "type" "Parameter",
                       "id" 11.0}
                      {"code" "P_PROB_B",
                       "value" 12.0,
                       "name" "Prob B",
                       "type" "Parameter",
                       "id" 12.0}
                      {"code" "P_CHECK_BOX_099",
                       "value" 13.0,
                       "name" "Check box 0,99",
                       "type" "Parameter",
                       "id" 13.0}]
        data [{"value" 1.0,
                 "id" 1.0}
                {"value" 2.0,
                 "id" 2.0}
                {"value" 3.0,
                 "id" 3.0}
                {"value" 4.0,
                 "id" 4.0}
                {"value" 5.0,
                 "id" 5.0}
                {"value" 6.0,
                 "id" 6.0}
                {"value" 7.0,
                 "id" 7.0}
                {"value" 8.0,
                 "id" 8.0}
                {"value" 9.0,
                 "id" 9.0}
                {"value" 10.0,
                 "id" 10.0}
                {"value" 11.0,
                 "id" 11.0}
                {"value" 12.0,
                 "id" 12.0}
                {"value" 13.0,
                 "id" 13.0}]]
    (update-sheet! wb "IN" data :by "id")
    (= result (get-sheet wb "IN"))))


