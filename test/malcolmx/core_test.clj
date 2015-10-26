(ns malcolmx.core-test
  (:require [clojure.test :refer :all]
            [malcolmx.core :refer :all])
  (:import (java.io File BufferedInputStream FileInputStream)))

 (deftest excel-file-test
   (is (= true (excel-file? (BufferedInputStream. (FileInputStream. (File. "test/malcolmx/Bolvanka.xlsx"))) "xlsx"))
       "Should determine file type")
   (is (= true (excel-file? (BufferedInputStream. (FileInputStream. (File. "test/malcolmx/Bolvanka.xls"))) "xls"))
       "Should determine file type")
   (is (= false (excel-file? (BufferedInputStream. (FileInputStream. (File. *file*))) "xls"))
       "Should not determine file type"))

(deftest get-sheets-names-test
  (is (= ["IN" "OUT" "Markets" "Dist_300" "Dist_80" "Total 5Set" "+&СУММ" "Лист1" "ВПР" "Prob"]
         (get-sheets-names (parse "test/malcolmx/Bolvanka.xls")))))

(deftest empty-cells-test
  (let [wb (parse "test/malcolmx/empty-cells.xlsx")
        out-result [{"id" 1.0, "outcome" 10.0, "param" 111.0}
                    {"id" 2.0, "outcome" 9.0, "param" "    "}
                    {"id" 3.0, "outcome" 8.0, "param" 222.0}
                    {"id" 4.0, "outcome" 7.0, "param" nil}
                    {"id" 5.0, "outcome" 6.0, "param" 333.0}
                    {"id" 6.0, "outcome" 5.0, "param" nil}
                    {"id" 7.0, "outcome" 4.0, "param" 444.0}
                    {"id" 8.0, "outcome" 3.0, "param" nil}
                    {"id" 9.0, "outcome" 2.0, "param" 555.0}
                    {"id" 10.0, "outcome" 1.0, "param" nil}]]
    (is (= out-result (get-sheet wb "OUT")))))

(deftest header-extraction-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        header ["id"
                "market"
                "outcome"
                "coef"
                "param"
                "m_code"
                "o_code"
                "mgp_code"
                "mn_code"
                "mgp_weight"
                "mn_weight"]]
    (is (= header (get-sheet-header wb "OUT")))))

(deftest super-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        in-result [{"code" "P_SCORE_A_PART_1",
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
        out-result [{"coef"    2.5302341930958367, "id" 1.0, "m_code" "MATCH_NORM.DIST_300",
                     "market"  "Norm.Dist_300", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "NORM_300", "mn_weight" 11.0, "o_code" "A", "outcome" "A",
                     "param"   999999.0}
                    {"coef"    1.460323733283845, "id" 2.0, "m_code" "MATCH_NORM.DIST_300",
                     "market"  "Norm.Dist_300", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "NORM_300", "mn_weight" 11.0, "o_code" "B", "outcome" "B",
                     "param"   999999.0}
                    {"coef"    3.3852250801427837, "id" 3.0, "m_code" "PART1_NORM.DIST_80",
                     "market"  "Norm.Dist_80", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "NORM_80", "mn_weight" 21.0, "o_code" "A", "outcome" "A",
                     "param"   999999.0}
                    {"coef"    1.2745369596148286, "id" 4.0, "m_code" "PART1_NORM.DIST_80",
                     "market"  "Norm.Dist_80", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "NORM_80", "mn_weight" 21.0, "o_code" "B", "outcome" "B",
                     "param"   999999.0}
                    {"coef"    0.9259259259259272, "id" 5.0, "m_code" "PART2_POISSON_80",
                     "market"  "Poisson_80", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "POISSON_80", "mn_weight" 31.0, "o_code" "A", "outcome" "A",
                     "param"   999999.0}
                    {"coef"      150.0, "id" 6.0, "m_code" "PART2_POISSON_80", "market" "Poisson_80",
                     "mgp_code"  "DISTRIBUTION", "mgp_weight" 1.0, "mn_code" "POISSON_80",
                     "mn_weight" 31.0, "o_code" "B", "outcome" "B",
                     "param"     999999.0}
                    {"coef"    0.9259259259259263, "id" 7.0, "m_code" "PART3_POISSON_300",
                     "market"  "Poisson_300", "mgp_code" "DISTRIBUTION", "mgp_weight" 1.0,
                     "mn_code" "POISSON_300", "mn_weight" 41.0, "o_code" "A", "outcome" "A",
                     "param"   999999.0}
                    {"coef"      150.0, "id" 8.0, "m_code" "PART3_POISSON_300", "market" "Poisson_300",
                     "mgp_code"  "DISTRIBUTION", "mgp_weight" 1.0, "mn_code" "POISSON_300",
                     "mn_weight" 41.0, "o_code" "B", "outcome" "B",
                     "param"     999999.0}
                    {"coef"    1.3419216317767044, "id" 9.0, "m_code" "MATCH_VPR_NORM.DIST_300",
                     "market"  "ВПР Norm.Dist_81", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "VPR_NORM_150", "mn_weight" 12.0, "o_code" "A", "outcome" "A",
                     "param"   190.0}
                    {"coef"    2.986857825567502, "id" 10.0, "m_code" "MATCH_VPR_NORM.DIST_300",
                     "market"  "ВПР Norm.Dist_81", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "VPR_NORM_150", "mn_weight" 12.0, "o_code" "B", "outcome" "B",
                     "param"   190.0}
                    {"coef"    3.561253561253567, "id" 11.0, "m_code" "MATCH_VPR_POISSON_81",
                     "market"  "ВПР Poisson_81", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "VPR_POISSON_81", "mn_weight" 13.0, "o_code" "A", "outcome" "A",
                     "param"   500.0}
                    {"coef"    1.2512512512512506, "id" 12.0, "m_code" "MATCH_VPR_POISSON_81",
                     "market"  "ВПР Poisson_81", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "VPR_POISSON_81", "mn_weight" 13.0, "o_code" "B", "outcome" "B",
                     "param"   500.0}
                    {"coef"    "#NUM!", "id" 13.0, "m_code" "MATCH_PROB",
                     "market"  "Match Probability", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "PROB", "mn_weight" 14.0, "o_code" "A", "outcome" "A",
                     "param"   999999.0}
                    {"coef"    "#NUM!", "id" 14.0, "m_code" "MATCH_PROB",
                     "market"  "Match Probability", "mgp_code" "OTHER", "mgp_weight" 100.0,
                     "mn_code" "PROB", "mn_weight" 14.0, "o_code" "B", "outcome" "B",
                     "param"   999999.0}
                    {"coef"    -1.187367238518843E-33, "id" 15.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "OVER", "outcome" "Over",
                     "param"   208.5}
                    {"coef"    1.187367238518843E-33, "id" 16.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "UNDER", "outcome" "Under",
                     "param"   208.5}
                    {"coef"    -3.2684236470440286E-33, "id" 17.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "OVER", "outcome" "Over",
                     "param"   88.5}
                    {"coef"    3.2684236470440286E-33, "id" 18.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "UNDER", "outcome" "Under",
                     "param"   88.5}
                    {"coef"    -3.0333987132024294E-33, "id" 19.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "OVER", "outcome" "Over",
                     "param"   89.5}
                    {"coef"    3.0333987132024294E-33, "id" 20.0, "m_code" "MATCH_TOTAL",
                     "market"  "Match Total", "mgp_code" "SIMPLE OPERATION", "mgp_weight" 2.0,
                     "mn_code" "TOTAL", "mn_weight" 15.0, "o_code" "UNDER", "outcome" "Under",
                     "param"   89.5}]
        data [{"value" 1.0,
                 "id" 1.0}
              {"value" 4.0,
               "id" 4.0}
              {"value" 2.0,
               "id" 2.0}
              {"value" 3.0,
               "id" 3.0}
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
    (is (= in-result (get-sheet wb "IN")))
    (is (= out-result (get-sheet wb "OUT")))))

;;;;;;;

(deftest create-cell!-test
  (let  [wb (parse "test/malcolmx/Bolvanka.xlsx")
         sheet-name "EventLog"
         sheet (.getSheet wb sheet-name)]
    (is (= "id" (get-cell-value wb sheet-name 0 0)))
    (create-cell! sheet 0 0)
    (is (= (nil? (get-cell-value wb sheet-name 0 0))))))

(deftest clean-all-rows!-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        sheet-name "EventLog"]
    (is (= 2 (get-last-row-index (.getSheet wb sheet-name))))
    (create-rows! wb sheet-name)
    (is (nil? (get-cell-value wb sheet-name 0 0)))
    (is (nil? (get-cell-value wb sheet-name 0 1)))
    (is (nil? (get-cell-value wb sheet-name 1 0)))
    (is (nil? (get-cell-value wb sheet-name 1 1)))
    (is (nil? (get-cell-value wb sheet-name 2 0)))
    (is (nil? (get-cell-value wb sheet-name 2 1)))
    (is (= 2 (get-last-row-index (.getSheet wb sheet-name))))))

(deftest clean-offset-rows!-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        sheet-name "EventLog"]
    (is (= 2 (get-last-row-index (.getSheet wb sheet-name))))
    (remove-rows! wb sheet-name 1)
    (is (= 0 (get-last-row-index (.getSheet wb sheet-name))))
    (is (= "id" (get-cell-value wb sheet-name 0 0)))
    (is (= "value" (get-cell-value wb sheet-name 0 1)))
    (is (nil? (get-cell-value wb sheet-name 1 0)))
    (is (nil? (get-cell-value wb sheet-name 1 1)))
    (is (nil? (get-cell-value wb sheet-name 2 0)))
    (is (nil? (get-cell-value wb sheet-name 2 1)))
    ))

(deftest set-cells-by-address!-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        sheet-name "EventLog"
        data (add-address [["a1" "b1" "c1"]
                           ["a2" "b2" "c2"]
                           ["a3" "b3" "c3"]
                           ["a4" "b4" "c4"]])]
    (is (= "id" (get-cell-value wb sheet-name 0 0)))
    (is (= "value" (get-cell-value wb sheet-name 0 1)))
    (is (= "a2" (get-cell-value wb sheet-name 1 0)))
    (is (= "b2" (get-cell-value wb sheet-name 1 1)))
    (is (= "a3" (get-cell-value wb sheet-name 2 0)))
    (is (= "b3" (get-cell-value wb sheet-name 2 1)))
    (is (nil?   (get-cell-value wb sheet-name 3 1)))
    (set-cells! wb sheet-name data)
    (is (= "a1" (get-cell-value wb sheet-name 0 0)))
    (is (= "b1" (get-cell-value wb sheet-name 0 1)))
    (is (= "c1" (get-cell-value wb sheet-name 0 2)))
    (is (= "a2" (get-cell-value wb sheet-name 1 0)))
    (is (= "b2" (get-cell-value wb sheet-name 1 1)))
    (is (= "c2" (get-cell-value wb sheet-name 1 2)))
    (is (= "a3" (get-cell-value wb sheet-name 2 0)))
    (is (= "b3" (get-cell-value wb sheet-name 2 1)))
    (is (= "c3" (get-cell-value wb sheet-name 2 2)))
    (is (= "a4" (get-cell-value wb sheet-name 3 0)))
    (is (= "b4" (get-cell-value wb sheet-name 3 1)))
    (is (= "c4" (get-cell-value wb sheet-name 3 2)))
    ))

(deftest append-rows!-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        sheet-name "EventLog"
        data [["a1" "b1" 1.0]
              ["a2" "b2" nil]
              ["a3" "b3" 3]
              ["a4" "b4" ""]]        ]
    (append-rows! wb sheet-name data)
    ;; head
    (is (= "id" (get-cell-value wb sheet-name 0 0)))
    (is (= "value" (get-cell-value wb sheet-name 0 1)))

    (is (= "a1" (get-cell-value wb sheet-name 1 0)))
    (is (= "b1" (get-cell-value wb sheet-name 1 1)))
    (is (= 1.0 (get-cell-value wb sheet-name 1 2)))
    (is (= "a2" (get-cell-value wb sheet-name 2 0)))
    (is (= "b2" (get-cell-value wb sheet-name 2 1)))
    (is (= "" (get-cell-value wb sheet-name 2 2)))
    (is (= "a3" (get-cell-value wb sheet-name 3 0)))
    (is (= "b3" (get-cell-value wb sheet-name 3 1)))
    (is (= 3.0 (get-cell-value wb sheet-name 3 2)))
    (is (= "a4" (get-cell-value wb sheet-name 4 0)))
    (is (= "b4" (get-cell-value wb sheet-name 4 1)))
    (is (= "" (get-cell-value wb sheet-name 4 2)))
    ))

(deftest set-rows!-test
  (let [wb (parse "test/malcolmx/Bolvanka.xlsx")
        sheet-name "EventLog"
        data [["a1" "b1" 1]
              ["a2" "b2" nil]
              ["a3" "b3" 1.0]
              ["a4" "b4" ""]]]
    (set-rows! wb sheet-name data)
    (is (= "a1" (get-cell-value wb sheet-name 0 0)))
    (is (= "b1" (get-cell-value wb sheet-name 0 1)))
    (is (= 1.0 (get-cell-value wb sheet-name 0 2)))
    (is (= "a2" (get-cell-value wb sheet-name 1 0)))
    (is (= "b2" (get-cell-value wb sheet-name 1 1)))
    (is (= "" (get-cell-value wb sheet-name 1 2)))
    (is (= "a3" (get-cell-value wb sheet-name 2 0)))
    (is (= "b3" (get-cell-value wb sheet-name 2 1)))
    (is (= 1.0 (get-cell-value wb sheet-name 2 2)))
    (is (= "a4" (get-cell-value wb sheet-name 3 0)))
    (is (= "b4" (get-cell-value wb sheet-name 3 1)))
    (is (= "" (get-cell-value wb sheet-name 3 2)))
    ))