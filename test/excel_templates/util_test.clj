(ns excel-templates.util-test
  (:require [clojure.test :refer :all]
            [excel-templates.util :refer :all]))

(deftest safe-sheet-name-test
  (testing "sanitizes the given sheet name with the default replacement (spaces)"
    (is (= "ABC " (safe-sheet-name "ABC*")))
    (is (= " ABC    " (safe-sheet-name "\\ABC/?:*"))))

  (testing "sanitizes the given sheet name with a given replacement"
    (is (= "ABCx" (safe-sheet-name "ABC*" \x)))
    (is (= "_ABC____" (safe-sheet-name "\\ABC/?:*" \_)))))

