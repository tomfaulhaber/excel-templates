(ns excel-templates.build-test
  (:import [java.io File]
           [org.apache.poi.openxml4j.exceptions InvalidFormatException])
  (:require [excel-templates.build :refer :all]
            [clojure.test :refer :all]))

(deftest rendering-to-file
  (testing "throws exceptions when asked"
    (let [temp-file (File/createTempFile "tmp" ".xlsx")]
      (is (thrown? InvalidFormatException
                   (render-to-file "test/no-content-type.xlsx" temp-file
                                   {:sheet-name "Some Sheet"
                                    0 [["Some Row"]]}
                                   {:throw-exceptions true})))))

  (testing "swallows exceptions by default"
    (let [temp-file (File/createTempFile "tmp" ".xlsx")]
      (is (nil? (render-to-file "test/no-content-type.xlsx" temp-file
                                {:sheet-name "Some Sheet"
                                 0 [["Some Row"]]}))))))
