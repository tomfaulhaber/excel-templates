(ns excel-templates.build
  (:import [java.io File FileOutputStream]
           [org.apache.poi.openxml4j.opc OPCPackage]
           [org.apache.poi.ss.usermodel Cell Row DateUtil WorkbookFactory]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFWorkbook])
  (:require [clojure.java.io :as io]
            [clojure.pprint :as pp]
            [clojure.set :as set]
            [clojure.java.shell :as sh]))

(defn indexed
  "For the collection coll with elements x0..xn, return a lazy sequence
   of pairs [0 x0]..[n xn]"
  [coll]
  (map (fn [i x] [i x]) (range) coll))

(defn cell-seq
  "Return a lazy seq of cells on the sheet in row major order (that is, across
   and then down)"
  [sheet]
  (apply
   concat
   (for [row-num (range (inc (.getLastRowNum sheet)))
         :let [row (.getRow sheet row-num)]
         :when row]
     (for [cell-num (range (inc (.getLastCellNum row)))
           :let [cell (.getCell row cell-num)]
           :when cell]
       cell))))

(defn formula?
  "Return true if src-cell has a formula"
  [cell]
  (= (.getCellType cell) Cell/CELL_TYPE_FORMULA))

(defn has-formula?
  "returns true if *any* of the cells on the sheet are calculated"
  [sheet]
  (some formula? (cell-seq sheet)))

(defn get-val
  "Get the value from a cell depending on the type"
  [cell]
  ;; I don't know why case doesn't work here, but it wasn't matching
  (let [cell-type (.getCellType cell)]
   (cond
     (= cell-type Cell/CELL_TYPE_STRING)
     (-> cell .getRichStringCellValue .getString)

     (= cell-type Cell/CELL_TYPE_NUMERIC)
     (if (DateUtil/isCellDateFormatted cell)
       (.getDateCellValue cell)
       (.getNumericCellValue cell))

     (= cell-type Cell/CELL_TYPE_BOOLEAN)
     (.getBooleanCellValue cell)

     (= cell-type Cell/CELL_TYPE_FORMULA)
     (.getCellFormula cell)

     (= cell-type Cell/CELL_TYPE_BLANK)
     nil

     :else (do (println (str "returning nil because type is " (.getCellType cell))) nil))))

(defn set-val
  "Set the value in the given cell, depending on the type"
  [wb cell val]
  (try
    (.setCellValue
     cell
     (cond
      (string? val) (.createRichTextString (.getCreationHelper wb) val)
      (number? val) (double val)
      :else val))
    (catch Exception e
      (throw (Exception. (pp/cl-format nil
                                       "Unable to assign value of type ~s to cell at (~d,~d)"
                                       (class val) (.getRowIndex cell) (.getColumnIndex cell)))))))

(defn set-formula
  "Set the formula for the given cell"
  [wb cell formula]
  (.setCellFormula cell formula)
  (-> wb .getCreationHelper .createFormulaEvaluator (.evaluateFormulaCell cell)))

(defn copy-row
  "Copy a single row of data from the template to the output, and the styles with them"
  [wb src-row dst-row]
  (when src-row
    (let [ncols (inc (.getLastCellNum src-row))]
     (doseq [cell-num (range ncols)]
       (when-let [src-cell (.getCell src-row cell-num Row/RETURN_BLANK_AS_NULL)]
         (let [dst-cell (.createCell dst-row cell-num)]
           (if (formula? src-cell)
             (set-formula wb dst-cell (get-val src-cell))
             (set-val wb dst-cell (get-val src-cell)))))))))

(defn inject-data-row
  "Take the data from the collection data-row at set the cell values in the target row accordingly.
If there are any nil values in the source collection, the corresponding cells are not modified."
  [data-row wb src-row dst-row]
  (let [src-cols (inc (.getLastCellNum src-row))
        data-cols (count data-row)
        ncols (max src-cols data-cols)]
    (doseq [cell-num (range ncols)]
      (let [data-val (nth data-row cell-num nil)
            src-cell (some-> src-row (.getCell cell-num))
            src-val (some-> src-cell get-val)]
        (when-let [val (or data-val src-val)]
          (let [dst-cell (or (.getCell dst-row cell-num)
                             (.createCell dst-row cell-num))]
            (if (and (not data-val) (formula? src-cell))
              (set-formula wb dst-cell val)
              (set-val wb dst-cell val))))))))

(defn copy-styles
  "Copy the styles from one row to another. We don't really copy, but rather
   assume that the styles with the same index in the source and destination
   are the same. This works since the destination is a copy of the source."
  [wb src-row dst-row]
  (let [ncols (inc (.getLastCellNum src-row))]
    (doseq [cell-num (range ncols)]
      (when-let [src-cell (.getCell src-row cell-num)]
        (let [dst-cell (or (.getCell dst-row cell-num)
                           (.createCell dst-row cell-num))]
          (.setCellStyle dst-cell
                         (->> src-cell .getCellStyle .getIndex (.getCellStyleAt wb)))))))
  (.setHeight dst-row (.getHeight src-row)))

(defn get-template
  "Get a file by its pathname or, if that's not found, use the resources"
  [template-file]
  (let [f (io/file template-file)]
    (if (.exists f)
      f
      (-> template-file io/resource io/file))))


(defn build-base-output
  "Build an output file with all the rows stripped out.
   This keeps all the styles and annotations while letting us write the
   spreadsheet using streaming so it can be arbitrarily large"
  [template-file output-file]
  (let [tmpfile (File/createTempFile "excel-template" ".xlsx")]
    (try
      (io/copy template-file tmpfile)
      (let [wb (XSSFWorkbook. (.getPath tmpfile))]
        (doseq [sheet-num (range (.getNumberOfSheets wb))]
          (let [sheet (.getSheetAt wb sheet-num)
                nrows (inc (.getLastRowNum sheet))]
            (doseq [row-num (reverse (range nrows))]
              (when-let [row (.getRow sheet row-num)]
                (.removeRow sheet row)))))
        ;; Write the resulting output Workbook
        (with-open [fos (FileOutputStream. output-file)]
          (.write wb fos)))
      (finally
        (io/delete-file tmpfile)))))

(defn get-all-sheet-names
  [wb]
  (map #(.getSheetName wb %) (range (.getNumberOfSheets wb))))

(defn save-workbook!
  [workbook file]
  (with-open [fos (FileOutputStream. file)]
    (.write workbook fos)))

(defn pad-sheet-rows!
  "Make sure that the sheet has at least the rows it needs to handle the
  incoming replacements."
  [sheet replacements]
  (let [row-nums (keys (replacements (.getSheetName sheet)))]
    (doseq [row-num (range (inc (apply max row-nums)))]
      (or (.getRow sheet row-num)
          (.createRow sheet row-num)))))

(defn create-missing-sheets!
  "Add new sheets to the workbook, saves the file."
  [excel-file replacements]
  (let [temp-file (File/createTempFile "add-sheets" ".xlsx")
        sheet-names (keys replacements)]
    (try
      (with-open [package (OPCPackage/open excel-file)]
        (let [workbook (XSSFWorkbook. package)
              all-sheet-names (get-all-sheet-names workbook)
              sheet-name-strings (filter string? sheet-names)
              sheet-name-numbers (filter number? sheet-names)]

          ;; For sheets specified by name that aren't in the workbook, create
          ;; sheets of the appropriate name.
          (doseq [sheet (set/difference (set sheet-name-strings)
                                        (set all-sheet-names))]
            (pad-sheet-rows! (.createSheet workbook sheet) replacements))

          ;; TODO If a sheet specified by number was greater than the number of
          ;; sheets in the workbook, add sheets to fill that difference.
          ;; TODO use when-let?
          #_(when (seq (sheet-name-numbers))
              (let [max-sheet-num (apply max sheet-name-numbers)
                    sheets-to-add (- max-sheet-num (dec (.getNumberOfSheets workbook)))]
                (doseq [sheet-index (range sheets-to-add)]
                  (.createSheet workbook ("Sheet" sheet-index)))))

          (save-workbook! workbook temp-file)))
      (io/copy temp-file excel-file)
      (finally
        (io/delete-file temp-file)))))

(defn render-to-file
  "Build a report based on a spreadsheet template"
  [template-file output-file replacements]
  (let [tmpfile (File/createTempFile "excel-output" ".xlsx")
        tmpcopy (File/createTempFile "excel-template-copy" ".xlsx")]
    (try
      ;; We copy the template file because otherwise POI will modify it even
      ;; though it's our input file. That's annoying from a source code
      ;; control perspective.
      (io/copy (get-template template-file) tmpcopy)
      (create-missing-sheets! tmpcopy replacements)
      (build-base-output tmpcopy tmpfile)
      (let [replacements (if (-> replacements first val map?)
                           replacements
                           {0 replacements})]
        (with-open [pkg (OPCPackage/open tmpcopy)]
          (let [template (XSSFWorkbook. pkg)
                intermediate-files (for [index (range (dec (.getNumberOfSheets template)))]
                                     (File/createTempFile (str "excel-intermediate-" index) ".xlsx"))
                inputs  (vec (concat [tmpfile]          intermediate-files))
                outputs (vec (concat intermediate-files [output-file]     ))]
            (try
              (doseq [sheet-num (range (.getNumberOfSheets template))]
                (let [src-sheet (.getSheetAt template sheet-num)
                      sheet-name (.getSheetName src-sheet)
                      sheet-data (or (get replacements sheet-name) (get replacements sheet-num) {})
                      nrows (inc (.getLastRowNum src-sheet))
                      src-has-formula? (has-formula? src-sheet)
                      wb (XSSFWorkbook. (.getPath (nth inputs sheet-num)))
                      wb (if src-has-formula? wb (SXSSFWorkbook. wb))]
                  (try
                    (let [sheet (.getSheetAt wb sheet-num)]
                      ;; loop through the rows of the template, copying
                      ;; from the template or injecting data rows as
                      ;; appropriate
                      (loop [src-row-num 0
                             dst-row-num 0]
                        (when (< src-row-num nrows)
                          (let [src-row (.getRow src-sheet src-row-num)]
                            (if-let [data-rows (get sheet-data src-row-num)]
                              (do
                                (doseq [[index data-row] (indexed data-rows)]
                                  (let [new-row (.createRow sheet (+ dst-row-num index))]
                                    (inject-data-row data-row wb src-row new-row)
                                    (copy-styles wb src-row new-row)))
                                (recur (inc src-row-num) (+ dst-row-num (count data-rows))))
                              (do
                                (when src-row
                                  (let [new-row (.createRow sheet dst-row-num)]
                                    (copy-row wb src-row new-row)
                                    (copy-styles wb src-row new-row)))
                                (recur (inc src-row-num) (inc dst-row-num))))))))
                    ;; Write the resulting output Workbook
                    (with-open [fos (FileOutputStream. (nth outputs sheet-num))]
                      (.write wb fos))
                    (catch Exception e (.printStackTrace e))
                    (finally
                      (when-not has-formula?
                        (.dispose wb))))))

              (catch Exception e (.printStackTrace e))
              (finally
                (doseq [f intermediate-files] (io/delete-file f)))))))
      (catch Exception e (.printStackTrace e))
      (finally
        (io/delete-file tmpfile)
        (io/delete-file tmpcopy)))))

(defn render-to-stream
  "Build a report based on a spreadsheet template, write it to the output
  stream."
  [template-file output-stream replacements]
  (let [temp-output-file (File/createTempFile "for-stream-output" ".xlsx")]
    (try
      (render-to-file template-file temp-output-file replacements)
      (io/copy (io/input-stream temp-output-file) output-stream)
      (finally (io/delete-file temp-output-file)))))

;; This works, even though foo.xlsx only contains Sheet1, Sheet2.
(comment (let [template-file "foo.xlsx"
               output-file "/tmp/bar.xlsx"
               data {"Sheet1" {2 [[nil "foo"]]}
                     "Sheet2" {2 [[nil "bar"]]}
                     "Sheet3" {2 [[nil "baz"]]}
                     "Another Sheet" {5 [[nil "foozle doozle"]]}}]
           (render-to-file template-file output-file data)
           (sh "libreoffice" "--calc" output-file)))
