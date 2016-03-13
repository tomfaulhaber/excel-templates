(ns excel-templates.build
  (:import [java.io File FileOutputStream FileInputStream]
           [java.util Calendar]
           [java.util.zip ZipEntry ZipFile ZipOutputStream]
           [org.apache.poi.openxml4j.opc OPCPackage]
           [org.apache.poi.ss.usermodel Cell Row DateUtil WorkbookFactory]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFWorkbook])
  (:require [clojure.java.io :as io]
            [clojure.java.shell :as sh]
            [clojure.pprint :as pp]
            [clojure.set :as set]
            [clojure.string :as str]
            [clojure.xml :as xml]
            [excel-templates.charts :as c]
            [excel-templates.formulas :as fo]))

(defn create-temp-xlsx-file
  "Create a temp file with correct headers to be opened by OPCPackage"
  [prefix]
  (let [tmpfile (File/createTempFile prefix ".xlsx")
        wb (XSSFWorkbook.)]
    (with-open [fos (FileOutputStream. (.getPath tmpfile))]
      (.write wb fos))
    tmpfile))

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


(defprotocol IExcelValue
  (excel-val [val wb]))

(extend-protocol IExcelValue
  Object
  (excel-val [val wb] val)

  String
  (excel-val [val wb]
    (.createRichTextString (.getCreationHelper wb) val))

  Number
  (excel-val [val wb]
    (double val))

  org.joda.time.DateTime
  (excel-val [val _]
    (let [[year mon day hour min sec]
          (-> val bean ((juxt :year :monthOfYear :dayOfMonth
                              :hourOfDay :minuteOfHour :secondOfMinute)))]
      (doto (Calendar/getInstance)
        (.set year (dec mon) day hour min sec)))))


(defn set-val
  "Set the value in the given cell, depending on the type"
  [wb cell val]
  (try
    (.setCellValue cell (excel-val val wb))
    (catch Exception e
      (throw (Exception. (pp/cl-format nil
                                       "Unable to assign value of type ~s to cell at (~d,~d)"
                                       (class val) (.getRowIndex cell) (.getColumnIndex cell)))))))

(defn set-formula
  "Set the formula for the given cell"
  [wb cell formula]
  (fo/debug-println
   (str "Set formula: (" (.getRowIndex cell) ", " (.getColumnIndex cell)
        "): " formula) )
  (.setCellFormula cell formula)
  (-> wb .getCreationHelper .createFormulaEvaluator (.evaluateFormulaCell cell)))

(defn copy-row
  "Copy a single row of data from the template to the output, and the styles with them"
  [translation-table wb sheet src-row dst-row]
  (when src-row
    (let [ncols (inc (.getLastCellNum src-row))]
     (doseq [cell-num (range ncols)]
       (when-let [src-cell (.getCell src-row cell-num Row/RETURN_BLANK_AS_NULL)]
         (let [dst-cell (.createCell dst-row cell-num)
               val (get-val src-cell)]
           (if (formula? src-cell)
             (let [target [(.getRowNum dst-row) cell-num]
                   formula (fo/translate-formula translation-table wb sheet target val)]
               (set-formula wb dst-cell formula))
             (set-val wb dst-cell val))))))))

(defn inject-data-row
  "Take the data from the collection data-row and set the cell values in the target row accordingly.
If there are any nil values in the source collection, the corresponding cells are not modified."
  [data-row translation-table wb sheet src-row dst-row]
  (let [src-cols (inc (if src-row (.getLastCellNum src-row) -1))
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
              (let [target [(.getRowNum dst-row) cell-num]
                    formula (fo/translate-formula translation-table wb sheet target val)]
                (set-formula wb dst-cell formula))
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
      (-> template-file io/resource io/input-stream))))


(defn build-base-output
  "Build an output file with all the rows stripped out.
   This keeps all the styles and annotations while letting us write the
   spreadsheet using streaming so it can be arbitrarily large"
  [template-file output-file]
  (let [tmpfile (create-temp-xlsx-file "excel-template")]
    (try
      (io/copy template-file tmpfile)
      (with-open [pkg (OPCPackage/open tmpfile)]
        (let [wb (XSSFWorkbook. pkg)]
          (doseq [sheet-num (range (.getNumberOfSheets wb))]
            (let [sheet (.getSheetAt wb sheet-num)
                  nrows (inc (.getLastRowNum sheet))]
              (doseq [row-num (reverse (range nrows))]
                (when-let [row (.getRow sheet row-num)]
                  (.removeRow sheet row)))))
          ;; Write the resulting output Workbook
          (with-open [fos (FileOutputStream. output-file)]
            (.write wb fos))))
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

;; TODO - this scenario can still be broken, for example if the first sheet
;; data doesn't have a sheet-name, but the second has the same sheet name
;; as the template. May be better modeled loop/recur, to maintain state.
(defn add-sheet-names
  "For replacements that don't have an explicit sheet name, add a unique one."
  [replacements]
  (into {} (for [[template-name sheet-data] replacements]
             [template-name
              (map-indexed
                (fn [i m] (if (:sheet-name m)
                            m
                            (assoc m :sheet-name (if (= i 0)
                                                   template-name
                                                   (str template-name "-" i)))))
                sheet-data)])))

(defn get-sheet-index
  [workbook sheet-name]
  (->> sheet-name
       (.getSheet workbook)
       (.getSheetIndex workbook)))

(defn clone-sheet!
  [workbook template-sheet-name dest-sheet-name target-loc]
  (->> template-sheet-name
       (get-sheet-index workbook)
       (.cloneSheet workbook))

  (->> (str template-sheet-name " (2)")
       (get-sheet-index workbook)
       (#(.setSheetName workbook % dest-sheet-name)))

  (.setSheetOrder workbook dest-sheet-name target-loc))

(defn remove-sheet!
  [workbook sheet-name]
  (.removeSheetAt workbook (get-sheet-index workbook sheet-name)))

(defn get-sheet-names
  "Get all the sheet names in a workbook in the order they appear"
  [wb]
  (for [index (range (.getNumberOfSheets wb))]
    (.getSheetName wb index)))

(defn get-sheets
  "Get all the sheet objects in a workbook in the order they appear"
  [wb]
  (for [index (range (.getNumberOfSheets wb))]
    (.getSheetAt wb index)))

(defn extract-charts
  "Get all the charts on the various sheets and return their XML as a dictionary
  keyed by sheet name. Then, delete the charts from the original workbook.

  This is necessary because POI has a bug cloning sheets with charts on them.
  See https://github.com/tomfaulhaber/excel-templates/issues/7 for info."
  [workbook]
  (->> workbook
       get-sheets
       (mapcat c/get-chart-data)))

(defn charts-from-file
  "Extract the chart info from a file"
  [excel-file]
  ;; Owing to the genius that is POI, looking at the file to get the
  ;; charts out can corrupt the file. Rather than fight this, we just
  ;; copy the file before manipulating it and then throw away the copy.
  (let [temp-file (File/createTempFile "excel-template-chart" ".xlsx")]
    (try
      (io/copy excel-file temp-file)
      (with-open [package (OPCPackage/open temp-file)]
       (let [workbook (XSSFWorkbook. package)]
         (extract-charts workbook)))
      (finally
        (io/delete-file temp-file)))))

(defn xml-to-str
  "Generate an XML string from parsed XML using clojure.xml"
  [xml-data]
  (-> (with-out-str (-> xml-data xml/emit))
      (str/replace #"^.*\n" "")
      (str/replace #"(\r?\n|\r)" "")))

(defn filter-zip-file
  "Filter zip file copies the zip file deleting and modifying the entries according to the rules
  provided"
  [input-file output-file rules]
  (let [buf-size 65536
        buf (byte-array buf-size)]
    (with-open [zip-file (ZipFile. input-file)
                output (ZipOutputStream. (io/output-stream output-file))]
      (doseq [entry (enumeration-seq (.entries zip-file))
              :let [entry-name (.getName entry)
                    rule (rules entry-name)
                    new-entry (ZipEntry. (or (:rename rule) entry-name))]]
        (when-not (:delete rule)
          (with-open [entry-data (.getInputStream zip-file entry)]
            (if-let [edit-fn (:edit rule)]
              (let [xml-data (xml/parse entry-data)
                    new-data (edit-fn xml-data)
                    data-str (xml-to-str new-data)
                    bytes (.getBytes data-str "UTF-8")
                    len (alength bytes)]
                (.putNextEntry output new-entry)
                (.write output bytes 0 len)
                (.closeEntry output))
              (do
                (.putNextEntry output new-entry)
                (loop []
                 (let [bytes-read (.read entry-data buf 0 buf-size)]
                   (when (pos? bytes-read)
                     (.write output buf 0 bytes-read)
                     (recur))))
                (.closeEntry output)))))))))


(defmacro memfnx
  "Like memfn, but gives n-args arguments so I don't have to name them."
  [method n-args]
  `(memfn ~method ~@(repeatedly n-args #(gensym "arg"))))

(defn insert-charts
  "Add the charts back into the sheets"
  [workbook chart-data src-sheet dst-sheet src-index target-loc]
  (when-let [charts (filter #(= src-sheet (:sheet %)) chart-data)]
    (doseq [{:keys [chart-location chart-xml chart-sheet?]} charts
            :when (not chart-sheet?)]
      (when-not chart-location
        (throw (RuntimeException. "The POI library only supports TwoCellAnchors. See https://github.com/tomfaulhaber/excel-templates/issues/29 for information and workaround.")))
      (let [target (.getSheet workbook dst-sheet)
            new-xml (c/chart-change-sheet target src-index target-loc chart-xml)
            drawing (.createDrawingPatriarch target)
            anchor-points (map #(get-in chart-location %)
                               [[:from :col-off] [:from :row-off]
                                [:to   :col-off] [:to   :row-off]
                                [:from :col    ] [:from :row    ]
                                [:to   :col    ] [:to   :row    ]])
            anchor (apply (memfnx createAnchor 8) drawing anchor-points)
            new-chart (.createChart drawing anchor)]
        (c/set-xml new-chart new-xml)))))

;; TODO validate that only valid templates are named in replacements.
;; use all-sheet-names and (keys replacements)

(defn create-missing-sheets!
  "Updates the excel file with any missing sheets, referred to by :sheet-name
  in the replacements."
  [excel-file replacements chart-data]
  (let [temp-file (File/createTempFile "excel-add-sheets" ".xlsx")]
    (try
      (with-open [package (OPCPackage/open excel-file)]
        (let [workbook (XSSFWorkbook. package)]
          (loop [src-index 0
                 src-sheets (get-sheet-names workbook)
                 dst-index 0
                 chart-data chart-data
                 seen-sheets #{}]
            (when src-sheets
              (let [[src-sheet & src-sheets] src-sheets
                    seen-sheets (conj seen-sheets src-sheet)]
                (if (contains? replacements src-sheet)
                  (let [dst-sheets (map :sheet-name (replacements src-sheet))
                        self-offset (.indexOf dst-sheets src-sheet)
                        replace-self? (>= self-offset 0)]
                    (doseq [[dst-offset dst-sheet] (map-indexed vector dst-sheets)
                            :let [target-loc (+ dst-index dst-offset
                                                (if (and replace-self? (<= self-offset dst-offset))
                                                  0 1))]]
                      (when (not= src-sheet dst-sheet)
                        (clone-sheet! workbook src-sheet dst-sheet target-loc))
                      (insert-charts workbook chart-data src-sheet dst-sheet src-index target-loc))

                    ;; Update any worksheets that point at this one
                    ;; Currently this does charts only
                    (when (and (seq dst-sheets)
                               (not (and replace-self?
                                         (= 1 (count dst-sheets)))))
                      (let [these (set (conj dst-sheets src-sheet))]
                        (doseq [sheet (get-sheets workbook)
                                :when (not (these (.getSheetName sheet)))]
                          (c/expand-charts sheet src-sheet dst-sheets))))

                    (let [new-chart-data (c/expand-chart-data
                                          workbook src-sheet dst-sheets
                                          (filter #(not (seen-sheets (:sheet %))) chart-data))]
                      (when-not replace-self?
                        (remove-sheet! workbook src-sheet))
                      (recur (inc src-index)
                             src-sheets
                             (+ dst-index (count dst-sheets))
                             new-chart-data
                             seen-sheets)))
                  (do
                    (insert-charts workbook chart-data src-sheet src-sheet src-index dst-index)
                    (recur (inc src-index) src-sheets (inc dst-index) chart-data seen-sheets))))))

          (save-workbook! workbook temp-file)))
      (io/copy temp-file excel-file)
      ;; (io/copy temp-file (io/file "/tmp/debug.xlsx"))
      (finally
        (io/delete-file temp-file)))))

(defn replacements-by-sheet-name
  "Convert replacements to a map of concrete sheet name -> sheet data map.

  {'Sheet1' [{:sheet-name 'Sheet1-1' ...} {:sheet-name 'Sheet1-2' ...}]}

  =>

  {'Sheet1-1' {...}
   'Sheet1-2' {...}}"
  [replacements]
  (into {} (for [[template-name sheet-datas] replacements
                 {:keys [sheet-name] :as sheet-data} sheet-datas]
             [sheet-name (dissoc sheet-data :sheet-name)])))

(defn row-map?
  "Returns true if this structure is a row-map. The difference
   between a row map and a worksheet map is that a row map doesn't
   contain any other maps and a worksheet map does."
  [data]
  (letfn [(has-map? [x] (or (map? x) (some map? x)))]
    (not (some has-map? (vals data)))))

(defn normalize
  "Convert replacements to their verbose form.
    * TODO allow single map of replacements to default sheet (0))
    * make sheet-datas vectors
    * add explicity :sheet-name to each sheet-data"
  [replacements]
  (->> replacements
       ;; This method no longer works to support default sheet -- possibly use
       ;; schema here.
       (#(if (row-map? %) {0 %} %))
       (map (juxt key (comp #(if (map? %) (vector %) %) val)))
       add-sheet-names))

(defn zip
  "mix together two seqs like (zip [a b c] [x y z]) => [[a x] [b y] [c z]]"
  [s1 s2]
  (map (fn [x1 x2] [x1 x2]) s1 s2))

(defn expand-replacements
  "Expand any keys of the form [start end] in the replacement map to being individual row
   keys"
  [replacements]
  (fo/map-values
   (fn [v]
     (into
      {}
      (mapcat (fn [[rows repls :as entry]]
                (if (coll? rows)
                  (let [[start end] rows
                        span (- end start -1)
                        repls (concat repls (repeat (- span (count repls)) nil))]
                    (concat (zip (range start end) (map vector repls))
                            [[end (drop (dec span) repls)]]))
                  [entry]))
              v)))
   replacements))

(defn render-to-file
  "Build a report based on a spreadsheet template"
  [template-file output-file replacements]
  (let [tmpfile (File/createTempFile "excel-output" ".xlsx")
        tmpcopy0 (File/createTempFile "excel-template-firstcopy" ".xlsx")
        tmpcopy (File/createTempFile "excel-template-copy" ".xlsx")
        replacements (normalize replacements)]
    (try
      ;; We copy the template file because otherwise POI will modify it even
      ;; though it's our input file. That's annoying from a source code
      ;; control perspective.
      (with-open [template (get-template template-file)]
        (io/copy template tmpcopy0))
      (let [chart-data (charts-from-file tmpcopy0)]
        (filter-zip-file tmpcopy0 tmpcopy (c/chart-removal-rules chart-data))
        (create-missing-sheets! tmpcopy replacements chart-data))
      (build-base-output tmpcopy tmpfile)
      (let [replacements (replacements-by-sheet-name replacements)
            translation-table (fo/build-translation-tables replacements)
            replacements (expand-replacements replacements)]
        (with-open [pkg (OPCPackage/open tmpcopy)]
          (let [template (XSSFWorkbook. pkg)
                intermediate-files (for [index (range (dec (.getNumberOfSheets template)))]
                                     (create-temp-xlsx-file (str "excel-intermediate-" index)))
                inputs  (vec (concat [tmpfile]          intermediate-files))
                outputs (vec (concat intermediate-files [output-file]     ))]
            (try
              (doseq [sheet-num (range (.getNumberOfSheets template))]
                (with-open [input-pkg (OPCPackage/open (nth inputs sheet-num))]
                  (let [src-sheet (.getSheetAt template sheet-num)
                        sheet-name (.getSheetName src-sheet)
                        sheet-data (or (get replacements sheet-name)
                                       (get replacements sheet-num)
                                       {})
                        nrows (inc (apply max (.getLastRowNum src-sheet)
                                          (when (count sheet-data) (keys sheet-data))))
                        src-has-formula? (or (has-formula? src-sheet)
                                             (c/has-chart? src-sheet))
                        wb (XSSFWorkbook. input-pkg)
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
                                      (inject-data-row data-row translation-table wb sheet src-row new-row)
                                      (when src-row
                                        (copy-styles wb src-row new-row))))
                                  (recur (inc src-row-num) (+ dst-row-num (count data-rows))))
                                (do
                                  (when src-row
                                    (let [new-row (.createRow sheet dst-row-num)]
                                      (copy-row translation-table wb sheet src-row new-row)
                                      (copy-styles wb src-row new-row)))
                                  (recur (inc src-row-num) (inc dst-row-num)))))))
                        (c/transform-charts sheet translation-table))
                      ;; Update cached results for each formula:
                      ;; https://poi.apache.org/spreadsheet/eval.html
                      (-> wb .getCreationHelper .createFormulaEvaluator .evaluateAll)
                      ;; Write the resulting output Workbook
                      (with-open [fos (FileOutputStream. (nth outputs sheet-num))]
                        (.write wb fos))
                      (finally
                        (when-not src-has-formula?
                          (.dispose wb)))))))
              (finally
                (doseq [f intermediate-files] (io/delete-file f)))))))
      (finally
        (io/delete-file tmpfile)
        (io/delete-file tmpcopy0)
        (io/delete-file tmpcopy)))))

(defn render-to-stream
  "Build a report based on a spreadsheet template, write it to the output
  stream."
  [template-file output-stream replacements]
  (let [temp-output-file (create-temp-xlsx-file "for-stream-output")]
    (try
      (render-to-file template-file temp-output-file replacements)
      (with-open [input-stream (io/input-stream temp-output-file)]
        (io/copy input-stream output-stream))
      (finally
        (io/delete-file temp-output-file )))))

(comment (let [template-file "foo.xlsx"
               output-file "/tmp/bar.xlsx"

               data {"Sheet1" [{2 [[nil "foo"]]}
                               {2 [[nil "bar"]] :sheet-name "Sheet1-a"}
                               {2 [[nil "baz"]] :sheet-name "Sheet1-b"}]
                     "Sheet2" [{2 [[nil "qux"]]}]}]
           (render-to-file template-file output-file data)
           (sh "libreoffice" "--calc" output-file))

         (let [template-file "foo.xlsx"
               output-file "/tmp/bar.xlsx"
               data {"Sheet1" {2 [[nil "old!"]]}}]
           (render-to-file template-file output-file data)
           (sh "libreoffice" "--calc" output-file)))
