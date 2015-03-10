(ns excel-templates.formulas
  ;; TODO remove unnecessary imports here
  (:import [java.io File FileOutputStream]
           [org.apache.poi.openxml4j.opc OPCPackage]
           [org.apache.poi.ss.formula FormulaParser FormulaRenderer FormulaType]
           [org.apache.poi.ss.usermodel Cell Row DateUtil WorkbookFactory]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFWorkbook XSSFEvaluationWorkbook]))

;;; Translate row and column from templates into target space

(defn map-values
  "Build a new map that has [k v] -> [k (f v)]. I don't know why clojure doesn't have this"
  [f m]
  (into {} (map (fn [[k v]] [k (f v)]) m)))

(defn sheet-number
  "Find the sheet number of this sheet object within its workbook"
  [sheet]
  (let [wkb (.getWorkbook sheet)
        sheet-names (map #(.getSheetName (.getSheetAt wkb %))
                         (range (.getNumberOfSheets wkb)))]
    (.indexOf sheet-names (.getSheetName sheet))))

(defn build-forward-table
  "Build the translation table to from source to dest"
  [data-seq]
  (sort-by first (map-values count data-seq)))

(defn build-reverse-table
  "Build the translation from destination back to source"
  [xlate-seq]
  (letfn [(add-delta [[a b c] [a' b']]
            [a' b' (dec (- a' a))])
          (to-reverse [[a b] [a' b' d]]
            [(+ a b d) b'])]
    (let [with-deltas (next (reductions add-delta [-1 0 0] xlate-seq))
          reverse-seq (next (reductions to-reverse [0 0] with-deltas))]
      reverse-seq)))

(defn build-translation-tables
  "Build the translation tables for all the worksheets mentioned in the data map"
  [data-map]
  (map-values
   (fn [data-seq]
     (let [src->dst (build-forward-table data-seq)
           dst->src (build-reverse-table src->dst)]
       {:src->dst src->dst, :dst->src dst->src}))
   data-map))

(defn src->dst
  "Translate a cell location from the source template to the target,
   accounting for cell motion and expansion. Returns a [row col] pair."
  [translation-table worksheet [row col] [target-row target-col] max?]
  ;; currently only handles row expansion
  (let [sheet-table (or (translation-table (.getSheetName worksheet))
                        (translation-table (sheet-number worksheet))
                        nil)
        forward (:src->dst sheet-table)
        reverse (:dst->src sheet-table)
        [repls-lt [[maybe-this-row maybe-this-count]]]
        (partition-by #(< (first %) row) forward)
        this-row-count (if (= maybe-this-row row) maybe-this-count 1)
        prev-rows (+ (reduce + (for [[_ count] repls-lt] count))
                     (- row (count repls-lt)))
        [start size] (last (take-while #(<= (first %) prev-rows) reverse))]
    [(+ prev-rows
        (if max?
          (if (and start (< target-row (+ start size)))
           (- target-row start)
           (- this-row-count 1))
          0))
     col]))

(defprotocol PtgTranslator
  "Protocol to handle translating various PTG classes by modifying cell references appropriately"
  (translate-ptg [ptg translation-table sheet target-cell]))

(extend-protocol PtgTranslator
  org.apache.poi.ss.formula.ptg.Ptg
  (translate-ptg [ptg translation-table sheet target-cell]
    ptg)
  org.apache.poi.ss.formula.ptg.RefPtg
  (translate-ptg [ptg translation-table sheet target-cell]
    ;; TODO: same translation for both relative and absolute refs?
    (let [src-row (.getRow ptg)
          src-col (.getColumn ptg)
          [row col] (src->dst translation-table sheet [src-row src-col]  target-cell true)]
      (doto ptg
        (.setRow row)
        (.setColumn col))))
  org.apache.poi.ss.formula.ptg.AreaPtg
  (translate-ptg [ptg translation-table sheet target-cell]
    ;; TODO: same translation for both relative and absolute refs?
    (let [first-src-row (.getFirstRow ptg)
          first-src-col (.getFirstColumn ptg)
          [first-row first-col] (src->dst translation-table sheet [first-src-row first-src-col]  target-cell false)
          last-src-row (.getLastRow ptg)
          last-src-col (.getLastColumn ptg)
          [last-row last-col] (src->dst translation-table sheet [last-src-row last-src-col]  target-cell true)]
      (doto ptg
        (.setFirstRow first-row)
        (.setFirstColumn first-col)
        (.setLastRow last-row)
        (.setLastColumn last-col)))))

(defn parse
  "Parse a formula into a PTG array"
  [workbook sheet-num formula-string]
  (let [evwb (XSSFEvaluationWorkbook/create workbook)]
    (FormulaParser/parse formula-string evwb FormulaType/CELL sheet-num)))

(defn render
  "Render a PTG array back to a formula string"
  [workbook ptgs]
  (let [evwb (XSSFEvaluationWorkbook/create workbook)]
    (FormulaRenderer/toFormulaString evwb ptgs)))

(defn translate-formula
  "Translate a formula from the source sheet to the output workbook based on the translation table"
  [translation-table workbook sheet target-cell formula]
  (let [ptgs (parse workbook (sheet-number sheet) formula)]
    (doseq [ptg ptgs]
      (translate-ptg ptg translation-table sheet target-cell))
    (render workbook ptgs)))
