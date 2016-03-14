(ns excel-templates.formulas
  ;; TODO remove unnecessary imports here
  (:import [java.io File FileOutputStream]
           [org.apache.poi.openxml4j.opc OPCPackage]
           [org.apache.poi.ss.formula FormulaParser FormulaRenderer FormulaType]
           [org.apache.poi.ss.usermodel Cell Row DateUtil WorkbookFactory]
           [org.apache.poi.xssf.streaming SXSSFWorkbook]
           [org.apache.poi.xssf.usermodel XSSFWorkbook XSSFEvaluationWorkbook]))


;;; Support for debugging

(def ^:dynamic *debug-prints*
  "Set to true to enable the trace prints that show how formulas are being translated"
  false)

(defn debug-println
  "Print the arguments if *debug-prints* is bound to true."
  [& args]
  (when *debug-prints*
    (apply println args)))


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
  (letfn [(build-entry [[k v]]
            (if (coll? k)
              (let [[start end] k]
                [start (count v) (inc (- end start))])
              [k (count v) 1]))]
    (let [result (sort-by first (map build-entry data-seq))]
      (if (zero? (ffirst result))
        result
        (concat [[0 1 1]] result)))))

;;; I don't think we need this anymore
(defn build-reverse-table
  "Build the translation from destination back to source"
  [xlate-seq]
  (letfn [(add-delta [[row count span diff] [row' count' span']]
            [row' count' span' (dec (- row' row))])
          (to-reverse [[row count span] [row' count' span' diff]]
            [(+ row count diff) count' span'])]
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

;;; There are three cell addresses that we need to consider when we're
;;; translating a reference:
;;;  1) The (row, col) of the reference in the template (the src addr)
;;;  2) The (row, col) of the reference in the output sheet (the dst addr)
;;;  3) The (row, col) of the cell containing the formula in the output sheet
;;;     (the target addr)
;;;
;;; The target address is used to decide if the reference in the formula
;;; is in the same range that we're generating data for. Note that not
;;; all references have a target address (charts and references to other
;;; sheets for instance).

(defn src->dst
  "Translate a cell location from the source template to the target,
   accounting for cell motion and expansion. Returns a [row col] pair."
  [translation-table worksheet [row col] [target-row target-col] max? abs?]
  ;; currently only handles row expansion
  (let [sheet-table (if (string? worksheet) ; we can use a string for tests
                      (translation-table worksheet)
                      (or (translation-table (.getSheetName worksheet))
                          (translation-table (sheet-number worksheet))))
        forward (:src->dst sheet-table)
        reverse (:dst->src sheet-table)
        _ (debug-println "R:" row target-row max? abs?)
        _ (debug-println (partition-by (fn [[start _ span]] (< row (+ start span))) forward))
        [repls-lt [[src-base src-count src-span :as element]]]
        (partition-by (fn [[start _ span]] (< row (+ start span))) forward)

        in-range? (and src-base (<= src-base row))
        [src-base src-count src-span] (if in-range? element [row 1 1])
        src-offset (- row src-base)
        _ (debug-println "S:" in-range? src-base src-count src-span src-offset)

        dst-base (+ (reduce + (for [[_ count] repls-lt] count))
                    (- src-base (count repls-lt))
                    (- (reduce + (for [[_ _ span] repls-lt] (dec span)))))
        target-in-range? (and (<= dst-base target-row) (< target-row (+ dst-base src-count)))
        target-offset (min (dec src-span) (- target-row dst-base)) ; only meaningful when target-in-range?
        dst-row (if target-in-range?
                  (if (and (not abs?) max?)
                    (- target-row (- src-span src-offset 1) (- target-offset src-span -1))
                    (+ dst-base src-offset))
                  (if max?
                    (+ dst-base src-count (- src-offset src-span))
                    (+ dst-base src-offset)))
        _ (debug-println "D:" target-in-range? dst-base target-offset dst-row)]
    [dst-row col]))

(defprotocol PtgTranslator
  "Protocol to handle translating various PTG classes by modifying cell references appropriately"
  (translate-ptg [ptg translation-table sheet target-cell]))

(extend-protocol PtgTranslator
  org.apache.poi.ss.formula.ptg.Ptg
  (translate-ptg [ptg translation-table sheet target-cell]
    ptg)
  org.apache.poi.ss.formula.ptg.RefPtg
  (translate-ptg [ptg translation-table sheet target-cell]
    (let [src-row (.getRow ptg)
          src-col (.getColumn ptg)
          abs?    (not (.isRowRelative ptg))
          [row col] (src->dst translation-table sheet [src-row src-col]  target-cell true abs?)]
      (doto ptg
        (.setRow row)
        (.setColumn col))))
  org.apache.poi.ss.formula.ptg.AreaPtg
  (translate-ptg [ptg translation-table sheet target-cell]
    (let [first-src-row (.getFirstRow ptg)
          first-src-col (.getFirstColumn ptg)
          first-abs?    (not (.isFirstRowRelative ptg))
          last-src-row  (.getLastRow ptg)
          last-src-col  (.getLastColumn ptg)
          last-abs?     (not (.isLastRowRelative ptg))
          adjust-first? (and (not first-abs?) (not= first-src-col last-src-col))
          [first-row first-col] (src->dst translation-table sheet [first-src-row first-src-col]
                                          target-cell adjust-first? first-abs?)
          [last-row last-col]   (src->dst translation-table sheet [last-src-row last-src-col]
                                          target-cell true last-abs?)]
      (doto ptg
        (.setFirstRow first-row)
        (.setFirstColumn first-col)
        (.setLastRow last-row)
        (.setLastColumn last-col))))
  org.apache.poi.ss.formula.ptg.Area3DPtg
  (translate-ptg [ptg translation-table sheet target-cell]
    (let [ex-sheet (.getSheetAt (.getWorkbook sheet) (.getExternSheetIndex ptg))
          first-src-row (.getFirstRow ptg)
          first-src-col (.getFirstColumn ptg)
          first-abs?    (not (.isFirstRowRelative ptg))
          [first-row first-col] (src->dst translation-table ex-sheet [first-src-row first-src-col]  target-cell false first-abs?)
          last-src-row (.getLastRow ptg)
          last-src-col (.getLastColumn ptg)
          last-abs?    (not (.isLastRowRelative ptg))
          [last-row last-col] (src->dst translation-table ex-sheet [last-src-row last-src-col]  target-cell true last-abs?)]
      (doto ptg
        (.setFirstRow first-row)
        (.setFirstColumn first-col)
        (.setLastRow last-row)
        (.setLastColumn last-col)))))

(defprotocol PtgRelocator
  "Protocol to handle translating various PTG classes by modifying cell references appropriately"
  (relocate-ptg [ptg old-index new-index]))

(extend-protocol PtgRelocator
  org.apache.poi.ss.formula.ptg.Ptg
  (relocate-ptg [ptg old-index new-index]
    ptg)

  org.apache.poi.ss.formula.ptg.Area3DPtg
  (relocate-ptg [ptg old-index new-index]
    (when (= old-index (.getExternSheetIndex ptg))
      (.setExternSheetIndex ptg new-index)))

  org.apache.poi.ss.formula.ptg.Ref3DPtg
  (relocate-ptg [ptg old-index new-index]
    (when (= old-index (.getExternSheetIndex ptg))
      (.setExternSheetIndex ptg new-index))))

(defprotocol PtgExternalSheets
  "Protocol to find any external sheet references in the formula and return them"
  (external-sheets-ptg [ptg]))

(extend-protocol PtgExternalSheets
  org.apache.poi.ss.formula.ptg.Ptg
  (external-sheets-ptg [ptg]
    nil)

  org.apache.poi.ss.formula.ptg.Area3DPtg
  (external-sheets-ptg [ptg]
    [(.getExternSheetIndex ptg)])

  org.apache.poi.ss.formula.ptg.Ref3DPtg
  (external-sheets-ptg [ptg]
    [(.getExternSheetIndex ptg)]))

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

(defn relocate-formula
  "Relocate any references to the sheet at old-index in the formula to refer to new-index"
  [workbook sheet old-index new-index formula]
  (let [ptgs (parse workbook (sheet-number sheet) formula)]
    (doseq [ptg ptgs]
      (relocate-ptg ptg old-index new-index))
    (render workbook ptgs)))

(defn external-sheets
  "Find any references to external sheets in the formula (note: all referenced sheets need
   to actually exist)."
  [workbook sheet formula]
  (let [ptgs (parse workbook (sheet-number sheet) formula)]
    (set (map #(->> % (.getSheetAt workbook) .getSheetName)
              (mapcat external-sheets-ptg ptgs)))))
