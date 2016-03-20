(ns excel-templates.charts
  (:import [org.apache.commons.lang3 StringEscapeUtils]
           [org.apache.poi.xssf.usermodel XSSFChartSheet]
           [org.openxmlformats.schemas.drawingml.x2006.chart CTChart$Factory])
  (:require [clojure.data.zip :as zf]
            [clojure.data.zip.xml :as zx]
            [clojure.set :as set]
            [clojure.string :as str]
            [clojure.xml :as xml]
            [clojure.walk :as walk]
            [clojure.zip :as zip]
            [excel-templates.formulas :as fo]))

;;; POI uses Java class wrappers based on Apache XML Beans to manage the contents of Office files.
;;; However the set of classes to describe charts is incredibly complex, so to avoid having a million
;;; special cases, I pull out the XML and edit that directly and then replace the original object.
;;; This works better because there are only a few common cases at the leaves of the XML tree that
;;; we need to transform.

;;; Manipulation of the POI objects for charts

(defmacro mjuxt
  "Like juxt, but for Java methods"
  [& methods]
  `(juxt ~@(map #(list 'memfn %) methods)))

;;; TODO createDrawingPatriarch should be replaced by getDrawingPatriarch when that's available in POI 3.12
(defn get-charts
  "Get the charts from a worksheet"
  [sheet]
  (-> sheet .createDrawingPatriarch .getCharts))

(defn has-chart?
  "Return true if the sheet has any charts on it"
  [sheet]
  (pos? (count (get-charts sheet))))

(defn get-xml
  "Get the XML representation of a chart"
  [chart]
  (-> chart .getCTChart .xmlText))

(defn set-xml
  "Set new XML for the chart"
  [chart xml-str]
  (let [new-chart (CTChart$Factory/parse xml-str)]
    (-> chart .getCTChart (.set new-chart))))

;;; XML transformation for charts

(defn parse-xml
  "Parse the XML string using clojure.xml and return a zipper"
  [xml-string]
  (-> xml-string
      (.getBytes (java.nio.charset.Charset/forName "UTF-8"))
      (java.io.ByteArrayInputStream.)
      xml/parse
      zip/xml-zip))

(defn escape-strings
  "Escape any illegal XML strings"
  [tree]
  (walk/postwalk
   #(if (and (map? %) (contains? % :content))
      (assoc % :content (seq (for [e (:content %)]
                               (if (string? e)
                                 (StringEscapeUtils/escapeXml11 e)
                                 e))))
      %)
   tree))

(defn emit-xml
  "Generate an XML string from a zipper using clojure.xml"
  [loc]
  (-> (with-out-str (-> loc zip/root escape-strings xml/emit))
      (str/replace #"^.*\n" "")
      (str/replace #"(\r?\n|\r)" "")))

(defn transform-formula
  "Transform a single chart formula according to the translation table"
  [sheet translation-table formula]
  (fo/translate-formula translation-table (.getWorkbook sheet) sheet [2000000 2000000] formula))

;;; tree-edit is based on a blog post by Ravi Kotecha at
;;; http://ravi.pckl.me/short/functional-xml-editing-using-zippers-in-clojure/

(defn tree-loc-edit
  "The rawer version of tree edit, this operates on a loc rather than a node.
   As a result, it allows for non-local manipulation of the tree."
  [zipper matcher editor & colls]
  (loop [loc zipper
         colls colls]
    (if (zip/end? loc)
      loc
      (if (matcher loc)
        (let [new-loc (apply editor loc (map first colls))]
          (recur (zip/next new-loc) (map next colls)))
        (recur (zip/next loc) colls)))))

(defn tree-edit
  "Take a zipper, a function that matches a pattern in the tree,
  and a function that edits the current location in the tree.  Examine the tree
  nodes in depth-first order, determine whether the matcher matches, and if so
  apply the editor.
  Optional colls are used as in clojure.core/map with one element of each coll passed
  as an argument to editor in sequence. These will be nil padded if necessary if
  the number of matches is longer that the length of the collection."
  [zipper matcher editor & colls]
  (apply
   tree-loc-edit
   zipper matcher
   (fn [loc & args]
     (apply zip/edit loc editor args))
   colls))

(defn formula?
  "Return true if the node at the loc is a formula"
  [loc]
  (= :c:f (-> loc zip/node :tag)))

(defn series?
  "Return true if the node at the loc is a series"
  [loc]
  (= :c:ser (-> loc zip/node :tag)))

(defn transform-xml
  "Transform the zipper representing the chart into a zipper with expansions"
  [sheet translation-table loc]
  (letfn [(editor [node]
            (assoc node
              :content [(->> node :content first (transform-formula sheet translation-table))]))]
    (tree-edit loc formula? editor)))

(defn chart-transform
  "Transform the formulas in the XML representation of a chart"
  [sheet translation-table chart-xml]
  (->> chart-xml parse-xml (transform-xml sheet translation-table) emit-xml))

;;; Combine the above to edit all charts in a sheet
(defn transform-charts
  "Transform the charts in a sheet according to the translation table"
  [sheet translation-table]
  ;; (println (str "Transforming sheet " (.getSheetName sheet) "(" (-> sheet .getWorkbook (.getSheetIndex sheet)) ")"))
  ;; (println (str "relations = " (-> sheet .createDrawingPatriarch .getRelations)))
  (doseq [chart (get-charts sheet)]
    ;; (println "xform chart")
    (->> chart get-xml (chart-transform sheet translation-table) (set-xml chart))))

(defn relocate-formula
  "Relocate a single chart formula from old-sheet to new-sheet"
  [sheet old-index new-index formula]
  (fo/relocate-formula (.getWorkbook sheet) sheet old-index new-index formula))

(defn relocate-xml
  "Find the formulas in the XML that refer to the sheet at old-index and make them point to the sheet at new-index"
  [sheet old-index new-index loc]
  (letfn [(editor [node]
            (assoc node
              :content [(->> node :content first (relocate-formula sheet old-index new-index))]))]
    (tree-edit loc formula? editor)))

(defn expand-series
  "Expand a single series into 0 or more destination series leaving the cursor such that zip/next
   will return the same result as when called."
  [sheet src-sheet dst-sheets series-loc]
  (let [series-formulas (mapcat :content (zx/xml-> series-loc zf/descendants (zx/tag= :c:f) zip/node))
        sheet-refs (reduce
                    set/union
                    (map
                     (partial fo/external-sheets (.getWorkbook sheet) sheet)
                     series-formulas))
        wb (.getWorkbook sheet)
        src-index (.getSheetIndex wb src-sheet)]
    (if (sheet-refs src-sheet)
      (let [series-xml (-> series-loc zip/node zip/xml-zip)
            base-loc (zip/remove series-loc)]
        (letfn [(add-series [loc dst-sheet]
                  (let [dst-index (.getSheetIndex wb dst-sheet)
                        new-xml (zip/node (relocate-xml sheet src-index dst-index series-xml))]
                    (zip/right (zip/insert-right loc new-xml))))]
          (reduce add-series base-loc dst-sheets)))
      series-loc)))

(defn reindex-series
  "After modifying a chart, make sure that the indices and order of the series is correct"
  [key values chart-xml]
  (letfn [(editor [loc index]
            (zip/edit (zx/xml1-> loc (zx/tag= key)) assoc-in [:attrs :val] (str index)))]
    (tree-loc-edit chart-xml series? editor values)))

(defn expand-all-series
  "Replicate the various series"
  [sheet src-sheet dst-sheets chart-xml]
  (tree-loc-edit chart-xml series? (partial expand-series sheet src-sheet dst-sheets)))

(defn px
  "Put this in the middle of a thread op to print the current state and keep threading the operand"
  [x]
  (println "px:" x)
  x)

(defn expand-xml-str
  "Replace any series in a chart that references a sheet that's being cloned to point to all
   the clones. "
  [sheet src-sheet dst-sheets xml-str]
  (->> xml-str
       parse-xml
       (expand-all-series sheet src-sheet dst-sheets)
       zip/root
       zip/xml-zip
       (reindex-series :c:idx (range))
       zip/root
       zip/xml-zip
       (reindex-series :c:order (range))
       emit-xml))

(defn expand-charts
  "Replace any series in charts on the sheet that reference src-sheet with multiple series
   each referencing a single element of dst-sheets"
  [sheet src-sheet dst-sheets]
  (doseq [chart (get-charts sheet)]
    (->> chart
         get-xml
         (expand-xml-str sheet src-sheet dst-sheets)
         (set-xml chart))))

;;; When we duplicate a sheet with charts on it, we need to make sure
;;; that any charts on that sheet point to the new sheet in any places
;;; where they were pointing to the base sheet

(defn chart-change-sheet
  "Handle a single chart that was duplicated from an old sheet to a new sheet"
  [sheet old-index new-index chart-xml]
  (->> chart-xml parse-xml (relocate-xml sheet old-index new-index) emit-xml))

(defn change-sheet
  "Update any reference in the charts on this sheet that points to the base sheet to
   point to this sheet"
  [sheet old-index new-index]
  (println (str "Changing sheet " (.getSheetName sheet) "(" (-> sheet .getWorkbook (.getSheetIndex sheet)) ") from " old-index " to " new-index))
  (println (str "relations = " (-> sheet .createDrawingPatriarch .getRelations count)))
  (doseq [chart (get-charts sheet)]
    (println "found chart")
    (->> chart get-xml (chart-change-sheet sheet old-index new-index) (set-xml chart))))

;;; Code for copying charts when we're duplicating a sheet
;;; Because POI can't clone a sheet with charts on it, we have to do the
;;; following:
;;; 1) Get the data about all the charts that are on worksheets
;;; 2) Delete charts from the worksheets (leave charts on the chartsheets, because they're different)
;;; 3) Rename the charts on chartsheets to be chart1, chart2, etc. because of the way POI creates
;;;    new charts
;;; 4) Add the charts back onto the sheets after they've been cloned (we do all worksheets because
;;;    it's easier that restricting to only cloned ones).
;;; 5) Make any charts that point to the new cloned charts have the right references
;;; 6) Do the same for all the saved charts since some of them won't have been added back into the
;;;    sheets yet.
;;;
;;; The logic to do this is split between here and create-missing-sheets in build.clj

(defn chart-sheet?
  "Return true if this sheet is a chart sheet"
  [sheet]
  (instance? XSSFChartSheet sheet))

(defn anchors-by-id
  "Gets a map of anchor objects by ID that show where the graphic with that ID is on the sheet"
  [sheet]
  (let [anchors (-> sheet .createDrawingPatriarch .getCTDrawing .getTwoCellAnchorList)]
    (into {} (for [anchor anchors]
               [(-> anchor .getGraphicFrame .getGraphic .getGraphicData .getDomNode
                    .getChildNodes (.item 0) (.getAttribute "r:id"))
                anchor]))))

(defn get-part-id
  "Get the part id for a document part in the drawing patriarch"
  [sheet part]
  (.getRelationId (.createDrawingPatriarch sheet) part))

(defn get-charts-and-anchors
  "Get maps representing each chart on the sheet along with its anchor"
  [sheet]
  (let [anchors (anchors-by-id sheet)]
    (for [chart (get-charts sheet)]
      {:chart chart, :anchor (anchors (get-part-id sheet chart))})))

(defn new-anchor
  "Get an anchor for a duplicated chart based on an anchor pulled from the original"
  [sheet old-anchor]
  (let [from (.getFrom old-anchor)
        to (.getTo old-anchor)]
    (.createAnchor (.createDrawingPatriarch sheet)
                   (.getColOff from) (.getRowOff from)
                   (.getColOff to)   (.getRowOff to)
                   (.getCol from)    (.getRow from)
                   (.getCol to)      (.getRow to))))


(defn get-anchor-location
  "Get the location info from an anchor so we can create a new one later"
  [anchor]
  (when anchor
    (zipmap [:from :to]
            (map (comp (partial zipmap [:col-off :row-off :col :row])
                       (mjuxt getColOff getRowOff getCol getRow))
                 ((mjuxt getFrom getTo) anchor)))))

(defn part-path
  "Get the path to this part for this object in the zip file"
  [part]
  (-> part .getPackagePart .getPartName .getName (subs 1)))

(defn rels-path
  "Get the path to the relationship definitions for this object in the zip file"
  [part]
  (let [[_ head tail] (re-matches #"^(.*)/([^/]+)" (part-path part))]
    (str head "/_rels/" tail ".rels")))

(defn get-chart-data
  "Get all the data the we need to delete and then recreate the charts for this sheet"
  [sheet]
  (let [anchors (anchors-by-id sheet)]
    (for [chart (get-charts sheet)
          :let [drawing (.createDrawingPatriarch sheet)
                chart-id (get-part-id sheet chart)]]
      {:sheet          (.getSheetName sheet)
       :chart-sheet?   (chart-sheet? sheet)
       :chart-path     (part-path chart)
       :drawing-path   (part-path drawing)
       :drawing-rels   (rels-path drawing)
       :chart-id       chart-id
       :chart-location (get-anchor-location (anchors chart-id))
       :chart-xml      (get-xml chart)})))

(defn chart-sheets
  "filter the chart data for charts from chart sheets only"
  [chart-data]
  (filter :chart-sheet? chart-data))

(defn work-sheets
  "filter the chart data for charts from worksheets only"
  [chart-data]
  (filter (complement :chart-sheet?) chart-data))

(defn remove-charts
  "Returns a map of chart names to a map with :delete set to true. This will cause the chart objects to be
  dropped."
  [chart-data]
  (into {}
        (for [chart (work-sheets chart-data)] [(:chart-path chart) {:delete true}])))

(defn remove-drawing-rels
  "Returns a map of relationship sheets to a function that will remove the correct relationships on each one"
  [chart-data]
  (let [ids-by-rels (fo/map-values
                     #(set (map :chart-id %))
                     (group-by :drawing-rels (work-sheets chart-data)))]
    (fo/map-values
     (fn [id-set]
       {:edit
        (fn [xml-data]
           (assoc xml-data :content (filter #(not (id-set (get-in % [:attrs :Id])))
                                            (:content xml-data))))})
     ids-by-rels)))

(defn remove-anchors
  "Returns a map of drawing sheets to functions that will move the anchors corresponding to the charts"
  [chart-data]
  (let [ids-by-drawings (fo/map-values
                         #(set (map :chart-id %))
                         (group-by :drawing-path (work-sheets chart-data)))]
    (fo/map-values
     (fn [id-set]
       {:edit
        (fn [xml-data]
           (loop [data xml-data]
             (if-let [new-data (zx/xml1->
                                (zip/xml-zip data)
                                zf/descendants
                                (zx/tag= :c:chart)
                                #(boolean (id-set (zx/attr % :r:id)))
                                zf/ancestors
                                (zx/tag= :xdr:twoCellAnchor)
                                zip/remove
                                zip/root)]
               (recur new-data)
               data)))})
     ids-by-drawings)))

(defn drawing-rel
  "Change a chart reference to be relative to the drawing object"
  [link]
  (.replaceFirst link "^xl/" "../"))

(defn renumber-chart-sheets
  "Returns a map with instructions about how to renumber the charts on chart sheets so that they
   a 1..n with no holes so that POI can re-add the charts on worksheets correctly."
  [chart-data]
  (let [chart-sheet-data (->> chart-data
                              (filter :chart-sheet?)
                              (map-indexed #(assoc %2 :new-chart-path
                                                   (format "xl/charts/chart%d.xml" (inc %1)))))]
    (apply
     merge
     (for [c chart-sheet-data
           :let [path (:chart-path c)
                 rel-path (drawing-rel path)
                 new-path (:new-chart-path c)
                 rels (:drawing-rels c)]]
       {path  {:rename new-path}
        rels  {:edit (fn [xml-data]
                       (assoc xml-data
                              :content (map #(if (= rel-path (get-in % [:attrs :Target]))
                                               (assoc-in % [:attrs :Target] (drawing-rel new-path))
                                               %)
                                            (:content xml-data))))}}))))

(defn chart-removal-rules
  "Combine all the rules to remove charts from this workbook"
  [chart-data]
  (apply merge
         (map #(% chart-data)
              [remove-charts remove-drawing-rels remove-anchors renumber-chart-sheets])))

(defn expand-chart-data
  "Expand the xml charts that we've captured if they have references to sheets that are being cloned"
  [workbook src-sheet dst-sheets chart-data]
  (doall
   (for [{:keys [sheet chart-xml] :as chart} chart-data
         :let [sheet-obj (.getSheet workbook sheet)]
         :when sheet-obj] ;;; if sheet-obj is nil, this chart has already been added back
     (assoc chart :chart-xml (expand-xml-str sheet-obj src-sheet dst-sheets chart-xml)))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; When a chart refers to a series on a sheet that's been duplicated, duplicate the series to match

(defn chart-formulas
  "Get all the formulas in each chart on the sheet"
  [sheet]
  (for [chart (get-charts sheet)
        :let [chart-xml (-> chart get-xml parse-xml)]]
    (mapcat :content (zx/xml-> chart-xml zf/descendants (zx/tag= :c:f) zip/node))))
