(ns excel-templates.charts
  (:import [org.openxmlformats.schemas.drawingml.x2006.chart CTChart$Factory])
  (:require [clojure.data.zip :as zf]
            [clojure.data.zip.xml :as zx]
            [clojure.set :as set]
            [clojure.string :as str]
            [clojure.xml :as xml]
            [clojure.zip :as zip]
            [excel-templates.formulas :as fo]))

;;; POI uses Java class wrappers based on Apache XML Beans to manage the contents of Office files.
;;; However the set of classes to describe charts is incredibly complex, so to avoid having a million
;;; special cases, I pull out the XML and edit that directly and then replace the original object.
;;; This works better because there are only a few common cases at the leaves of the XML tree that
;;; we need to transform.

;;; Manipulation of the POI objects for charts

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

(defn emit-xml
  "Generate an XML string from a zipper using clojure.xml"
  [loc]
  (-> (with-out-str (-> loc zip/root xml/emit))
      (str/replace #"^.*\n" "")
      (str/replace "\n" "")))

(defn transform-formula
  "Transform a single chart formula according to the translation table"
  [sheet translation-table formula]
  (fo/translate-formula translation-table (.getWorkbook sheet) sheet [2000000 2000000] formula))

;;; tree-edit is based on a blog post by Ravi Kotecha at
;;; http://ravi.pckl.me/short/functional-xml-editing-using-zippers-in-clojure/

(defn tree-loc-edit
  "The rawer version of tree edit, this operates on a loc rather than a node.
   As a result, it allows for non-local manipulation of the tree."
  [zipper matcher editor]
  (loop [loc zipper]
    (if (zip/end? loc)
      loc
      (if (matcher loc)
        (let [new-loc (editor loc)]
          (recur (zip/next new-loc)))
        (recur (zip/next loc))))))

(defn tree-edit
  "Take a zipper, a function that matches a pattern in the tree,
  and a function that edits the current location in the tree.  Examine the tree
  nodes in depth-first order, determine whether the matcher matches, and if so
  apply the editor."
  [zipper matcher editor]
  (tree-loc-edit zipper matcher #(zip/edit % editor)))

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

(defn expand-all-series
  "Replicate the various series"
  [sheet src-sheet dst-sheets chart-xml]
  (tree-loc-edit chart-xml series? (partial expand-series sheet src-sheet dst-sheets)))

(defn px [x] (println "px:" x) x)
(defn expand-charts
  "Replace any series in charts on the sheet that reference src-sheet with multiple series
   each referencing a single element of dst-sheets"
  [sheet src-sheet dst-sheets]
  (doseq [chart (get-charts sheet)]
    (->> chart get-xml parse-xml (expand-all-series sheet src-sheet dst-sheets) emit-xml px (set-xml chart))))

;;; NOTE: Everything below here is for relocating charts when sheets are
;;; cloned, but we're not using this right now because POI has some
;;; fundamental problems with cloning sheets with drawing objects on them.
;;; I think I can work around this, but I don't have time right now.

;;; When we duplicate a sheet with charts on it, we need to make sure
;;; that any charts on that sheet point to the new sheet in any places
;;; where they were pointing to the base sheet



(defn chart-change-sheet
  "Handle a single chart that was duplicated from an old sheet to a new sheet"
  [sheet old-index new-index chart-xml]
  (->> chart-xml parse-xml (relocate-xml sheet old-index new-index) emit-xml))

(defn change-sheet
  "Update any reference in the charts on this sheet that point to the base sheet to
   point to this sheet"
  [sheet old-index new-index]
  (println (str "Changing sheet " (.getSheetName sheet) "(" (-> sheet .getWorkbook (.getSheetIndex sheet)) ") from " old-index " to " new-index))
  (println (str "relations = " (-> sheet .createDrawingPatriarch .getRelations count)))
  (doseq [chart (get-charts sheet)]
    (println "found chart")
    (->> chart get-xml (chart-change-sheet sheet old-index new-index) (set-xml chart))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; The following code represents an error in thinking on my part, I think
;; I'm leaving it here until I'm sure.

;;; Code for copying charts when we're duplicating a sheet

(defn anchors-by-id
  "Gets a map of anchor objects by ID that show where the graphic with that ID is on the sheet"
  [sheet]
  (let [anchors (-> sheet .createDrawingPatriarch bean :CTDrawing .getTwoCellAnchorList)]
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

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; When a chart refers to a series on a sheet that's been duplicated, duplicate the series to match

(defn chart-formulas
  "Get all the formulas in each chart on the sheet"
  [sheet]
  (for [chart (get-charts sheet)
        :let [chart-xml (-> chart get-xml parse-xml)]]
    (mapcat :content (zx/xml-> chart-xml zf/descendants (zx/tag= :c:f) zip/node))))
