(ns excel-templates.charts
  (:import [org.openxmlformats.schemas.drawingml.x2006.chart CTChart$Factory])
  (:require [clojure.data.zip :as zf]
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

(defn tree-edit
  "Take a zipper, a function that matches a pattern in the tree,
   and a function that edits the current location in the tree.  Examine the tree
   nodes in depth-first order, determine whether the matcher matches, and if so
   apply the editor."
  [zipper matcher editor]
  (loop [loc zipper]
    (if (zip/end? loc)
      loc
      (if (matcher loc)
        (let [new-loc (zip/edit loc editor)]
          (recur (zip/next new-loc)))
        (recur (zip/next loc))))))

(defn transform-xml
  "Transform the zipper representing the chart into a zipper with expansions"
  [sheet translation-table loc]
  (letfn [(formula? [loc]
            (= :c:f (-> loc zip/node :tag)))
          (editor [node]
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
  (doseq [chart (get-charts sheet)]
    (->> chart get-xml (chart-transform sheet translation-table) (set-xml chart))))
