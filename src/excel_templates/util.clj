(ns excel-templates.util
  (:import [org.apache.poi.ss.util WorkbookUtil]))

(defn safe-sheet-name
  "Sanitize the sheet name provided, replacing invalid characters.
   Excel doesn't allow certain special characters, like  [ ] * / ?"
  ([name]
   (WorkbookUtil/createSafeSheetName name))
  ([name replacement]
   (WorkbookUtil/createSafeSheetName name replacement)))
