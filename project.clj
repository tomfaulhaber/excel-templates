(defproject com.infolace/excel-templates "0.3.0"
  :description "Build Excel files by combining a template with plain old data"
  :url "https://github.com/tomfaulhaber/excel-templates"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.7.0"]
                 [cider/cider-nrepl   "0.9.1"]
                 [org.apache.poi/poi-ooxml "3.10-FINAL"]
                 [org.apache.poi/ooxml-schemas "1.1"]
                 [org.clojure/data.zip "0.1.1"]
                 [joda-time "2.7"]]
  :repl-options {:port 4001
                 :nrepl-middleware [cider.nrepl.middleware.apropos/wrap-apropos
                                    cider.nrepl.middleware.classpath/wrap-classpath
                                    cider.nrepl.middleware.complete/wrap-complete
                                    cider.nrepl.middleware.debug/wrap-debug
                                    cider.nrepl.middleware.format/wrap-format
                                    cider.nrepl.middleware.info/wrap-info
                                    cider.nrepl.middleware.inspect/wrap-inspect
                                    cider.nrepl.middleware.macroexpand/wrap-macroexpand
                                    cider.nrepl.middleware.ns/wrap-ns
                                    cider.nrepl.middleware.pprint/wrap-pprint
                                    cider.nrepl.middleware.refresh/wrap-refresh
                                    cider.nrepl.middleware.resource/wrap-resource
                                    cider.nrepl.middleware.stacktrace/wrap-stacktrace
                                    cider.nrepl.middleware.test/wrap-test
                                    cider.nrepl.middleware.trace/wrap-trace
                                    cider.nrepl.middleware.undef/wrap-undef]})
