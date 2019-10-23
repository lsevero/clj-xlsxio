(ns basic-read
  (:require [clj-xlsxio.core :as read]))

(defn basic-read
  []
  (let [xlsx (read/open "examples/test.xlsx")
        sheet (read/sheet-open xlsx)]
    (println "Basic read:")
    (pr (read/read-xlsx sheet))
    (println)
    (read/sheet-close sheet)
    (read/close xlsx)))
