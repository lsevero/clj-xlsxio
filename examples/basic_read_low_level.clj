(ns basic-read-low-level
  (:require [clj-xlsxio.low-level-read :as read]
            [clj-xlsxio.read :refer [skip-empty-rows]]))

(defn basic-read-low-level
  []
  (let [xlsx (read/open "examples/test.xlsx")
        sheet (read/sheet-open xlsx nil skip-empty-rows)]
    (println "Low level reading:")
    (read/sheet-next-row sheet)
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (println)
    (read/sheet-next-row sheet)
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (println)
    (read/sheet-next-row sheet)
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (println)
    (read/sheet-next-row sheet)
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (print (read/sheet-next-cell sheet) "\t")
    (println)
    (read/sheet-close sheet)
    (read/close xlsx)))
