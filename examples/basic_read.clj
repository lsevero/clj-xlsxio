(ns basic-read
  (:require [clj-xlsxio.core :as read]))

(defn basic-read
  []
  (let [xlsx (read/open "examples/test.xlsx")
        sheet (read/sheet-open xlsx nil read/SKIP_EMPTY_ROWS)]
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
