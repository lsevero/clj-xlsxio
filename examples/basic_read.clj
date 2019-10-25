(ns basic-read
  (:require [clj-xlsxio.read :as read]
            [clj-xlsxio.low-level-read :as low-level]))

(defn basic-read
  []
  (do
    (println "Basic reading:") 
    (pr (read/read-xlsx "examples/test.xlsx"))
    (println)
    (println "Reading skipping empty rows:")
    (pr (read/read-xlsx "examples/test.xlsx" :skip read/skip-empty-rows))
    (println)
    (println "Reading skipping empty cells:")
    (pr (read/read-xlsx "examples/test.xlsx" :skip read/skip-empty-cells))
    (println)
    (println "Reading skipping all empty:")
    (pr (read/read-xlsx "examples/test.xlsx" :skip read/skip-all-empty))
    (println)
    (println "Reading skipping all extra")
    (pr (read/read-xlsx "examples/test.xlsx" :skip read/skip-extra-cells))
    (println)
    (let [xlsx (low-level/open "examples/test.xlsx")
          sheet (low-level/sheet-open xlsx)]
      (println "Reading directly from the sheet object")
      (pr (read/read-xlsx sheet xlsx))
      (println)
      (low-level/sheet-close sheet)
      (low-level/close xlsx))

    (println "Xlsx rows to enumerated maps by column:")
    (pr (-> (read/read-xlsx "examples/test.xlsx") read/xlsx->enumerated-maps))
    (println)
    
    (println "Xlsx rows to excel enumerated maps by column:")
    (pr (-> (read/read-xlsx "examples/test.xlsx") read/xlsx->excel-enumerated-maps))
    (println)))
