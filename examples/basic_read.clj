(ns basic-read
  (:require [clj-xlsxio.read :as read]
            [clj-xlsxio.low-level-read :as low-level]
            [clojure.string :as st]
            [clojure.java.io :as io]))

(defn basic-read
  []
  (do
    (println "Basic reading:") 
    (prn (read/read-xlsx "examples/test.xlsx"))
    (println "Reading skipping empty rows:")
    (prn (read/read-xlsx "examples/test.xlsx" :skip read/skip-empty-rows))
    (println "Reading skipping empty cells:")
    (prn (read/read-xlsx "examples/test.xlsx" :skip read/skip-empty-cells))
    (println "Reading skipping all empty:")
    (prn (read/read-xlsx "examples/test.xlsx" :skip read/skip-all-empty))
    (println "Reading skipping all extra")
    (prn (read/read-xlsx "examples/test.xlsx" :skip read/skip-extra-cells))
    (let [xlsx (low-level/open "examples/test.xlsx")
          sheet (low-level/sheet-open xlsx)]
      (println "Reading directly from the sheet object")
      (prn (read/read-xlsx sheet xlsx))
      (low-level/sheet-close sheet)
      (low-level/close xlsx))

    (println "Read xlsx from java File")
    (prn (read/read-xlsx (io/file "examples/test.xlsx")))
    
    (println "Xlsx rows to enumerated maps by column:")
    (prn (-> (read/read-xlsx "examples/test.xlsx") read/xlsx->enumerated-maps))
    
    (println "Xlsx rows to excel enumerated maps by column:")
    (prn (-> (read/read-xlsx "examples/test.xlsx") read/xlsx->excel-enumerated-maps))
    
    (println "Xlsx rows to column title maps:")
    (prn (-> (read/read-xlsx "examples/test.xlsx") read/xlsx->column-title-maps))

    (println "Xlsx column coertion:")
    (prn (-> (read/read-xlsx "examples/coerce_test.xlsx") (read/coerce [(comp inc #(Long/parseLong %)) st/upper-case read/excel-date->java-date])))

    (println "Xlsx column coertion skipping first row (useful when the first row are titles):")
    (prn (-> (read/read-xlsx "examples/coerce_test.xlsx")
             (read/coerce [(comp inc #(Long/parseLong %)) st/upper-case read/excel-date->java-date]
                          :skip-first-row true)))

    (println "Listing sheets inside xlsx:")
    (prn (read/list-sheets "examples/test.xlsx"))

    (println "Listing sheets inside xlsx on java File:")
    (prn (read/list-sheets (io/file "examples/test.xlsx")))))
