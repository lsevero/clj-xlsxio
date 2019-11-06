(ns basic-write
  (:require [clj-xlsxio.write :as write])
  (:import [java.util Date]))

(defn basic-write []
  (println "Basic write test.")
  (println "You can the following classes to a xlsx file:")
  (println "String, Long, Double, Ratio, java.util.Date, org.joda.time.DateTime, java.time.LocalDateTime")
  (write/write-xlsx "/tmp/basic-write-test.xlsx"
                    [["abc" 1 1.1 "aaaaaaaaaa"] 
                     ["oahfdihifuh" 123456789012345 (Date.) 1/3]]))
