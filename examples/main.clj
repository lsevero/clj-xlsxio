(ns main
  (:gen-class)
  (:require [basic-read-low-level :refer :all]
            [basic-read :refer :all]
            ))

(defn -main
  []
  (basic-read)
  (basic-read-low-level)
  
  )
