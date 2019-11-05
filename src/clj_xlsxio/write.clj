(ns clj-xlsxio.write
  (:require [clj-xlsxio.low-level-write :refer :all])
  (:import [com.sun.jna Pointer]
           [java.io File]))

(defn write-row
  [^Pointer handle row]
  (do
    (doseq [cell row]
      (add-cell handle cell))
    (next-row handle)))

(defmulti write-xlsx (fn [x & args] (type x)))

(defmethod write-xlsx String
  [^String filename matrix & {:keys [sheetname] :or {sheetname nil}}]
  (let [handle (open filename sheetname)]
    (do
      (doseq [row matrix]
        (write-row handle row))
      (close handle))))

(defmethod write-xlsx File
  [^File file matrix & {:keys [sheetname] :or {sheetname nil}}]
  (write-xlsx (.getAbsolutePath file) sheetname))
