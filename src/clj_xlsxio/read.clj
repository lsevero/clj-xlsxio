(ns clj-xlsxio.read
  (:require [clj-xlsxio.low-level-read :refer :all])
  (:import [com.sun.jna Pointer]
           [java.util Date TimeZone]
           [org.joda.time DateTime]
           [java.time LocalDateTime Instant]
           [java.io File]))

(def ^:const skip-none 0)
(def ^:const skip-empty-rows 0x01)
(def ^:const skip-empty-cells 0x02)
(def ^:const skip-all-empty (bit-or skip-empty-rows skip-empty-cells))
(def ^:const skip-extra-cells 0x04)

(defn read-row
  ([sheet]
   (if-not (zero? (sheet-next-row sheet))
     (loop
       [res []]
       (if-let [cell-value (sheet-next-cell sheet)]
         (recur (conj res cell-value))
         res))
     nil))
  ([sheet xlsx]
   (if-not (zero? (sheet-next-row sheet))
     (loop
       [res []]
       (if-let [cell-value (sheet-next-cell sheet)]
         (recur (conj res cell-value))
         res))
     (do 
       (sheet-close sheet)
       (close xlsx)
       nil))))

(defmulti read-xlsx (fn [x & args] (type x)))

(defmethod read-xlsx Pointer
  ([sheet]
   (if-let [first-row (read-row sheet)]
     (lazy-seq (cons first-row (read-xlsx sheet)))
     nil))
  ([sheet xlsx]
   (if-let [first-row (read-row sheet)]
     (lazy-seq (cons first-row (read-xlsx sheet)))
     (do 
       (sheet-close sheet)
       (close xlsx)
       nil))))

(defmethod read-xlsx String
  [filename & {:keys [skip sheet] :or {skip skip-none sheet nil}}]
  (let [xlsx (open filename)
        sheet (sheet-open xlsx sheet skip)]
    (read-xlsx sheet xlsx)))

(defmethod read-xlsx File
  [file & {:keys [skip sheet] :or {skip skip-none sheet nil}}]
  (let [^String filename (.getAbsolutePath file)]
    (read-xlsx filename :sheet sheet :skip skip)))

(defn xlsx->enumerated-maps
  [lz-seq]
  (let [n-columns (-> lz-seq first count)]
       (map zipmap (repeat (range n-columns)) lz-seq)))

(defn- int->excel-column
  [^Long n]
  (loop [s []
         aux (inc n)]
    (if (>= (dec aux) 0)
        (recur (conj s (char (+ (int \A) (mod (dec aux) 26)))) (quot (dec aux) 26))
        (reduce str (reverse s)))))

(defn xlsx->excel-enumerated-maps
  [lz-seq]
  (let [n-columns (-> lz-seq first count)]
    (map zipmap (repeat (map (comp keyword str int->excel-column) (range n-columns))) lz-seq)))

(defn xlsx->column-title-maps
  "Takes the first row of a xlsx and enumerate every row with the column title"
  [lz-seq]
  (map zipmap
       (->> (first lz-seq)
            (map keyword)
            repeat)
    (rest lz-seq)))

(defn coerce
  "Coerce every row applying a vector of functions"
  [lz-seq fs & {:keys [skip-first-row] :or {skip-first-row false}}]
  (if skip-first-row
    (let [[head & tail] lz-seq]
      (cons head (map (fn [row] (mapv #(%1 %2) fs row)) tail)))
    (map (fn [row] (mapv #(%1 %2) fs row)) lz-seq)))

(defn excel-date->unix-timestamp
  ^Long
  [^String n-str]
  (let [n (Double/parseDouble n-str)]
    (long (* 86400 (- n 25569)))))

(defn excel-date->java-date
  ^Date
  [^String n-str]
  (Date. (* 1000 (excel-date->unix-timestamp n-str))))

(defn excel-date->joda-date
  ^DateTime
  [^String n-str]
  (DateTime. (* 1000 (excel-date->unix-timestamp n-str))))

(defn excel-date->java-localdatetime
  ^LocalDateTime
  [^String n-str]
  (LocalDateTime/ofInstant (Instant/ofEpochSecond (excel-date->unix-timestamp n-str))
                           (-> (TimeZone/getDefault) .toZoneId)))

(defprotocol ListSheets
  (list-sheets [this]))

(extend-protocol ListSheets
  Pointer (list-sheets [xlsx]
            (let [sheetlist (sheetlist-open xlsx)
                  res (loop [sheets []]
                        (if-let [sheetname (sheetlist-next sheetlist)]
                          (recur (conj sheets sheetname))
                          sheets))]
              (sheetlist-close sheetlist)
              res))
  String  (list-sheets [filename]
            (let [^Pointer xlsx (open filename)
                  res (list-sheets xlsx)]
              (close xlsx)
              res))
  File    (list-sheets [file]
            (let [^String filename (.getAbsolutePath file)]
              (list-sheets filename))))
