(ns clj-xlsxio.read
  (:require [clj-xlsxio.low-level-read :refer :all])
  (:import [com.sun.jna Pointer]
           [java.util Date]))

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
  ([filename & {:keys [skip] :or {skip skip-none}}]
   (let [xlsx (open filename)
         sheet (sheet-open xlsx skip)]
     (read-xlsx sheet xlsx))))

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

(defn coerce
  [lz-seq fs]
  (map (fn [row] (mapv #(%1 %2) fs row)) lz-seq))

(defn excel-date->unix-timestamp
  ^Long
  [^String n-str]
  (let [n (Long/parseLong n-str)]
    (* 86400 (- n 25569))))

(defn excel-date->java-date
  ^Date
  [^String n-str]
  (Date. (* 1000 (excel-date->unix-timestamp n-str))))
