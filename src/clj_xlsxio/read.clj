(ns clj-xlsxio.read
  (:require [clj-xlsxio.low-level-read :refer :all]
            [clojure.set :as set]
            [clojure.java.io :as io])
  (:import [com.sun.jna Pointer]
           [java.util Date TimeZone]
           [org.joda.time DateTime]
           [java.time LocalDateTime Instant]
           [java.io File]
           [java.io BufferedInputStream]))

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
  [^File file & {:keys [skip sheet] :or {skip skip-none sheet nil}}]
  (let [^String filename (.getAbsolutePath file)]
    (read-xlsx filename :sheet sheet :skip skip)))

(defmethod read-xlsx BufferedInputStream
  [^BufferedInputStream stream & {:keys [skip sheet] :or {skip skip-none sheet nil}}]
  (let [fname (str (gensym))
        fext (str (gensym))
        tmp (File/createTempFile fname fext)]
    (.deleteOnExit tmp)
    (with-open [in stream
                out (java.io.FileOutputStream. tmp)]
      (io/copy in out))
    (read-xlsx (.getAbsolutePath tmp) :sheet sheet :skip skip)))

(defn xlsx->enumerated-maps
  "Returns a lazy sequence of maps, keys are the number of the column in the excel format"
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
  "Returns a lazy sequence of maps, keys are the number of the column in the excel format"
  [lz-seq]
  (let [n-columns (-> lz-seq first count)]
    (map zipmap (repeat (map (comp keyword str int->excel-column) (range n-columns))) lz-seq)))

(defn xlsx->column-title-maps
  "Takes the first row of a xlsx and enumerate every row with the column title.
  Use the keyword arg :column-fn to pass a functions that are applied on each of the column names, must be a function of 1 arg. 
  "
  [lz-seq & {:keys [str-keys column-fn] :or {str-keys false column-fn nil}}]
  (map zipmap
       (->> (first lz-seq)
            (#(if column-fn (map column-fn %) %))
            (#(if str-keys % (map keyword %)))
            repeat)
       (rest lz-seq)))

(defn coerce
  "Coerce every row applying a vector of functions"
  [lz-seq fs & {:keys [skip-first-row] :or {skip-first-row false}}]
  (if skip-first-row
    (let [[head & tail] lz-seq]
      (cons head (map (fn [row] (mapv #(%1 %2) fs row)) tail)))
    (map (fn [row] (mapv #(%1 %2) fs row)) lz-seq)))

(defn coerce-map
  "Coerce one map based on a coercion map, extra keys are untouched.
  Extra keys in fs will be ignored."
  [m fs]
  (loop [new-m {}
         ks (into [] (set/intersection (set (keys m)) (set (keys fs))))]
        (if (empty? ks)
          (merge m new-m)
          (recur (assoc new-m (first ks) ((get fs (first ks)) (get m (first ks))))
                 (rest ks)))))

(comment (coerce-map {:d "extra key" :a "1" :b "10" :c "doasdjasodjas"}
                     {:a #(Long/parseLong %) :b excel-date->java-date :c #(count %)}))

(comment
  ;; test with string keys
  (coerce-map {"d" "extra key" "a" "1" "b" "10" "c" "doasdjasodjas"}
              {"a" #(Long/parseLong %) "b" excel-date->java-date "c" #(count %)}))

(comment (coerce-map {:d "extra key" :a "1" :b "10" :c "doasdjasodjas"}
                     {:a #(Long/parseLong %) :b excel-date->java-date :c #(count %) :e #(Double/parseDouble %)}))

(defn coerce-maps
  "coerce a list of maps based on a coercion map, extra keys will be untouched."
  [lz-seq fs & {:keys [skip-first-row] :or {skip-first-row false}}]
  (if skip-first-row
    (let [[head & tail] lz-seq]
      (cons head (map coerce-map tail (repeat fs))))
    (map coerce-map lz-seq (repeat fs))))

(comment (coerce-maps (repeat 5 {:d "an extra key"
                                 :another-extra-key-with-a-bizare-class-that-we-dont-want-to-mess-with-or-apply-a-coercion (Thread.)
                                 :a "1" :b "10" :c "doasdjasodjas"})
                      {:a #(Long/parseLong %) :b excel-date->java-date :c #(count %)}))

(defn excel-date->unix-timestamp
  "Takes a excel date and convert it to a unix timestamp"
  ^Long
  [^String n-str]
  (let [n (Double/parseDouble n-str)]
    (long (* 86400 (- n 25569)))))

(defn excel-date->java-date
  "Takes a excel date and convert it to a java Date object"
  ^Date
  [^String n-str]
  (Date. ^Long (* 1000 (excel-date->unix-timestamp n-str))))

(defn excel-date->joda-date
  "Takes a excel date and convert it to a joda time DateTime object"
  ^DateTime
  [^String n-str]
  (DateTime. (* 1000 (excel-date->unix-timestamp n-str))))

(defn excel-date->java-localdatetime
  "Takes a excel date and convert it to a java.time.LocalDateTime object"
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
              (list-sheets filename)))
  BufferedInputStream (list-sheets [stream]
                        (let [fname (str (gensym))
                              fext (str (gensym))
                              tmp (File/createTempFile fname fext)]
                          (.deleteOnExit tmp)
                          (with-open [in stream out (java.io.FileOutputStream. tmp)]
                            (io/copy in out))
                          (list-sheets (.getAbsolutePath tmp)))))
