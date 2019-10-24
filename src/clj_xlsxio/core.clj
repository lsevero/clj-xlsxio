(ns clj-xlsxio.core
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer Memory NativeLong Platform]
           [java.io FileNotFoundException]))

(let [dl (NativeLibrary/getInstance "dl")
      RTLD_LAZY	0x00001
      RTLD_NOW	0x00002
      RTLD_BINDING_MASK   0x3
      RTLD_NOLOAD	0x00004
      RTLD_DEEPBIND	0x00008
      RTLD_GLOBAL	0x00100]
  (defn dlopen
    ^Pointer
    [^String lib]
    (.invoke (.getFunction dl "dlopen") Void (to-array [lib (bit-or RTLD_LAZY RTLD_GLOBAL)]))))



(do
  (def z (NativeLibrary/getInstance "z"))
  (println z)
  (def expat (NativeLibrary/getInstance "expat"))
  (println expat)
  (def minizip (NativeLibrary/getInstance "minizip"))
  (println minizip)
  (def libxlsxio-read (NativeLibrary/getInstance "xlsxio_read"))
  ;(def libxlsxio-write (NativeLibrary/getInstance "xlsxio_write"))
)

(def ^:const skip-none 0)
(def ^:const skip-empty-rows 0x01)
(def ^:const skip-empty-cells 0x02)
(def ^:const skip-all-empty (bit-or skip-empty-rows skip-empty-cells))
(def ^:const skip-extra-cells 0x04)

(defn open
  ^Pointer
  [^String filename]
  (if (.exists (io/file filename))
    (.invoke (.getFunction libxlsxio-read "xlsxioread_open") Pointer (to-array [filename]))
    (throw (FileNotFoundException. (str "File " filename " does not exists.")))))

(defn sheet-open
  ^Pointer
  ([^Pointer xlsx]
   (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_open") Pointer (to-array [xlsx nil skip-none]))) 
  ([^Pointer xlsx ^Long opts]
   (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_open") Pointer (to-array [xlsx nil opts]))) 
  ([^Pointer xlsx ^String sheetname ^Long opts]
   (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_open") Pointer (to-array [xlsx sheetname opts]))))

(defn sheet-next-row
  ^Long
  [^Pointer sheet]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_next_row") Long (to-array [sheet])))

(defn sheet-next-cell
  ^String
  [^Pointer sheet]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_next_cell") String (to-array [sheet])))

(defn sheet-close
  ^Void
  [^Pointer sheet]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_close") Void (to-array [sheet])))

(defn close
  ^Void
  [^Pointer xlsx]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_close") Void (to-array [xlsx])))

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
