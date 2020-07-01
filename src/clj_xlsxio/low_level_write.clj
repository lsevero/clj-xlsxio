(ns clj-xlsxio.low-level-write
  (:refer-clojure :exclude [name])
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer NativeLong]
           [java.util Date]
           [org.joda.time DateTime]
           [java.time LocalDateTime ZoneOffset]
           [clojure.lang Ratio]))

(try
  (do
    (def z (NativeLibrary/getInstance "libz.so.1"))
    (def minizip (NativeLibrary/getInstance "libminizip.so.1"))
    (import [xlsxio.jna XlsxioWrite])
    (let [^NativeLibrary c (NativeLibrary/getInstance "c")
          setlocale (.getFunction c "setlocale")]
      (mapv #(.invoke setlocale String (to-array [% "C"])) (range 30))))
  (catch Exception e 
    (do
      (println "============================================================================
               We've had a problem loading the native libraries.
               This library has two dependencies:
               minizip and libz (which is used by minizip)
               All of them are bundled in this jar library,
               however if you are having issues loading the shared objects consider
               installing them on your system.
               It is important to notice that all of these bundled native libraries were
               compiled with GNU libc standard library. If you are on a system based on musl
               (like Alpine Linux ) or another standard library you WILL
               need to install those 2 dependencies on your system.
               ============================================================================")
      (pr e))))

(defn open
  ^Pointer
  ([^String filename ^String sheetname]
   (let [res (XlsxioWrite/xlsxiowrite_open filename sheetname)]
     (if (= res Pointer/NULL)
       (throw (RuntimeException. "Error on xlsxiowrite_open, returned NULL."))
       res)))
  ([^String filename]
   (open filename nil)))

(defn close
  ^long
  [^Pointer handle]
  (let [res (long (XlsxioWrite/xlsxiowrite_close handle))]
    (if-not (= res 0)
      (throw (RuntimeException. (str "Error on xlsxiowrite_close, returned " res)))
      res)))

(defn add-column
  ([^Pointer handle ^String name width]
   (XlsxioWrite/xlsxiowrite_add_column handle name width))
  ([^Pointer handle ^String name]
   (XlsxioWrite/xlsxiowrite_add_column handle name 0)))

(defn next-row
  [^Pointer handle]
  (XlsxioWrite/xlsxiowrite_next_row handle))

(defn add-cell-string
  [^Pointer handle ^String value]
  (XlsxioWrite/xlsxiowrite_add_cell_string handle value))

(defn add-cell-int
  [^Pointer handle value]
  (XlsxioWrite/xlsxiowrite_add_cell_int handle value))

(defn add-cell-float
  [^Pointer handle value]
  (XlsxioWrite/xlsxiowrite_add_cell_float handle value))

(defn add-cell-double
  [^Pointer handle value]
  (add-cell-float handle value))

(defn add-cell-datetime
  [^Pointer handle value]
  (XlsxioWrite/xlsxiowrite_add_cell_datetime handle value))

(defprotocol AddCell
  (add-cell-generic [generic handle]))

(extend-protocol AddCell
  String        (add-cell-generic [value handle]
                  (add-cell-string handle value))
  Double        (add-cell-generic [value handle]
                  (add-cell-float handle value))
  Ratio         (add-cell-generic [value handle]
                  (add-cell-float handle (double value)))
  Long          (add-cell-generic [value handle]
                  (add-cell-int handle (long value)))
  Date          (add-cell-generic [value handle]
                  (add-cell-datetime handle (NativeLong. (long (/ (.getTime value) 1000)))))
  DateTime      (add-cell-generic [value handle]
                  (add-cell-datetime handle (NativeLong. (long (/ (.getMillis value) 1000)))))
  LocalDateTime (add-cell-generic [value handle]
                  (add-cell-datetime handle (NativeLong. (long (.toEpochSecond value (ZoneOffset/ofHours 0)))))))

(defn add-cell
  [^Pointer handle value]
  (add-cell-generic value handle))
