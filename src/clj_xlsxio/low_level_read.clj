(ns clj-xlsxio.low-level-read
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer]
           [java.io FileNotFoundException]))

(try
  (do
    (def z (NativeLibrary/getInstance "libz.so.1"))
    (def expat (NativeLibrary/getInstance "libexpat.so.1"))
    (def minizip (NativeLibrary/getInstance "minizip"))
    (def ^NativeLibrary libxlsxio-read (NativeLibrary/getInstance "xlsxio_read")))
  (catch Exception e 
    (do
      (println "============================================================================
               We've had a problem loading the native libraries.
               This library has three dependencies:
               expat, minizip and libz (which is used by minizip)
               All of them are bundled in this jar library,
               however if you are having issues loading the shared objects consider
               installing them on your system.

               It is important to notice that all of these bundled native libraries were
               compiled with GNU libc standard library. If you are on a system based on musl
               (like Alpine Linux ) or another standard library you WILL
               need to install those 3 dependencies on your system.
               ============================================================================")
      (pr e))))

(defn open
  ^Pointer
  [^String filename]
  (if (.exists (io/file filename))
    (.invoke (.getFunction libxlsxio-read "xlsxioread_open") Pointer (to-array [filename]))
    (throw (FileNotFoundException. (str "File " filename " does not exists.")))))

(defn sheet-open
  ^Pointer
  ([^Pointer xlsx]
   (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_open") Pointer (to-array [xlsx nil 0]))) 
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

(defn sheetlist-open
  ^Pointer
  [^Pointer xlsx]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheetlist_open") Pointer (to-array [xlsx])))

(defn sheetlist-next
  ^String
  [^Pointer sheetlist]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheetlist_next") String (to-array [sheetlist])))

(defn sheetlist-close
  ^Void
  [^Pointer sheetlist]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheetlist_close") Void (to-array [sheetlist])))
