(ns clj-xlsxio.low-level-read
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer]
           [java.io FileNotFoundException File]
           [java.util.zip ZipFile]
           [java.nio.file Files]))

(try
  (do
    (def z (NativeLibrary/getInstance "libz.so.1"))
    (def expat (NativeLibrary/getInstance "libexpat.so.1"))
    (def minizip (NativeLibrary/getInstance "minizip"))
    (import [xlsxio.jna XlsxioRead]))
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
  [^String filename & {:keys [check-mime-type] :or {check-mime-type false}}]
  (if (.exists (io/file filename))
    (do
      (try
        (ZipFile. filename)
        (catch Exception e
          (throw (IllegalArgumentException. (str "File " filename " is not a valid xlsx file.")))))
      (when check-mime-type 
        (when-not (= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     (Files/probeContentType (.toPath ^File (io/file filename))))
          (throw (IllegalArgumentException. (str "File " filename " is not a valid xlsx file.")))))
      (XlsxioRead/xlsxioread_open filename))
    (throw (FileNotFoundException. (str "File " filename " does not exists.")))))

(defn sheet-open
  ^Pointer
  ([^Pointer xlsx]
  (XlsxioRead/xlsxioread_sheet_open xlsx nil 0)) 
  ([^Pointer xlsx opts]
  (XlsxioRead/xlsxioread_sheet_open xlsx nil opts)) 
  ([^Pointer xlsx ^String sheetname opts]
  (XlsxioRead/xlsxioread_sheet_open xlsx sheetname opts)))

(defn sheet-next-row
  ^long
  [^Pointer sheet]
  (long (XlsxioRead/xlsxioread_sheet_next_row sheet)))

(defn sheet-next-cell
  ^String
  [^Pointer sheet]
  (XlsxioRead/xlsxioread_sheet_next_cell sheet))

(defn sheet-close
  [^Pointer sheet]
  (XlsxioRead/xlsxioread_sheet_close sheet))

(defn close
  [^Pointer xlsx]
  (XlsxioRead/xlsxioread_close xlsx))

(defn sheetlist-open
  ^Pointer
  [^Pointer xlsx]
  (XlsxioRead/xlsxioread_sheetlist_open xlsx))

(defn sheetlist-next
  ^String
  [^Pointer sheetlist]
  (XlsxioRead/xlsxioread_sheetlist_next sheetlist))

(defn sheetlist-close
  [^Pointer sheetlist]
  (XlsxioRead/xlsxioread_sheetlist_close sheetlist))
