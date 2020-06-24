(ns clj-xlsxio.low-level-read
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer]
           [java.io FileNotFoundException File]
           [java.util.zip ZipFile]
           [java.nio.file Files]
           [xlsxio.jna XlsxioRead]))

(defn open
  ^Pointer
  [^String filename]
  (if (.exists (io/file filename))
    (do
      (try
        (ZipFile. filename)
        (catch Exception e
          (throw (IllegalArgumentException. (str "File " filename " is not a valid xlsx file.")))))
      (when-not (= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                   (Files/probeContentType (.toPath ^File (io/file filename))))
        (throw (IllegalArgumentException. (str "File " filename " is not a valid xlsx file."))))
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
