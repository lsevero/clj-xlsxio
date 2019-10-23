(ns clj-xlsxio.core
  (:require [clojure.java.io :as io])
  (:import [com.sun.jna NativeLibrary Pointer Memory NativeLong Platform]
           [java.io FileNotFoundException]))

(def libxlsxio-read (com.sun.jna.NativeLibrary/getInstance "xlsxio_read"))
;(def libxlsxio-write (com.sun.jna.NativeLibrary/getInstance "xlsxio_write"))

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
  [sheet]
  (if-not (zero? (sheet-next-row sheet))
    (loop
      [res []]
      (if-let [cell-value (sheet-next-cell sheet)]
        (recur (conj res cell-value))
        res))
    nil))

;(defprotocol ReadXlsx
  ;(read-xlsx [this]))

;(extend-protocol ReadXlsx
  ;com.sun.jna.Pointer
  ;(read-xlsx [this]
    ;(if-let [first-row (read-row this)]
      ;(lazy-seq (cons first-row (read-xlsx this)))
      ;nil))
  ;String
  ;(read-xlsx [filename]
    ;(let [xlsx (open filename)
          ;sheet (sheet-open xlsx)
          ;res (read-xlsx sheet)]
      ;(sheet-close sheet)
      ;(close xlsx)
      ;res)))

(defn read-xlsx
  [sheet]
  (if-let [first-row (read-row sheet)]
    (lazy-seq (cons first-row (read-xlsx sheet)))
    nil))
