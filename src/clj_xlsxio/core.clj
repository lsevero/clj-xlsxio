(ns clj-xlsxio.core
   (:import [com.sun.jna NativeLibrary Pointer Memory NativeLong Platform]))

(def libxlsxio-read (com.sun.jna.NativeLibrary/getInstance "xlsxio_read"))
(def libxlsxio-write (com.sun.jna.NativeLibrary/getInstance "xlsxio_write"))

(def SKIP_NONE 0)
(def SKIP_EMPTY_ROWS 0x01)
(def SKIP_EMPTY_CELLS 0x02)
(def SKIP_ALL_EMPTY (bit-or SKIP_EMPTY_ROWS SKIP_EMPTY_CELLS))
(def SKIP_EXTRA_CELLS 0x04)

(defn open
  ^Pointer
  [^String filename]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_open") Pointer (to-array [filename])))

(defn sheet-open
  ^Pointer
  [^Pointer xlsx ^String sheetname ^Long opts]
  (.invoke (.getFunction libxlsxio-read "xlsxioread_sheet_open") Pointer (to-array [xlsx sheetname opts])))

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
