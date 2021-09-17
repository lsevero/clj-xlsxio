package  xlsxio.jna;

import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.Platform;

public class XlsxioRead {
    public static native Pointer xlsxioread_open(String filename);
    public static native Pointer xlsxioread_sheet_open(Pointer handle, String sheetname, int flags);
    public static native void xlsxioread_sheet_close(Pointer sheethandle);
    public static native int xlsxioread_sheet_next_row(Pointer sheethandle);
    public static native String xlsxioread_sheet_next_cell(Pointer sheethandle);
    public static native void xlsxioread_close(Pointer handle);
    public static native Pointer xlsxioread_sheetlist_open(Pointer handle);
    public static native void xlsxioread_sheetlist_close(Pointer sheetlisthandle);
    public static native String xlsxioread_sheetlist_next(Pointer sheetlisthandle);

    static {
        if(Platform.isLinux() || Platform.isMac())
            Native.register("xlsxio_read");
        else
            throw new RuntimeException("Unrecognized OS. Supported OS are linux-arm64, linux-x86-64, darwin-x86-64");
    }

}
