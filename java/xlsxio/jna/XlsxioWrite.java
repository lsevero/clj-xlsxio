package  xlsxio.jna;

import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.Platform;
import com.sun.jna.NativeLong;

public class XlsxioWrite {
    public static native Pointer xlsxiowrite_open(String filename, String sheetname);
    public static native int xlsxiowrite_close(Pointer handle);
    public static native void xlsxiowrite_add_column(Pointer handle, String name, int width);
    public static native void xlsxiowrite_next_row(Pointer handle);
    public static native void xlsxiowrite_add_cell_string(Pointer handle, String value);
    public static native void xlsxiowrite_add_cell_int(Pointer handle, long value);
    public static native void xlsxiowrite_add_cell_float(Pointer handle, double value);
    public static native void xlsxiowrite_add_cell_datetime(Pointer handle, NativeLong value);

    static {
        if(Platform.isLinux())
            Native.register("xlsxio_write");
        else
            throw new RuntimeException("Unrecognized OS. Supported OS are linux-arm64, linux-x86-64");
    }
}
