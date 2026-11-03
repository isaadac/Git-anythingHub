package __PACKAGE__;

import android.content.Context;
import android.os.Bundle;
import android.os.CancellationSignal;
import android.os.ParcelFileDescriptor;
import android.print.PageRange;
import android.print.PrintAttributes;
import android.print.PrintDocumentAdapter;
import android.print.PrintManager;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

/**
 * Printing and PDF for the pages: Android's print dialog, which also offers "Save as PDF".
 * The call finishes when the dialog is done, so the page can put its screen layout back.
 */
@CapacitorPlugin(name = "IsadekPrint")
public class PrintPlugin extends Plugin {
    @PluginMethod
    public void print(PluginCall call) {
        String name = call.getString("name", "Isadek Office");
        boolean landscape = Boolean.TRUE.equals(call.getBoolean("landscape", false));
        getActivity().runOnUiThread(() -> {
            PrintManager manager = (PrintManager) getActivity().getSystemService(Context.PRINT_SERVICE);
            PrintDocumentAdapter page = getBridge().getWebView().createPrintDocumentAdapter(name);
            PrintDocumentAdapter adapter = new PrintDocumentAdapter() {
                @Override public void onStart() { page.onStart(); }
                @Override public void onLayout(PrintAttributes before, PrintAttributes after, CancellationSignal cancel, LayoutResultCallback cb, Bundle extras) { page.onLayout(before, after, cancel, cb, extras); }
                @Override public void onWrite(PageRange[] pages, ParcelFileDescriptor out, CancellationSignal cancel, WriteResultCallback cb) { page.onWrite(pages, out, cancel, cb); }
                @Override public void onFinish() { page.onFinish(); call.resolve(); }
            };
            PrintAttributes.Builder attrs = new PrintAttributes.Builder();
            if (landscape) attrs.setMediaSize(PrintAttributes.MediaSize.UNKNOWN_LANDSCAPE);
            manager.print(name, adapter, attrs.build());
        });
    }

    @PluginMethod
    public void pdf(PluginCall call) {
        // Android makes PDFs through the print dialog's "Save as PDF".
        print(call);
    }
}
