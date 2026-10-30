
// isadek-print: printing and PDF for the pages (added by scripts/add-platform.js).
@objc(IsadekPrintPlugin)
public class IsadekPrintPlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "IsadekPrintPlugin"
    public let jsName = "IsadekPrint"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "print", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "pdf", returnType: CAPPluginReturnPromise)
    ]

    @objc func print(_ call: CAPPluginCall) {
        DispatchQueue.main.async {
            guard let web = self.bridge?.webView else { return call.reject("No page to print") }
            let info = UIPrintInfo(dictionary: nil)
            info.jobName = call.getString("name") ?? "Isadek Office"
            info.outputType = .general
            info.orientation = (call.getBool("landscape") ?? false) ? .landscape : .portrait
            let controller = UIPrintInteractionController.shared
            controller.printInfo = info
            controller.printFormatter = web.viewPrintFormatter()
            controller.present(animated: true) { _, _, _ in call.resolve() }
        }
    }

    /// The page as a paginated PDF (A4), using the page's print styles.
    @objc func pdf(_ call: CAPPluginCall) {
        DispatchQueue.main.async {
            guard let web = self.bridge?.webView else { return call.reject("No page to export") }
            let landscape = call.getBool("landscape") ?? false
            let paper = landscape ? CGRect(x: 0, y: 0, width: 842, height: 595) : CGRect(x: 0, y: 0, width: 595, height: 842)
            let renderer = UIPrintPageRenderer()
            renderer.addPrintFormatter(web.viewPrintFormatter(), startingAtPageAt: 0)
            renderer.setValue(NSValue(cgRect: paper), forKey: "paperRect")
            renderer.setValue(NSValue(cgRect: paper.insetBy(dx: 36, dy: 36)), forKey: "printableRect")
            let data = NSMutableData()
            UIGraphicsBeginPDFContextToData(data, paper, nil)
            renderer.prepare(forDrawingPages: NSRange(location: 0, length: renderer.numberOfPages))
            for i in 0..<renderer.numberOfPages {
                UIGraphicsBeginPDFPage()
                renderer.drawPage(at: i, in: UIGraphicsGetPDFContextBounds())
            }
            UIGraphicsEndPDFContext()
            call.resolve(["data": data.base64EncodedString()])
        }
    }
}

class IsadekViewController: CAPBridgeViewController {
    override func capacitorDidLoad() {
        bridge?.registerPluginInstance(IsadekPrintPlugin())
    }
}
