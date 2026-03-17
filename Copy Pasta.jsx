#target aftereffects
#targetengine "copy_pasta_engine"

(function CopyPastaPanel(thisObj) {
    var SCRIPT_NAME = "Copy Pasta";
    var SCRIPT_VERSION = "1.8";
    var TEMP_FOLDER_NAME = "CopyPastaTemp";
    var IMPORT_FOLDER_NAME = "CopyPasta Imports";
    var WINDOWS_HELPER_EXE_NAME = "copy_pasta_clipboard_helper.exe";
    var WINDOWS_HELPER_SRC_NAME = "copy_pasta_clipboard_helper.cs";
    var THEME = {
        bg: [0.08, 0.1, 0.14, 1.0],
        surface: [0.12, 0.16, 0.21, 1.0],
        accent: [0.2, 0.67, 0.96, 1.0],
        textMain: [0.93, 0.95, 0.98, 1.0],
        textMuted: [0.66, 0.72, 0.79, 1.0],
        statusOk: [0.47, 0.83, 0.98, 1.0],
        statusError: [1.0, 0.45, 0.45, 1.0]
    };

    function isWindows() {
        return $.os && $.os.toLowerCase().indexOf("windows") !== -1;
    }

    function isMac() {
        return $.os && $.os.toLowerCase().indexOf("mac") !== -1;
    }

    function shQuote(str) {
        return "'" + str.replace(/'/g, "'\\''") + "'";
    }

    function psQuote(str) {
        return str.replace(/'/g, "''");
    }

    function appleQuote(str) {
        return str.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    }

    function writeTextFile(fileObj, text) {
        if (!fileObj.open("w")) {
            throw new Error("Cannot write file: " + fileObj.fsName);
        }
        try {
            fileObj.encoding = "UTF-8";
            fileObj.lineFeed = "Unix";
            fileObj.write(text);
        } finally {
            fileObj.close();
        }
    }

    function readTextFile(fileObj) {
        if (!fileObj || !fileObj.exists) return "";
        if (!fileObj.open("r")) return "";
        try {
            fileObj.encoding = "UTF-8";
            return fileObj.read();
        } catch (e) {
            return "";
        } finally {
            fileObj.close();
        }
    }

    function getWindowsPowerShellExe() {
        var sysRoot = $.getenv("SystemRoot");
        if (!sysRoot || sysRoot === "") {
            sysRoot = "C:/Windows";
        }
        return new File(sysRoot + "/System32/WindowsPowerShell/v1.0/powershell.exe");
    }

    function runPowerShellWindows(psCode) {
        var tempScript = null;
        var out = "";
        var cmd = "";

        try {
            tempScript = uniqueFile(getTempFolder(), "copy_pasta_clipboard", "ps1");
            writeTextFile(tempScript, psCode);

            var exeFile = getWindowsPowerShellExe();
            var exePath = exeFile.exists ? exeFile.fsName : "powershell";
            cmd = '"' + exePath + '" -NoProfile -STA -ExecutionPolicy Bypass -File "' + tempScript.fsName + '" 2>&1';
            out = system.callSystem(cmd);
            if (out && out !== "") return out;

            cmd = 'powershell -NoProfile -STA -ExecutionPolicy Bypass -File "' + tempScript.fsName + '" 2>&1';
            out = system.callSystem(cmd);
            if (out && out !== "") return out;

            return "";
        } catch (err) {
            return "ERR:" + err.toString();
        } finally {
            try {
                if (tempScript && tempScript.exists) tempScript.remove();
            } catch (cleanupErr) {}
        }
    }

    function getWindowsHelperFolder() {
        return ensureFolder(new Folder(Folder.userData.fsName + "/CopyPastaHelper"));
    }

    function getWindowsHelperExe() {
        return new File(getWindowsHelperFolder().fsName + "/" + WINDOWS_HELPER_EXE_NAME);
    }

    function getWindowsHelperSource() {
        return new File(getWindowsHelperFolder().fsName + "/" + WINDOWS_HELPER_SRC_NAME);
    }

    function getWindowsCscCandidates() {
        var winDir = $.getenv("WINDIR") || $.getenv("SystemRoot") || "C:/Windows";
        return [
            new File(winDir + "/Microsoft.NET/Framework64/v4.0.30319/csc.exe"),
            new File(winDir + "/Microsoft.NET/Framework/v4.0.30319/csc.exe")
        ];
    }

    function buildWindowsClipboardHelperSource() {
        return [
            "using System;",
            "using System.Drawing;",
            "using System.Drawing.Imaging;",
            "using System.IO;",
            "using System.Net;",
            "using System.Text;",
            "using System.Text.RegularExpressions;",
            "using System.Threading;",
            "using System.Windows.Forms;",
            "",
            "namespace CopyPastaClipboardHelper",
            "{",
            "    internal static class Program",
            "    {",
            "        [STAThread]",
            "        private static int Main(string[] args)",
            "        {",
            "            if (args == null || args.Length < 3)",
            "            {",
            "                return WriteResult(null, \"ERR:Arguments expected: <set|get> <imagePath> <resultPath>\");",
            "            }",
            "",
            "            string mode = (args[0] ?? \"\").Trim().ToLowerInvariant();",
            "            string imagePath = args[1] ?? \"\";",
            "            string resultPath = args[2] ?? \"\";",
            "",
            "            try",
            "            {",
            "                if (mode == \"set\")",
            "                {",
            "                    return SetClipboardImage(imagePath, resultPath);",
            "                }",
            "",
            "                if (mode == \"get\")",
            "                {",
            "                    return GetClipboardImage(imagePath, resultPath);",
            "                }",
            "",
            "                return WriteResult(resultPath, \"ERR:Unknown mode\");",
            "            }",
            "            catch (Exception ex)",
            "            {",
            "                return WriteResult(resultPath, \"ERR:\" + ex.Message);",
            "            }",
            "        }",
            "",
            "        private static int SetClipboardImage(string imagePath, string resultPath)",
            "        {",
            "            if (!File.Exists(imagePath))",
            "            {",
            "                return WriteResult(resultPath, \"ERR:Input image file not found\");",
            "            }",
            "",
            "            Exception last = null;",
            "            for (int i = 0; i < 15; i++)",
            "            {",
            "                try",
            "                {",
            "                    using (Image img = Image.FromFile(imagePath))",
            "                    using (Bitmap bmp = new Bitmap(img))",
            "                    {",
            "                        Clipboard.SetImage(bmp);",
            "                    }",
            "",
            "                    return WriteResult(resultPath, \"OK\");",
            "                }",
            "                catch (Exception ex)",
            "                {",
            "                    last = ex;",
            "                    Thread.Sleep(120);",
            "                }",
            "            }",
            "",
            "            return WriteResult(resultPath, \"ERR:\" + (last != null ? last.Message : \"Clipboard write failed\"));",
            "        }",
            "",
            "        private static int GetClipboardImage(string outputPath, string resultPath)",
            "        {",
            "            Exception last = null;",
            "            for (int i = 0; i < 15; i++)",
            "            {",
            "                try",
            "                {",
            "                    if (TrySaveClipboardImage(outputPath))",
            "                    {",
            "                        return WriteResult(resultPath, \"OK\");",
            "                    }",
            "                }",
            "                catch (Exception ex)",
            "                {",
            "                    last = ex;",
            "                }",
            "",
            "                Thread.Sleep(120);",
            "            }",
            "",
            "            string formats = GetClipboardFormats();",
            "            if (!string.IsNullOrEmpty(formats))",
            "            {",
            "                return WriteResult(resultPath, \"NO_IMAGE:\" + formats);",
            "            }",
            "",
            "            if (last != null)",
            "            {",
            "                return WriteResult(resultPath, \"ERR:\" + last.Message);",
            "            }",
            "",
            "            return WriteResult(resultPath, \"NO_IMAGE\");",
            "        }",
            "",
            "        private static bool TrySaveClipboardImage(string outputPath)",
            "        {",
            "            if (Clipboard.ContainsImage())",
            "            {",
            "                using (Image img = Clipboard.GetImage())",
            "                {",
            "                    if (img != null)",
            "                    {",
            "                        using (Bitmap bmp = new Bitmap(img))",
            "                        {",
            "                            bmp.Save(outputPath, ImageFormat.Png);",
            "                        }",
            "                        return true;",
            "                    }",
            "                }",
            "            }",
            "",
            "            IDataObject data = Clipboard.GetDataObject();",
            "            if (data == null)",
            "            {",
            "                return false;",
            "            }",
            "",
            "            if (data.GetDataPresent(DataFormats.Bitmap))",
            "            {",
            "                object b = data.GetData(DataFormats.Bitmap);",
            "                if (b is Image)",
            "                {",
            "                    using (Bitmap bmp = new Bitmap((Image)b))",
            "                    {",
            "                        bmp.Save(outputPath, ImageFormat.Png);",
            "                    }",
            "                    return true;",
            "                }",
            "            }",
            "",
            "            string[] fmts = data.GetFormats();",
            "            if (fmts != null)",
            "            {",
            "                foreach (string fmt in fmts)",
            "                {",
            "                    if (fmt == null) continue;",
            "                    if (fmt.IndexOf(\"PNG\", StringComparison.OrdinalIgnoreCase) < 0) continue;",
            "",
            "                    object png = data.GetData(fmt);",
            "                    if (png is MemoryStream)",
            "                    {",
            "                        using (MemoryStream ms = (MemoryStream)png)",
            "                        {",
            "                            if (SaveStreamToFile(ms, outputPath)) return true;",
            "                        }",
            "                    }",
            "                    else if (png is byte[])",
            "                    {",
            "                        File.WriteAllBytes(outputPath, (byte[])png);",
            "                        return true;",
            "                    }",
            "                    else if (png is Stream)",
            "                    {",
            "                        using (Stream s = (Stream)png)",
            "                        {",
            "                            if (SaveStreamToFile(s, outputPath)) return true;",
            "                        }",
            "                    }",
            "                }",
            "            }",
            "",
            "            if (data.GetDataPresent(DataFormats.FileDrop))",
            "            {",
            "                string[] files = data.GetData(DataFormats.FileDrop) as string[];",
            "                if (files != null)",
            "                {",
            "                    foreach (string f in files)",
            "                    {",
            "                        if (string.IsNullOrEmpty(f) || !File.Exists(f)) continue;",
            "                        try",
            "                        {",
            "                            using (Image img2 = Image.FromFile(f))",
            "                            using (Bitmap bmp2 = new Bitmap(img2))",
            "                            {",
            "                                bmp2.Save(outputPath, ImageFormat.Png);",
            "                            }",
            "                            return true;",
            "                        }",
            "                        catch { }",
            "                    }",
            "                }",
            "            }",
            "",
            "            string url = ExtractImageUrlFromData(data);",
            "            if (!string.IsNullOrEmpty(url))",
            "            {",
            "                if (TryDownloadImageAsPng(url, outputPath))",
            "                {",
            "                    return true;",
            "                }",
            "            }",
            "",
            "            return false;",
            "        }",
            "",
            "        private static bool SaveStreamToFile(Stream stream, string path)",
            "        {",
            "            if (stream == null) return false;",
            "            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read))",
            "            {",
            "                try { if (stream.CanSeek) stream.Position = 0; } catch { }",
            "                stream.CopyTo(fs);",
            "            }",
            "            return true;",
            "        }",
            "",
            "        private static string ExtractImageUrlFromData(IDataObject data)",
            "        {",
            "            string html = null;",
            "            string text = null;",
            "",
            "            try { if (data.GetDataPresent(\"HTML Format\")) html = data.GetData(\"HTML Format\") as string; } catch { }",
            "            try { if (string.IsNullOrEmpty(text) && data.GetDataPresent(DataFormats.UnicodeText)) text = data.GetData(DataFormats.UnicodeText) as string; } catch { }",
            "            try { if (string.IsNullOrEmpty(text) && data.GetDataPresent(DataFormats.Text)) text = data.GetData(DataFormats.Text) as string; } catch { }",
            "",
            "            string fromHtml = ExtractUrl(html);",
            "            if (!string.IsNullOrEmpty(fromHtml)) return fromHtml;",
            "",
            "            string fromText = ExtractUrl(text);",
            "            if (!string.IsNullOrEmpty(fromText)) return fromText;",
            "",
            "            try",
            "            {",
            "                string clipText = Clipboard.GetText(TextDataFormat.UnicodeText);",
            "                string fromClip = ExtractUrl(clipText);",
            "                if (!string.IsNullOrEmpty(fromClip)) return fromClip;",
            "            }",
            "            catch { }",
            "",
            "            return null;",
            "        }",
            "",
            "        private static string ExtractUrl(string text)",
            "        {",
            "            if (string.IsNullOrEmpty(text)) return null;",
            "            Match m = Regex.Match(text, @\"https?://[^\\s\"\"'<>]+\", RegexOptions.IgnoreCase);",
            "            if (!m.Success) return null;",
            "            return m.Value.Replace(\"&amp;\", \"&\");",
            "        }",
            "",
            "        private static bool TryDownloadImageAsPng(string url, string outputPath)",
            "        {",
            "            try",
            "            {",
            "                try",
            "                {",
            "                    ServicePointManager.SecurityProtocol =",
            "                        SecurityProtocolType.Tls12 |",
            "                        SecurityProtocolType.Tls11 |",
            "                        SecurityProtocolType.Tls;",
            "                }",
            "                catch { }",
            "",
            "                using (WebClient wc = new WebClient())",
            "                {",
            "                    wc.Headers.Add(\"User-Agent\", \"CopyPastaHelper/1.0\");",
            "                    byte[] bytes = wc.DownloadData(url);",
            "                    using (MemoryStream ms = new MemoryStream(bytes))",
            "                    using (Image img = Image.FromStream(ms))",
            "                    using (Bitmap bmp = new Bitmap(img))",
            "                    {",
            "                        bmp.Save(outputPath, ImageFormat.Png);",
            "                        return true;",
            "                    }",
            "                }",
            "            }",
            "            catch",
            "            {",
            "                return false;",
            "            }",
            "        }",
            "",
            "        private static string GetClipboardFormats()",
            "        {",
            "            try",
            "            {",
            "                IDataObject data = Clipboard.GetDataObject();",
            "                if (data == null) return null;",
            "                string[] formats = data.GetFormats();",
            "                if (formats == null || formats.Length == 0) return null;",
            "                return string.Join(\", \", formats);",
            "            }",
            "            catch",
            "            {",
            "                return null;",
            "            }",
            "        }",
            "",
            "        private static int WriteResult(string resultPath, string text)",
            "        {",
            "            string msg = string.IsNullOrEmpty(text) ? \"ERR:Unknown helper error\" : text;",
            "            try",
            "            {",
            "                if (!string.IsNullOrEmpty(resultPath))",
            "                {",
            "                    File.WriteAllText(resultPath, msg, Encoding.UTF8);",
            "                }",
            "            }",
            "            catch { }",
            "",
            "            Console.WriteLine(msg);",
            "            return msg.StartsWith(\"OK\", StringComparison.OrdinalIgnoreCase) ? 0 : 1;",
            "        }",
            "    }",
            "}"
        ].join("\r\n");
    }

    function ensureWindowsClipboardHelper() {
        var exe = getWindowsHelperExe();
        if (exe.exists) {
            return { ok: true, file: exe, message: "" };
        }

        var src = getWindowsHelperSource();
        try {
            writeTextFile(src, buildWindowsClipboardHelperSource());
        } catch (writeErr) {
            return { ok: false, file: null, message: "Cannot write helper source: " + writeErr.toString() };
        }

        var candidates = getWindowsCscCandidates();
        var csc = null;
        var i;
        for (i = 0; i < candidates.length; i++) {
            if (candidates[i].exists) {
                csc = candidates[i];
                break;
            }
        }

        if (!csc) {
            return { ok: false, file: null, message: "C# compiler not found (csc.exe)." };
        }

        var compileCmd =
            '"' + csc.fsName + '"' +
            ' /nologo /target:exe' +
            ' /out:"' + exe.fsName + '"' +
            ' /reference:System.Windows.Forms.dll' +
            ' /reference:System.Drawing.dll' +
            ' "' + src.fsName + '"';

        var compileOut = system.callSystem(compileCmd);
        if (exe.exists) {
            return { ok: true, file: exe, message: "" };
        }

        return {
            ok: false,
            file: null,
            message: "Clipboard helper compile failed." + (compileOut ? (" " + compileOut.replace(/\r?\n/g, " ")) : "")
        };
    }

    function runWindowsClipboardHelper(mode, imagePath) {
        var helperInfo = ensureWindowsClipboardHelper();
        if (!helperInfo.ok || !helperInfo.file) {
            return { ok: false, noImage: false, unavailable: true, message: helperInfo.message || "Clipboard helper unavailable." };
        }

        var resultFile = uniqueFile(getTempFolder(), "copy_pasta_helper_result", "txt");
        var cmd =
            '"' + helperInfo.file.fsName + '"' +
            ' "' + mode + '"' +
            ' "' + imagePath + '"' +
            ' "' + resultFile.fsName + '"';

        var out = system.callSystem(cmd);
        var status = readTextFile(resultFile);

        try { if (resultFile.exists) resultFile.remove(); } catch (cleanupErr2) {}

        if (isBlankText(status)) {
            status = out;
        }

        if (!isBlankText(status)) {
            status = status.replace(/^\uFEFF/, "");
            status = status.replace(/\r?\n/g, " ");
            status = status.replace(/^\s+|\s+$/g, "");
        }

        if (status && status.indexOf("OK") === 0) {
            return { ok: true, noImage: false, unavailable: false, message: "" };
        }

        if (status && status.indexOf("NO_IMAGE:") === 0) {
            return {
                ok: false,
                noImage: true,
                unavailable: false,
                message: "Clipboard does not contain a readable image format. Available formats: " + status.substring(9)
            };
        }

        if (status && status.indexOf("NO_IMAGE") === 0) {
            return {
                ok: false,
                noImage: true,
                unavailable: false,
                message: "Clipboard does not contain an image."
            };
        }

        if (isBlankText(status)) {
            return {
                ok: false,
                noImage: false,
                unavailable: true,
                message: "Clipboard helper returned no status."
            };
        }

        return {
            ok: false,
            noImage: false,
            unavailable: false,
            message: status
        };
    }

    function ensureFolder(folder) {
        if (!folder.exists) {
            if (!folder.create()) {
                throw new Error("Cannot create folder: " + folder.fsName);
            }
        }
        return folder;
    }

    function getTempFolder() {
        return ensureFolder(new Folder(Folder.temp.fsName + "/" + TEMP_FOLDER_NAME));
    }

    function getImportFolder() {
        var baseFolder;
        if (app.project && app.project.file) {
            baseFolder = app.project.file.parent;
        } else {
            baseFolder = Folder.myDocuments;
        }
        return ensureFolder(new Folder(baseFolder.fsName + "/" + IMPORT_FOLDER_NAME));
    }

    function uniqueFile(folder, prefix, extension) {
        var stamp = (new Date()).getTime();
        var rand = Math.floor(Math.random() * 100000);
        return new File(folder.fsName + "/" + prefix + "_" + stamp + "_" + rand + "." + extension);
    }

    function getActiveComp() {
        if (!app.project) return null;
        var item = app.project.activeItem;
        return (item && item instanceof CompItem) ? item : null;
    }

    function hasScriptFileNetworkAccess() {
        try {
            return app.preferences.getPrefAsLong("Main Pref Section", "Pref_SCRIPTING_FILE_NETWORK_SECURITY") === 1;
        } catch (e) {
            return true;
        }
    }

    function isBlankText(str) {
        return !str || str.replace(/\s+/g, "") === "";
    }

    function getUIFont(style, size) {
        var font = null;

        try { font = ScriptUI.newFont("Poppins", style, size); } catch (e1) {}
        if (!font) {
            try { font = ScriptUI.newFont("Segoe UI", style, size); } catch (e2) {}
        }
        if (!font) {
            try { font = ScriptUI.newFont("Arial", style, size); } catch (e3) {}
        }

        return font;
    }

    function setTextColor(control, rgba) {
        try {
            var g = control.graphics;
            g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, rgba, 1);
        } catch (e) {}
    }

    function setBackgroundColor(control, rgba) {
        try {
            var g = control.graphics;
            g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, rgba);
        } catch (e) {}
    }

    function isShapeLayer(layer) {
        return !!layer && layer.matchName === "ADBE Vector Layer";
    }

    function isImageLayer(layer) {
        if (!layer || !(layer instanceof AVLayer)) return false;

        var src = layer.source;
        if (!src || !(src instanceof FootageItem)) return false;
        if (!src.mainSource || !(src.mainSource instanceof FileSource)) return false;

        return src.mainSource.isStill === true;
    }

    function getFirstCopyableLayer(comp) {
        if (!comp || !comp.selectedLayers || comp.selectedLayers.length === 0) {
            return null;
        }

        var i;
        for (i = 0; i < comp.selectedLayers.length; i++) {
            if (isShapeLayer(comp.selectedLayers[i]) || isImageLayer(comp.selectedLayers[i])) {
                return comp.selectedLayers[i];
            }
        }

        return "INVALID_TYPE";
    }

    function shiftLayerToTimeZero(layer, sourceTime) {
        try { layer.startTime -= sourceTime; } catch (e1) {}
        try { layer.inPoint -= sourceTime; } catch (e2) {}
        try { layer.outPoint -= sourceTime; } catch (e3) {}
    }

    function saveCompFrameToPng(comp, time, outputFile) {
        if (typeof comp.saveFrameToPng !== "function") {
            throw new Error("This After Effects version does not support saveFrameToPng.");
        }

        comp.saveFrameToPng(time, outputFile);

        if (!outputFile.exists) {
            throw new Error("Failed to render the selected layer.");
        }
    }

    function gatherCaptureDependencies(layer, keepMap) {
        if (!layer || keepMap[layer.index]) return;
        keepMap[layer.index] = true;

        try {
            if (layer.parent) {
                gatherCaptureDependencies(layer.parent, keepMap);
            }
        } catch (e1) {}

        try {
            if (layer.trackMatteType !== TrackMatteType.NO_TRACK_MATTE && layer.index > 1) {
                gatherCaptureDependencies(layer.containingComp.layer(layer.index - 1), keepMap);
            }
        } catch (e2) {}
    }

    function captureLayerViaDuplicateComp(comp, layer, outputFile) {
        var captureComp = null;

        try {
            captureComp = comp.duplicate();
            captureComp.name = "__CopyPasta_Capture__";

            var targetLayer = captureComp.layer(layer.index);
            if (!targetLayer) {
                throw new Error("Could not isolate selected layer in capture composition.");
            }

            var keep = {};
            gatherCaptureDependencies(targetLayer, keep);

            var i;
            for (i = 1; i <= captureComp.numLayers; i++) {
                if (!keep[i]) {
                    try { captureComp.layer(i).enabled = false; } catch (e3) {}
                }
            }

            for (i = 1; i <= captureComp.numLayers; i++) {
                try { captureComp.layer(i).solo = false; } catch (e4) {}
            }
            for (i = 1; i <= captureComp.numLayers; i++) {
                if (keep[i]) {
                    try { captureComp.layer(i).solo = true; } catch (e5) {}
                }
            }

            captureComp.time = comp.time;
            saveCompFrameToPng(captureComp, captureComp.time, outputFile);
        } finally {
            if (captureComp) {
                try { captureComp.remove(); } catch (cleanupErr1) {}
            }
        }
    }

    function captureLayerViaCopyToComp(comp, layer, outputFile) {
        var tempComp = null;

        try {
            var fps = comp.frameRate > 0 ? comp.frameRate : 25;
            var duration = 1 / fps;

            tempComp = app.project.items.addComp(
                "__CopyPasta_Capture__",
                comp.width,
                comp.height,
                comp.pixelAspect,
                duration,
                fps
            );

            layer.copyToComp(tempComp);
            var copiedLayer = tempComp.layer(1);
            shiftLayerToTimeZero(copiedLayer, comp.time);
            tempComp.time = 0;

            saveCompFrameToPng(tempComp, 0, outputFile);
        } finally {
            if (tempComp) {
                try { tempComp.remove(); } catch (cleanupErr2) {}
            }
        }
    }

    function captureLayerToPng(comp, layer, outputFile) {
        var primaryError = null;

        try {
            captureLayerViaDuplicateComp(comp, layer, outputFile);
            return;
        } catch (err1) {
            primaryError = err1;
        }

        try {
            if (outputFile.exists) outputFile.remove();
        } catch (cleanupErr3) {}

        try {
            captureLayerViaCopyToComp(comp, layer, outputFile);
            return;
        } catch (err2) {
            throw new Error("Primary capture failed: " + primaryError.toString() + "\nFallback capture failed: " + err2.toString());
        }
    }

    function setClipboardImageWindows(filePath) {
        var helperResult = runWindowsClipboardHelper("set", filePath);
        if (helperResult.ok) {
            return { ok: true, message: "" };
        }

        var helperMessage = helperResult.message || "";

        var p = psQuote(filePath);
        var resultFile = uniqueFile(getTempFolder(), "copy_pasta_result_set", "txt");
        var r = psQuote(resultFile.fsName);
        var ps =
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8\n" +
            "$OutputEncoding=[System.Text.Encoding]::UTF8\n" +
            "$ErrorActionPreference='Stop'\n" +
            "$path='" + p + "'\n" +
            "$resultPath='" + r + "'\n" +
            "function Write-Result([string]$s){ try{ Set-Content -LiteralPath $resultPath -Value $s -Encoding UTF8 -Force }catch{}; Write-Output $s }\n" +
            "if(-not (Test-Path -LiteralPath $path)){ Write-Result 'ERR:Rendered image file not found'; exit 0 }\n" +
            "$ok=$false\n" +
            "$err=''\n" +
            "try{\n" +
            "  Add-Type -AssemblyName System.Windows.Forms | Out-Null\n" +
            "  Add-Type -AssemblyName System.Drawing | Out-Null\n" +
            "  for($i=0;$i -lt 14 -and -not $ok;$i++){\n" +
            "    $img=$null\n" +
            "    try{\n" +
            "      $img=[System.Drawing.Image]::FromFile($path)\n" +
            "      [System.Windows.Forms.Clipboard]::SetImage($img)\n" +
            "      $ok=$true\n" +
            "    }catch{\n" +
            "      $err=$_.Exception.Message\n" +
            "      Start-Sleep -Milliseconds 120\n" +
            "    }finally{\n" +
            "      if($img -ne $null){ try{$img.Dispose()}catch{} }\n" +
            "    }\n" +
            "  }\n" +
            "}catch{ if($err -eq ''){$err=$_.Exception.Message} }\n" +
            "if(-not $ok){\n" +
            "  try{\n" +
            "    Add-Type -AssemblyName System.Windows.Forms | Out-Null\n" +
            "    Add-Type -AssemblyName System.Drawing | Out-Null\n" +
            "    $img3=[System.Drawing.Image]::FromFile($path)\n" +
            "    $dobj=New-Object System.Windows.Forms.DataObject\n" +
            "    $dobj.SetData([System.Windows.Forms.DataFormats]::Bitmap,$img3)\n" +
            "    [System.Windows.Forms.Clipboard]::SetDataObject($dobj,$true)\n" +
            "    try{$img3.Dispose()}catch{}\n" +
            "    $ok=$true\n" +
            "  }catch{\n" +
            "    if($err -eq ''){$err=$_.Exception.Message}\n" +
            "  }\n" +
            "}\n" +
            "if(-not $ok){\n" +
            "  try{\n" +
            "    Add-Type -AssemblyName PresentationCore | Out-Null\n" +
            "    Add-Type -AssemblyName WindowsBase | Out-Null\n" +
            "    for($i=0;$i -lt 12 -and -not $ok;$i++){\n" +
            "      try{\n" +
            "        $bitmap=New-Object System.Windows.Media.Imaging.BitmapImage\n" +
            "        $bitmap.BeginInit()\n" +
            "        $bitmap.CacheOption=[System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad\n" +
            "        $bitmap.UriSource=New-Object System.Uri($path)\n" +
            "        $bitmap.EndInit()\n" +
            "        $bitmap.Freeze()\n" +
            "        [System.Windows.Clipboard]::SetImage($bitmap)\n" +
            "        $ok=$true\n" +
            "      }catch{\n" +
            "        $err=$_.Exception.Message\n" +
            "        Start-Sleep -Milliseconds 120\n" +
            "      }\n" +
            "    }\n" +
            "  }catch{ if($err -eq ''){$err=$_.Exception.Message} }\n" +
            "}\n" +
            "if($ok){ Write-Result 'OK' } else { if($err -eq ''){$err='Clipboard image write failed'}; Write-Result ('ERR:' + $err) }\n";

        var out = runPowerShellWindows(ps);
        if (isBlankText(out)) {
            out = readTextFile(resultFile);
        }

        try { if (resultFile.exists) resultFile.remove(); } catch (cleanupResult1) {}

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, message: "" };
        }

        if (isBlankText(out)) {
            var verifyFile = null;
            try {
                verifyFile = uniqueFile(getTempFolder(), "copy_pasta_verify", "png");
                var verify = getClipboardImageWindows(verifyFile.fsName);
                if (verify.ok) {
                    return { ok: true, message: "" };
                }
                var verifyMessage = verify.message || "Clipboard verification failed.";
                if (!isBlankText(helperMessage)) {
                    verifyMessage = "Helper: " + helperMessage + " | " + verifyMessage;
                }
                return { ok: false, message: verifyMessage };
            } catch (verifyErr) {
                var noOutputMessage = "No output from PowerShell and clipboard verification failed: " + verifyErr.toString();
                if (!isBlankText(helperMessage)) {
                    noOutputMessage = "Helper: " + helperMessage + " | " + noOutputMessage;
                }
                return { ok: false, message: noOutputMessage };
            } finally {
                try {
                    if (verifyFile && verifyFile.exists) verifyFile.remove();
                } catch (cleanupErr1) {}
            }
        }

        var finalMessage = out || "Unknown clipboard error.";
        if (!isBlankText(helperMessage)) {
            finalMessage = "Helper: " + helperMessage + " | PowerShell: " + finalMessage;
        }

        return { ok: false, message: finalMessage };
    }

    function getClipboardImageWindows(filePath) {
        var helperResult = runWindowsClipboardHelper("get", filePath);
        if (helperResult.ok) {
            return { ok: true, noImage: false, message: "" };
        }
        if (helperResult.noImage) {
            return { ok: false, noImage: true, message: helperResult.message || "Clipboard does not contain an image." };
        }

        var helperMessage = helperResult.message || "";

        var p = psQuote(filePath);
        var resultFile = uniqueFile(getTempFolder(), "copy_pasta_result_get", "txt");
        var r = psQuote(resultFile.fsName);
        var ps =
            "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8\n" +
            "$OutputEncoding=[System.Text.Encoding]::UTF8\n" +
            "$ErrorActionPreference='Stop'\n" +
            "$outPath='" + p + "'\n" +
            "$resultPath='" + r + "'\n" +
            "$saved=$false\n" +
            "$err=''\n" +
            "$lastFormats=''\n" +
            "function Write-Result([string]$s){ try{ Set-Content -LiteralPath $resultPath -Value $s -Encoding UTF8 -Force }catch{}; Write-Output $s }\n" +
            "function Save-PngStream($stream, $path){\n" +
            "  if($stream -eq $null){ return $false }\n" +
            "  $fs=[System.IO.File]::Open($path,[System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)\n" +
            "  try{ $stream.Position=0 }catch{}\n" +
            "  $stream.CopyTo($fs)\n" +
            "  $fs.Dispose()\n" +
            "  try{ $stream.Dispose() }catch{}\n" +
            "  return $true\n" +
            "}\n" +
            "function Try-DownloadImageToPng($url, $path){\n" +
            "  if([string]::IsNullOrWhiteSpace($url)){ return $false }\n" +
            "  $url=$url.Trim()\n" +
            "  if($url -notmatch '^https?://'){ return $false }\n" +
            "  $tmp=[System.IO.Path]::GetTempFileName()\n" +
            "  try{\n" +
            "    $ProgressPreference='SilentlyContinue'\n" +
            "    Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -ErrorAction Stop | Out-Null\n" +
            "    $img=[System.Drawing.Image]::FromFile($tmp)\n" +
            "    $img.Save($path,[System.Drawing.Imaging.ImageFormat]::Png)\n" +
            "    $img.Dispose()\n" +
            "    return $true\n" +
            "  }catch{\n" +
            "    return $false\n" +
            "  }finally{\n" +
            "    try{ if(Test-Path -LiteralPath $tmp){ Remove-Item -LiteralPath $tmp -Force } }catch{}\n" +
            "  }\n" +
            "}\n" +
            "function Extract-ImageUrl($text){\n" +
            "  if([string]::IsNullOrWhiteSpace($text)){ return $null }\n" +
            "  $m=[regex]::Match($text,'https?://[^\\s\"''<>]+')\n" +
            "  if($m.Success){\n" +
            "    $u=$m.Value\n" +
            "    $u=$u -replace '&amp;','&'\n" +
            "    return $u\n" +
            "  }\n" +
            "  return $null\n" +
            "}\n" +
            "try{\n" +
            "  Add-Type -AssemblyName System.Windows.Forms | Out-Null\n" +
            "  Add-Type -AssemblyName System.Drawing | Out-Null\n" +
            "  for($i=0;$i -lt 14 -and -not $saved;$i++){\n" +
            "    try{\n" +
            "      if([System.Windows.Forms.Clipboard]::ContainsImage()){\n" +
            "        $img=[System.Windows.Forms.Clipboard]::GetImage()\n" +
            "        if($img -ne $null){\n" +
            "          $img.Save($outPath,[System.Drawing.Imaging.ImageFormat]::Png)\n" +
            "          $img.Dispose()\n" +
            "          $saved=$true\n" +
            "          break\n" +
            "        }\n" +
            "      }\n" +
            "      $data=[System.Windows.Forms.Clipboard]::GetDataObject()\n" +
            "      if($data -ne $null -and -not $saved){\n" +
            "        $formats=$data.GetFormats()\n" +
            "        try{ $lastFormats=($formats -join ', ') }catch{}\n" +
            "        foreach($fmt in $formats){\n" +
            "          if($fmt -match 'PNG'){\n" +
            "            $png=$data.GetData($fmt)\n" +
            "            if($png -is [System.IO.Stream]){\n" +
            "              if(Save-PngStream $png $outPath){ $saved=$true; break }\n" +
            "            }elseif($png -is [byte[]]){\n" +
            "              [System.IO.File]::WriteAllBytes($outPath,$png)\n" +
            "              $saved=$true\n" +
            "              break\n" +
            "            }\n" +
            "          }\n" +
            "        }\n" +
            "      }\n" +
            "      if($data -ne $null -and -not $saved -and $data.GetDataPresent([System.Windows.Forms.DataFormats]::Bitmap)){\n" +
            "        $bmp=$data.GetData([System.Windows.Forms.DataFormats]::Bitmap)\n" +
            "        if($bmp -is [System.Drawing.Image]){\n" +
            "          $bmp.Save($outPath,[System.Drawing.Imaging.ImageFormat]::Png)\n" +
            "          $bmp.Dispose()\n" +
            "          $saved=$true\n" +
            "          break\n" +
            "        }\n" +
            "      }\n" +
            "      if($data -ne $null -and -not $saved -and $data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)){\n" +
            "        $files=$data.GetData([System.Windows.Forms.DataFormats]::FileDrop)\n" +
            "        if($files -ne $null){\n" +
            "          foreach($f in $files){\n" +
            "            if([System.IO.File]::Exists($f)){\n" +
            "              try{\n" +
            "                $img2=[System.Drawing.Image]::FromFile($f)\n" +
            "                $img2.Save($outPath,[System.Drawing.Imaging.ImageFormat]::Png)\n" +
            "                $img2.Dispose()\n" +
            "                $saved=$true\n" +
            "                break\n" +
            "              }catch{}\n" +
            "            }\n" +
            "          }\n" +
            "        }\n" +
            "      }\n" +
            "    }catch{\n" +
            "      $err=$_.Exception.Message\n" +
            "      Start-Sleep -Milliseconds 120\n" +
            "    }\n" +
            "  }\n" +
            "}catch{ if($err -eq ''){$err=$_.Exception.Message} }\n" +
            "if(-not $saved){\n" +
            "  try{\n" +
            "    Add-Type -AssemblyName PresentationCore | Out-Null\n" +
            "    Add-Type -AssemblyName WindowsBase | Out-Null\n" +
            "    for($i=0;$i -lt 10 -and -not $saved;$i++){\n" +
            "      try{\n" +
            "        $src=[System.Windows.Clipboard]::GetImage()\n" +
            "        if($src -ne $null){\n" +
            "          $encoder=New-Object System.Windows.Media.Imaging.PngBitmapEncoder\n" +
            "          $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($src))\n" +
            "          $fs=[System.IO.File]::Open($outPath,[System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)\n" +
            "          $encoder.Save($fs)\n" +
            "          $fs.Dispose()\n" +
            "          $saved=$true\n" +
            "          break\n" +
            "        }\n" +
            "      }catch{\n" +
            "        $err=$_.Exception.Message\n" +
            "        Start-Sleep -Milliseconds 120\n" +
            "      }\n" +
            "    }\n" +
            "  }catch{ if($err -eq ''){$err=$_.Exception.Message} }\n" +
            "}\n" +
            "if(-not $saved){\n" +
            "  try{\n" +
            "    $urlCandidate=$null\n" +
            "    $d=[System.Windows.Forms.Clipboard]::GetDataObject()\n" +
            "    if($d -ne $null){\n" +
            "      if($d.GetDataPresent('HTML Format')){\n" +
            "        $html=$d.GetData('HTML Format')\n" +
            "        $urlCandidate=Extract-ImageUrl $html\n" +
            "      }\n" +
            "      if($urlCandidate -eq $null -and $d.GetDataPresent([System.Windows.Forms.DataFormats]::UnicodeText)){\n" +
            "        $txt=$d.GetData([System.Windows.Forms.DataFormats]::UnicodeText)\n" +
            "        $urlCandidate=Extract-ImageUrl $txt\n" +
            "      }\n" +
            "      if($urlCandidate -eq $null -and $d.GetDataPresent([System.Windows.Forms.DataFormats]::Text)){\n" +
            "        $txt2=$d.GetData([System.Windows.Forms.DataFormats]::Text)\n" +
            "        $urlCandidate=Extract-ImageUrl $txt2\n" +
            "      }\n" +
            "    }\n" +
            "    if($urlCandidate -eq $null){\n" +
            "      try{\n" +
            "        $clipTxt=Get-Clipboard -Raw\n" +
            "        $urlCandidate=Extract-ImageUrl $clipTxt\n" +
            "      }catch{}\n" +
            "    }\n" +
            "    if($urlCandidate -ne $null){\n" +
            "      if(Try-DownloadImageToPng $urlCandidate $outPath){\n" +
            "        $saved=$true\n" +
            "      }elseif($err -eq ''){\n" +
            "        $err='Clipboard had URL data but image download/convert failed'\n" +
            "      }\n" +
            "    }\n" +
            "  }catch{ if($err -eq ''){$err=$_.Exception.Message} }\n" +
            "}\n" +
            "if($saved){ Write-Result 'OK' } elseif($err -ne ''){ Write-Result ('ERR:' + $err) } elseif($lastFormats -ne ''){ Write-Result ('NO_IMAGE:' + $lastFormats) } else { Write-Result 'NO_IMAGE' }\n";

        var out = runPowerShellWindows(ps);
        if (isBlankText(out)) {
            out = readTextFile(resultFile);
        }

        try { if (resultFile.exists) resultFile.remove(); } catch (cleanupResult2) {}

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, noImage: false, message: "" };
        }

        var savedFile = new File(filePath);
        if (savedFile.exists && savedFile.length > 0) {
            return { ok: true, noImage: false, message: "" };
        }

        if (out && out.indexOf("NO_IMAGE") !== -1) {
            var details = out.replace(/\r?\n/g, " ");
            if (details.indexOf("NO_IMAGE:") !== -1) {
                details = details.substring(details.indexOf("NO_IMAGE:") + 9);
                return { ok: false, noImage: true, message: "Clipboard does not contain a readable image format. Available formats: " + details };
            }
            return { ok: false, noImage: true, message: "Clipboard does not contain an image." };
        }

        if (isBlankText(out)) {
            var noStatusMsg = "PowerShell returned no status and no clipboard image file was created. PowerShell execution may be blocked by system policy.";
            if (!isBlankText(helperMessage)) {
                noStatusMsg = "Helper: " + helperMessage + " | " + noStatusMsg;
            }
            return { ok: false, noImage: false, message: noStatusMsg };
        }

        var finalMessage = out || "Clipboard read failed.";
        if (!isBlankText(helperMessage)) {
            finalMessage = "Helper: " + helperMessage + " | PowerShell: " + finalMessage;
        }

        return { ok: false, noImage: false, message: finalMessage };
    }

    function convertToPngMac(inputPath, outputPath) {
        var cmdConvert = "/usr/bin/sips -s format png " +
            shQuote(inputPath) +
            " --out " +
            shQuote(outputPath) +
            " >/dev/null";

        var out = system.callSystem(cmdConvert);
        if (!(new File(outputPath)).exists) {
            return { ok: false, message: out || "Clipboard image conversion failed." };
        }

        return { ok: true, message: "" };
    }

    function writeClipboardDataMac(typeExpression, outputPath) {
        var setPathLine = 'set outPath to POSIX file "' + appleQuote(outputPath) + '"';
        var cmd = "/usr/bin/osascript" +
            " -e " + shQuote("try") +
            " -e " + shQuote("set clipData to the clipboard as " + typeExpression) +
            " -e " + shQuote("on error") +
            " -e " + shQuote('return "NO_IMAGE"') +
            " -e " + shQuote("end try") +
            " -e " + shQuote(setPathLine) +
            " -e " + shQuote("set fRef to open for access outPath with write permission") +
            " -e " + shQuote("try") +
            " -e " + shQuote("set eof fRef to 0") +
            " -e " + shQuote("write clipData to fRef") +
            " -e " + shQuote("close access fRef") +
            " -e " + shQuote('return "OK"') +
            " -e " + shQuote("on error errMsg") +
            " -e " + shQuote("try") +
            " -e " + shQuote("close access fRef") +
            " -e " + shQuote("end try") +
            " -e " + shQuote('return "ERR:" & errMsg') +
            " -e " + shQuote("end try");

        var out = system.callSystem(cmd);

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, noImage: false, message: "" };
        }
        if (out && out.indexOf("NO_IMAGE") !== -1) {
            return { ok: false, noImage: true, message: "Clipboard does not contain this image format." };
        }

        return { ok: false, noImage: false, message: out || "Could not read clipboard image data." };
    }

    function setClipboardImageMac(filePath) {
        var tiffFile = new File(filePath.replace(/\.png$/i, ".tiff"));

        var cmdConvert = "/usr/bin/sips -s format tiff " +
            shQuote(filePath) +
            " --out " +
            shQuote(tiffFile.fsName) +
            " >/dev/null";

        system.callSystem(cmdConvert);

        var applescriptLine = 'set the clipboard to (read (POSIX file "' + appleQuote(tiffFile.fsName) + '") as TIFF picture)';
        var cmdClip = "/usr/bin/osascript" +
            " -e " + shQuote("try") +
            " -e " + shQuote(applescriptLine) +
            " -e " + shQuote('return "OK"') +
            " -e " + shQuote("on error errMsg") +
            " -e " + shQuote('return "ERR:" & errMsg') +
            " -e " + shQuote("end try");

        var out = system.callSystem(cmdClip);

        try { if (tiffFile.exists) tiffFile.remove(); } catch (cleanupErr) {}

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, message: "" };
        }

        return { ok: false, message: out || "Unknown clipboard error." };
    }

    function getClipboardImageMac(filePath) {
        var tiffPath = filePath.replace(/\.png$/i, ".tiff");
        var jpgPath = filePath.replace(/\.png$/i, ".jpg");

        var tiffResult = writeClipboardDataMac("TIFF picture", tiffPath);
        if (tiffResult.ok) {
            var convertTiff = convertToPngMac(tiffPath, filePath);
            try {
                var tiffFile = new File(tiffPath);
                if (tiffFile.exists) tiffFile.remove();
            } catch (cleanupErr1) {}

            if (convertTiff.ok) {
                return { ok: true, noImage: false, message: "" };
            }
            return { ok: false, noImage: false, message: convertTiff.message };
        }

        var pngResult = writeClipboardDataMac("«class PNGf»", filePath);
        if (pngResult.ok && (new File(filePath)).exists) {
            return { ok: true, noImage: false, message: "" };
        }

        var jpgResult = writeClipboardDataMac("JPEG picture", jpgPath);
        if (jpgResult.ok) {
            var convertJpg = convertToPngMac(jpgPath, filePath);
            try {
                var jpgFile = new File(jpgPath);
                if (jpgFile.exists) jpgFile.remove();
            } catch (cleanupErr2) {}

            if (convertJpg.ok) {
                return { ok: true, noImage: false, message: "" };
            }
            return { ok: false, noImage: false, message: convertJpg.message };
        }

        if (tiffResult.noImage && pngResult.noImage && jpgResult.noImage) {
            return { ok: false, noImage: true, message: "Clipboard does not contain an image." };
        }

        var errMsg = tiffResult.message;
        if (!errMsg && !pngResult.noImage) errMsg = pngResult.message;
        if (!errMsg && !jpgResult.noImage) errMsg = jpgResult.message;

        return { ok: false, noImage: false, message: errMsg || "Could not read clipboard image." };
    }

    function copyImageFileToClipboard(fileObj) {
        if (isWindows()) return setClipboardImageWindows(fileObj.fsName);
        if (isMac()) return setClipboardImageMac(fileObj.fsName);
        return { ok: false, message: "Unsupported OS. Clipboard image copy is available on Windows/macOS only." };
    }

    function saveClipboardImageToFile(fileObj) {
        if (isWindows()) return getClipboardImageWindows(fileObj.fsName);
        if (isMac()) return getClipboardImageMac(fileObj.fsName);
        return { ok: false, noImage: false, message: "Unsupported OS. Clipboard image paste is available on Windows/macOS only." };
    }

    function importFileIntoComp(fileObj, comp) {
        var importOptions = new ImportOptions(fileObj);
        var footageItem = app.project.importFile(importOptions);
        var layer = comp.layers.add(footageItem);
        layer.startTime = comp.time;
        return layer;
    }

    function setStatus(statusText, message, isError) {
        statusText.text = message;
        try {
            var g = statusText.graphics;
            var color = isError ? THEME.statusError : THEME.statusOk;
            g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, color, 1);
        } catch (e) {}
    }

    function doCopy(statusText) {
        if (!hasScriptFileNetworkAccess()) {
            alert("Copy Pasta: Scripting file/network access is disabled.\nEnable it in Preferences > Scripting & Expressions > Allow Scripts to Write Files and Access Network.");
            setStatus(statusText, "Enable scripting file/network access.", true);
            return;
        }

        var comp = getActiveComp();
        if (!comp) {
            alert("Copy Pasta: No active composition.\nOpen a composition and try again.");
            setStatus(statusText, "No active composition.", true);
            return;
        }

        if (!comp.selectedLayers || comp.selectedLayers.length === 0) {
            alert("Copy Pasta: No layer selected.\nSelect an Image Layer or Shape Layer, then click Copy.");
            setStatus(statusText, "No layer selected.", true);
            return;
        }

        var layer = getFirstCopyableLayer(comp);
        if (layer === "INVALID_TYPE") {
            alert("Copy Pasta: Unsupported layer type.\nCopy works with still Image Layers and Shape Layers.");
            setStatus(statusText, "Unsupported layer type.", true);
            return;
        }

        if (comp.time < layer.inPoint || comp.time >= layer.outPoint) {
            alert("Copy Pasta: The selected layer is not visible at the current time indicator.\nMove the playhead into the layer's visible range and try again.");
            setStatus(statusText, "Layer not visible at current time.", true);
            return;
        }

        app.beginUndoGroup("Copy Pasta - Copy");
        var tempPng = null;

        try {
            tempPng = uniqueFile(getTempFolder(), "copy_pasta_capture", "png");
            captureLayerToPng(comp, layer, tempPng);

            var clipResult = copyImageFileToClipboard(tempPng);
            if (!clipResult.ok) {
                alert("Copy Pasta: Could not copy image to system clipboard.\n\n" + clipResult.message);
                setStatus(statusText, "Copy failed.", true);
                return;
            }

            setStatus(statusText, "Copied to system clipboard.", false);
        } catch (err) {
            alert("Copy Pasta: Copy failed.\n\n" + err.toString());
            setStatus(statusText, "Copy failed.", true);
        } finally {
            try {
                if (tempPng && tempPng.exists) tempPng.remove();
            } catch (cleanupError) {}
            app.endUndoGroup();
        }
    }

    function doPaste(statusText) {
        if (!hasScriptFileNetworkAccess()) {
            alert("Copy Pasta: Scripting file/network access is disabled.\nEnable it in Preferences > Scripting & Expressions > Allow Scripts to Write Files and Access Network.");
            setStatus(statusText, "Enable scripting file/network access.", true);
            return;
        }

        var comp = getActiveComp();
        if (!comp) {
            alert("Copy Pasta: No active composition.\nOpen a composition and try again.");
            setStatus(statusText, "No active composition.", true);
            return;
        }

        app.beginUndoGroup("Copy Pasta - Paste");

        try {
            var outputFile = uniqueFile(getImportFolder(), "pasted_image", "png");
            var readResult = saveClipboardImageToFile(outputFile);

            if (!readResult.ok) {
                if (readResult.noImage) {
                    alert("Copy Pasta: Clipboard does not contain a readable image.\n\n" + (readResult.message || ""));
                } else {
                    alert("Copy Pasta: Could not read image from system clipboard.\n\n" + readResult.message);
                }
                setStatus(statusText, "Paste failed.", true);
                return;
            }

            if (!outputFile.exists) {
                alert("Copy Pasta: Clipboard read succeeded, but no image file was created.");
                setStatus(statusText, "Paste failed.", true);
                return;
            }

            importFileIntoComp(outputFile, comp);
            setStatus(statusText, "Pasted as new layer.", false);
        } catch (err) {
            alert("Copy Pasta: Paste failed.\n\n" + err.toString());
            setStatus(statusText, "Paste failed.", true);
        } finally {
            app.endUndoGroup();
        }
    }

    function applyWindowStyle(win) {
        setBackgroundColor(win, THEME.bg);
    }

    function buildUI(thisObj) {
        var pal = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", SCRIPT_NAME, undefined, { resizeable: true });

        if (!pal) return null;

        pal.orientation = "column";
        pal.alignChildren = ["fill", "top"];
        pal.spacing = 12;
        pal.margins = 16;

        applyWindowStyle(pal);

        var accentBar = pal.add("panel", undefined, "");
        accentBar.alignment = ["fill", "top"];
        accentBar.preferredSize = [-1, 3];
        accentBar.margins = [0, 0, 0, 0];
        setBackgroundColor(accentBar, THEME.accent);

        var buttonGroup = pal.add("group");
        buttonGroup.orientation = "column";
        buttonGroup.alignChildren = ["fill", "top"];
        buttonGroup.spacing = 10;

        var copyBtn = buttonGroup.add("button", undefined, "Copy");
        copyBtn.preferredSize = [280, 48];
        copyBtn.helpTip = "Copy selected image/shape layer to clipboard.";

        var pasteBtn = buttonGroup.add("button", undefined, "Paste");
        pasteBtn.preferredSize = [280, 48];
        pasteBtn.helpTip = "Paste clipboard image into active composition.";

        var copyFont = getUIFont("BOLD", 20);
        var pasteFont = getUIFont("REGULAR", 20);
        try { if (copyFont) copyBtn.graphics.font = copyFont; } catch (e1) {}
        try { if (pasteFont) pasteBtn.graphics.font = pasteFont; } catch (e2) {}
        setTextColor(copyBtn, THEME.textMain);
        setTextColor(pasteBtn, THEME.textMain);

        var statusPanel = pal.add("panel", undefined, "");
        statusPanel.alignment = ["fill", "top"];
        statusPanel.margins = [10, 10, 10, 10];
        setBackgroundColor(statusPanel, THEME.surface);

        var statusText = statusPanel.add("statictext", undefined, "Ready (v" + SCRIPT_VERSION + ").");
        statusText.alignment = ["fill", "top"];
        statusText.justify = "center";
        try {
            var statusFont = getUIFont("REGULAR", 11);
            if (statusFont) statusText.graphics.font = statusFont;
        } catch (e3) {}
        setTextColor(statusText, THEME.textMuted);

        var tagText = pal.add("statictext", undefined, "@g.fnaa");
        tagText.alignment = ["fill", "top"];
        tagText.justify = "center";
        try {
            var tagFont = getUIFont("REGULAR", 10);
            if (tagFont) tagText.graphics.font = tagFont;
        } catch (e4) {}
        setTextColor(tagText, [0.56, 0.76, 0.9, 1.0]);

        copyBtn.onClick = function () {
            doCopy(statusText);
        };

        pasteBtn.onClick = function () {
            doPaste(statusText);
        };

        pal.onResizing = pal.onResize = function () {
            this.layout.resize();
        };

        pal.layout.layout(true);
        return pal;
    }

    var panel = buildUI(thisObj);
    if (panel && panel instanceof Window) {
        panel.center();
        panel.show();
    }
})(this);
