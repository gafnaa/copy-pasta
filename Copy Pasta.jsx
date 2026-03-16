#target aftereffects
#targetengine "copy_pasta_engine"

(function CopyPastaPanel(thisObj) {
    var SCRIPT_NAME = "Copy Pasta";
    var TEMP_FOLDER_NAME = "CopyPastaTemp";
    var IMPORT_FOLDER_NAME = "CopyPasta Imports";

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
        var p = psQuote(filePath);
        var ps =
            "$ErrorActionPreference='Stop';" +
            "Add-Type -AssemblyName System.Windows.Forms;" +
            "Add-Type -AssemblyName System.Drawing;" +
            "$img=$null;" +
            "$path='" + p + "';" +
            "for($i=0;$i -lt 12 -and $img -eq $null;$i++){" +
            "try{$img=[System.Drawing.Image]::FromFile($path)}catch{Start-Sleep -Milliseconds 120}" +
            "};" +
            "if($img -eq $null){Write-Output 'ERR:Could not open rendered image';exit}" +
            "$setOk=$false;" +
            "for($i=0;$i -lt 10 -and -not $setOk;$i++){" +
            "try{[System.Windows.Forms.Clipboard]::SetImage($img);$setOk=$true}catch{Start-Sleep -Milliseconds 120}" +
            "};" +
            "$img.Dispose();" +
            "if($setOk){Write-Output 'OK'}else{Write-Output 'ERR:Clipboard busy or inaccessible'}";

        var cmd = "powershell -NoProfile -STA -ExecutionPolicy Bypass -Command \"" + ps + "\"";
        var out = system.callSystem(cmd);

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, message: "" };
        }

        return { ok: false, message: out || "Unknown clipboard error." };
    }

    function getClipboardImageWindows(filePath) {
        var p = psQuote(filePath);
        var ps =
            "$ErrorActionPreference='Stop';" +
            "Add-Type -AssemblyName System.Windows.Forms;" +
            "Add-Type -AssemblyName System.Drawing;" +
            "$outPath='" + p + "';" +
            "$saved=$false;" +
            "for($i=0;$i -lt 10 -and -not $saved;$i++){" +
            "try{" +
            "$data=[System.Windows.Forms.Clipboard]::GetDataObject();" +
            "if([System.Windows.Forms.Clipboard]::ContainsImage()){" +
            "$img=[System.Windows.Forms.Clipboard]::GetImage();" +
            "if($img -ne $null){$img.Save($outPath,[System.Drawing.Imaging.ImageFormat]::Png);$img.Dispose();$saved=$true}" +
            "};" +
            "if(-not $saved -and $data -ne $null -and $data.GetDataPresent('PNG')){" +
            "$png=$data.GetData('PNG');" +
            "if($png -is [System.IO.MemoryStream]){[System.IO.File]::WriteAllBytes($outPath,$png.ToArray());$png.Dispose();$saved=$true}" +
            "elseif($png -is [byte[]]){[System.IO.File]::WriteAllBytes($outPath,$png);$saved=$true}" +
            "};" +
            "if(-not $saved -and $data -ne $null -and $data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)){" +
            "$files=$data.GetData([System.Windows.Forms.DataFormats]::FileDrop);" +
            "if($files -ne $null){foreach($f in $files){if([System.IO.File]::Exists($f)){try{$img2=[System.Drawing.Image]::FromFile($f);$img2.Save($outPath,[System.Drawing.Imaging.ImageFormat]::Png);$img2.Dispose();$saved=$true;break}catch{}}}}" +
            "};" +
            "}catch{Start-Sleep -Milliseconds 120}" +
            "};" +
            "if($saved){Write-Output 'OK'}else{Write-Output 'NO_IMAGE'}";

        var cmd = "powershell -NoProfile -STA -ExecutionPolicy Bypass -Command \"" + ps + "\"";
        var out = system.callSystem(cmd);

        if (out && out.indexOf("OK") !== -1) {
            return { ok: true, noImage: false, message: "" };
        }
        if (out && out.indexOf("NO_IMAGE") !== -1) {
            return { ok: false, noImage: true, message: "Clipboard does not contain an image." };
        }

        return { ok: false, noImage: false, message: out || "Unknown clipboard error." };
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
            var color = isError ? [1.0, 0.35, 0.35, 1.0] : [0.65, 0.9, 0.65, 1.0];
            g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, color, 1);
        } catch (e) {}
    }

    function doCopy(statusText) {
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
                    alert("Copy Pasta: Clipboard does not contain an image.");
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
        try {
            var g = win.graphics;
            g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, [0.12, 0.12, 0.12, 1.0]);
        } catch (e) {}
    }

    function buildUI(thisObj) {
        var pal = (thisObj instanceof Panel)
            ? thisObj
            : new Window("palette", SCRIPT_NAME, undefined, { resizeable: true });

        if (!pal) return null;

        pal.orientation = "column";
        pal.alignChildren = ["fill", "top"];
        pal.spacing = 10;
        pal.margins = 16;

        applyWindowStyle(pal);

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

        try {
            copyBtn.graphics.font = ScriptUI.newFont("Segoe UI", "BOLD", 16);
            pasteBtn.graphics.font = ScriptUI.newFont("Segoe UI", "BOLD", 16);
        } catch (e1) {}

        var statusPanel = pal.add("panel", undefined, "");
        statusPanel.alignment = ["fill", "top"];
        statusPanel.margins = [8, 8, 8, 8];

        var statusText = statusPanel.add("statictext", undefined, "Ready.");
        statusText.alignment = ["fill", "top"];
        statusText.justify = "center";
        try { statusText.graphics.font = ScriptUI.newFont("Segoe UI", "REGULAR", 11); } catch (e2) {}

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
