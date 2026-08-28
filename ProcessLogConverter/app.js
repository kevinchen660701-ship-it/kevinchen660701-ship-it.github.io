(function () {
  "use strict";

  const VERSION = "1.4.5.0";
  const PROCESS_LOG_FILE_NAME = "ProcessLog.ini";
  const SERVER_IDX_NC3 = 2;
  const RECIPE_LEVEL_STEP_NUM = 10;
  const RECIPE_IDX_NC3_DIS_G1 = 7;
  const RECIPE_IDX_NC3_DIS_G2 = 8;
  const RECIPE_IDX_NC1_TOTAL_STEP = 6;
  const RECIPE_IDX_NC1_TARGET_THICK = 7;
  const WHEEL_BEFORE_GRD_DATA_LEN = 21;
  const WHEEL_BEFORE_GRD_CHUCK_MAX = 7;
  const WHEEL_BEFORE_GRD_LAST_WAFER_MAX = 10;
  const WHEEL_BEFORE_GRD_LAST_WAFER_MIN = 11;
  const WHEEL_BEFORE_GRD_INCOME_THICK = 12;
  const WHEEL_AFTER_GRD_WEAR_SINGLE = 3;
  const WHEEL_AFTER_GRD_WEAR_TOTAL = 4;
  const WHEEL_AFTER_GRD_TARGET_CONFIRM = 11;
  const WHEEL_AFTER_GRD_WHEEL_THICK = 12;
  const SAMP_LEN = 26;
  const SAMP_IDX_S1_LOADING = 7;
  const SAMP_IDX_T2 = 15;
  const AI_AIR_CUT_BASE_DATA_LEN = 5;
  const AI_AIR_CUT_EXTENDED_DATA_LEN = 9;
  const MAX_PROFILE_STRING_LENGTH = 2999;

  const RECIPE_INIT_STEP0_NC12 = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";
  const RECIPE_INIT_STEP1_10_NC12 = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,";
  const RECIPE_INIT_STEP0_NC3 = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";
  const WHEEL_INIT_STEP0_10 = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0";
  const WAFER_LOG_DATA_INIT = "0,1,1,1";
  const WAFER_LOG_RUN_TIME_INIT = "0,0,0";
  const WAFER_LOG_CASS_INIT = "0,0,0";
  const WAFER_LOG_GRD_TIME_INIT = "0,0,0";
  const WAFER_LOG_PEAK_INIT = "0,0,0,0,0,0,0";
  const AI_AIR_CUT_DATA_INIT = "0,0,0,0,0";
  const WAFER_LOG_INCLINE_SV_INIT = "0,0,0,0";
  const PROCESS_LOG_RUN_DATA_INIT = "0,0,0,0";
  const PROCESS_LOG_ID_INIT = "0,0,0";

  const RELEASE_NOTES = `FSG-2300 ProcessLog Converter
Release Notes

[1.4.5.0] 2026-06-23
Added
Item 1
  Added GrdDistCollect.csv for extended AirCut data.
Item 2
  Collects G1 and G2 wafer rows into separate blocks.
Item 3
  Includes chuck, wafer ID, and the four new GrdDist values.
Item 4
  Skips old wafer logs without meaningful new values.

[1.4.4.0] 2026-06-23
Fixed
Item 1
  Treated blank extended AirCut fields as missing data.
Item 2
  Treated all-zero extended AirCut fields as missing data.
Item 3
  Hid the four extended AirCut columns for old logs with zero padding.
Item 4
  Continued showing extended AirCut columns when any new value is non-zero.

[1.4.3.0] 2026-06-23
Changed
Item 1
  Hid extended AirCutS1LoadingRec columns for older ini files.
Item 2
  Showed the four new AirCut columns only when data exists.
Item 3
  Kept original five-column AirCut output for old logs.
Item 4
  Preserved extended output for new logs with the extra values.

[1.4.2.0] 2026-06-23
Added
Item 1
  Added four extended AirCutS1LoadingRec output columns.
Item 2
  Added S1ReachZmch(mm).
Item 3
  Added StartZmch-EndZmch(um) and GrdDist-RemoveValue(um).
Item 4
  Added S1ReachZabs-WaferIncomeMax(um) with old ini padding.

[1.4.1.0] 2026-06-09
Added
Item 1
  Added S1OutputPower(KW) to wafer sampling CSV output.
Item 2
  Placed the new column after CGLevelComp(um).
Item 3
  Preserved the extra sample value when provided by the ini file.
Item 4
  Padded older sampling rows so the new column stays aligned.

[1.4.0.0] 2026-06-02
Fixed
Item 1
  Reduced memory growth during large folder conversion.
Item 2
  Released cached ini data after each wafer is processed.
Item 3
  Changed ini parsing to stream lines instead of reading all lines.
Item 4
  Disabled Prefer 32-bit so 64-bit Windows can use more memory.

[1.3.9.0] 2026-05-19
Changed
Item 1
  Tightened spacing between Release Note versions.
Item 2
  Kept only one blank line between different versions.
Item 3
  Preserved compact item spacing within each version.
Item 4
  Updated bundled RELEASE_NOTES.txt display content.

[1.3.8.0] 2026-05-19
Changed
Item 1
  Tightened Release Note item spacing.
Item 2
  Removed blank lines between items in the same version.
Item 3
  Kept blank lines only between different versions.
Item 4
  Improved readability in the Release Note window.

[1.3.7.0] 2026-05-19
Changed
Item 1
  Reduced repeated file metadata checks during conversion.
Item 2
  Reused cached ini content after the first file load.
Item 3
  Improved speed for heavy wafer and sample lookups.
Item 4
  Kept conversion output format unchanged.

[1.3.6.0] 2026-05-19
Changed
Item 1
  Improved the Release Note text layout.
Item 2
  Changed the Release Note window to RichTextBox.
Item 3
  Enabled cleaner word wrapping in the note window.
Item 4
  Reworked this text file into shorter readable lines.

[1.3.5.0] 2026-05-19
Changed
Item 1
  Added progress updates while collecting S1 Loading data.
Item 2
  Added status updates while writing S1LoadingCollect.csv.
Item 3
  Added status updates while writing S1LoadingCollect.xlsx.
Item 4
  Improved remaining time refresh during long S1 output steps.

[1.3.4.0] 2026-05-19
Changed
Item 1
  Renamed the S1 Loading csv output file.
Item 2
  S1LoadingMatch.csv is now S1LoadingCollect.csv.
Item 3
  Renamed the S1 Loading xlsx output file.
Item 4
  S1LoadingMatch.xlsx is now S1LoadingCollect.xlsx.

[1.3.3.0] 2026-05-19
Changed
Item 1
  Added per-conversion ini file caching.
Item 2
  Reused wafer file lists instead of scanning folders twice.
Item 3
  Built G1 and G2 S1 Loading columns in one wafer pass.
Item 4
  Precomputed XLSX column names while streaming worksheets.

[1.3.2.0] 2026-05-12
Changed
Item 1
  Improved S1LoadingMatch generation speed.
Item 2
  Reused one S1 Loading data set for CSV and XLSX output.
Item 3
  Changed XLSX worksheet writing to stream directly to the file.
Item 4
  Reduced duplicate wafer ini reads during S1LoadingMatch output.

[1.3.1.0] 2026-05-12
Added
Item 1
  Remember the last selected source folder.
Item 2
  Remember the last selected output base folder.
Item 3
  Load saved paths automatically when the application starts.
Item 4
  Store path settings in FSG-2300_ProcessLogConverter.ini.

[1.3.0.0] 2026-05-12
Added
Item 1
  Added S1LoadingMatch.xlsx output.
Item 2
  Added separate Excel worksheet tabs for G1 and G2.
Item 3
  Kept S1LoadingMatch.csv for compatibility.
Item 4
  Updated documentation and release notes.

[1.2.1.0] 2026-05-12
Changed
Item 1
  Changed output folder format to:
  yyyyMMdd_HHmmss_ProcessLogCsv
Item 2
  Updated Release Note text to use item-style entries.
Item 3
  Shortened release note lines for readability.
Item 4
  Release Note window now wraps long lines automatically.

[1.2.0.0] 2026-05-12
Added
Item 1
  Added S1LoadingMatch.csv output for each ProcessLog folder.
Item 2
  S1LoadingMatch.csv separates G1 and G2 data blocks.
Item 3
  Each block arranges wafer S1 Loading values by wafer ID.
  The first row shows G1_Chuck or G2_Chuck.
  The second row shows Wafer_ID.
Item 4
  Added RELEASE_NOTES.txt.
  Added a main form Release Note button.
Documentation
Item 1
  Updated README.md with S1LoadingMatch behavior.
Item 2
  Updated release note files for the UI display feature.
Item 3
  Updated output folder format documentation.
Item 4
  Kept release note lines short for easier reading.

[1.1.1.0] 2026-05-12
Fixed
Item 1
  Added missing units to the WaferBatchLog.csv header row.
Item 2
  Updated BatchLog headers to align with FSG-2300 BatchView.
Item 3
  Added units for T2, Income, TarSV, TarPV, and LaMin.
Item 4
  Added units for peak load, wear, and wheel thickness.
Documentation
Item 1
  Added README.md.
Item 2
  Added release notes for version tracking.

[1.1.0.0] 2026-05-05
Added
Item 1
  Added version display on the main form.
Item 2
  Renamed the project to FSG-2300_ProcessLogConverter.
Item 3
  Added estimated remaining time while converting.
Item 4
  Added timestamped output folder creation.
  Added Cancel button during conversion.
  Added executable icon.

[1.0.0.0] 2026-05-05
Added
Item 1
  Initial Visual Studio 2013 / .NET Framework 4.5 project.
Item 2
  Added source folder selection.
Item 3
  Added ProcessLog conversion to CSV.
Item 4
  Added recursive folder conversion.
  Added progress bar.
`;

  const ui = {
    selectFolderButton: document.getElementById("selectFolderButton"),
    convertButton: document.getElementById("convertButton"),
    cancelButton: document.getElementById("cancelButton"),
    releaseNoteButton: document.getElementById("releaseNoteButton"),
    progressBar: document.getElementById("progressBar"),
    versionLabel: document.getElementById("versionLabel"),
    sourceFolderLabel: document.getElementById("sourceFolderLabel"),
    outputFolderLabel: document.getElementById("outputFolderLabel"),
    remainTimeLabel: document.getElementById("remainTimeLabel"),
    statusLabel: document.getElementById("statusLabel"),
    folderInput: document.getElementById("folderInput"),
    releaseNoteDialog: document.getElementById("releaseNoteDialog"),
    releaseNoteText: document.getElementById("releaseNoteText"),
    closeReleaseNoteButton: document.getElementById("closeReleaseNoteButton")
  };

  let sourceFolder = null;
  let outputBaseFolder = null;
  let savedOutputHandle = null;
  let outputFolderName = "";
  let isBusy = false;
  let cancelRequested = false;
  let convertStartTime = 0;

  ui.versionLabel.textContent = "版本：" + VERSION;
  loadSettings();
  restoreDirectoryHandles();

  ui.selectFolderButton.addEventListener("click", selectFolder);
  ui.convertButton.addEventListener("click", convertFiles);
  ui.cancelButton.addEventListener("click", cancelConvert);
  ui.releaseNoteButton.addEventListener("click", showReleaseNotes);
  ui.closeReleaseNoteButton.addEventListener("click", () => ui.releaseNoteDialog.close());
  ui.folderInput.addEventListener("change", handleFallbackFolderSelection);

  async function selectFolder() {
    if (isBusy) {
      return;
    }

    try {
      sourceFolder = null;
      ui.folderInput.value = "";
      ui.folderInput.click();
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      showError(error);
    }
  }

  async function handleFallbackFolderSelection(event) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) {
      return;
    }

    try {
      await clearDirectoryHandle("sourceHandle");
      const selectedSourceFolder = await buildSourceFromFileList(files);
      const hasProcessLog = hasProcessLogFolder(selectedSourceFolder);
      if (hasProcessLog) {
        sourceFolder = selectedSourceFolder;
        saveSourceFolderName(sourceFolder.displayName);
        ui.sourceFolderLabel.textContent = "來源資料夾：" + sourceFolder.displayName;
      } else {
        clearSourceFolderName();
        ui.sourceFolderLabel.textContent = "來源資料夾：未選取";
        sourceFolder = null;
      }
      ui.convertButton.disabled = false;
      ui.statusLabel.textContent = hasProcessLog
        ? "已選取資料夾，可開始轉換"
        : buildProcessLogNotFoundMessage(selectedSourceFolder, "剛選的資料夾");
    } catch (error) {
      showError(error);
    }
  }

  async function convertFiles() {
    if (isBusy) {
      return;
    }

    if (!await prepareSourceForConversion()) {
      return;
    }

    let outputDirectoryHandle = null;
    outputFolderName = makeOutputFolderName();
    try {
      if ("showDirectoryPicker" in window) {
        outputBaseFolder = await pickOutputDirectory();
        savedOutputHandle = outputBaseFolder;
        await saveDirectoryHandle("outputHandle", outputBaseFolder);
        saveOutputFolderName(outputBaseFolder.name);
        ui.outputFolderLabel.textContent = "輸出資料夾：" + outputBaseFolder.name + "\\" + outputFolderName;
        outputDirectoryHandle = outputBaseFolder;
      } else {
        ui.outputFolderLabel.textContent = "輸出資料夾：" + outputFolderName + ".zip";
      }
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      showError(error);
      return;
    }

    setBusy(true);
    cancelRequested = false;
    convertStartTime = performance.now();
    ui.progressBar.value = 0;
    ui.remainTimeLabel.textContent = "預估剩餘時間：計算中";
    ui.statusLabel.textContent = "正在轉換...";

    try {
      const result = await convertFolder(sourceFolder, reportProgress, isCancellationRequested);
      throwIfCancellationRequested();

      if (outputDirectoryHandle) {
        await writeOutputsToDirectory(outputDirectoryHandle, outputFolderName, result.outputs, reportWriteProgress);
      } else {
        downloadZip(outputFolderName + ".zip", result.outputs.map((file) => ({
          path: outputFolderName + "/" + file.path,
          data: file.data,
          binary: file.binary
        })));
      }

      ui.progressBar.value = 100;
      ui.remainTimeLabel.textContent = "完成時間：" + formatDuration(performance.now() - convertStartTime);
      ui.statusLabel.textContent = "轉換完成：" + result.folderCount + " 個資料夾，" + result.outputFileCount + " CSV";
      window.alert("轉換完成，共轉換 " + result.folderCount + " 個資料夾，產生 " + result.outputFileCount + " 個 CSV 檔。\r\n\r\n輸出資料夾：" + outputFolderName);
    } catch (error) {
      if (error && error.name === "OperationCanceled") {
        ui.progressBar.value = 0;
        ui.remainTimeLabel.textContent = "預估剩餘時間：--:--:--";
        ui.statusLabel.textContent = "轉換已取消";
        window.alert("轉換已取消。");
      } else {
        ui.progressBar.value = 0;
        ui.remainTimeLabel.textContent = "預估剩餘時間：--:--:--";
        ui.statusLabel.textContent = "轉換失敗";
        showError(error);
      }
    } finally {
      setBusy(false);
    }
  }

  async function prepareSourceForConversion() {
    if (sourceFolder && hasProcessLogFolder(sourceFolder)) {
      return true;
    }

    if (sourceFolder && !hasProcessLogFolder(sourceFolder)) {
      ui.statusLabel.textContent = "目前來源資料夾找不到 ProcessLog.ini，請按「選取資料夾」重新選 Desktop\\Ini 底下的原始來源資料夾";
      return false;
    }

    window.alert("請先按「選取資料夾」選取來源資料夾。");
    return false;
  }

  async function pickOutputDirectory() {
    const options = {
      id: "fsg2300-output-folder",
      mode: "readwrite",
      startIn: savedOutputHandle || "desktop"
    };
    try {
      return await window.showDirectoryPicker(options);
    } catch (error) {
      if (error && error.name === "TypeError") {
        return await window.showDirectoryPicker({ mode: "readwrite" });
      }
      throw error;
    }
  }

  function cancelConvert() {
    if (!isBusy) {
      return;
    }
    cancelRequested = true;
    ui.cancelButton.disabled = true;
    ui.statusLabel.textContent = "正在取消，請稍候...";
  }

  function showReleaseNotes() {
    ui.releaseNoteText.textContent = RELEASE_NOTES;
    if (typeof ui.releaseNoteDialog.showModal === "function") {
      ui.releaseNoteDialog.showModal();
    } else {
      window.alert(RELEASE_NOTES);
    }
  }

  function setBusy(value) {
    isBusy = value;
    ui.selectFolderButton.disabled = value;
    ui.convertButton.disabled = value;
    ui.cancelButton.disabled = !value;
    ui.cancelButton.hidden = !value;
    ui.releaseNoteButton.disabled = value;
    ui.releaseNoteButton.hidden = value;
    document.body.classList.toggle("busy", value);
  }

  function reportProgress(percent, message) {
    const safePercent = Math.max(0, Math.min(100, percent));
    ui.progressBar.value = safePercent;
    if (message) {
      ui.statusLabel.textContent = message;
    }
    ui.remainTimeLabel.textContent = "預估剩餘時間：" + getRemainingTimeText(safePercent);
  }

  function reportWriteProgress(index, total, path) {
    const percent = total <= 0 ? 100 : Math.round((index * 100) / total);
    reportProgress(percent, "正在寫入 " + path);
  }

  function getRemainingTimeText(percent) {
    if (percent <= 0 || !convertStartTime) {
      return "計算中";
    }
    if (percent >= 100) {
      return "00:00:00";
    }
    const elapsedMs = performance.now() - convertStartTime;
    const totalMs = elapsedMs * 100 / percent;
    return formatDuration(Math.max(0, totalMs - elapsedMs));
  }

  function formatDuration(milliseconds) {
    const totalSeconds = Math.floor(milliseconds / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return pad2(hours) + ":" + pad2(minutes) + ":" + pad2(seconds);
  }

  function pad2(value) {
    return String(value).padStart(2, "0");
  }

  function makeOutputFolderName() {
    const now = new Date();
    return now.getFullYear()
      + pad2(now.getMonth() + 1)
      + pad2(now.getDate())
      + "_"
      + pad2(now.getHours())
      + pad2(now.getMinutes())
      + pad2(now.getSeconds())
      + "_ProcessLogCsv";
  }

  function isCancellationRequested() {
    return cancelRequested;
  }

  function throwIfCancellationRequested() {
    if (cancelRequested) {
      const error = new Error("Operation canceled.");
      error.name = "OperationCanceled";
      throw error;
    }
  }

  function showError(error) {
    const message = isStaleFileSystemError(error)
      ? "輸出資料夾狀態被系統更新，瀏覽器寫入失敗。請確認輸出資料夾或正在寫入的 xlsx/csv 沒有被檔案總管、Excel 或同步軟體占用後，再按「轉換檔案」重試。"
      : (error && error.message ? error.message : String(error || "Unknown error."));
    window.alert(message);
  }

  function isStaleFileSystemError(error) {
    const message = String(error && error.message || "").toLowerCase();
    return message.includes("state cached")
      || message.includes("state had changed")
      || message.includes("read from disk");
  }

  function loadSettings() {
    try {
      const sourceName = localStorage.getItem("FSG2300.SourceFolderName");
      const outputName = localStorage.getItem("FSG2300.OutputBaseFolderName");
      if (sourceName) {
        ui.sourceFolderLabel.textContent = "來源資料夾：未選取";
        ui.statusLabel.textContent = "已記錄上次來源資料夾：" + sourceName + "；請重新按「選取資料夾」載入檔案";
      }
      if (outputName) {
        ui.outputFolderLabel.textContent = "輸出資料夾：" + outputName;
      }
    } catch {
      ui.statusLabel.textContent = "設定檔讀取失敗，請重新選取資料夾";
    }
  }

  function saveSourceFolderName(name) {
    localStorage.setItem("FSG2300.SourceFolderName", name || "");
    localStorage.setItem("FSG2300.HasSourceFolder", name ? "1" : "");
  }

  function clearSourceFolderName() {
    localStorage.removeItem("FSG2300.SourceFolderName");
    localStorage.removeItem("FSG2300.HasSourceFolder");
  }

  function saveOutputFolderName(name) {
    localStorage.setItem("FSG2300.OutputBaseFolderName", name || "");
  }

  async function clearSourceFolderRecord() {
    await clearDirectoryHandle("sourceHandle");
    sourceFolder = null;
    clearSourceFolderName();
    ui.sourceFolderLabel.textContent = "來源資料夾：未選取";
  }

  async function restoreDirectoryHandles() {
    await clearDirectoryHandle("sourceHandle");
    if (!("showDirectoryPicker" in window)) {
      return;
    }
    try {
      const outputHandle = await getDirectoryHandle("outputHandle");
      if (outputHandle) {
        savedOutputHandle = outputHandle;
      }
      if (outputHandle && await hasPermission(outputHandle, "readwrite")) {
        outputBaseFolder = outputHandle;
        ui.outputFolderLabel.textContent = "輸出資料夾：" + outputHandle.name;
      } else if (outputHandle) {
        ui.outputFolderLabel.textContent = "輸出資料夾：" + outputHandle.name;
      }
    } catch {
      ui.statusLabel.textContent = "設定檔讀取失敗，請重新選取資料夾";
    }
  }

  async function hasPermission(handle, mode) {
    if (!handle || typeof handle.queryPermission !== "function") {
      return false;
    }
    const permission = await handle.queryPermission({ mode });
    return permission === "granted";
  }

  async function ensurePermission(handle, mode) {
    if (!handle) {
      return false;
    }
    if (typeof handle.queryPermission !== "function" || typeof handle.requestPermission !== "function") {
      return true;
    }
    if (await handle.queryPermission({ mode }) === "granted") {
      return true;
    }
    return await handle.requestPermission({ mode }) === "granted";
  }

  function openSettingsDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open("FSG2300ProcessLogConverter", 1);
      request.onupgradeneeded = () => {
        request.result.createObjectStore("settings");
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function saveDirectoryHandle(key, handle) {
    if (!("indexedDB" in window)) {
      return;
    }
    try {
      const db = await openSettingsDb();
      await transactionRequest(db, "readwrite", (store) => store.put(handle, key));
      db.close();
    } catch {
      ui.statusLabel.textContent = "設定檔儲存失敗";
    }
  }

  async function getDirectoryHandle(key) {
    if (!("indexedDB" in window)) {
      return null;
    }
    const db = await openSettingsDb();
    try {
      return await transactionRequest(db, "readonly", (store) => store.get(key));
    } finally {
      db.close();
    }
  }

  async function clearDirectoryHandle(key) {
    if (!("indexedDB" in window)) {
      return;
    }
    try {
      const db = await openSettingsDb();
      await transactionRequest(db, "readwrite", (store) => store.delete(key));
      db.close();
    } catch {
      ui.statusLabel.textContent = "設定檔清除失敗，請重新選取資料夾";
    }
  }

  function transactionRequest(db, mode, operation) {
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("settings", mode);
      const request = operation(transaction.objectStore("settings"));
      let result;
      request.onsuccess = () => {
        result = request.result;
      };
      request.onerror = () => reject(request.error);
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  }

  async function buildSourceFromDirectoryHandle(rootHandle) {
    const files = [];
    await walkDirectoryHandle(rootHandle, "", files);
    return makeSource(rootHandle.name, files);
  }

  async function walkDirectoryHandle(directoryHandle, dirPath, files) {
    for await (const [name, handle] of directoryHandle.entries()) {
      const relativePath = dirPath ? dirPath + "/" + name : name;
      if (handle.kind === "directory") {
        await walkDirectoryHandle(handle, relativePath, files);
      } else if (handle.kind === "file") {
        files.push({
          name,
          relativePath,
          dirPath,
          getText: async () => readFileText(await handle.getFile())
        });
      }
    }
  }

  async function buildSourceFromFileList(fileList) {
    const firstPath = fileList[0].webkitRelativePath || fileList[0].name;
    const rootName = firstPath.includes("/") ? firstPath.split("/")[0] : "SelectedFolder";
    const files = fileList.map((file) => {
      const browserPath = file.webkitRelativePath || file.name;
      const relativePath = browserPath.startsWith(rootName + "/")
        ? browserPath.slice(rootName.length + 1)
        : browserPath;
      return {
        name: file.name,
        relativePath,
        dirPath: dirname(relativePath),
        getText: async () => readFileText(file)
      };
    });
    return makeSource(rootName, files);
  }

  function makeSource(rootName, files) {
    const byPath = new Map();
    const byDir = new Map();
    for (const file of files) {
      const normalizedPath = normalizePath(file.relativePath);
      const dirPath = dirname(normalizedPath);
      const entry = {
        name: basename(normalizedPath),
        relativePath: normalizedPath,
        dirPath,
        getText: file.getText
      };
      byPath.set(normalizedPath.toLowerCase(), entry);
      if (!byDir.has(dirPath.toLowerCase())) {
        byDir.set(dirPath.toLowerCase(), []);
      }
      byDir.get(dirPath.toLowerCase()).push(entry);
    }
    for (const list of byDir.values()) {
      list.sort(comparePath);
    }
    return {
      displayName: rootName,
      rootName,
      files,
      byPath,
      byDir
    };
  }

  async function readFileText(file) {
    const bytes = await file.arrayBuffer();
    const view = new Uint8Array(bytes);
    if (view.length >= 3 && view[0] === 0xef && view[1] === 0xbb && view[2] === 0xbf) {
      return new TextDecoder("utf-8").decode(view.subarray(3));
    }
    return new TextDecoder("big5").decode(view);
  }

  function hasProcessLogFolder(source) {
    return Boolean(source) && getProcessLogFolders(source).length > 0;
  }

  function getProcessLogFolders(source) {
    const folders = [];
    for (const file of source.byPath.values()) {
      if (isProcessLogFile(file)) {
        folders.push(file.dirPath);
      }
    }
    return Array.from(new Set(folders)).sort(compareText);
  }

  function isProcessLogFile(file) {
    return isProcessLogFileName(file && (file.relativePath || file.name));
  }

  function isProcessLogFileName(name) {
    return normalizeFileNameForMatch(basename(name)) === normalizeFileNameForMatch(PROCESS_LOG_FILE_NAME);
  }

  function normalizeFileNameForMatch(name) {
    return String(name || "")
      .normalize("NFKC")
      .replace(/[\u200b-\u200f\ufeff]/g, "")
      .trim()
      .toLowerCase();
  }

  function buildProcessLogNotFoundMessage(source, prefix) {
    const files = source ? Array.from(source.byPath.values()).sort(comparePath) : [];
    const preview = files.slice(0, 5).map((file) => file.relativePath).join("、");
    const iniPreview = files
      .filter((file) => normalizeFileNameForMatch(file.relativePath).includes(".ini"))
      .slice(0, 8)
      .map((file) => file.relativePath)
      .join("、");
    const suffix = files.length > 0
      ? "；實際掃到 " + files.length + " 個檔案：" + preview
      : "；實際掃到 0 個檔案";
    return prefix + "找不到 ProcessLog.ini，請選 Desktop\\Ini 底下的原始來源資料夾" + suffix
      + (iniPreview ? "；掃到的 ini：" + iniPreview : "");
  }

  function getWaferFiles(source, folderPath) {
    const list = source.byDir.get(folderPath.toLowerCase()) || [];
    return list
      .filter((file) => file.name.toLowerCase().endsWith(".ini"))
      .filter((file) => !isProcessLogFile(file))
      .sort(comparePath);
  }

  async function convertFolder(source, progress, isCancelled) {
    const converter = new ProcessLogConverter(source);
    const logFolders = getProcessLogFolders(source);
    const workItems = logFolders.map((folderPath) => ({
      folderPath,
      waferFiles: getWaferFiles(source, folderPath)
    }));

    if (workItems.length === 0) {
      throw new Error("No ProcessLog.ini was found in the selected folder or its subfolders.");
    }

    let totalSteps = 0;
    for (const item of workItems) {
      totalSteps += (item.waferFiles.length * 2) + 5;
    }

    let completed = 0;
    let outputFileCount = 0;
    const outputs = [];
    report(progress, completed, totalSteps, "找到 " + workItems.length + " 個 ProcessLog 資料夾");

    for (const item of workItems) {
      if (isCancelled()) {
        throwIfCancellationRequested();
      }
      const outputFolder = getOutputFolder(source, item.folderPath);
      outputFileCount += await converter.convertSingleFolder(item.folderPath, outputFolder, item.waferFiles, outputs, (message) => {
        completed++;
        report(progress, completed, totalSteps, message);
      }, (message) => {
        report(progress, completed, totalSteps, message);
      }, isCancelled);
      converter.clearCache();
      await nextFrame();
    }

    report(progress, totalSteps, totalSteps, "全部轉換完成");
    return {
      outputs,
      outputFileCount,
      folderCount: workItems.length
    };
  }

  function getOutputFolder(source, logFolder) {
    if (!logFolder) {
      return source.rootName;
    }
    return normalizePath(logFolder);
  }

  function report(progress, completed, total, message) {
    if (!progress) {
      return;
    }
    const percent = total <= 0 ? 0 : Math.round((completed * 100) / total);
    progress(Math.max(0, Math.min(100, percent)), message);
  }

  class ProcessLogConverter {
    constructor(source) {
      this.source = source;
      this.cache = new Map();
    }

    clearCache() {
      this.cache.clear();
    }

    removeFromCache(filePath) {
      this.cache.delete(filePath.toLowerCase());
    }

    async convertSingleFolder(sourceFolder, outputFolder, waferFiles, outputs, stepCompleted, stepProgress, isCancelled) {
      throwIfCancellationRequested();
      const processLogPath = joinPath(sourceFolder, PROCESS_LOG_FILE_NAME);
      if (!this.source.byPath.has(processLogPath.toLowerCase())) {
        throw new Error("ProcessLog.ini not found.");
      }

      const folderName = sourceFolder ? basename(sourceFolder) : this.source.rootName;
      const processLogResult = await this.convertProcessLog(processLogPath);
      outputs.push(textOutput(joinPath(outputFolder, "ProcessLog.csv"), processLogResult.content));
      stepCompleted("已轉換 " + folderName + "\\ProcessLog.csv");

      const batchRows = [];
      const grdDistData = { g1Rows: [], g2Rows: [] };
      for (const waferFile of waferFiles) {
        if (isCancelled()) {
          throwIfCancellationRequested();
        }
        const nameOnly = basenameWithoutExtension(waferFile.relativePath);
        outputs.push(textOutput(joinPath(outputFolder, nameOnly + ".csv"), await this.convertWaferLog(waferFile.relativePath, processLogResult.isCancelGrd)));
        const batchRow = await this.buildBatchRow(waferFile.relativePath, processLogResult.isCancelGrd);
        batchRows.push(batchRow);
        await this.addGrdDistCollectRows(grdDistData, waferFile.relativePath, batchRow.waferNumber, processLogResult.isCancelGrd);
        stepCompleted("已轉換 " + folderName + "\\" + waferFile.name);
        await nextFrame();
      }

      outputs.push(textOutput(joinPath(outputFolder, "WaferBatchLog.csv"), this.writeWaferBatchLog(batchRows, processLogResult.isCancelGrd)));
      stepCompleted("已產生 " + folderName + "\\WaferBatchLog.csv");

      let wroteGrdDistCollect = false;
      if (grdDistData.g1Rows.length > 0 || grdDistData.g2Rows.length > 0) {
        stepProgress("正在輸出 " + folderName + "\\GrdDistCollect.csv");
        outputs.push(textOutput(joinPath(outputFolder, "GrdDistCollect.csv"), this.writeGrdDistCollect(grdDistData)));
        wroteGrdDistCollect = true;
        stepCompleted("已產生 " + folderName + "\\GrdDistCollect.csv");
      }

      stepProgress("正在整理 " + folderName + "\\S1LoadingCollect");
      const s1LoadingData = await this.buildS1LoadingMatchData(waferFiles, processLogResult.isCancelGrd, folderName, stepCompleted, isCancelled);
      outputs.push(textOutput(joinPath(outputFolder, "S1LoadingCollect.csv"), this.writeS1LoadingMatch(s1LoadingData, processLogResult.isCancelGrd, folderName, stepProgress, isCancelled)));
      stepCompleted("已產生 " + folderName + "\\S1LoadingCollect.csv");

      outputs.push(binaryOutput(joinPath(outputFolder, "S1LoadingCollect.xlsx"), this.writeS1LoadingMatchWorkbook(s1LoadingData, processLogResult.isCancelGrd, folderName, stepProgress, isCancelled)));
      stepCompleted("已產生 " + folderName + "\\S1LoadingCollect.xlsx");

      return waferFiles.length + 4 + (wroteGrdDistCollect ? 1 : 0);
    }

    async convertProcessLog(processLogPath) {
      const lines = [];
      const isCancelGrd = [false, false];

      lines.push("[Info]");
      lines.push("LotStartTime,LotEndTime,LotTotalRunTime,LotWaferAmount");
      lines.push(await this.read(processLogPath, "Info", "RunData", PROCESS_LOG_RUN_DATA_INIT));
      lines.push("MachineModel,SerialNumber,LotID,OpID");
      lines.push(await this.read(processLogPath, "Info", "ID", PROCESS_LOG_ID_INIT));

      lines.push("[WheelData]");
      lines.push(",ID,TotalThick(um),SurplusThick(um),TotalWear(um),pcs_OK,pcs_NG,DressCount,GaugeSetCount,AvgWear(um),pcs_Grind");
      lines.push("G1Wheel," + await this.read(processLogPath, "Info", "G1WheelID", "") + "," + await this.read(processLogPath, "Info", "G1WheelData", ""));
      lines.push("G2Wheel," + await this.read(processLogPath, "Info", "G2WheelID", "") + "," + await this.read(processLogPath, "Info", "G2WheelData", ""));

      lines.push("");
      lines.push("[Recipe]");
      lines.push("Number,Name");
      lines.push(await this.read(processLogPath, "HMI", "Number", "") + "," + await this.read(processLogPath, "HMI", "Name", ""));
      lines.push("Reserved,WaferSize,Reserved,TotalPart,G1FinalThickness(um),G2FinalThickness(um),Reserved,CancelG1,CancelG2,Reserved,GrdMode,CancelChuck1,CancelChuck2,CancelChuck3,Reserved,UseRemoveValue");
      const nc3Step0 = await this.read(processLogPath, "NC3", "Step0", RECIPE_INIT_STEP0_NC3);
      lines.push(nc3Step0);

      const step0Values = splitCsv(nc3Step0);
      isCancelGrd[0] = valueAt(step0Values, RECIPE_IDX_NC3_DIS_G1) === "1";
      isCancelGrd[1] = valueAt(step0Values, RECIPE_IDX_NC3_DIS_G2) === "1";

      await this.writeRecipe(lines, processLogPath, "NC1", "[G1_Recipe]");
      await this.writeRecipe(lines, processLogPath, "NC2", "[G2_Recipe]");
      await this.writeParameterBackup(lines, processLogPath);

      return { content: finishLines(lines), isCancelGrd };
    }

    async writeRecipe(lines, iniPath, section, title) {
      lines.push("");
      lines.push(title);
      lines.push(",Reserved,WaferSize,Reserved,IncomeThickness(um),S1direction,S2direction,TotalStep,FinalTarget(um),AirCutDist(um),Reserve,Reserve,SparkOutTime(sec),Reserve,SparkOutS2speed(RPM),EscapeCutDist(um),EscapeCutFeed(um/sec),EscapeCutS2speed(RPM)");
      lines.push("," + await this.read(iniPath, section, "Step0", RECIPE_INIT_STEP0_NC12));
      lines.push(",S1speed(RPM),S2speed(RPM),Zfeedrate(um/sec),TargetThickness(um),OnlineSelect,Compensation(um),SparkOut(Sec),UpDistance(um),UpSpeed(um/sec),RemoveValue(um)");
      for (let i = 0; i < RECIPE_LEVEL_STEP_NUM; i++) {
        const step = "Step" + (i + 1);
        lines.push(step + "," + await this.read(iniPath, section, step, RECIPE_INIT_STEP1_10_NC12));
      }
    }

    async writeParameterBackup(lines, iniPath) {
      lines.push("");
      lines.push("[ParameterBackup]");
      lines.push(buildParameterHeader(8000, 360));
      for (let i = 0; i < 3; i++) {
        lines.push("NC" + (i + 1) + "," + await this.read(iniPath, "Info", "Param8000_NC" + (i + 1), ""));
      }

      lines.push("");
      lines.push(buildParameterHeader(8360, 360));
      for (let i = 0; i < 3; i++) {
        lines.push("NC" + (i + 1) + "," + await this.read(iniPath, "Info", "Param8360_NC" + (i + 1), ""));
      }

      lines.push("");
      lines.push(buildParameterHeader(0, 360));
      lines.push("HMI," + await this.read(iniPath, "Info", "Param_HMI", ""));
    }

    async convertWaferLog(waferPath, isCancelGrd) {
      await this.getIni(waferPath);
      const lines = [];
      lines.push("[WaferInfo]");
      lines.push("CassetteID,CassetteOpID");
      lines.push(this.readCached(waferPath, "BlockInfo", "ID", ","));
      lines.push("");

      lines.push("ChuckNumber,G1GrdSts,G2GrdSts,Reserved,InclineNumSV,G1_V1MchCoordSV(mm),G1_V2MchCoordSV(mm),G2_V1MchCoordSV(mm),G2_V2MchCoordSV(mm)");
      lines.push(convertGrdStatus(this.readCached(waferPath, "BlockInfo", "Data", WAFER_LOG_DATA_INIT)) + this.readCached(waferPath, "BlockInfo", "InclineSV", WAFER_LOG_INCLINE_SV_INIT));
      lines.push("");
      lines.push("CassetteNumber,SlotNumber,WaferNumber,RobotGetTime,RobotPutTime,TotalRunTime");
      lines.push(this.readCached(waferPath, "BlockInfo", "Cassette", WAFER_LOG_CASS_INIT) + "," + this.readCached(waferPath, "BlockInfo", "RunTime", WAFER_LOG_RUN_TIME_INIT));
      lines.push("");
      lines.push(",Income(um),TargetSV(um),TargetPV(um),LastWaferMax(um),LastWaferMin(um),PeakS1(%),PeakS2(%),PeakZ(%),ZautoUpCount,AlarmZupCount,S1WarningSV(%),S1AlarmSV(%),StartTime,EndTime,RunTime");

      for (let i = 0; i < SERVER_IDX_NC3; i++) {
        if (!isCancelGrd[i]) {
          lines.push(this.buildWaferSummaryLine(waferPath, i));
        }
      }

      lines.push("");
      const hasExtendedAirCutData = this.hasExtendedAirCutData(waferPath, isCancelGrd);
      const airCutDataLen = hasExtendedAirCutData ? AI_AIR_CUT_EXTENDED_DATA_LEN : AI_AIR_CUT_BASE_DATA_LEN;
      let airCutHeader = ",S1NullLoading(%),S1CompLoading(%),S1ReachZabs(mm),S1ReachLoading(%),S1ReachThickness(um)";
      if (hasExtendedAirCutData) {
        airCutHeader += ",S1ReachZmch(mm),StartZmch-EndZmch(um),GrdDist-RemoveValue(um),S1ReachZabs-WaferIncomeMax(um)";
      }
      lines.push(airCutHeader);
      for (let i = 0; i < SERVER_IDX_NC3; i++) {
        if (!isCancelGrd[i]) {
          const key = i === 0 ? "G1AirCutS1LoadingRec" : "G2AirCutS1LoadingRec";
          lines.push("G" + (i + 1) + "," + toFixedCsv(this.readCached(waferPath, "BlockInfo", key, AI_AIR_CUT_DATA_INIT), airCutDataLen));
        }
      }

      for (let i = 0; i < SERVER_IDX_NC3; i++) {
        if (!isCancelGrd[i]) {
          this.writeWaferRecipeWheelAndSampling(lines, waferPath, i);
        }
      }
      return finishLines(lines);
    }

    buildWaferSummaryLine(waferPath, grindIndex) {
      const gnum = "G" + (grindIndex + 1);
      const runStepLen = this.getRunStepLenCached(waferPath, gnum + "_Recipe");
      const step0 = splitCsv(this.readCached(waferPath, gnum + "_Recipe", "Step0", RECIPE_INIT_STEP0_NC12));
      const targetSv = valueAt(step0, RECIPE_IDX_NC1_TARGET_THICK);

      const before = splitCsv(this.readCached(waferPath, gnum + "_Wheel", "BeforeGrd", WHEEL_INIT_STEP0_10));
      const income = valueAt(before, WHEEL_BEFORE_GRD_INCOME_THICK);
      const lastWaferMax = valueAt(before, WHEEL_BEFORE_GRD_LAST_WAFER_MAX);
      const lastWaferMin = valueAt(before, WHEEL_BEFORE_GRD_LAST_WAFER_MIN);

      let targetPv = "";
      if (runStepLen > 0) {
        const after = splitCsv(this.readCached(waferPath, gnum + "_Wheel", "AfterGrdStep" + runStepLen, WHEEL_INIT_STEP0_10));
        targetPv = valueAt(after, WHEEL_AFTER_GRD_TARGET_CONFIRM);
      }

      const peakKey = grindIndex === 0 ? "G1PeakLoading" : "G2PeakLoading";
      const timeKey = grindIndex === 0 ? "G1_Time" : "G2_Time";
      return gnum + "," + income + "," + targetSv + "," + targetPv + "," + lastWaferMax + "," + lastWaferMin + ","
        + this.readCached(waferPath, "BlockInfo", peakKey, WAFER_LOG_PEAK_INIT) + ","
        + this.readCached(waferPath, "BlockInfo", timeKey, WAFER_LOG_GRD_TIME_INIT) + ",";
    }

    writeWaferRecipeWheelAndSampling(lines, waferPath, grindIndex) {
      const gnum = "G" + (grindIndex + 1);
      const recipeSection = gnum + "_Recipe";
      const runStepLen = this.getRunStepLenCached(waferPath, recipeSection);

      lines.push("");
      lines.push("[" + recipeSection + "]");
      lines.push(",Reserved,WaferSize,Reserved,IncomeThickness(um),S1direction,S2direction,TotalStep,FinalTarget(um),AirCutDist(um),Reserve,Reserve,SparkOutTime(sec),Reserve,SparkOutS2speed(RPM),EscapeCutDist(um),EscapeCutFeed(um/sec),EscapeCutS2speed(RPM)");
      lines.push("," + this.readCached(waferPath, recipeSection, "Step0", RECIPE_INIT_STEP0_NC12));
      lines.push(",S1speed(RPM),S2speed(RPM),Zfeedrate(um/sec),TargetThickness(um),OnlineSelect,Compensation(um),SparkOut(Sec),UpDistance(um),UpSpeed(um/sec),RemoveValue(um)");
      for (let j = 0; j < runStepLen; j++) {
        const step = "Step" + (j + 1);
        lines.push(step + "," + this.readCached(waferPath, recipeSection, step, RECIPE_INIT_STEP1_10_NC12));
      }

      const wheelSection = gnum + "_Wheel";
      lines.push("");
      lines.push("[" + wheelSection + "]");
      lines.push(",ChuckNumber,WheelTotal(um),WheelSurplus(um),WearTotal(um),WearSingle(um),WaferMax(um),WaferMin(um),ChuckMax(um),ChuckMin(um),Z_G92(mm),IncomeThickness(um),V1MchCoordPV(mm),V2MchCoordPV(mm),WaterFlowSV(L/min),P8606,P8607,P8608,StartStep,AirCutPV(um)");
      lines.push("BeforeGrd," + removeBeforeGrdLastWaferColumns(this.readCached(waferPath, wheelSection, "BeforeGrd", WHEEL_INIT_STEP0_10)));
      lines.push(",Step,ChuckMax(um),ChuckMin(um),WearSingle(um),WearTotal(um),Zabs(mm),Zmach(mm),WaferMax(um),WaferMin(um),WearRAW(NoBlock)(um),WearRate(%),TargetConfirm(um),SurplusWheel(um)");
      for (let j = 0; j < runStepLen; j++) {
        const step = "AfterGrdStep" + (j + 1);
        lines.push("AfterGrd," + this.readCached(waferPath, wheelSection, step, WHEEL_INIT_STEP0_10));
      }

      lines.push("");
      for (let j = 0; j < runStepLen; j++) {
        const sampleSection = gnum + "_Step" + (j + 1);
        const sampleLen = toInt(this.readCached(waferPath, sampleSection, "0", "0"));
        if (sampleLen <= 0) {
          continue;
        }
        lines.push("[" + sampleSection + "]");
        lines.push("Time,WaferMax(um),WaferMin(um),WaferOnline(um),Zmch(mm),Zabs(mm),Zabs-WaferOnline(mm),S1Loading(%),S2Loading(%),ZLoading(%),WaterFlow(L/min),S1speed(RPM),S2speed(RPM),Zfeedrate(um/sec),Aloading(%),SpTemp(C),WaferMinLatch(um),WaferAVG(um),GrdVacIn(kPa),GrdVacOut(kPa),BruVacIn(kPa),BruVacOut(kPa),CGWear(um),CGLevelComp(um),S1OutputPower(KW)");
        for (let k = 0; k < sampleLen; k++) {
          lines.push(removeT2Column(this.readCached(waferPath, sampleSection, String(k + 1), "")));
        }
      }
    }

    async buildBatchRow(waferPath, isCancelGrd) {
      await this.getIni(waferPath);
      const row = { waferNumber: "", grindData: ["", ""] };
      const cassette = splitCsv(this.readCached(waferPath, "BlockInfo", "Cassette", WAFER_LOG_CASS_INIT));
      row.waferNumber = valueAt(cassette, 2);
      if (!row.waferNumber) {
        row.waferNumber = basenameWithoutExtension(waferPath);
      }
      for (let i = 0; i < SERVER_IDX_NC3; i++) {
        row.grindData[i] = isCancelGrd[i] ? "" : this.buildBatchGrindData(waferPath, i);
      }
      return row;
    }

    buildBatchGrindData(waferPath, grindIndex) {
      const gnum = "G" + (grindIndex + 1);
      const runStepLen = this.getRunStepLenCached(waferPath, gnum + "_Recipe");
      const blockData = splitCsv(this.readCached(waferPath, "BlockInfo", "Data", WAFER_LOG_DATA_INIT));
      const chuckNum = valueAt(blockData, 0);
      const result = valueAt(blockData, 1 + grindIndex) === "0" ? "OK" : "NG";

      const recipe = splitCsv(this.readCached(waferPath, gnum + "_Recipe", "Step0", RECIPE_INIT_STEP0_NC12));
      const targetSv = valueAt(recipe, RECIPE_IDX_NC1_TARGET_THICK);

      const before = splitCsv(this.readCached(waferPath, gnum + "_Wheel", "BeforeGrd", WHEEL_INIT_STEP0_10));
      const t2Position = valueAt(before, WHEEL_BEFORE_GRD_CHUCK_MAX);
      const income = valueAt(before, WHEEL_BEFORE_GRD_INCOME_THICK);
      const lastWaferMin = valueAt(before, WHEEL_BEFORE_GRD_LAST_WAFER_MIN);

      let targetPv = "";
      let totalWear = "";
      let wheel = "";
      let singleWear = 0;
      if (runStepLen > 0) {
        const after = splitCsv(this.readCached(waferPath, gnum + "_Wheel", "AfterGrdStep" + runStepLen, WHEEL_INIT_STEP0_10));
        targetPv = valueAt(after, WHEEL_AFTER_GRD_TARGET_CONFIRM);
        totalWear = valueAt(after, WHEEL_AFTER_GRD_WEAR_TOTAL);
        wheel = valueAt(after, WHEEL_AFTER_GRD_WHEEL_THICK);
      }

      for (let j = 0; j < runStepLen; j++) {
        const after = splitCsv(this.readCached(waferPath, gnum + "_Wheel", "AfterGrdStep" + (j + 1), WHEEL_INIT_STEP0_10));
        singleWear += toDouble(valueAt(after, WHEEL_AFTER_GRD_WEAR_SINGLE));
      }

      const peak = splitCsv(this.readCached(waferPath, "BlockInfo", grindIndex === 0 ? "G1PeakLoading" : "G2PeakLoading", WAFER_LOG_PEAK_INIT));
      const time = this.readCached(waferPath, "BlockInfo", grindIndex === 0 ? "G1_Time" : "G2_Time", WAFER_LOG_GRD_TIME_INIT);
      const id = splitCsv(this.readCached(waferPath, "BlockInfo", "ID", ","));

      return [
        chuckNum,
        t2Position,
        income,
        targetSv,
        targetPv,
        lastWaferMin,
        valueAt(peak, 0),
        valueAt(peak, 1),
        valueAt(peak, 2),
        roundOne(singleWear),
        totalWear,
        wheel,
        result,
        time,
        valueAt(id, 0),
        valueAt(id, 1)
      ].join(",");
    }

    writeWaferBatchLog(rows, isCancelGrd) {
      const lines = [];
      lines.push("G1_WaferNum,G1_ChuckNum,G1_T2(um),G1_Income(um),G1_TarSV(um),G1_TarPV(um),G1_LaMin(um),G1_S1Peak(%),G1_S2Peak(%),G1_ZPeak(%),G1_SWear(um),G1_TWear(um),G1_Wheel(um),G1_Result,G1_StartTime,G1_EndTime,G1_RunTime,G1_CassetteID,G1_CassetteOpID,G2_WaferNum,G2_ChuckNum,G2_T2(um),G2_Income(um),G2_TarSV(um),G2_TarPV(um),G2_LaMin(um),G2_S1Peak(%),G2_S2Peak(%),G2_ZPeak(%),G2_SWear(um),G2_TWear(um),G2_Wheel(um),G2_Result,G2_StartTime,G2_EndTime,G2_RunTime,G2_CassetteID,G2_CassetteOpID");
      for (const row of rows) {
        const g1 = isCancelGrd[0] ? emptyBatchSide(row.waferNumber) : row.waferNumber + "," + row.grindData[0];
        const g2 = isCancelGrd[1] ? emptyBatchSide(row.waferNumber) : row.waferNumber + "," + row.grindData[1];
        lines.push(g1 + "," + g2);
      }
      return finishLines(lines);
    }

    async addGrdDistCollectRows(data, waferPath, waferNumber, isCancelGrd) {
      await this.getIni(waferPath);
      const blockData = splitCsv(this.readCached(waferPath, "BlockInfo", "Data", WAFER_LOG_DATA_INIT));
      const chuckNumber = valueAt(blockData, 0);
      const waferId = waferNumber || basenameWithoutExtension(waferPath);

      for (let grindIndex = 0; grindIndex < SERVER_IDX_NC3; grindIndex++) {
        if (isCancelGrd[grindIndex]) {
          continue;
        }

        const key = grindIndex === 0 ? "G1AirCutS1LoadingRec" : "G2AirCutS1LoadingRec";
        const values = splitCsv(this.readCached(waferPath, "BlockInfo", key, AI_AIR_CUT_DATA_INIT));
        if (!hasMeaningfulExtendedAirCutValues(values)) {
          continue;
        }

        const row = {
          chuckNumber,
          waferId,
          s1ReachZmch: valueAt(values, 5),
          startZmchEndZmch: valueAt(values, 6),
          grdDistRemoveValue: valueAt(values, 7),
          s1ReachZabsWaferIncomeMax: valueAt(values, 8)
        };

        if (grindIndex === 0) {
          data.g1Rows.push(row);
        } else {
          data.g2Rows.push(row);
        }
      }
    }

    writeGrdDistCollect(data) {
      const lines = [];
      let wroteBlock = false;
      if (data.g1Rows.length > 0) {
        writeGrdDistCollectBlock(lines, "G1_Chuck", data.g1Rows);
        wroteBlock = true;
      }
      if (data.g2Rows.length > 0) {
        if (wroteBlock) {
          lines.push("");
        }
        writeGrdDistCollectBlock(lines, "G2_Chuck", data.g2Rows);
      }
      return finishLines(lines);
    }

    async buildS1LoadingMatchData(waferFiles, isCancelGrd, folderName, stepCompleted, isCancelled) {
      const data = { g1Columns: [], g2Columns: [] };
      for (let waferIndex = 0; waferIndex < waferFiles.length; waferIndex++) {
        if (isCancelled()) {
          throwIfCancellationRequested();
        }
        const waferFile = waferFiles[waferIndex];
        const columns = await this.buildS1LoadingColumns(waferFile.relativePath, isCancelGrd);
        if (!isCancelGrd[0]) {
          data.g1Columns.push(columns[0]);
        }
        if (!isCancelGrd[1]) {
          data.g2Columns.push(columns[1]);
        }
        stepCompleted("正在整理 " + folderName + "\\S1LoadingCollect " + basenameWithoutExtension(waferFile.relativePath) + " (" + (waferIndex + 1) + "/" + waferFiles.length + ")");
        await nextFrame();
      }
      return data;
    }

    writeS1LoadingMatch(data, isCancelGrd, folderName, stepProgress, isCancelled) {
      const lines = [];
      let wroteBlock = false;
      for (let grindIndex = 0; grindIndex < SERVER_IDX_NC3; grindIndex++) {
        if (isCancelGrd[grindIndex]) {
          continue;
        }
        if (wroteBlock) {
          lines.push("");
        }
        const grindName = "G" + (grindIndex + 1);
        const columns = grindIndex === 0 ? data.g1Columns : data.g2Columns;
        lines.push(grindName + "_Chuck," + columns.map((column) => column.chuckNumber).join(","));
        lines.push("Wafer_ID," + columns.map((column) => column.waferId).join(","));
        const maxRows = columns.length === 0 ? 0 : Math.max(...columns.map((column) => column.values.length));
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
          if ((rowIndex % 100) === 0) {
            if (isCancelled()) {
              throwIfCancellationRequested();
            }
            stepProgress("正在輸出 " + folderName + "\\S1LoadingCollect.csv " + grindName + " row " + (rowIndex + 1) + "/" + maxRows);
          }
          const values = columns.map((column) => rowIndex < column.values.length ? column.values[rowIndex] : "");
          lines.push("," + values.join(","));
        }
        wroteBlock = true;
      }
      return finishLines(lines);
    }

    writeS1LoadingMatchWorkbook(data, isCancelGrd, folderName, stepProgress, isCancelled) {
      const entries = [
        {
          path: "[Content_Types].xml",
          data: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
            + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            + "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            + "</Types>"
        },
        {
          path: "_rels/.rels",
          data: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "</Relationships>"
        },
        {
          path: "xl/workbook.xml",
          data: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<sheets><sheet name=\"G1\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"G2\" sheetId=\"2\" r:id=\"rId2\"/></sheets>"
            + "</workbook>"
        },
        {
          path: "xl/_rels/workbook.xml.rels",
          data: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
            + "</Relationships>"
        },
        {
          path: "xl/worksheets/sheet1.xml",
          data: this.buildS1LoadingWorksheetXml("G1", "G1_Chuck", data.g1Columns, folderName, stepProgress, isCancelled)
        },
        {
          path: "xl/worksheets/sheet2.xml",
          data: this.buildS1LoadingWorksheetXml("G2", "G2_Chuck", data.g2Columns, folderName, stepProgress, isCancelled)
        }
      ];
      return createZipBlob(entries);
    }

    buildS1LoadingWorksheetXml(sheetName, chuckHeader, columns, folderName, stepProgress, isCancelled) {
      const excelColumns = buildExcelColumns(columns.length + 1);
      const parts = [];
      parts.push("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      parts.push("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
      parts.push("<row r=\"1\">");
      parts.push(inlineStringCell(excelColumns[0], 1, chuckHeader));
      for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
        parts.push(inlineStringCell(excelColumns[columnIndex + 1], 1, columns[columnIndex].chuckNumber));
      }
      parts.push("</row>");
      parts.push("<row r=\"2\">");
      parts.push(inlineStringCell(excelColumns[0], 2, "Wafer_ID"));
      for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
        parts.push(inlineStringCell(excelColumns[columnIndex + 1], 2, columns[columnIndex].waferId));
      }
      parts.push("</row>");

      const maxRows = columns.length === 0 ? 0 : Math.max(...columns.map((column) => column.values.length));
      for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
        if ((rowIndex % 100) === 0) {
          if (isCancelled()) {
            throwIfCancellationRequested();
          }
          stepProgress("正在輸出 " + folderName + "\\S1LoadingCollect.xlsx " + sheetName + " row " + (rowIndex + 1) + "/" + maxRows);
        }
        const excelRow = rowIndex + 3;
        parts.push("<row r=\"" + excelRow + "\">");
        parts.push(inlineStringCell(excelColumns[0], excelRow, ""));
        for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
          const column = columns[columnIndex];
          parts.push(inlineStringCell(excelColumns[columnIndex + 1], excelRow, rowIndex < column.values.length ? column.values[rowIndex] : ""));
        }
        parts.push("</row>");
      }
      parts.push("</sheetData></worksheet>");
      return parts.join("");
    }

    async buildS1LoadingColumns(waferPath, isCancelGrd) {
      await this.getIni(waferPath);
      const blockData = splitCsv(this.readCached(waferPath, "BlockInfo", "Data", WAFER_LOG_DATA_INIT));
      const waferId = basenameWithoutExtension(waferPath);
      const chuckNumber = valueAt(blockData, 0);
      const columns = [
        { waferId, chuckNumber, values: [] },
        { waferId, chuckNumber, values: [] }
      ];

      for (let grindIndex = 0; grindIndex < SERVER_IDX_NC3; grindIndex++) {
        const column = columns[grindIndex];
        if (isCancelGrd[grindIndex]) {
          continue;
        }
        const gnum = "G" + (grindIndex + 1);
        const runStepLen = this.getRunStepLenCached(waferPath, gnum + "_Recipe");
        for (let stepIndex = 0; stepIndex < runStepLen; stepIndex++) {
          const sampleSection = gnum + "_Step" + (stepIndex + 1);
          const sampleLen = toInt(this.readCached(waferPath, sampleSection, "0", "0"));
          for (let sampleIndex = 0; sampleIndex < sampleLen; sampleIndex++) {
            const sampleValues = splitCsv(this.readCached(waferPath, sampleSection, String(sampleIndex + 1), ""));
            column.values.push(valueAt(sampleValues, SAMP_IDX_S1_LOADING));
          }
        }
      }
      return columns;
    }

    hasExtendedAirCutData(waferPath, isCancelGrd) {
      for (let i = 0; i < SERVER_IDX_NC3; i++) {
        if (isCancelGrd[i]) {
          continue;
        }
        const key = i === 0 ? "G1AirCutS1LoadingRec" : "G2AirCutS1LoadingRec";
        const values = splitCsv(this.readCached(waferPath, "BlockInfo", key, AI_AIR_CUT_DATA_INIT));
        if (hasMeaningfulExtendedAirCutValues(values)) {
          return true;
        }
      }
      return false;
    }

    async getRunStepLen(iniPath, recipeSection) {
      const step0 = splitCsv(await this.read(iniPath, recipeSection, "Step0", RECIPE_INIT_STEP0_NC12));
      return Math.max(0, toInt(valueAt(step0, RECIPE_IDX_NC1_TOTAL_STEP)));
    }

    getRunStepLenCached(iniPath, recipeSection) {
      const step0 = splitCsv(this.readCached(iniPath, recipeSection, "Step0", RECIPE_INIT_STEP0_NC12));
      return Math.max(0, toInt(valueAt(step0, RECIPE_IDX_NC1_TOTAL_STEP)));
    }

    async read(filePath, section, key, defaultValue) {
      const ini = await this.getIni(filePath);
      const value = ini.get(section, key);
      return limitProfileString(value === undefined ? defaultValue : value);
    }

    readCached(filePath, section, key, defaultValue) {
      const ini = this.cache.get(filePath.toLowerCase());
      if (!ini) {
        throw new Error("INI cache was not loaded: " + filePath);
      }
      const value = ini.get(section, key);
      return limitProfileString(value === undefined ? defaultValue : value);
    }

    async getIni(filePath) {
      const key = filePath.toLowerCase();
      if (this.cache.has(key)) {
        return this.cache.get(key);
      }
      const file = this.source.byPath.get(key);
      if (!file) {
        return new IniFile("");
      }
      const text = await file.getText();
      const ini = new IniFile(text);
      this.cache.set(key, ini);
      return ini;
    }
  }

  class IniFile {
    constructor(text) {
      this.sections = new Map();
      this.parse(text || "");
    }

    parse(text) {
      let currentSection = null;
      const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line || line.startsWith(";") || line.startsWith("#")) {
          continue;
        }
        if (line.startsWith("[") && line.endsWith("]") && line.length > 2) {
          const sectionName = line.slice(1, -1).trim().toLowerCase();
          if (!this.sections.has(sectionName)) {
            this.sections.set(sectionName, new Map());
          }
          currentSection = this.sections.get(sectionName);
          continue;
        }
        if (!currentSection) {
          continue;
        }
        const equalsIndex = line.indexOf("=");
        if (equalsIndex <= 0) {
          continue;
        }
        const key = line.slice(0, equalsIndex).trim().toLowerCase();
        const value = line.slice(equalsIndex + 1);
        currentSection.set(key, value);
      }
    }

    get(section, key) {
      const currentSection = this.sections.get(String(section || "").toLowerCase());
      if (!currentSection) {
        return undefined;
      }
      return currentSection.get(String(key || "").toLowerCase());
    }
  }

  function textOutput(path, data) {
    return { path: normalizePath(path), data, binary: false };
  }

  function binaryOutput(path, data) {
    return { path: normalizePath(path), data, binary: true };
  }

  function finishLines(lines) {
    return lines.join("\r\n") + "\r\n";
  }

  function splitCsv(value) {
    return String(value || "").split(",");
  }

  function pad(values, length) {
    if (values.length >= length) {
      return values;
    }
    const padded = values.slice();
    while (padded.length < length) {
      padded.push("");
    }
    return padded;
  }

  function toFixedCsv(data, length) {
    const values = splitCsv(data);
    const fixedValues = [];
    for (let i = 0; i < length; i++) {
      fixedValues.push(valueAt(values, i));
    }
    return fixedValues.join(",");
  }

  function valueAt(values, index) {
    if (index < 0 || index >= values.length) {
      return "";
    }
    return values[index];
  }

  function toInt(value) {
    const number = parseInt(String(value || "").trim(), 10);
    return Number.isFinite(number) ? number : 0;
  }

  function toDouble(value) {
    const number = Number.parseFloat(String(value || "").trim());
    return Number.isFinite(number) ? number : 0;
  }

  function roundOne(value) {
    return (Math.round(value * 10) / 10).toString();
  }

  function limitProfileString(value) {
    if (value == null || value.length <= MAX_PROFILE_STRING_LENGTH) {
      return value == null ? "" : value;
    }
    return value.slice(0, MAX_PROFILE_STRING_LENGTH);
  }

  function buildParameterHeader(startAddress, count) {
    let output = ",";
    for (let i = 0; i < count; i++) {
      output += "P" + String(startAddress + i).padStart(4, "0") + ",";
    }
    return output;
  }

  function convertGrdStatus(data) {
    const values = splitCsv(data);
    let output = "";
    for (let i = 0; i < values.length; i++) {
      if (i === 1 || i === 2) {
        if (values[i] === "0") {
          output += "OK,";
        } else if (values[i] === "1") {
          output += "NG,";
        } else {
          output += "PASS,";
        }
      } else {
        output += values[i] + ",";
      }
    }
    return output;
  }

  function removeBeforeGrdLastWaferColumns(data) {
    const values = pad(splitCsv(data), WHEEL_BEFORE_GRD_DATA_LEN);
    let output = "";
    for (let i = 0; i < WHEEL_BEFORE_GRD_DATA_LEN; i++) {
      if (i !== WHEEL_BEFORE_GRD_LAST_WAFER_MAX && i !== WHEEL_BEFORE_GRD_LAST_WAFER_MIN) {
        output += values[i] + ",";
      }
    }
    return output;
  }

  function removeT2Column(data) {
    const values = splitCsv(data);
    const outputValues = [];
    for (let i = 0; i < SAMP_LEN; i++) {
      if (i !== SAMP_IDX_T2) {
        outputValues.push(valueAt(values, i));
      }
    }
    return outputValues.join(",");
  }

  function hasMeaningfulExtendedAirCutValues(values) {
    for (let i = AI_AIR_CUT_BASE_DATA_LEN; i < values.length; i++) {
      if (isMeaningfulExtendedValue(values[i])) {
        return true;
      }
    }
    return false;
  }

  function isMeaningfulExtendedValue(value) {
    if (String(value || "").trim() === "") {
      return false;
    }
    const number = Number.parseFloat(String(value).trim());
    if (Number.isFinite(number)) {
      return Math.abs(number) > 0.0000001;
    }
    return true;
  }

  function emptyBatchSide(waferNumber) {
    return waferNumber + ",,,,,,,,,,,,,,,,,,";
  }

  function writeGrdDistCollectBlock(lines, chuckHeader, rows) {
    lines.push(chuckHeader + ",Wafer_ID,S1ReachZmch(mm),StartZmch-EndZmch(um),GrdDist-RemoveValue(um),S1ReachZabs-WaferIncomeMax(um)");
    for (const row of rows) {
      lines.push([
        row.chuckNumber,
        row.waferId,
        row.s1ReachZmch,
        row.startZmchEndZmch,
        row.grdDistRemoveValue,
        row.s1ReachZabsWaferIncomeMax
      ].join(","));
    }
  }

  function inlineStringCell(columnName, rowNumber, value) {
    return "<c r=\"" + columnName + rowNumber + "\" t=\"inlineStr\"><is><t>" + escapeXml(value) + "</t></is></c>";
  }

  function buildExcelColumns(count) {
    const columns = [];
    for (let i = 0; i < count; i++) {
      columns.push(getExcelColumnName(i + 1));
    }
    return columns;
  }

  function getExcelColumnName(columnNumber) {
    let columnName = "";
    while (columnNumber > 0) {
      const modulo = (columnNumber - 1) % 26;
      columnName = String.fromCharCode("A".charCodeAt(0) + modulo) + columnName;
      columnNumber = Math.floor((columnNumber - modulo) / 26);
    }
    return columnName;
  }

  function escapeXml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  async function writeOutputsToDirectory(outputBaseHandle, outputRootName, outputs, progress) {
    for (let i = 0; i < outputs.length; i++) {
      throwIfCancellationRequested();
      const output = outputs[i];
      progress(i + 1, outputs.length, output.path);
      await writeOutputFileWithRetry(outputBaseHandle, outputRootName, output);
      await nextFrame();
    }
  }

  async function writeOutputFileWithRetry(outputBaseHandle, outputRootName, output) {
    try {
      await writeOutputFile(outputBaseHandle, outputRootName, output);
    } catch (error) {
      if (!isStaleFileSystemError(error)) {
        throw error;
      }
      await nextFrame();
      await writeOutputFile(outputBaseHandle, outputRootName, output);
    }
  }

  async function writeOutputFile(outputBaseHandle, outputRootName, output) {
    const parts = [outputRootName].concat(normalizePath(output.path).split("/").filter(Boolean));
    const fileName = parts.pop();
    let dir = outputBaseHandle;
    for (const part of parts) {
      dir = await dir.getDirectoryHandle(part, { create: true });
    }
    const fileHandle = await dir.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(outputToBlob(output));
    await writable.close();
  }

  function outputToBlob(output) {
    if (output.binary) {
      return output.data;
    }
    return new Blob(["\ufeff", output.data], { type: "text/csv;charset=utf-8" });
  }

  function downloadZip(fileName, outputs) {
    const blob = createZipBlob(outputs.map((output) => ({
      path: output.path,
      data: output.binary ? output.data : "\ufeff" + output.data,
      binary: output.binary
    })));
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function createZipBlob(entries) {
    const bytes = createZipBytes(entries);
    const blob = new Blob([bytes], { type: "application/zip" });
    blob._zipBytes = bytes;
    return blob;
  }

  function createZipBytes(entries) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = encoder.encode(normalizePath(entry.path));
      const dataBytes = entry.binary
        ? blobToUint8ArraySyncLike(entry.data)
        : encoder.encode(String(entry.data || ""));
      const crc = crc32(dataBytes);
      const localHeader = makeLocalFileHeader(nameBytes, dataBytes.length, crc);
      localParts.push(localHeader, dataBytes);
      centralParts.push(makeCentralDirectoryHeader(nameBytes, dataBytes.length, crc, offset));
      offset += localHeader.length + dataBytes.length;
    }

    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const end = makeEndOfCentralDirectory(entries.length, centralSize, offset);
    return concatUint8Arrays([...localParts, ...centralParts, end]);
  }

  function blobToUint8ArraySyncLike(blobOrBytes) {
    if (blobOrBytes instanceof Uint8Array) {
      return blobOrBytes;
    }
    if (blobOrBytes instanceof ArrayBuffer) {
      return new Uint8Array(blobOrBytes);
    }
    if (blobOrBytes && blobOrBytes._zipBytes instanceof Uint8Array) {
      return blobOrBytes._zipBytes;
    }
    throw new Error("Binary ZIP data is not available.");
  }

  function concatUint8Arrays(parts) {
    const size = parts.reduce((sum, part) => sum + part.length, 0);
    const output = new Uint8Array(size);
    let offset = 0;
    for (const part of parts) {
      output.set(part, offset);
      offset += part.length;
    }
    return output;
  }

  function makeLocalFileHeader(nameBytes, size, crc) {
    const header = new Uint8Array(30 + nameBytes.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 0, true);
    view.setUint16(8, 0, true);
    view.setUint16(10, 0, true);
    view.setUint16(12, 0, true);
    view.setUint32(14, crc, true);
    view.setUint32(18, size, true);
    view.setUint32(22, size, true);
    view.setUint16(26, nameBytes.length, true);
    view.setUint16(28, 0, true);
    header.set(nameBytes, 30);
    return header;
  }

  function makeCentralDirectoryHeader(nameBytes, size, crc, offset) {
    const header = new Uint8Array(46 + nameBytes.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x02014b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, 20, true);
    view.setUint16(8, 0, true);
    view.setUint16(10, 0, true);
    view.setUint16(12, 0, true);
    view.setUint16(14, 0, true);
    view.setUint32(16, crc, true);
    view.setUint32(20, size, true);
    view.setUint32(24, size, true);
    view.setUint16(28, nameBytes.length, true);
    view.setUint16(30, 0, true);
    view.setUint16(32, 0, true);
    view.setUint16(34, 0, true);
    view.setUint16(36, 0, true);
    view.setUint32(38, 0, true);
    view.setUint32(42, offset, true);
    header.set(nameBytes, 46);
    return header;
  }

  function makeEndOfCentralDirectory(entryCount, centralSize, centralOffset) {
    const header = new Uint8Array(22);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x06054b50, true);
    view.setUint16(4, 0, true);
    view.setUint16(6, 0, true);
    view.setUint16(8, entryCount, true);
    view.setUint16(10, entryCount, true);
    view.setUint32(12, centralSize, true);
    view.setUint32(16, centralOffset, true);
    view.setUint16(20, 0, true);
    return header;
  }

  const CRC_TABLE = makeCrcTable();

  function makeCrcTable() {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function normalizePath(path) {
    return String(path || "").replace(/\\/g, "/").replace(/^\/+/, "").replace(/\/+/g, "/");
  }

  function joinPath() {
    return Array.from(arguments).filter((part) => part !== "").map(normalizePath).join("/");
  }

  function dirname(path) {
    const normalized = normalizePath(path);
    const index = normalized.lastIndexOf("/");
    return index < 0 ? "" : normalized.slice(0, index);
  }

  function basename(path) {
    const normalized = normalizePath(path);
    const index = normalized.lastIndexOf("/");
    return index < 0 ? normalized : normalized.slice(index + 1);
  }

  function basenameWithoutExtension(path) {
    const name = basename(path);
    const index = name.lastIndexOf(".");
    return index < 0 ? name : name.slice(0, index);
  }

  function compareText(left, right) {
    return String(left).localeCompare(String(right), undefined, { sensitivity: "accent" });
  }

  function comparePath(left, right) {
    const leftPath = typeof left === "string" ? left : left.relativePath;
    const rightPath = typeof right === "string" ? right : right.relativePath;
    return compareText(leftPath.toLowerCase(), rightPath.toLowerCase());
  }

  function nextFrame() {
    return new Promise((resolve) => setTimeout(resolve, 0));
  }

  window.FSG2300ConverterForTests = {
    ProcessLogConverter,
    convertFolder,
    makeSource,
    IniFile,
    helpers: {
      splitCsv,
      valueAt,
      removeT2Column,
      removeBeforeGrdLastWaferColumns,
      convertGrdStatus,
      createZipBlob,
      crc32
    }
  };
})();
