/**
 * Office2AI — 辦公文件轉 AI 低 Token 最佳化工具
 * 支援自動辨識 Word (.docx)、PowerPoint (.pptx)、Excel (.xlsx/.xls/.csv) 與純文字/JSON
 */

const state = {
  activeMode: "ai", // "ai" | "pdf2ppt"
  files: [], // Array of { id, name, rawSize, detectedType, parsedData, error }
  selectedFileId: null, // null means "view all merged" in AI tab
  selectedPdfId: null, // selected PDF file id in PDF2PPT tab
};

// DOM Elements - Header & Mode Tabs
const tabBtnAi = document.getElementById("tabBtnAi");
const tabBtnPdf2Ppt = document.getElementById("tabBtnPdf2Ppt");
const panelAi = document.getElementById("panelAi");
const panelPdf2Ppt = document.getElementById("panelPdf2Ppt");

// DOM Elements - Tab 1 (AI Converter)
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileListEl = document.getElementById("fileList");
const fileCountEl = document.getElementById("fileCount");
const outputPreview = document.getElementById("outputPreview");
const activeViewBadge = document.getElementById("activeViewBadge");

const metricRawSize = document.getElementById("metricRawSize");
const metricCleanSize = document.getElementById("metricCleanSize");
const metricSizeSaved = document.getElementById("metricSizeSaved");
const metricTokens = document.getElementById("metricTokens");
const metricTokenHint = document.getElementById("metricTokenHint");

const btnDemo = document.getElementById("btnDemo");
const btnClearAll = document.getElementById("btnClearAll");
const btnCopy = document.getElementById("btnCopy");
const btnCopyPrompt = document.getElementById("btnCopyPrompt");
const btnDownload = document.getElementById("btnDownload");

// DOM Elements - Tab 2 (PDF to PPT)
const pdf2pptDropzone = document.getElementById("pdf2pptDropzone");
const pdf2pptFileInput = document.getElementById("pdf2pptFileInput");
const pdf2pptFileListEl = document.getElementById("pdf2pptFileList");
const pdf2pptFileCountEl = document.getElementById("pdf2pptFileCount");
const pptMetricPages = document.getElementById("pptMetricPages");
const pptMetricBlocks = document.getElementById("pptMetricBlocks");
const pptMetricMode = document.getElementById("pptMetricMode");
const pptMetricHint = document.getElementById("pptMetricHint");
const pptActiveFileBadge = document.getElementById("pptActiveFileBadge");
const btnExportPptx = document.getElementById("btnExportPptx");
const pptSlidesPreview = document.getElementById("pptSlidesPreview");

// Settings Elements
const settingIds = [
  "outputFormat",
  "tableFormat",
  "optCollapseWhitespace",
  "optPdfSmartClean",
  "optTrimEmptyCells",
  "optRoundFloats",
  "optIncludeNotes",
  "optImagePlaceholder",
];

const pptSettingIds = [
  "pptConvertMode",
  "pptLayout",
  "pptScale",
  "optPptMergeBlocks",
  "optPptIncludeNotes",
  "optPptHideBgText",
];

function switchMode(mode) {
  state.activeMode = mode;
  if (tabBtnAi && tabBtnPdf2Ppt && panelAi && panelPdf2Ppt) {
    const isAi = mode === "ai";
    tabBtnAi.classList.toggle("active", isAi);
    tabBtnAi.setAttribute("aria-selected", String(isAi));
    tabBtnPdf2Ppt.classList.toggle("active", !isAi);
    tabBtnPdf2Ppt.setAttribute("aria-selected", String(!isAi));
    panelAi.classList.toggle("hidden", !isAi);
    panelPdf2Ppt.classList.toggle("hidden", isAi);
  }
  renderAll();
}

// Initialize Event Listeners
function init() {
  if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  if (tabBtnAi && tabBtnPdf2Ppt) {
    tabBtnAi.addEventListener("click", () => switchMode("ai"));
    tabBtnPdf2Ppt.addEventListener("click", () => switchMode("pdf2ppt"));
  }

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") fileInput.click();
  });

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", async (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files?.length) {
      await handleFiles(Array.from(e.dataTransfer.files));
    }
  });

  fileInput.addEventListener("change", async (e) => {
    if (e.target.files?.length) {
      await handleFiles(Array.from(e.target.files));
      fileInput.value = "";
    }
  });

  // Tab 2: PDF-to-PPT Dropzone Listeners
  if (pdf2pptDropzone && pdf2pptFileInput) {
    pdf2pptDropzone.addEventListener("click", () => pdf2pptFileInput.click());
    pdf2pptDropzone.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") pdf2pptFileInput.click();
    });

    pdf2pptDropzone.addEventListener("dragover", (e) => {
      e.preventDefault();
      pdf2pptDropzone.classList.add("dragover");
    });

    pdf2pptDropzone.addEventListener("dragleave", () => {
      pdf2pptDropzone.classList.remove("dragover");
    });

    pdf2pptDropzone.addEventListener("drop", async (e) => {
      e.preventDefault();
      pdf2pptDropzone.classList.remove("dragover");
      if (e.dataTransfer.files?.length) {
        await handleFiles(Array.from(e.dataTransfer.files));
      }
    });

    pdf2pptFileInput.addEventListener("change", async (e) => {
      if (e.target.files?.length) {
        await handleFiles(Array.from(e.target.files));
        pdf2pptFileInput.value = "";
      }
    });
  }

  settingIds.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("change", () => renderAll());
  });

  pptSettingIds.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("change", () => renderAll());
  });

  activeViewBadge.addEventListener("click", () => {
    state.selectedFileId = null;
    renderAll();
  });

  btnClearAll.addEventListener("click", () => {
    state.files = [];
    state.selectedFileId = null;
    state.selectedPdfId = null;
    renderAll();
    showToast("已清空所有檔案");
  });

  btnCopy.addEventListener("click", () => {
    copyText(outputPreview.value, "已複製 AI 最佳化內容至剪貼簿！");
  });

  btnCopyPrompt.addEventListener("click", () => {
    const promptPrefix =
      "以下為經過結構化壓縮的辦公文件內容（其中表格採用緊湊分隔格式以節省 Token）。請根據這些資料回答我的問題或進行摘要分析：\n\n";
    copyText(promptPrefix + outputPreview.value, "已複製「AI 提示詞 + 轉換內容」至剪貼簿！");
  });

  btnDownload.addEventListener("click", downloadResult);
  if (btnExportPptx) {
    btnExportPptx.addEventListener("click", () => {
      const targetPdf = getActivePdfFile();
      if (targetPdf) exportPdfToPptx(targetPdf);
    });
  }
  btnDemo.addEventListener("click", loadDemoFiles);
}

/**
 * 自動辨識檔案格式（結合 Magic Bytes、ZIP 內部結構與副檔名）
 */
async function detectFileFormat(file, arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer.slice(0, 8));
  const ext = (file.name.split(".").pop() || "").toLowerCase();

  // Check PDF signature (%PDF- -> 0x25 0x50 0x44 0x46 0x2D)
  const isPdf =
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46 &&
    bytes[4] === 0x2d;
  if (isPdf || ext === "pdf") {
    return { type: "pdf", label: "PDF" };
  }

  // Check ZIP signature (PK\x03\x04 -> 0x50 0x4B 0x03 0x04)
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  if (isZip && window.JSZip) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      if (zip.file("word/document.xml")) return { type: "docx", label: "Word", zip };
      if (zip.file("ppt/presentation.xml")) return { type: "pptx", label: "PPT", zip };
      if (zip.file("xl/workbook.xml")) return { type: "xlsx", label: "Excel", zip };
    } catch (_) {
      // Fallback to extension
    }
  }

  // Check OLE2 Compound Document signature (.xls, .doc, .ppt) -> D0 CF 11 E0
  const isOle2 = bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
  if (isOle2 && ext === "xls") {
    return { type: "xlsx", label: "Excel" };
  }

  if (ext === "docx") return { type: "docx", label: "Word" };
  if (ext === "pptx") return { type: "pptx", label: "PPT" };
  if (["xlsx", "xls", "csv"].includes(ext)) return { type: ext === "csv" ? "csv" : "xlsx", label: "Excel" };
  if (["json", "md", "txt"].includes(ext)) return { type: ext, label: "Text" };

  return { type: "txt", label: "Text" };
}

/**
 * 處理使用者丟入的檔案清單
 */
async function handleFiles(fileList) {
  for (const file of fileList) {
    const id = "file_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const detected = await detectFileFormat(file, arrayBuffer);
      let parsedData = null;

      if (detected.type === "pdf") {
        parsedData = await parsePdf(arrayBuffer);
      } else if (detected.type === "docx") {
        parsedData = await parseDocx(arrayBuffer, detected.zip);
      } else if (detected.type === "pptx") {
        parsedData = await parsePptx(arrayBuffer, detected.zip);
      } else if (detected.type === "xlsx" || detected.type === "csv") {
        parsedData = parseExcel(arrayBuffer, detected.type);
      } else {
        parsedData = parseTextFile(arrayBuffer, detected.type);
      }

      state.files.push({
        id,
        name: file.name,
        rawSize: file.size || arrayBuffer.byteLength,
        detectedType: detected.type,
        typeLabel: detected.label,
        parsedData,
      });
    } catch (err) {
      console.error("解析檔案失敗:", file.name, err);
      state.files.push({
        id,
        name: file.name,
        rawSize: file.size || 0,
        detectedType: "error",
        typeLabel: "Error",
        error: err.message || "無法解析此檔案格式",
      });
    }
  }

  renderAll();
  showToast(`已成功載入並轉換 ${fileList.length} 個檔案`);
}

/* =========================================================
 * 1. Word (.docx) 解析器：直接解析 OpenXML 結構，剔除樣式與圖片肥大流量
 * ========================================================= */
async function parseDocx(arrayBuffer, existingZip) {
  const zip = existingZip || (await JSZip.loadAsync(arrayBuffer));
  const docXmlFile = zip.file("word/document.xml");
  if (!docXmlFile) throw new Error("找不到 word/document.xml，確認是否為有效的 .docx 檔案");

  const xmlText = await docXmlFile.async("string");
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, "application/xml");

  // Parse hyperlinks mapping if available
  const relsMap = {};
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (relsFile) {
    const relsXml = parser.parseFromString(await relsFile.async("string"), "application/xml");
    const relNodes = relsXml.getElementsByTagName("Relationship");
    for (const rel of relNodes) {
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      const type = rel.getAttribute("Type") || "";
      if (id && target && type.includes("hyperlink")) {
        relsMap[id] = target;
      }
    }
  }

  const blocks = [];
  const body = xmlDoc.getElementsByTagName("w:body")[0] || xmlDoc.documentElement;

  for (const child of Array.from(body.childNodes)) {
    const nodeName = child.nodeName;
    if (nodeName === "w:p") {
      const pBlock = parseDocxParagraph(child, relsMap);
      if (pBlock) blocks.push(pBlock);
    } else if (nodeName === "w:tbl") {
      const rows = parseDocxTable(child, relsMap);
      if (rows.length > 0) {
        blocks.push({ kind: "table", rows });
      }
    }
  }

  return { docType: "docx", blocks };
}

function parseDocxParagraph(pNode, relsMap) {
  let headingLevel = 0;
  let isList = false;
  let listLevel = 0;

  const pPr = pNode.getElementsByTagName("w:pPr")[0];
  if (pPr) {
    const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
    if (pStyle) {
      const val = pStyle.getAttribute("w:val") || "";
      const match = val.match(/^(?:Heading|標題\s*)([1-6])$/i);
      if (match) headingLevel = parseInt(match[1], 10);
      else if (/^Title$/i.test(val)) headingLevel = 1;
      else if (/^Subtitle$/i.test(val)) headingLevel = 2;
    }
    const outlineLvl = pPr.getElementsByTagName("w:outlineLvl")[0];
    if (!headingLevel && outlineLvl) {
      const lvl = parseInt(outlineLvl.getAttribute("w:val"), 10);
      if (!isNaN(lvl) && lvl >= 0 && lvl <= 5) headingLevel = lvl + 1;
    }

    const numPr = pPr.getElementsByTagName("w:numPr")[0];
    if (numPr) {
      isList = true;
      const ilvl = numPr.getElementsByTagName("w:ilvl")[0];
      if (ilvl) listLevel = parseInt(ilvl.getAttribute("w:val") || "0", 10) || 0;
    }
  }

  let hasImage = false;
  if (pNode.getElementsByTagName("w:drawing").length > 0 || pNode.getElementsByTagName("w:pict").length > 0) {
    hasImage = true;
  }

  const textParts = [];
  const collectRuns = (parent) => {
    for (const node of Array.from(parent.childNodes)) {
      if (node.nodeName === "w:r") {
        for (const rChild of Array.from(node.childNodes)) {
          if (rChild.nodeName === "w:t") {
            textParts.push(rChild.textContent || "");
          } else if (rChild.nodeName === "w:tab") {
            textParts.push("\t");
          } else if (rChild.nodeName === "w:br" || rChild.nodeName === "w:cr") {
            textParts.push("\n");
          }
        }
      } else if (node.nodeName === "w:hyperlink") {
        const rId = node.getAttribute("r:id");
        const beforeLen = textParts.length;
        collectRuns(node);
        const linkText = textParts.splice(beforeLen).join("");
        if (rId && relsMap[rId] && linkText && linkText !== relsMap[rId]) {
          textParts.push(`[${linkText}](${relsMap[rId]})`);
        } else {
          textParts.push(linkText);
        }
      }
    }
  };

  collectRuns(pNode);
  const text = textParts.join("").trim();

  if (!text && !hasImage) return null;
  return {
    kind: "paragraph",
    text,
    headingLevel,
    isList,
    listLevel,
    hasImage,
  };
}

function parseDocxTable(tblNode, relsMap) {
  const rows = [];
  const trNodes = Array.from(tblNode.childNodes).filter((n) => n.nodeName === "w:tr");
  for (const tr of trNodes) {
    const rowCells = [];
    const tcNodes = Array.from(tr.childNodes).filter((n) => n.nodeName === "w:tc");
    for (const tc of tcNodes) {
      const pNodes = Array.from(tc.childNodes).filter((n) => n.nodeName === "w:p");
      const cellTexts = pNodes
        .map((p) => {
          const parsed = parseDocxParagraph(p, relsMap);
          return parsed ? parsed.text : "";
        })
        .filter(Boolean);
      rowCells.push(cellTexts.join(" / "));
    }
    if (rowCells.some((c) => c !== "")) {
      rows.push(rowCells);
    }
  }
  return rows;
}

/* =========================================================
 * 2. PowerPoint (.pptx) 解析器：按投影片順序提取標題、層級清單、表格與備忘稿
 * ========================================================= */
async function parsePptx(arrayBuffer, existingZip) {
  const zip = existingZip || (await JSZip.loadAsync(arrayBuffer));
  const parser = new DOMParser();

  const slideFiles = [];
  zip.forEach((relativePath) => {
    const match = relativePath.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (match) {
      slideFiles.push({ num: parseInt(match[1], 10), path: relativePath });
    }
  });

  slideFiles.sort((a, b) => a.num - b.num);
  const slides = [];

  for (const slideFile of slideFiles) {
    const xmlStr = await zip.file(slideFile.path).async("string");
    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");

    let title = "";
    const bullets = [];
    const tables = [];
    let imageCount = 0;

    const spTree = xmlDoc.getElementsByTagName("p:spTree")[0] || xmlDoc.documentElement;
    for (const node of Array.from(spTree.childNodes)) {
      if (node.nodeName === "p:sp") {
        // Check if shape is title placeholder
        const ph = node.getElementsByTagName("p:ph")[0];
        const phType = ph ? ph.getAttribute("type") || "" : "";
        const isTitle = phType === "title" || phType === "ctrTitle";
        // Ignore slide number / footer / date placeholders to save tokens
        if (phType === "sldNum" || phType === "dt" || phType === "ftr") continue;

        const aParas = Array.from(node.getElementsByTagName("a:p"));
        for (const aP of aParas) {
          const pPr = aP.getElementsByTagName("a:pPr")[0];
          const lvl = pPr ? parseInt(pPr.getAttribute("lvl") || "0", 10) : 0;
          const tNodes = Array.from(aP.getElementsByTagName("a:t"));
          const line = tNodes.map((t) => t.textContent || "").join("").trim();
          if (!line) continue;

          if (isTitle && !title) {
            title = line;
          } else {
            bullets.push({ level: lvl, text: line });
          }
        }
      } else if (node.nodeName === "p:graphicFrame") {
        const tbl = node.getElementsByTagName("a:tbl")[0];
        if (tbl) {
          const rows = [];
          const trs = Array.from(tbl.childNodes).filter((n) => n.nodeName === "a:tr");
          for (const tr of trs) {
            const row = [];
            const tcs = Array.from(tr.childNodes).filter((n) => n.nodeName === "a:tc");
            for (const tc of tcs) {
              const tNodes = Array.from(tc.getElementsByTagName("a:t"));
              row.push(tNodes.map((t) => t.textContent || "").join(" ").trim());
            }
            if (row.some((c) => c !== "")) rows.push(row);
          }
          if (rows.length > 0) tables.push(rows);
        }
      } else if (node.nodeName === "p:pic") {
        imageCount++;
      }
    }

    // Speaker Notes
    let notes = "";
    const notesPath = `ppt/notesSlides/notesSlide${slideFile.num}.xml`;
    if (zip.file(notesPath)) {
      const notesXmlStr = await zip.file(notesPath).async("string");
      const notesDoc = parser.parseFromString(notesXmlStr, "application/xml");
      const noteShapes = Array.from(notesDoc.getElementsByTagName("p:sp"));
      const noteLines = [];
      for (const sp of noteShapes) {
        const ph = sp.getElementsByTagName("p:ph")[0];
        const phType = ph ? ph.getAttribute("type") || "" : "";
        if (phType === "sldNum" || phType === "sldImg" || phType === "hdr" || phType === "ftr") continue;
        const aParas = Array.from(sp.getElementsByTagName("a:p"));
        for (const aP of aParas) {
          const line = Array.from(aP.getElementsByTagName("a:t"))
            .map((t) => t.textContent || "")
            .join("")
            .trim();
          if (line) noteLines.push(line);
        }
      }
      notes = noteLines.join("\n");
    }

    slides.push({
      slideNumber: slideFile.num,
      title,
      bullets,
      tables,
      imageCount,
      notes,
    });
  }

  return { docType: "pptx", slides };
}

/* =========================================================
 * 3. Excel (.xlsx / .xls / .csv) 解析器：自動清理空儲存格與浮點數誤差
 * ========================================================= */
function parseExcel(arrayBuffer, type) {
  let workbook;
  if (type === "csv") {
    const text = new TextDecoder("utf-8").decode(arrayBuffer);
    workbook = XLSX.read(text, { type: "string" });
  } else {
    workbook = XLSX.read(arrayBuffer, { type: "array" });
  }

  const sheets = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    sheets.push({
      sheetName,
      rows: rawRows,
    });
  }

  return { docType: "excel", sheets };
}

/* =========================================================
 * 4. PDF (.pdf) 解析器：座標排序、字型與真實色彩吸取、重複頁首頁尾濾除與斷行段落合併
 * ========================================================= */
function rgbToHex(r, g, b) {
  const clamp = (n) => Math.max(0, Math.min(255, Math.round(n)));
  return [clamp(r), clamp(g), clamp(b)]
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function medianOfSorted(arr) {
  if (!arr || arr.length === 0) return 0;
  return arr[Math.floor(arr.length / 2)];
}

/**
 * 解析 PDF 文字項目的字型家族 (Font Family)、粗體與斜體，對應為 PowerPoint 支援之字型
 */
function resolvePdfFontProps(item, styles, text) {
  const rawName = String(item.fontName || "");
  const styleObj = (styles && styles[rawName]) || {};
  const rawFamily = String(styleObj.fontFamily || "");
  const combined = `${rawName} ${rawFamily}`;
  const tx = item.transform || [1, 0, 0, 1, 0, 0];

  const isBold = /bold|black|heavy|semibold|demi|extrab|w[6-9]|medi/i.test(combined);
  const isItalic = /italic|oblique/i.test(combined) || Math.abs(tx[2]) > 0.08;
  const hasCjk = /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(text || "");

  let fontFace = hasCjk ? "Microsoft JhengHei" : "Arial";
  if (/kai|楷/i.test(combined)) {
    fontFace = "DFKai-SB";
  } else if (/ming|song|細明|明體|宋/i.test(combined)) {
    fontFace = "PMingLiU";
  } else if (/mono|courier|consolas|menlo|code/i.test(combined)) {
    fontFace = "Consolas";
  } else if (/serif|times|georgia|garamond|minion|cambria/i.test(combined) && !/sans/i.test(combined)) {
    fontFace = hasCjk ? "PMingLiU" : "Times New Roman";
  } else if (/calibri/i.test(combined)) {
    fontFace = hasCjk ? "Microsoft JhengHei" : "Calibri";
  }

  return { fontFace, isBold, isItalic };
}

/**
 * 從已渲染的 PDF Canvas 像素中精確吸取該文字區塊的「原始文字顏色 (colorHex)」與「局部背景色 (bgHex)」
 */
function sampleTextAndBgColorsFromCanvas(ctx, xPt, yTopPt, widthPt, heightPt, scale, canvasW, canvasH) {
  const defaultResult = { colorHex: "0F172A", bgHex: "FFFFFF", bgRgb: [255, 255, 255] };
  if (!ctx || typeof ctx.getImageData !== "function") return defaultResult;

  try {
    const pad = 2;
    const x0 = Math.max(0, Math.floor(xPt * scale) - pad);
    const y0 = Math.max(0, Math.floor(yTopPt * scale) - pad);
    const x1 = Math.min(canvasW - 1, Math.ceil((xPt + widthPt) * scale) + pad);
    const y1 = Math.min(canvasH - 1, Math.ceil((yTopPt + heightPt) * scale) + pad);
    const w = x1 - x0 + 1;
    const h = y1 - y0 + 1;
    if (w <= 2 || h <= 2) return defaultResult;

    const imgData = ctx.getImageData(x0, y0, w, h).data;

    // 1. Sample perimeter pixels (top & bottom edges) to find local background color
    const bgRs = [];
    const bgGs = [];
    const bgBs = [];
    for (let cx = 0; cx < w; cx++) {
      const topIdx = cx * 4;
      const botIdx = ((h - 1) * w + cx) * 4;
      bgRs.push(imgData[topIdx], imgData[botIdx]);
      bgGs.push(imgData[topIdx + 1], imgData[botIdx + 1]);
      bgBs.push(imgData[topIdx + 2], imgData[botIdx + 2]);
    }
    bgRs.sort((a, b) => a - b);
    bgGs.sort((a, b) => a - b);
    bgBs.sort((a, b) => a - b);
    const bgR = medianOfSorted(bgRs);
    const bgG = medianOfSorted(bgGs);
    const bgB = medianOfSorted(bgBs);

    // 2. Find interior pixels with highest contrast from background (the core strokes of the text)
    const candidates = [];
    for (let cy = 1; cy < h - 1; cy++) {
      for (let cx = 1; cx < w - 1; cx++) {
        const idx = (cy * w + cx) * 4;
        const r = imgData[idx];
        const g = imgData[idx + 1];
        const b = imgData[idx + 2];
        const dist = Math.hypot(r - bgR, g - bgG, b - bgB);
        if (dist > 28) {
          candidates.push({ r, g, b, dist });
        }
      }
    }

    if (candidates.length === 0) {
      return {
        colorHex: bgR + bgG + bgB < 380 ? "FFFFFF" : "0F172A",
        bgHex: rgbToHex(bgR, bgG, bgB),
        bgRgb: [bgR, bgG, bgB],
      };
    }

    // Sort by contrast distance descending and take top 25% core stroke pixels (avoiding anti-aliased edges)
    candidates.sort((a, b) => b.dist - a.dist);
    const topCount = Math.max(1, Math.ceil(candidates.length * 0.25));
    const core = candidates.slice(0, topCount);
    const fgRs = core.map((c) => c.r).sort((a, b) => a - b);
    const fgGs = core.map((c) => c.g).sort((a, b) => a - b);
    const fgBs = core.map((c) => c.b).sort((a, b) => a - b);

    return {
      colorHex: rgbToHex(medianOfSorted(fgRs), medianOfSorted(fgGs), medianOfSorted(fgBs)),
      bgHex: rgbToHex(bgR, bgG, bgB),
      bgRgb: [bgR, bgG, bgB],
    };
  } catch (_) {
    return defaultResult;
  }
}

async function parsePdf(arrayBuffer) {
  if (!window.pdfjsLib) {
    throw new Error("PDF.js 尚未載入，請確認網路連線或重新整理頁面");
  }

  const pdfBytes = new Uint8Array(arrayBuffer.slice(0));
  const loadingTask = window.pdfjsLib.getDocument({
    data: new Uint8Array(pdfBytes),
    useSystemFonts: true,
    disableFontFace: false,
  });
  const pdfDoc = await loadingTask.promise;

  const rawPages = [];
  const allFontSizes = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.0 });
    const pageWidthPt = viewport.width || 612;
    const pageHeightPt = viewport.height || 792;
    const textContent = await page.getTextContent();
    const styles = textContent.styles || {};

    // Render page at 1.2x scale for thumbnail preview AND pixel-accurate text color sampling
    let thumbDataUrl = "";
    let sampleCtx = null;
    let sampleW = 0;
    let sampleH = 0;
    const sampleScale = 1.25;
    try {
      const thumbCanvas = document.createElement("canvas");
      if (thumbCanvas && typeof thumbCanvas.getContext === "function") {
        const thumbViewport = page.getViewport({ scale: sampleScale });
        sampleW = Math.ceil(thumbViewport.width);
        sampleH = Math.ceil(thumbViewport.height);
        thumbCanvas.width = sampleW;
        thumbCanvas.height = sampleH;
        sampleCtx = thumbCanvas.getContext("2d");
        if (sampleCtx) {
          sampleCtx.fillStyle = "#ffffff";
          sampleCtx.fillRect(0, 0, sampleW, sampleH);
          await page.render({ canvasContext: sampleCtx, viewport: thumbViewport }).promise;
          thumbDataUrl = thumbCanvas.toDataURL("image/png");
        }
      }
    } catch (_) {
      // Ignore thumbnail rendering errors in headless environments
    }

    // Extract text items with coordinates, font family, bold, italic, and font size
    const items = [];
    for (const item of textContent.items) {
      if (!item.str || !item.str.trim()) continue;
      const tx = item.transform || [1, 0, 0, 1, 0, 0];
      const x = tx[4];
      const y = tx[5];
      const fontSize = Math.round(Math.hypot(tx[0], tx[1]) || item.height || 12);
      const itemWidth = item.width || item.str.length * fontSize * 0.52;
      const fontProps = resolvePdfFontProps(item, styles, item.str);

      allFontSizes.push(fontSize);
      items.push({
        str: item.str,
        x,
        y,
        width: itemWidth,
        fontSize,
        fontFace: fontProps.fontFace,
        isBold: fontProps.isBold,
        isItalic: fontProps.isItalic,
      });
    }

    // Sort top-to-bottom (descending Y in PDF coordinate space), then left-to-right (ascending X)
    items.sort((a, b) => {
      const tol = Math.max(3.5, Math.min(a.fontSize, b.fontSize) * 0.42);
      if (Math.abs(b.y - a.y) > tol) return b.y - a.y;
      return a.x - b.x;
    });

    // Group items into visual lines (for AI text extraction) AND tight horizontal segments (for PPT editable text boxes)
    const lineGroups = [];
    for (const item of items) {
      const lastGroup = lineGroups[lineGroups.length - 1];
      if (
        lastGroup &&
        Math.abs(lastGroup.y - item.y) <= Math.max(3.5, Math.min(lastGroup.fontSize, item.fontSize) * 0.42)
      ) {
        lastGroup.items.push(item);
        lastGroup.fontSize = Math.max(lastGroup.fontSize, item.fontSize);
      } else {
        lineGroups.push({
          y: item.y,
          fontSize: item.fontSize,
          items: [item],
        });
      }
    }

    const lines = [];
    const pptSegments = [];

    for (const group of lineGroups) {
      group.items.sort((a, b) => a.x - b.x);
      let lineText = "";
      let prevEnd = null;
      let prevChar = "";
      const minX = group.items[0].x;
      let maxEnd = minX;

      // Also split into tight PPT segments when there is a large horizontal gap (columns/tables) or font change
      let curSeg = null;

      for (const it of group.items) {
        const text = it.str;
        const itemEnd = it.x + (it.width || text.length * it.fontSize * 0.52);

        if (prevEnd !== null) {
          const gap = it.x - prevEnd;
          const firstChar = text.charAt(0);
          const isCjkPair =
            /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(prevChar) &&
            /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(firstChar);

          if (gap > group.fontSize * 2.2) {
            lineText += "  ";
          } else if (gap > group.fontSize * 0.25 && !isCjkPair && !lineText.endsWith(" ") && !text.startsWith(" ")) {
            lineText += " ";
          }

          // Check if item belongs to current PPT segment or starts a new column/segment
          const sameStyle =
            curSeg &&
            Math.abs(curSeg.fontSize - it.fontSize) <= 1.5 &&
            curSeg.isBold === it.isBold &&
            curSeg.fontFace === it.fontFace;
          if (curSeg && sameStyle && gap <= group.fontSize * 1.35) {
            if (gap > group.fontSize * 0.25 && !isCjkPair && !curSeg.text.endsWith(" ") && !text.startsWith(" ")) {
              curSeg.text += " ";
            }
            curSeg.text += text;
            curSeg.maxX = Math.max(curSeg.maxX, itemEnd);
          } else {
            if (curSeg) pptSegments.push(curSeg);
            curSeg = {
              text,
              fontSize: it.fontSize,
              fontFace: it.fontFace,
              isBold: it.isBold,
              isItalic: it.isItalic,
              y: it.y,
              minX: it.x,
              maxX: itemEnd,
            };
          }
        } else {
          curSeg = {
            text,
            fontSize: it.fontSize,
            fontFace: it.fontFace,
            isBold: it.isBold,
            isItalic: it.isItalic,
            y: it.y,
            minX: it.x,
            maxX: itemEnd,
          };
        }

        lineText += text;
        prevEnd = itemEnd;
        if (prevEnd > maxEnd) maxEnd = prevEnd;
        prevChar = text.charAt(text.length - 1);
      }
      if (curSeg) pptSegments.push(curSeg);

      const trimmedLine = lineText.trim();
      if (trimmedLine) {
        const yTopPt = Math.max(0, pageHeightPt - group.y - group.fontSize * 0.88);
        const widthPt = Math.max(group.fontSize, maxEnd - minX);
        const heightPt = group.fontSize * 1.25;
        const sampled = sampleTextAndBgColorsFromCanvas(
          sampleCtx,
          Math.max(0, minX),
          yTopPt,
          widthPt,
          heightPt,
          sampleScale,
          sampleW,
          sampleH
        );

        lines.push({
          text: trimmedLine,
          fontSize: group.fontSize,
          fontFace: group.items[0].fontFace,
          isBold: group.items[0].isBold,
          isItalic: group.items[0].isItalic,
          colorHex: sampled.colorHex,
          bgHex: sampled.bgHex,
          y: group.y,
          xPt: Math.max(0, minX),
          yTopPt,
          widthPt,
          heightPt,
        });
      }
    }

    // Finalize tight PPT segments with exact bounding boxes and sampled colors
    const finalizedSegments = pptSegments
      .map((seg) => {
        const text = seg.text.trim();
        if (!text) return null;
        const xPt = Math.max(0, seg.minX);
        const yTopPt = Math.max(0, pageHeightPt - seg.y - seg.fontSize * 0.88);
        const widthPt = Math.max(seg.fontSize * 0.8, seg.maxX - seg.minX);
        const heightPt = seg.fontSize * 1.22;
        const sampled = sampleTextAndBgColorsFromCanvas(
          sampleCtx,
          xPt,
          yTopPt,
          widthPt,
          heightPt,
          sampleScale,
          sampleW,
          sampleH
        );
        return {
          text,
          fontSize: seg.fontSize,
          fontFace: seg.fontFace,
          isBold: seg.isBold,
          isItalic: seg.isItalic,
          colorHex: sampled.colorHex,
          bgHex: sampled.bgHex,
          xPt,
          yTopPt,
          widthPt,
          heightPt,
        };
      })
      .filter(Boolean);

    rawPages.push({
      pageNumber: pageNum,
      pageWidthPt,
      pageHeightPt,
      thumbDataUrl,
      lines,
      pptSegments: finalizedSegments,
    });
  }

  // Determine median body font size for heading detection
  allFontSizes.sort((a, b) => a - b);
  const bodyFontSize = allFontSizes.length > 0 ? allFontSizes[Math.floor(allFontSizes.length / 2)] : 12;

  // Detect repeated running headers/footers across pages (when PDF has >= 2 pages)
  const edgeSignatures = new Map();
  const normalizeSig = (s) => s.replace(/\d+/g, "#").replace(/\s+/g, " ").trim().toLowerCase();

  if (rawPages.length >= 2) {
    for (const p of rawPages) {
      if (p.lines.length > 0) {
        const firstSig = normalizeSig(p.lines[0].text);
        if (firstSig) edgeSignatures.set("top:" + firstSig, (edgeSignatures.get("top:" + firstSig) || 0) + 1);
      }
      if (p.lines.length > 1) {
        const lastSig = normalizeSig(p.lines[p.lines.length - 1].text);
        if (lastSig) edgeSignatures.set("bot:" + lastSig, (edgeSignatures.get("bot:" + lastSig) || 0) + 1);
      }
    }
  }

  const isStandalonePageNumber = (str) =>
    /^(?:[-–—]?\s*\d+\s*[-–—]?|page\s*\d+(?:\s*of\s*\d+)?|\d+\s*\/\s*\d+|第\s*\d+\s*頁(?:\s*[，,\/]\s*共\s*\d+\s*頁)?)$/i.test(
      str.trim()
    );

  const pages = rawPages.map((p) => {
    const markedLines = p.lines.map((line, idx) => {
      const isTop = idx === 0;
      const isBottom = idx === p.lines.length - 1;
      const sig = normalizeSig(line.text);
      const repeatedTop = isTop && (edgeSignatures.get("top:" + sig) || 0) >= 2;
      const repeatedBot = isBottom && (edgeSignatures.get("bot:" + sig) || 0) >= 2;
      const isPageNum = (isTop || isBottom) && isStandalonePageNumber(line.text);

      return {
        ...line,
        isHeaderFooterOrPageNum: repeatedTop || repeatedBot || isPageNum,
        isHeading: line.fontSize >= bodyFontSize * 1.25 && line.text.length <= 80,
      };
    });

    const markedSegments = (p.pptSegments || []).map((seg) => ({
      ...seg,
      isHeading: seg.isBold || (seg.fontSize >= bodyFontSize * 1.25 && seg.text.length <= 80),
    }));

    return {
      pageNumber: p.pageNumber,
      pageWidthPt: p.pageWidthPt,
      pageHeightPt: p.pageHeightPt,
      thumbDataUrl: p.thumbDataUrl,
      lines: markedLines,
      pptSegments: markedSegments,
    };
  });

  return { docType: "pdf", pages, bodyFontSize, pdfBytes };
}

function formatPdfPages(parsedData, settings) {
  const pageBlocks = [];

  for (const page of parsedData.pages) {
    let lines = page.lines;
    if (settings.pdfSmartClean) {
      lines = lines.filter((l) => !l.isHeaderFooterOrPageNum);
    }
    if (lines.length === 0) continue;

    if (!settings.pdfSmartClean) {
      const rawPageText = lines
        .map((l) => (settings.collapseWhitespace ? l.text.replace(/[ \t]+/g, " ").trim() : l.text))
        .join("\n");
      if (parsedData.pages.length > 1) {
        pageBlocks.push(`## Page ${page.pageNumber}\n${rawPageText}`);
      } else {
        pageBlocks.push(rawPageText);
      }
      continue;
    }

    // Smart paragraph & heading reconstruction
    const paragraphs = [];
    let currentPara = "";

    const flushPara = () => {
      if (currentPara.trim()) {
        paragraphs.push(currentPara.trim());
        currentPara = "";
      }
    };

    for (const line of lines) {
      const cleanLine = settings.collapseWhitespace ? line.text.replace(/[ \t]+/g, " ").trim() : line.text.trim();
      if (!cleanLine) continue;

      const isBullet = /^(?:[-•*▪‣]|\d+[.)]|[(（]\d+[)）]|[一二三四五六七八九十]+、)\s*/.test(cleanLine);

      if (line.isHeading) {
        flushPara();
        paragraphs.push(`### ${cleanLine}`);
      } else if (isBullet) {
        flushPara();
        paragraphs.push(cleanLine.startsWith("-") ? cleanLine : `- ${cleanLine}`);
      } else {
        if (!currentPara) {
          currentPara = cleanLine;
        } else {
          // Join wrapped lines: handle English hyphenation ("trans-\nform" -> "transform") and CJK vs Latin spacing
          if (/[a-zA-Z]-$/.test(currentPara) && /^[a-zA-Z]/.test(cleanLine)) {
            currentPara = currentPara.slice(0, -1) + cleanLine;
          } else {
            const lastChar = currentPara.charAt(currentPara.length - 1);
            const firstChar = cleanLine.charAt(0);
            const isCjkJoin =
              /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(lastChar) ||
              /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(firstChar);
            currentPara += (isCjkJoin ? "" : " ") + cleanLine;
          }
          // If line ends with terminal punctuation, check if it looks like the end of a paragraph
          if (/[。！？!?.：:]$/.test(cleanLine) && cleanLine.length < 42) {
            flushPara();
          }
        }
      }
    }
    flushPara();

    const bodyText = paragraphs.join("\n\n");
    if (bodyText) {
      if (parsedData.pages.length > 1) {
        pageBlocks.push(`## Page ${page.pageNumber}\n${bodyText}`);
      } else {
        pageBlocks.push(bodyText);
      }
    }
  }

  return pageBlocks.join("\n\n");
}

/* =========================================================
 * 5. 純文字 / JSON / Markdown 解析器
 * ========================================================= */
function parseTextFile(arrayBuffer, type) {
  const rawText = new TextDecoder("utf-8").decode(arrayBuffer);
  if (type === "json") {
    try {
      const parsed = JSON.parse(rawText);
      return { docType: "json", data: parsed, rawText };
    } catch (_) {
      // Fallback to raw text
    }
  }
  return { docType: "text", rawText };
}

/* =========================================================
 * 格式化引擎：根據設定將解析後的結構轉為 AI 極致省 Token 輸出
 * ========================================================= */
function getSettings() {
  return {
    outputFormat: document.getElementById("outputFormat").value,
    tableFormat: document.getElementById("tableFormat").value,
    collapseWhitespace: document.getElementById("optCollapseWhitespace").checked,
    pdfSmartClean: document.getElementById("optPdfSmartClean").checked,
    trimEmptyCells: document.getElementById("optTrimEmptyCells").checked,
    roundFloats: document.getElementById("optRoundFloats").checked,
    includeNotes: document.getElementById("optIncludeNotes").checked,
    imagePlaceholder: document.getElementById("optImagePlaceholder").checked,
  };
}

function cleanCellValue(val, settings) {
  if (val === null || val === undefined) return "";
  if (typeof val === "number") {
    if (settings.roundFloats && !Number.isInteger(val)) {
      return Number(val.toFixed(6)).toString();
    }
    return String(val);
  }
  let str = String(val);
  // Check numeric string with floating point precision noise
  if (settings.roundFloats && /^-?\d+\.\d{7,}$/.test(str.trim())) {
    const num = parseFloat(str);
    if (!isNaN(num)) str = Number(num.toFixed(6)).toString();
  }
  if (settings.collapseWhitespace) {
    str = str.replace(/\s+/g, " ").trim();
  }
  return str;
}

function processMatrix(rows, settings) {
  if (!rows || rows.length === 0) return [];
  let cleaned = rows.map((row) => row.map((cell) => cleanCellValue(cell, settings)));

  if (settings.trimEmptyCells) {
    // Filter out completely empty rows
    cleaned = cleaned.filter((row) => row.some((cell) => cell !== ""));
    if (cleaned.length === 0) return [];

    // Find non-empty columns
    const maxCols = Math.max(...cleaned.map((r) => r.length));
    const nonEmptyCols = [];
    for (let c = 0; c < maxCols; c++) {
      if (cleaned.some((r) => (r[c] || "") !== "")) {
        nonEmptyCols.push(c);
      }
    }
    cleaned = cleaned.map((r) => nonEmptyCols.map((c) => r[c] ?? ""));
  }
  return cleaned;
}

function formatTable(rows, settings) {
  const matrix = processMatrix(rows, settings);
  if (matrix.length === 0) return "";

  if (settings.tableFormat === "csv") {
    return matrix
      .map((row) =>
        row
          .map((cell) => {
            if (/[",\n]/.test(cell)) {
              return `"${cell.replace(/"/g, '""')}"`;
            }
            return cell;
          })
          .join(",")
      )
      .join("\n");
  }

  if (settings.tableFormat === "tsv") {
    return matrix.map((row) => row.map((c) => c.replace(/\t/g, " ")).join("\t")).join("\n");
  }

  // Standard Markdown Table
  const colCount = Math.max(...matrix.map((r) => r.length));
  const escapeMd = (s) => s.replace(/\|/g, "\\|").replace(/\n/g, " ");
  const header = matrix[0];
  const headerLine = "| " + Array.from({ length: colCount }, (_, i) => escapeMd(header[i] || "")).join(" | ") + " |";
  const sepLine = "| " + Array.from({ length: colCount }, () => "---").join(" | ") + " |";
  const bodyLines = matrix
    .slice(1)
    .map((r) => "| " + Array.from({ length: colCount }, (_, i) => escapeMd(r[i] || "")).join(" | ") + " |");

  return [headerLine, sepLine, ...bodyLines].join("\n");
}

function formatSingleFileContent(fileObj, settings) {
  if (fileObj.error) {
    return `[錯誤] ${fileObj.name}: ${fileObj.error}`;
  }

  const { parsedData } = fileObj;
  let content = "";

  if (parsedData.docType === "pdf") {
    content = formatPdfPages(parsedData, settings);
  } else if (parsedData.docType === "docx") {
    const lines = [];
    for (const block of parsedData.blocks) {
      if (block.kind === "paragraph") {
        let text = block.text;
        if (settings.collapseWhitespace) {
          text = text.replace(/[ \t]+/g, " ").trim();
        }
        if (block.hasImage && settings.imagePlaceholder) {
          text = text ? `${text} [圖片]` : "[圖片]";
        }
        if (!text) continue;

        if (block.headingLevel > 0) {
          lines.push(`${"#".repeat(block.headingLevel)} ${text}`);
        } else if (block.isList) {
          const indent = "  ".repeat(Math.min(block.listLevel, 4));
          lines.push(`${indent}- ${text}`);
        } else {
          lines.push(text);
        }
      } else if (block.kind === "table") {
        const tblStr = formatTable(block.rows, settings);
        if (tblStr) lines.push(tblStr);
      }
    }
    content = lines.join("\n\n");
  } else if (parsedData.docType === "pptx") {
    const slideSections = [];
    for (const slide of parsedData.slides) {
      const sLines = [];
      const headerTitle = slide.title ? `: ${slide.title}` : "";
      sLines.push(`## Slide ${slide.slideNumber}${headerTitle}`);

      for (const b of slide.bullets) {
        const text = settings.collapseWhitespace ? b.text.replace(/\s+/g, " ").trim() : b.text;
        const indent = "  ".repeat(Math.min(b.level, 4));
        sLines.push(`${indent}- ${text}`);
      }

      if (slide.imageCount > 0 && settings.imagePlaceholder) {
        sLines.push(`[圖片 x${slide.imageCount}]`);
      }

      for (const tbl of slide.tables) {
        const tblStr = formatTable(tbl, settings);
        if (tblStr) sLines.push(tblStr);
      }

      if (settings.includeNotes && slide.notes) {
        const cleanNotes = settings.collapseWhitespace ? slide.notes.replace(/\s+/g, " ").trim() : slide.notes;
        sLines.push(`> [備忘稿] ${cleanNotes}`);
      }

      if (sLines.length > 1 || slide.title) {
        slideSections.push(sLines.join("\n"));
      }
    }
    content = slideSections.join("\n\n");
  } else if (parsedData.docType === "excel") {
    const sheetSections = [];
    for (const sheet of parsedData.sheets) {
      const tblStr = formatTable(sheet.rows, settings);
      if (!tblStr) continue;
      if (parsedData.sheets.length > 1) {
        sheetSections.push(`## Sheet: ${sheet.sheetName}\n${tblStr}`);
      } else {
        sheetSections.push(tblStr);
      }
    }
    content = sheetSections.join("\n\n");
  } else if (parsedData.docType === "json") {
    content = settings.collapseWhitespace
      ? JSON.stringify(parsedData.data)
      : JSON.stringify(parsedData.data, null, 2);
  } else if (parsedData.docType === "text") {
    content = parsedData.rawText;
    if (settings.collapseWhitespace) {
      content = content
        .split("\n")
        .map((line) => line.replace(/[ \t]+/g, " ").trimEnd())
        .join("\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
    }
  }

  return content.trim();
}

function buildFinalOutput(filesToRender, settings) {
  if (filesToRender.length === 0) return "";

  if (settings.outputFormat === "json") {
    const payload = filesToRender.map((f) => ({
      file: f.name,
      type: f.detectedType,
      content: formatSingleFileContent(f, settings),
    }));
    return settings.collapseWhitespace ? JSON.stringify(payload) : JSON.stringify(payload, null, 2);
  }

  if (settings.outputFormat === "xml") {
    return filesToRender
      .map((f) => {
        const body = formatSingleFileContent(f, settings);
        return `<document name="${f.name}" type="${f.detectedType}">\n${body}\n</document>`;
      })
      .join("\n\n");
  }

  // Default: Markdown
  if (filesToRender.length === 1) {
    const f = filesToRender[0];
    return `# 📄 ${f.name}\n\n${formatSingleFileContent(f, settings)}`;
  }

  return filesToRender
    .map((f) => `# 📄 ${f.name} (${f.typeLabel})\n\n${formatSingleFileContent(f, settings)}`)
    .join("\n\n---\n\n");
}

/* =========================================================
 * Token 估算與流量計算
 * ========================================================= */
function estimateTokens(text) {
  if (!text) return 0;
  // CJK characters typically take ~1 to 1.5 tokens in modern tokenizers (o200k_base / cl100k_base / Gemini)
  // ASCII words/punctuation average ~1 token per 4 chars
  const cjkMatches = text.match(/[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]/g);
  const cjkCount = cjkMatches ? cjkMatches.length : 0;
  const nonCjkLength = text.length - cjkCount;
  return Math.ceil(cjkCount * 1.2 + nonCjkLength / 4);
}

function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

/* =========================================================
 * UI 渲染更新
 * ========================================================= */
function renderAll() {
  const settings = getSettings();
  fileCountEl.textContent = state.files.length;

  const hasFiles = state.files.length > 0;
  btnClearAll.disabled = !hasFiles;
  btnCopy.disabled = !hasFiles;
  btnCopyPrompt.disabled = !hasFiles;
  btnDownload.disabled = !hasFiles;

  // Render File List
  if (!hasFiles) {
    fileListEl.innerHTML = `<li class="empty-hint">尚未放入任何檔案</li>`;
  } else {
    fileListEl.innerHTML = "";
    state.files.forEach((f) => {
      const li = document.createElement("li");
      li.className = "file-item" + (state.selectedFileId === f.id ? " active" : "");

      const singleOut = formatSingleFileContent(f, settings);
      const cleanBytes = new Blob([singleOut]).size;
      const tokens = estimateTokens(singleOut);

      let badgeClass = "badge-text";
      if (f.detectedType === "pdf") badgeClass = "badge-pdf";
      else if (f.detectedType === "docx") badgeClass = "badge-word";
      else if (f.detectedType === "pptx") badgeClass = "badge-ppt";
      else if (f.detectedType === "xlsx" || f.detectedType === "csv") badgeClass = "badge-excel";

      li.innerHTML = `
        <div class="file-info">
          <span class="file-type-tag ${badgeClass}">${f.typeLabel}</span>
          <div class="file-meta">
            <div class="file-name" title="${f.name}">${f.name}</div>
            <div class="file-stats">${formatBytes(f.rawSize)} → ${formatBytes(cleanBytes)} · 約 ${tokens.toLocaleString()} tokens</div>
          </div>
        </div>
        <div class="file-actions">
          <button class="file-download" title="以原檔名下載此檔案" data-id="${f.id}">⬇</button>
          <button class="file-remove" title="移除此檔案" data-id="${f.id}">✕</button>
        </div>
      `;

      li.addEventListener("click", (e) => {
        if (e.target.closest(".file-remove") || e.target.closest(".file-download")) return;
        state.selectedFileId = state.selectedFileId === f.id ? null : f.id;
        renderAll();
      });

      li.querySelector(".file-download").addEventListener("click", (e) => {
        e.stopPropagation();
        downloadFiles([f]);
      });

      li.querySelector(".file-remove").addEventListener("click", (e) => {
        e.stopPropagation();
        state.files = state.files.filter((item) => item.id !== f.id);
        if (state.selectedFileId === f.id) state.selectedFileId = null;
        if (state.selectedPdfId === f.id) state.selectedPdfId = null;
        renderAll();
      });

      fileListEl.appendChild(li);
    });
  }

  // Determine active files to preview
  const targetFiles = state.selectedFileId
    ? state.files.filter((f) => f.id === state.selectedFileId)
    : state.files;

  if (state.selectedFileId && targetFiles.length === 1) {
    activeViewBadge.textContent = `單檔檢視：${targetFiles[0].name} (點擊切回全部)`;
  } else {
    activeViewBadge.textContent = `合併檢視全部檔案 (${state.files.length})`;
  }

  // Update download button label to show the preserved filename
  if (targetFiles.length === 1) {
    const dlName = getDownloadFileName(targetFiles, settings.outputFormat);
    btnDownload.textContent = `💾 下載 ${dlName}`;
    btnDownload.title = `以原檔名下載：${dlName}`;
  } else if (targetFiles.length > 1) {
    const dlName = getDownloadFileName(targetFiles, settings.outputFormat);
    btnDownload.textContent = `💾 下載檔案 (${targetFiles.length})`;
    btnDownload.title = `下載合併結果 (${dlName})，或點擊左側單一檔案旁的 ⬇ 個別以原檔名下載`;
  } else {
    btnDownload.textContent = `💾 下載檔案`;
    btnDownload.title = "";
  }

  const finalOutput = buildFinalOutput(targetFiles, settings);
  outputPreview.value = finalOutput;

  // Update Metrics
  const totalRawBytes = targetFiles.reduce((acc, f) => acc + f.rawSize, 0);
  const totalCleanBytes = new Blob([finalOutput]).size;
  const totalTokens = estimateTokens(finalOutput);

  metricRawSize.textContent = formatBytes(totalRawBytes);
  metricCleanSize.textContent = formatBytes(totalCleanBytes);
  metricTokens.textContent = totalTokens.toLocaleString();

  if (totalRawBytes > 0) {
    const savedRatio = Math.max(0, ((totalRawBytes - totalCleanBytes) / totalRawBytes) * 100);
    metricSizeSaved.textContent = `體積與流量減少 ${savedRatio.toFixed(1)}%`;
  } else {
    metricSizeSaved.textContent = "流量節省 0%";
  }

  // Compare with unoptimized Markdown table + uncollapsed whitespace
  if (hasFiles) {
    const unoptimizedOutput = buildFinalOutput(targetFiles, {
      ...settings,
      tableFormat: "markdown",
      collapseWhitespace: false,
      pdfSmartClean: false,
      trimEmptyCells: false,
      roundFloats: false,
      imagePlaceholder: true,
    });
    const unoptTokens = estimateTokens(unoptimizedOutput);
    if (unoptTokens > totalTokens) {
      const tokenDiffPct = (((unoptTokens - totalTokens) / unoptTokens) * 100).toFixed(0);
      metricTokenHint.textContent = `比一般未壓縮轉換再省 ${tokenDiffPct}% Token`;
    } else {
      metricTokenHint.textContent = `已去除所有 XML 與排版冗餘標籤`;
    }
  } else {
    metricTokenHint.textContent = "支援中英混合估算";
  }

  // Also update Tab 2 (PDF to PPT)
  renderPdf2PptPanel();
}

/* =========================================================
 * TAB 2: PDF 轉 PPT 簡報引擎（雙重混合 / 可編輯文字框 / 高保真圖片）
 * ========================================================= */
function getPptSettings() {
  const modeEl = document.getElementById("pptConvertMode");
  const layoutEl = document.getElementById("pptLayout");
  const scaleEl = document.getElementById("pptScale");
  const mergeEl = document.getElementById("optPptMergeBlocks");
  const notesEl = document.getElementById("optPptIncludeNotes");
  const hideBgEl = document.getElementById("optPptHideBgText");

  return {
    mode: modeEl ? modeEl.value : "hybrid",
    layout: layoutEl ? layoutEl.value : "auto",
    scale: scaleEl ? parseFloat(scaleEl.value) || 2 : 2,
    mergeBlocks: mergeEl ? mergeEl.checked : true,
    includeNotes: notesEl ? notesEl.checked : true,
    hideBgText: hideBgEl ? hideBgEl.checked : true,
  };
}

function getPdfFiles() {
  return state.files.filter((f) => f.detectedType === "pdf" && !f.error && f.parsedData);
}

function getActivePdfFile() {
  const pdfFiles = getPdfFiles();
  if (pdfFiles.length === 0) return null;
  if (state.selectedPdfId) {
    const found = pdfFiles.find((f) => f.id === state.selectedPdfId);
    if (found) return found;
  }
  state.selectedPdfId = pdfFiles[0].id;
  return pdfFiles[0];
}

/**
 * 將單頁 PDF 的文字區塊重組為 PPT 可編輯文字方塊（保留原始字型、顏色、粗斜體與精確座標）
 */
function getPptTextBlocksForPage(page, bodyFontSize, mergeBlocks) {
  const sourceList =
    page && page.pptSegments && page.pptSegments.length > 0
      ? page.pptSegments
      : (page && page.lines) || [];
  if (sourceList.length === 0) return [];

  const normalized = sourceList.map((item) => ({
    text: item.text,
    fontSize: item.fontSize || 12,
    fontFace: item.fontFace || "Microsoft JhengHei",
    isBold: Boolean(item.isBold || item.isHeading),
    isItalic: Boolean(item.isItalic),
    isHeading: Boolean(item.isHeading),
    isHeaderFooterOrPageNum: Boolean(item.isHeaderFooterOrPageNum),
    colorHex: item.colorHex || "0F172A",
    bgHex: item.bgHex || "FFFFFF",
    xPt: item.xPt ?? 40,
    yTopPt: item.yTopPt ?? 40,
    widthPt: item.widthPt ?? 200,
    heightPt: item.heightPt ?? (item.fontSize || 12) * 1.25,
  }));

  if (!mergeBlocks) {
    return normalized;
  }

  // Merge vertically adjacent lines ONLY when they share the same font, size, color, and left alignment
  const blocks = [];
  for (const item of normalized) {
    const prev = blocks[blocks.length - 1];
    if (prev) {
      const prevBottom = prev.yTopPt + prev.heightPt;
      const verticalGap = item.yTopPt - prevBottom;
      const sameStyle =
        Math.abs(prev.fontSize - item.fontSize) <= 1.5 &&
        prev.fontFace === item.fontFace &&
        prev.isBold === item.isBold &&
        prev.colorHex === item.colorHex;
      const alignedLeft = Math.abs(prev.xPt - item.xPt) <= prev.fontSize * 1.5;

      if (sameStyle && alignedLeft && verticalGap >= -3 && verticalGap <= prev.fontSize * 0.95) {
        prev.text += "\n" + item.text;
        const rightEdge = Math.max(prev.xPt + prev.widthPt, item.xPt + item.widthPt);
        prev.xPt = Math.min(prev.xPt, item.xPt);
        prev.widthPt = rightEdge - prev.xPt;
        prev.heightPt = Math.max(prev.heightPt + item.heightPt, item.yTopPt + item.heightPt - prev.yTopPt);
        continue;
      }
    }
    blocks.push({ ...item });
  }

  return blocks;
}

/**
 * 智慧抹除底圖上的原始文字像素（以文字上下邊緣的真實背景色進行垂直插值修補，100% 保留圖案與漸層原色、絕不變淡）
 */
function eraseTextRegionsOnCanvas(ctx, segments, scale, canvasW, canvasH) {
  if (!ctx || typeof ctx.getImageData !== "function" || !segments || segments.length === 0) return;

  for (const seg of segments) {
    try {
      const padX = Math.max(2, Math.round(seg.fontSize * scale * 0.12));
      const padY = Math.max(2, Math.round(seg.fontSize * scale * 0.15));
      const x0 = Math.max(0, Math.floor(seg.xPt * scale) - padX);
      const y0 = Math.max(0, Math.floor(seg.yTopPt * scale) - padY);
      const x1 = Math.min(canvasW - 1, Math.ceil((seg.xPt + seg.widthPt) * scale) + padX);
      const y1 = Math.min(canvasH - 1, Math.ceil((seg.yTopPt + seg.heightPt) * scale) + padY);
      const w = x1 - x0 + 1;
      const h = y1 - y0 + 1;
      if (w <= 2 || h <= 2) continue;

      const img = ctx.getImageData(x0, y0, w, h);
      const data = img.data;

      // For each vertical column x, linearly interpolate between top boundary pixel and bottom boundary pixel
      for (let cx = 0; cx < w; cx++) {
        const topIdx = cx * 4;
        const botIdx = ((h - 1) * w + cx) * 4;
        const rTop = data[topIdx];
        const gTop = data[topIdx + 1];
        const bTop = data[topIdx + 2];
        const rBot = data[botIdx];
        const gBot = data[botIdx + 1];
        const bBot = data[botIdx + 2];

        for (let cy = 1; cy < h - 1; cy++) {
          const t = cy / (h - 1);
          const idx = (cy * w + cx) * 4;
          data[idx] = Math.round(rTop + (rBot - rTop) * t);
          data[idx + 1] = Math.round(gTop + (gBot - gTop) * t);
          data[idx + 2] = Math.round(bTop + (bBot - bTop) * t);
          data[idx + 3] = 255;
        }
      }

      ctx.putImageData(img, x0, y0);
    } catch (_) {
      // Ignore individual region errors
    }
  }
}

function renderPdf2PptPanel() {
  if (!pdf2pptFileListEl || !pptSlidesPreview) return;

  const pdfFiles = getPdfFiles();
  const pptSettings = getPptSettings();
  pdf2pptFileCountEl.textContent = pdfFiles.length;

  // Update Mode Metric Display
  if (pptSettings.mode === "hybrid") {
    pptMetricMode.textContent = "原色圖案+文字框";
    pptMetricHint.textContent = "維持 100% 圖案原色 + 原色可編輯文字框";
  } else if (pptSettings.mode === "editable") {
    pptMetricMode.textContent = "純文字方塊";
    pptMetricHint.textContent = "依原座標、字型與顏色建立文字框";
  } else {
    pptMetricMode.textContent = "純圖片不跑版";
    pptMetricHint.textContent = "高畫質頁面背景 + 備忘稿全文";
  }

  if (pdfFiles.length === 0) {
    pdf2pptFileListEl.innerHTML = `<li class="empty-hint">尚未放入 PDF 檔案（在第一分頁載入的 PDF 也會自動同步至此）</li>`;
    pptMetricPages.textContent = "0 頁";
    pptMetricBlocks.textContent = "0 個";
    pptActiveFileBadge.textContent = "請先拖放 PDF 檔案";
    btnExportPptx.disabled = true;
    btnExportPptx.textContent = "📊 下載 PowerPoint (.pptx)";
    pptSlidesPreview.innerHTML = `
      <div class="slides-empty-state">
        <div class="empty-icon">📊</div>
        <h3>將 PDF 拖放到左側，即可預覽投影片並匯出 .pptx</h3>
        <p>支援保留原檔案名稱下載，亦可點擊右上角「載入測試範例」立即體驗 PDF 轉 PPT 效果</p>
      </div>
    `;
    return;
  }

  const activePdf = getActivePdfFile();

  // Render PDF File Queue List
  pdf2pptFileListEl.innerHTML = "";
  pdfFiles.forEach((f) => {
    const li = document.createElement("li");
    li.className = "file-item" + (activePdf && activePdf.id === f.id ? " active" : "");
    const pageCount = f.parsedData.pages.length;
    const pptxName = `${getBaseFileName(f.name)}.pptx`;

    li.innerHTML = `
      <div class="file-info">
        <span class="file-type-tag badge-pdf">PDF</span>
        <div class="file-meta">
          <div class="file-name" title="${f.name}">${f.name}</div>
          <div class="file-stats">共 ${pageCount} 頁 · 輸出檔名：${pptxName}</div>
        </div>
      </div>
      <div class="file-actions">
        <button class="file-download" title="立即轉換並下載 ${pptxName}">📊</button>
        <button class="file-remove" title="移除此檔案">✕</button>
      </div>
    `;

    li.addEventListener("click", (e) => {
      if (e.target.closest(".file-remove") || e.target.closest(".file-download")) return;
      state.selectedPdfId = f.id;
      renderAll();
    });

    li.querySelector(".file-download").addEventListener("click", (e) => {
      e.stopPropagation();
      exportPdfToPptx(f);
    });

    li.querySelector(".file-remove").addEventListener("click", (e) => {
      e.stopPropagation();
      state.files = state.files.filter((item) => item.id !== f.id);
      if (state.selectedFileId === f.id) state.selectedFileId = null;
      if (state.selectedPdfId === f.id) state.selectedPdfId = null;
      renderAll();
    });

    pdf2pptFileListEl.appendChild(li);
  });

  // Render Active PDF Slide Cards & Metrics
  const pages = activePdf.parsedData.pages;
  const bodyFontSize = activePdf.parsedData.bodyFontSize || 12;
  let totalBlocks = 0;

  const slideCardsHtml = pages
    .map((page) => {
      const blocks = getPptTextBlocksForPage(page, bodyFontSize, pptSettings.mergeBlocks);
      totalBlocks += blocks.length;

      const thumbHtml = page.thumbDataUrl
        ? `<img src="${page.thumbDataUrl}" alt="Slide ${page.pageNumber}" />`
        : `<div style="padding:36px 12px;color:#475569;font-size:0.8rem;text-align:center;">第 ${page.pageNumber} 頁預覽</div>`;

      const blockItemsHtml =
        blocks.length > 0
          ? blocks
              .map((b) => {
                const escaped = b.text
                  .replace(/&/g, "&amp;")
                  .replace(/</g, "&lt;")
                  .replace(/>/g, "&gt;")
                  .replace(/\n/g, " ↵ ");
                const weightLabel = b.isBold ? "粗體 " : "";
                return `
                  <div class="slide-text-box-item">
                    <span class="slide-text-box-meta" title="字型：${b.fontFace} / 顏色：#${b.colorHex}">
                      <span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:#${b.colorHex};border:1px solid rgba(255,255,255,0.4);vertical-align:middle;margin-right:3px;"></span>[${b.fontFace} ${weightLabel}${b.fontSize}pt]
                    </span>
                    <span>${escaped}</span>
                  </div>
                `;
              })
              .join("")
          : `<div class="slide-text-box-item"><span>（此頁為純圖片或無文字圖層，將以高畫質影像呈現）</span></div>`;

      const modeBadgeText =
        pptSettings.mode === "image"
          ? `高畫質背景圖 + 備忘稿文字`
          : `原色可編輯文字方塊 x ${blocks.length}`;

      return `
        <div class="slide-card">
          <div class="slide-thumb-wrap">
            <span class="slide-num-tag">Slide ${page.pageNumber}</span>
            ${thumbHtml}
          </div>
          <div class="slide-details">
            <div class="slide-details-header">
              <h4>第 ${page.pageNumber} 頁投影片 (${Math.round(page.pageWidthPt)} × ${Math.round(page.pageHeightPt)} pt)</h4>
              <span class="slide-box-badge">${modeBadgeText}</span>
            </div>
            <div class="slide-text-list">
              ${blockItemsHtml}
            </div>
          </div>
        </div>
      `;
    })
    .join("");

  pptSlidesPreview.innerHTML = slideCardsHtml;
  pptMetricPages.textContent = `${pages.length} 頁`;
  pptMetricBlocks.textContent = pptSettings.mode === "image" ? `0 個 (純圖+備忘)` : `${totalBlocks} 個`;
  pptActiveFileBadge.textContent = `預覽中：${activePdf.name}`;

  const outPptxName = `${getBaseFileName(activePdf.name)}.pptx`;
  btnExportPptx.disabled = false;
  btnExportPptx.textContent = `📊 下載 ${outPptxName}`;
  btnExportPptx.title = `將 ${activePdf.name} 轉換為 ${outPptxName}`;
}

/**
 * 將指定 PDF 檔案依高解析度渲染並轉換為可編輯 .pptx 簡報下載（維持圖案原色、還原文字顏色與字型）
 */
async function exportPdfToPptx(pdfFileObj) {
  if (!pdfFileObj || !pdfFileObj.parsedData) return;

  const pptSettings = getPptSettings();
  const outFileName = `${getBaseFileName(pdfFileObj.name)}.pptx`;
  const origBtnText = btnExportPptx ? btnExportPptx.textContent : "";

  try {
    if (btnExportPptx) {
      btnExportPptx.disabled = true;
      btnExportPptx.textContent = "⏳ 正在生成 PPTX...";
    }
    showToast(`正在將 ${pdfFileObj.name} 轉換為 PowerPoint 簡報...`);

    const { pages, bodyFontSize, pdfBytes } = pdfFileObj.parsedData;
    const firstPage = pages[0] || { pageWidthPt: 612, pageHeightPt: 792 };

    // Determine Slide Dimensions in Inches
    let slideW = 10;
    let slideH = 5.625;
    if (pptSettings.layout === "16x9") {
      slideW = 10;
      slideH = 5.625;
    } else if (pptSettings.layout === "16x10") {
      slideW = 10;
      slideH = 6.25;
    } else if (pptSettings.layout === "4x3") {
      slideW = 10;
      slideH = 7.5;
    } else {
      // Auto match PDF page aspect ratio (base width = 10 inches)
      const ratio = (firstPage.pageHeightPt || 792) / (firstPage.pageWidthPt || 612);
      slideW = 10;
      slideH = Number(Math.max(4, Math.min(14, 10 * ratio)).toFixed(3));
    }

    // Render high-res page images (and cleanly inpaint/erase old raster text in hybrid mode without dimming graphics!)
    const pageImages = [];
    if ((pptSettings.mode === "image" || pptSettings.mode === "hybrid") && pdfBytes && window.pdfjsLib) {
      const loadingTask = window.pdfjsLib.getDocument({
        data: new Uint8Array(pdfBytes),
        useSystemFonts: true,
        disableFontFace: false,
      });
      const pdfDoc = await loadingTask.promise;

      for (let i = 1; i <= pdfDoc.numPages; i++) {
        const page = await pdfDoc.getPage(i);
        const vp = page.getViewport({ scale: pptSettings.scale });
        const canvas = document.createElement("canvas");
        if (canvas && typeof canvas.getContext === "function") {
          canvas.width = Math.ceil(vp.width);
          canvas.height = Math.ceil(vp.height);
          const ctx = canvas.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          await page.render({ canvasContext: ctx, viewport: vp }).promise;

          // If in hybrid mode and optPptHideBgText ("智慧抹除底圖原文字") is checked:
          // Erase ONLY the text glyph regions by vertical edge interpolation so 100% of graphics stay vivid!
          if (pptSettings.mode === "hybrid" && pptSettings.hideBgText) {
            const segs = pages[i - 1]?.pptSegments || pages[i - 1]?.lines || [];
            eraseTextRegionsOnCanvas(ctx, segs, pptSettings.scale, canvas.width, canvas.height);
          }

          pageImages.push(canvas.toDataURL("image/png"));
        } else {
          pageImages.push(pages[i - 1]?.thumbDataUrl || "");
        }
      }
    }

    // Use PptxGenJS if available, otherwise fallback to native JSZip OpenXML builder
    if (window.PptxGenJS) {
      const pptx = new window.PptxGenJS();
      pptx.defineLayout({ name: "PDF_CUSTOM", width: slideW, height: slideH });
      pptx.layout = "PDF_CUSTOM";
      pptx.title = getBaseFileName(pdfFileObj.name);

      pages.forEach((page, idx) => {
        const slide = pptx.addSlide();
        const pW = page.pageWidthPt || 612;
        const pH = page.pageHeightPt || 792;
        const blocks = getPptTextBlocksForPage(page, bodyFontSize || 12, pptSettings.mergeBlocks);
        const imgData = pageImages[idx] || page.thumbDataUrl || "";

        // 1. Add 100% Original Vivid Background Layer for "image" or "hybrid" mode (NO dimming overlay!)
        if ((pptSettings.mode === "image" || pptSettings.mode === "hybrid") && imgData) {
          slide.addImage({
            data: imgData,
            x: 0,
            y: 0,
            w: slideW,
            h: slideH,
          });
        }

        // 2. Add Editable Text Boxes with exact original color, font family, bold/italic, and 0 margin
        if (pptSettings.mode === "editable" || pptSettings.mode === "hybrid") {
          const ptScale = (slideW * 72) / pW;
          for (const b of blocks) {
            const xIn = Number(Math.max(0, Math.min(slideW - 0.3, (b.xPt / pW) * slideW)).toFixed(3));
            const yIn = Number(Math.max(0, Math.min(slideH - 0.2, (b.yTopPt / pH) * slideH)).toFixed(3));
            // Give a slight width buffer so single-line text boxes don't prematurely wrap in PowerPoint
            const wIn = Number(
              Math.max(0.5, Math.min(slideW - xIn, ((b.widthPt + b.fontSize * 1.1) / pW) * slideW)).toFixed(3)
            );
            const hIn = Number(
              Math.max(0.24, Math.min(slideH - yIn, ((b.heightPt + b.fontSize * 0.35) / pH) * slideH)).toFixed(3)
            );
            const scaledFontPt = Math.max(7, Math.min(64, Math.round(b.fontSize * ptScale * 0.94)));

            slide.addText(b.text, {
              x: xIn,
              y: yIn,
              w: wIn,
              h: hIn,
              fontFace: b.fontFace || "Microsoft JhengHei",
              fontSize: scaledFontPt,
              bold: Boolean(b.isBold),
              italic: Boolean(b.isItalic),
              color: b.colorHex || "0F172A",
              valign: "top",
              margin: 0,
              wrap: true,
            });
          }
        }

        // 3. Add Speaker Notes if enabled
        if (pptSettings.includeNotes || pptSettings.mode === "image") {
          const noteText = page.lines.map((l) => l.text).join("\n");
          if (noteText) {
            slide.addNotes(noteText);
          }
        }
      });

      await pptx.writeFile({ fileName: outFileName });
    } else {
      // Built-in JSZip OpenXML PPTX fallback generator
      const blob = await buildPptxArchiveWithJSZip(pages, bodyFontSize, slideW, slideH, pptSettings, pageImages);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outFileName;
      a.click();
      URL.revokeObjectURL(url);
    }

    showToast(`已成功下載 ${outFileName}！`);
  } catch (err) {
    console.error("匯出 PPTX 失敗:", err);
    showToast(`匯出 PPTX 失敗：${err.message || "未知錯誤"}`);
  } finally {
    if (btnExportPptx) {
      btnExportPptx.disabled = false;
      btnExportPptx.textContent = origBtnText || `📊 下載 ${outFileName}`;
    }
  }
}

/**
 * 原生 JSZip OpenXML (.pptx) 備援產生器（支援原始字型、文字顏色與無邊距文字方塊）
 */
async function buildPptxArchiveWithJSZip(pages, bodyFontSize, slideW, slideH, pptSettings, pageImages) {
  const zip = new JSZip();
  const emuPerInch = 914400;
  const cx = Math.round(slideW * emuPerInch);
  const cy = Math.round(slideH * emuPerInch);

  const escapeXml = (s) =>
    String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const slideOverrides = pages
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    )
    .join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      ${slideOverrides}
    </Types>`
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships>`
  );

  const sldIdLst = pages.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:sldIdLst>${sldIdLst}</p:sldIdLst>
      <p:sldSz cx="${cx}" cy="${cy}"/>
      <p:notesSz cx="${cy}" cy="${cx}"/>
    </p:presentation>`
  );

  const presRels = pages
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    )
    .join("");

  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${presRels}
    </Relationships>`
  );

  pages.forEach((page, idx) => {
    const slideNum = idx + 1;
    const pW = page.pageWidthPt || 612;
    const pH = page.pageHeightPt || 792;
    const ptScale = (slideW * 72) / pW;
    const blocks = getPptTextBlocksForPage(page, bodyFontSize || 12, pptSettings.mergeBlocks);

    let shapesXml = "";
    let shapeId = 2;

    if (pptSettings.mode === "editable" || pptSettings.mode === "hybrid") {
      blocks.forEach((b) => {
        const xEmu = Math.round((b.xPt / pW) * cx);
        const yEmu = Math.round((b.yTopPt / pH) * cy);
        const wEmu = Math.round(Math.max(0.5 * emuPerInch, ((b.widthPt + b.fontSize * 1.1) / pW) * cx));
        const hEmu = Math.round(Math.max(0.24 * emuPerInch, ((b.heightPt + b.fontSize * 0.35) / pH) * cy));
        const fontHundredths = Math.max(700, Math.round(b.fontSize * ptScale * 0.94 * 100));
        const colorHex = b.colorHex || "0F172A";
        const fontFace = escapeXml(b.fontFace || "Microsoft JhengHei");

        const parasXml = b.text
          .split("\n")
          .map(
            (line) =>
              `<a:p><a:r><a:rPr lang="zh-TW" sz="${fontHundredths}" b="${b.isBold ? 1 : 0}" i="${b.isItalic ? 1 : 0}"><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill><a:latin typeface="${fontFace}"/><a:ea typeface="${fontFace}"/></a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
          )
          .join("");

        shapesXml += `
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="${shapeId++}" name="TextBox ${shapeId}"/>
              <p:cNvSpPr txBox="1"/>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="${xEmu}" y="${yEmu}"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:noFill/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0"/>
              <a:lstStyle/>
              ${parasXml}
            </p:txBody>
          </p:sp>
        `;
      });
    }

    zip.file(
      `ppt/slides/slide${slideNum}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${shapesXml}
          </p:spTree>
        </p:cSld>
      </p:sld>`
    );
  });

  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

/* =========================================================
 * 輔助功能：複製、下載、產生真實測試檔案 (PDF / Word / PPT / Excel)
 * ========================================================= */
async function copyText(text, msg) {
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    showToast(msg);
  } catch (_) {
    outputPreview.select();
    document.execCommand("copy");
    showToast(msg);
  }
}

function getBaseFileName(fileName) {
  if (!fileName) return "document";
  const lastDot = fileName.lastIndexOf(".");
  return lastDot > 0 ? fileName.slice(0, lastDot) : fileName;
}

function getOutputExtension(outputFormat) {
  return outputFormat === "json" ? "json" : outputFormat === "xml" ? "xml" : "md";
}

function getDownloadFileName(filesToDownload, outputFormat) {
  const ext = getOutputExtension(outputFormat);
  if (!filesToDownload || filesToDownload.length === 0) {
    return `document.${ext}`;
  }
  const firstBase = getBaseFileName(filesToDownload[0].name);
  if (filesToDownload.length === 1) {
    return `${firstBase}.${ext}`;
  }
  return `${firstBase}_等${filesToDownload.length}個檔案.${ext}`;
}

function downloadFiles(filesToDownload) {
  const settings = getSettings();
  const text = buildFinalOutput(filesToDownload, settings);
  if (!text) return;
  const downloadName = getDownloadFileName(filesToDownload, settings.outputFormat);
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = downloadName;
  a.click();
  URL.revokeObjectURL(url);
  showToast(`已下載 ${downloadName}`);
}

function downloadResult() {
  const targetFiles = state.selectedFileId
    ? state.files.filter((f) => f.id === state.selectedFileId)
    : state.files;
  downloadFiles(targetFiles);
}

let toastTimer = null;
function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.add("hidden");
  }, 2600);
}

/**
 * 動態建立一個符合 PDF 1.4 規範的雙頁測試 PDF 二進位檔（含重複頁首頁尾、標題字級與斷行段落）
 */
function buildSamplePdfBytes() {
  const page1Stream = [
    "q 0.11 0.23 0.54 rg 40 695 532 42 re f Q",
    "q 0.93 0.96 1.00 rg 40 595 532 70 re f Q",
    "BT /F1 9 Tf 0.45 0.50 0.58 rg 50 760 Td (ACME Enterprise Confidential Report - 2026) Tj ET",
    "BT /F1 18 Tf 1 1 1 rg 52 710 Td (Executive Summary: AI Document Pipeline) Tj ET",
    "BT /F1 11 Tf 0.10 0.15 0.25 rg 50 672 Td (Traditional PDF files contain repetitive running headers, page numbers,) Tj ET",
    "BT /F1 11 Tf 0.10 0.15 0.25 rg 50 656 Td (and hard line-breaks at the end of every visual margin that waste tokens.) Tj ET",
    "BT /F1 12 Tf 0.08 0.38 0.75 rg 52 636 Td (- Automated header and footer deduplication across all pages) Tj ET",
    "BT /F1 12 Tf 0.82 0.16 0.16 rg 52 614 Td (- Exact color, font family, and editable text-box reconstruction) Tj ET",
    "BT /F1 9 Tf 0.45 0.50 0.58 rg 260 40 Td (Page 1 of 2) Tj ET",
  ].join("\n");

  const page2Stream = [
    "q 0.09 0.45 0.33 rg 40 695 532 42 re f Q",
    "BT /F1 9 Tf 0.45 0.50 0.58 rg 50 760 Td (ACME Enterprise Confidential Report - 2026) Tj ET",
    "BT /F1 16 Tf 1 1 1 rg 52 710 Td (Benchmark & Token Savings Analysis) Tj ET",
    "BT /F1 11 Tf 0.10 0.15 0.25 rg 50 672 Td (By merging broken lines and stripping repeated page metadata, LLM context) Tj ET",
    "BT /F1 11 Tf 0.10 0.15 0.25 rg 50 656 Td (comprehension improves while reducing prompt token consumption by 25-45%.) Tj ET",
    "BT /F1 9 Tf 0.45 0.50 0.58 rg 260 40 Td (Page 2 of 2) Tj ET",
  ].join("\n");

  const objects = [
    "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
    "2 0 obj\n<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>\nendobj\n",
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>\nendobj\n",
    "4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >>\nendobj\n",
    "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>\nendobj\n",
    `6 0 obj\n<< /Length ${page1Stream.length} >>\nstream\n${page1Stream}\nendstream\nendobj\n`,
    `7 0 obj\n<< /Length ${page2Stream.length} >>\nstream\n${page2Stream}\nendstream\nendobj\n`,
  ];

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (const obj of objects) {
    offsets.push(pdf.length);
    pdf += obj;
  }
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= objects.length; i++) {
    pdf += String(offsets[i]).padStart(10, "0") + " 00000 n \n";
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new TextEncoder().encode(pdf);
}

/**
 * 動態在瀏覽器記憶體中生成真實的 .pdf, .docx, .pptx, .xlsx 測試檔案，驗證完整解析流程
 */
async function loadDemoFiles() {
  btnDemo.disabled = true;
  btnDemo.textContent = "產生範例中...";
  try {
    const demoFiles = [];

    // 0. 建立真實 .pdf 檔案（雙頁，含重複頁首頁尾與斷行）
    const pdfBytes = buildSamplePdfBytes();
    demoFiles.push(new File([pdfBytes], "AI架構白皮書_雙頁報告.pdf", { type: "application/pdf" }));

    // 1. 建立真實 .docx 檔案
    const docxZip = new JSZip();
    docxZip.file(
      "word/document.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
            <w:r><w:t>2026 年度智慧辦公系統升級企劃書</w:t></w:r>
          </w:p>
          <w:p>
            <w:r><w:t>本文件說明部門導入 AI 文件前處理管線之核心目標，透過剔除冗餘 OpenXML 樣式標籤與裝飾圖檔，大幅降低 LLM API 呼叫成本與延遲。</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
            <w:r><w:t>核心效益指標</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
            <w:r><w:t>平均減少 85% 以上的原始檔案體積與傳輸流量</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
            <w:r><w:t>表格轉換為緊湊 CSV 後，比傳統 Markdown 表格再省 40% Token</w:t></w:r>
          </w:p>
          <w:tbl>
            <w:tr>
              <w:tc><w:p><w:r><w:t>階段</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>交付項目</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>預計完成日</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Phase 1</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>Word/PPT/Excel 自動辨識引擎</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>2026-Q4</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Phase 2</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>企業內部知識庫 RAG 串接</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>2027-Q1</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>`
    );
    // Add simulated padding XML metadata to mimic real Word bloat
    docxZip.file("word/styles.xml", "<w:styles>" + "<!-- style bloat -->".repeat(800) + "</w:styles>");
    const docxBlob = await docxZip.generateAsync({ type: "blob" });
    demoFiles.push(new File([docxBlob], "專案企劃書_v2.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));

    // 2. 建立真實 .pptx 檔案
    const pptxZip = new JSZip();
    pptxZip.file("ppt/presentation.xml", `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`);
    pptxZip.file(
      "ppt/slides/slide1.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
            <p:txBody><a:p><a:r><a:t>Q4 營運成果與 AI 轉型簡報</a:t></a:r></a:p></p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
            <p:txBody>
              <a:p><a:pPr lvl="0"/><a:r><a:t>本季營收成長達 28%，超越年度目標</a:t></a:r></a:p>
              <a:p><a:pPr lvl="1"/><a:r><a:t>亞太區企業客戶續約率提升至 94%</a:t></a:r></a:p>
              <a:p><a:pPr lvl="0"/><a:r><a:t>客服與文件自動化系統每月節省 1,200 工時</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:pic/>
        </p:spTree></p:cSld>
      </p:sld>`
    );
    pptxZip.file(
      "ppt/notesSlides/notesSlide1.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
      <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
            <p:txBody><a:p><a:r><a:t>向董事會強調：亞太區成長主要來自新推出的自動化分析模組。</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld>
      </p:notes>`
    );
    pptxZip.file("ppt/theme/theme1.xml", "<a:theme>" + "<!-- theme bloat -->".repeat(1000) + "</a:theme>");
    const pptxBlob = await pptxZip.generateAsync({ type: "blob" });
    demoFiles.push(new File([pptxBlob], "Q4營運簡報.pptx", { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));

    // 3. 建立真實 .xlsx 檔案（包含空白列、空白欄與浮點數精度誤差）
    const wb = XLSX.utils.book_new();
    const wsData = [
      ["部門名稱", "", "Q3 預算 (萬)", "Q4 實際支出 (萬)", "達成率", ""],
      ["", "", "", "", "", ""], // 空白列（將被自動剔除）
      ["研發部", "", 120.5, 118.2, 0.98091286307, ""],
      ["行銷部", "", 85.0, 82.45, 0.30000000000000004, ""],
      ["客服部", "", 45.0, 41.2, 0.91555555555, ""],
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, "部門預算表");
    const emptyWs = XLSX.utils.aoa_to_sheet([["", ""], ["", ""]]);
    XLSX.utils.book_append_sheet(wb, emptyWs, "空白工作表");
    const xlsxBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    demoFiles.push(new File([xlsxBuffer], "部門財務報表.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));

    await handleFiles(demoFiles);
  } finally {
    btnDemo.disabled = false;
    btnDemo.textContent = "載入測試範例";
  }
}

init();
