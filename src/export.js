import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const CAT_COLORS = ["3b82f6", "8b5cf6", "10b981", "f59e0b", "ec4899", "6366f1"];
const thin = { style: "thin", color: { argb: "FFcbd5e1" } };
const bdr = { top: thin, bottom: thin, left: thin, right: thin };

function sc(ws, r, c, value, { font, fill, align, fmt, wrap, noBorder } = {}) {
  const cell = ws.getCell(r, c);
  if (value !== undefined && value !== null) cell.value = value;
  if (font) cell.font = font;
  if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + fill } };
  if (fmt) cell.numFmt = fmt;
  cell.alignment = align || { horizontal: "center", vertical: "middle", wrapText: !!wrap };
  if (!noBorder) cell.border = bdr;
  return cell;
}

function colLetter(colIdx) {
  let result = "";
  let n = colIdx - 1;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export async function exportToExcel(categories, modelName, modelDescription, numOptions, optionNames) {
  const wb = new ExcelJS.Workbook();
  const name = modelName || "Weighted Factor Model";
  const desc = modelDescription || "";

  const processed = categories.map(cat => {
    const hasSub = cat.criteria.length > 1 ||
      (cat.criteria.length === 1 && cat.criteria[0].name.trim() !== "" && cat.criteria[0].name.trim() !== cat.name.trim());
    return { ...cat, hasSub };
  });

  // ============ OVERVIEW SHEET ============
  const wsO = wb.addWorksheet("Overview", { properties: { tabColor: { argb: "FF1a1a2e" } } });
  wsO.getColumn(1).width = 4;
  wsO.getColumn(2).width = 24;
  wsO.getColumn(3).width = 70;

  // Title
  wsO.mergeCells("B2:C2");
  sc(wsO, 2, 2, name, { font: { name: "Arial", size: 20, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });

  // Purpose
  let oRow = 3;
  if (desc) {
    wsO.mergeCells("B3:C3");
    sc(wsO, 3, 2, desc, { font: { name: "Arial", size: 11, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
    oRow = 4;
  }

  // Spacer
  oRow += 1;

  // How to use
  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "How to use this workbook", { font: { name: "Arial", size: 13, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
  oRow += 1;

  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "1.  Review the Weights tab and adjust criteria weights if needed (blue cells are editable).", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
  oRow += 1;

  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "2.  Go to the Input tab and score each option against every criterion (blue cells).", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
  oRow += 1;

  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "3.  Check the Results tab to see your options ranked by weighted z-score.", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
  oRow += 2;

  // Tab guide
  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "Tabs in this workbook", { font: { name: "Arial", size: 13, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });
  oRow += 1;

  const tabs = [
    { tab: "Overview", desc: "This page. Summary and instructions." },
    { tab: "Input", desc: "Enter your raw scores here. Blue cells are editable. Criteria scores are calculated automatically for criteria with sub-criteria." },
    { tab: "Z-Score Normalisation", desc: "Standardised view of your data. Each criterion is normalised to mean=0, std=1, removing scale differences. Weighted totals and ranks are calculated here." },
    { tab: "Results", desc: "Clean ranked output. Options sorted from best to worst based on weighted z-scores." },
    { tab: "Definitions", desc: "Reference sheet listing each criterion with its definition and source. Fill in as needed." },
    { tab: "Weights", desc: "Master weight structure. All other sheets pull their weights from here. Edit the blue cells to adjust." },
  ];

  // Header row
  sc(wsO, oRow, 2, "Tab", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });
  sc(wsO, oRow, 3, "Description", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
  oRow += 1;

  tabs.forEach(({ tab, desc: d }) => {
    sc(wsO, oRow, 2, tab, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "f8fafc", align: { horizontal: "left", vertical: "middle" } });
    sc(wsO, oRow, 3, d, { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle", wrapText: true } });
    oRow += 1;
  });

  oRow += 1;
  wsO.mergeCells("B" + oRow + ":C" + oRow);
  sc(wsO, oRow, 2, "Generated by the Weighted Factor Model Builder \u00A9 Jacco Rubens", { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" }, noBorder: true });

  // Hide gridlines on Overview for a clean look
  wsO.views = [{ showGridLines: false }];

  // ============ WEIGHTS SHEET ============
  const wsW = wb.addWorksheet("Weights", { properties: { tabColor: { argb: "FF10b981" } } });
  wsW.mergeCells("A1:F1");
  sc(wsW, 1, 1, name + " - Weight Structure", { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
  let wInstrRow = 2;
  if (desc) {
    sc(wsW, 2, 1, desc, { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });
    wInstrRow = 3;
  }
  sc(wsW, wInstrRow, 1, "Edit weights here (blue cells). All other sheets reference this tab.", { font: { name: "Arial", size: 9, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });

  const wHdrRow = wInstrRow + 1;
  const wHeaders = ["Criteria", "Criteria Weight", "Sub-criterion", "Sub-criteria Weight", "Effective Weight", "Check"];
  wHeaders.forEach((h, i) => sc(wsW, wHdrRow, i + 1, h, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "10b981" }));
  [22, 16, 26, 16, 18, 30].forEach((w, i) => { wsW.getColumn(i + 1).width = w; });

  const catWeightCells = {};
  const critWeightCells = {};
  const catRowRanges = {};
  let wr = wHdrRow + 1;

  processed.forEach((cat, ci) => {
    const catStart = wr;
    if (cat.hasSub) {
      cat.criteria.forEach((crit, cri) => {
        if (cri === 0) {
          sc(wsW, wr, 1, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "f1f5f9", align: { horizontal: "left", vertical: "middle" } });
          sc(wsW, wr, 2, cat.weight / 100, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF0000FF" } }, fill: "eff6ff", fmt: "0%" });
          catWeightCells[ci] = "B" + wr;
        } else {
          sc(wsW, wr, 1, "", { fill: "f1f5f9" });
          sc(wsW, wr, 2, "", { fill: "f1f5f9" });
        }
        sc(wsW, wr, 3, crit.name, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });
        sc(wsW, wr, 4, crit.weight / 100, { font: { name: "Arial", size: 10, color: { argb: "FF0000FF" } }, fill: "eff6ff", fmt: "0%" });
        critWeightCells[ci + "-" + cri] = "D" + wr;
        wr++;
      });
    } else {
      sc(wsW, wr, 1, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "f1f5f9", align: { horizontal: "left", vertical: "middle" } });
      sc(wsW, wr, 2, cat.weight / 100, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF0000FF" } }, fill: "eff6ff", fmt: "0%" });
      catWeightCells[ci] = "B" + wr;
      sc(wsW, wr, 3, "(single criterion)", { font: { name: "Arial", size: 10, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });
      sc(wsW, wr, 4, 1, { font: { name: "Arial", size: 10, color: { argb: "FF94a3b8" } }, fmt: "0%" });
      critWeightCells[ci + "-0"] = "D" + wr;
      wr++;
    }
    catRowRanges[ci] = [catStart, wr - 1];
    if (cat.hasSub && cat.criteria.length > 1) {
      wsW.mergeCells(catStart, 1, wr - 1, 1);
      wsW.mergeCells(catStart, 2, wr - 1, 2);
    }
  });

  const firstDataRow = wHdrRow + 1;
  processed.forEach((cat, ci) => {
    const [s, e] = catRowRanges[ci];
    const cwRef = catWeightCells[ci];
    const critRange = "D" + s + ":D" + e;
    for (let row = s; row <= e; row++) {
      sc(wsW, row, 5, { formula: "IFERROR(" + cwRef + "*(D" + row + "/SUM(" + critRange + ")),0)" }, { font: { name: "Arial", size: 10, bold: true }, fmt: "0.0%" });
    }
    sc(wsW, s, 6, { formula: 'IF(SUM(' + critRange + ')=0,"Empty",IF(ABS(SUM(' + critRange + ')-1)<0.001,"OK","Sum: "&TEXT(SUM(' + critRange + '),"0%")))' }, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFdc2626" } }, fill: "fefce8" });
    for (let row = s + 1; row <= e; row++) sc(wsW, row, 6, "", { fill: "fefce8" });
  });

  wr += 1;
  const catSumParts = processed.map((_, ci) => catWeightCells[ci]).join("+");
  sc(wsW, wr, 1, "TOTAL CHECK", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "e2e8f0", align: { horizontal: "left", vertical: "middle" } });
  sc(wsW, wr, 2, { formula: catSumParts }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0%" });
  sc(wsW, wr, 3, "", { fill: "e2e8f0" }); sc(wsW, wr, 4, "", { fill: "e2e8f0" });
  sc(wsW, wr, 5, { formula: "SUM(E" + firstDataRow + ":E" + (wr - 2) + ")" }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0.0%" });
  sc(wsW, wr, 6, { formula: 'IF(ABS(' + catSumParts + '-1)<0.001,"All weights balanced","Criteria weights sum to "&TEXT(' + catSumParts + ',"0%")&" (need 100%)")' }, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFdc2626" } }, fill: "fefce8" });

  // Conditional formatting: formula-based for Google Sheets compatibility
  // Apply per-cell rules for each check cell
  processed.forEach((cat, ci) => {
    const [s] = catRowRanges[ci];
    const cellRef = "F" + s;
    // Green when OK
    wsW.addConditionalFormatting({
      ref: cellRef,
      rules: [{
        type: "expression",
        formulae: ['ISNUMBER(SEARCH("OK",F' + s + '))'],
        style: { font: { color: { argb: "FF16a34a" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFdcfce7" } } },
        priority: 1,
      }],
    });
    // Red when not OK
    wsW.addConditionalFormatting({
      ref: cellRef,
      rules: [{
        type: "expression",
        formulae: ['NOT(ISNUMBER(SEARCH("OK",F' + s + ')))'],
        style: { font: { color: { argb: "FFdc2626" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFfee2e2" } } },
        priority: 2,
      }],
    });
  });
  // Total check row
  const totalCheckRef = "F" + wr;
  wsW.addConditionalFormatting({
    ref: totalCheckRef,
    rules: [{
      type: "expression",
      formulae: ['ISNUMBER(SEARCH("balanced",F' + wr + '))'],
      style: { font: { color: { argb: "FF16a34a" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFdcfce7" } } },
      priority: 1,
    }],
  });
  wsW.addConditionalFormatting({
    ref: totalCheckRef,
    rules: [{
      type: "expression",
      formulae: ['NOT(ISNUMBER(SEARCH("balanced",F' + wr + ')))'],
      style: { font: { color: { argb: "FFdc2626" } }, fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFfee2e2" } } },
      priority: 2,
    }],
  });

  // ============ SHARED CONSTANTS ============
  // Input tab has no total/rank; Z-Score tab has them
  const R_CAT = 4, R_CW = 5, R_CRIT = 6, R_CW2 = 7, R_EW = 8, R_HDR = 9, R_OPT = 10;

  // ============ INPUT SHEET (data entry, category scores but no total/rank) ============
  function buildInputSheet() {
    const ws = wb.addWorksheet("Input", { properties: { tabColor: { argb: "FF3b82f6" } } });
    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, name + " - Data Input", { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
    let instrRow = 2;
    if (desc) {
      sc(ws, 2, 1, desc, { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });
      instrRow = 3;
    }
    sc(ws, instrRow, 1, "Enter data in the blue cells. Weights are pulled from the Weights tab.", { font: { name: "Arial", size: 9, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });

    const OC = 1;
    const DC = 2;

    // Column map (with category score columns for multi-criteria categories)
    let col = DC;
    const catColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = col;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      col += numCrit;
      const cscoreCol = cat.hasSub ? col : null;
      if (cat.hasSub) col++;
      catColInfo.push({ critStart, critEnd: critStart + numCrit - 1, cscoreCol, ci });
    });

    // Label col
    for (let rr = R_CAT; rr <= R_EW; rr++) sc(ws, rr, OC, "", { fill: "f8fafc" });
    sc(ws, R_CAT, OC, "Criteria", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW, OC, "Criteria weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CRIT, OC, "Sub-criterion", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle", wrapText: true } });
    sc(ws, R_CW2, OC, "Sub-criteria weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_EW, OC, "Effective weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_HDR, OC, "Option", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
    ws.getColumn(OC).width = 20;

    catColInfo.forEach(({ critStart, critEnd, cscoreCol, ci }) => {
      const cat = processed[ci];
      const color = CAT_COLORS[ci % CAT_COLORS.length];
      const [s_w] = catRowRanges[ci];
      const cwRef = catWeightCells[ci];
      const mergeEnd = cscoreCol || critEnd;

      if (mergeEnd > critStart) ws.mergeCells(R_CAT, critStart, R_CAT, mergeEnd);
      sc(ws, R_CAT, critStart, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: color });
      if (mergeEnd > critStart) ws.mergeCells(R_CW, critStart, R_CW, mergeEnd);
      sc(ws, R_CW, critStart, { formula: "Weights!" + cwRef }, { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });

      if (cat.hasSub) {
        cat.criteria.forEach((crit, j) => {
          const cc = critStart + j;
          sc(ws, R_CRIT, cc, crit.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc", align: { horizontal: "center", vertical: "middle", wrapText: true } });
          sc(ws, R_CW2, cc, { formula: "Weights!" + critWeightCells[ci + "-" + j] }, { font: { name: "Arial", size: 9, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });
          sc(ws, R_EW, cc, { formula: "Weights!E" + (s_w + j) }, { font: { name: "Arial", size: 9 }, fill: "f1f5f9", fmt: "0.0%" });
          sc(ws, R_HDR, cc, "", { fill: "e2e8f0" });
          ws.getColumn(cc).width = 16;
        });
        sc(ws, R_CRIT, cscoreCol, cat.name + " Score", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF" + color } }, fill: "f0fdf4", wrap: true });
        sc(ws, R_CW2, cscoreCol, "", { fill: "f0fdf4" });
        sc(ws, R_EW, cscoreCol, "", { fill: "f0fdf4" });
        sc(ws, R_HDR, cscoreCol, "", { fill: "f0fdf4" });
        ws.getColumn(cscoreCol).width = 14;
      } else {
        const cc = critStart;
        sc(ws, R_CRIT, cc, cat.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc", align: { horizontal: "center", vertical: "middle", wrapText: true } });
        sc(ws, R_CW2, cc, "", { fill: "f1f5f9" });
        sc(ws, R_EW, cc, { formula: "Weights!E" + s_w }, { font: { name: "Arial", size: 9 }, fill: "f1f5f9", fmt: "0.0%" });
        sc(ws, R_HDR, cc, "", { fill: "e2e8f0" });
        ws.getColumn(cc).width = 16;
      }
    });

    for (let opt = 0; opt < numOptions; opt++) {
      const r = R_OPT + opt;
      const optName = (optionNames[opt] && optionNames[opt].trim()) ? optionNames[opt].trim() : "Option " + (opt + 1);
      sc(ws, r, OC, optName, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });
      catColInfo.forEach(({ critStart, cscoreCol, ci }) => {
        const cat = processed[ci];
        const [s_w, e_w] = catRowRanges[ci];
        const numCrit = cat.hasSub ? cat.criteria.length : 1;
        for (let j = 0; j < numCrit; j++) {
          sc(ws, r, critStart + j, null, { font: { name: "Arial", size: 10, color: { argb: "FF0000FF" } }, fill: "eff6ff" });
        }
        if (cat.hasSub) {
          const parts = cat.criteria.map((_, j) => colLetter(critStart + j) + r + "*Weights!" + critWeightCells[ci + "-" + j]);
          const critWtSum = "SUM(Weights!D" + s_w + ":D" + e_w + ")";
          sc(ws, r, cscoreCol, { formula: "IFERROR((" + parts.join("+") + ")/" + critWtSum + ",0)" }, { font: { name: "Arial", size: 10, bold: true }, fill: "f0fdf4", fmt: "0.00" });
        }
      });
    }

    // Conditional formatting: per-column color scale on data cells (red-yellow-green)
    catColInfo.forEach(({ critStart, ci }) => {
      const cat = processed[ci];
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      for (let j = 0; j < numCrit; j++) {
        const cc = critStart + j;
        const cl = colLetter(cc);
        const ref = cl + R_OPT + ":" + cl + (R_OPT + numOptions - 1);
        ws.addConditionalFormatting({ ref, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 1 }] });
      }
    });

    ws.views = [{ state: "frozen", xSplit: DC - 1, ySplit: R_OPT - 1 }];
  }
  // We need to know the Input sheet column layout for cross-refs
  function buildZScoreSheet() {
    const ws = wb.addWorksheet("Z-Score Normalisation", { properties: { tabColor: { argb: "FF8b5cf6" } } });
    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, name + " - Z-Score Normalisation", { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
    let instrRow = 2;
    if (desc) {
      sc(ws, 2, 1, desc, { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });
      instrRow = 3;
    }
    sc(ws, instrRow, 1, "Scores standardised per criterion (mean=0, std=1), then weighted.", { font: { name: "Arial", size: 9, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });

    const OC = 1, TC = 2, RC = 3, DC = 4;

    // Column map for z-score sheet
    let col = DC;
    const zCatColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = col;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      col += numCrit;
      const cscoreCol = cat.hasSub ? col : null;
      if (cat.hasSub) col++;
      zCatColInfo.push({ critStart, critEnd: critStart + numCrit - 1, cscoreCol, ci });
    });

    // Input sheet column map (now has cat score cols too, starts at col 2)
    let iCol = 2;
    const iCatColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = iCol;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      iCol += numCrit;
      if (cat.hasSub) iCol++; // skip category score column
      iCatColInfo.push({ critStart, ci });
    });

    // Label columns A-C
    for (let rr = R_CAT; rr <= R_EW; rr++) {
      sc(ws, rr, OC, "", { fill: "f8fafc" });
      sc(ws, rr, TC, "", { fill: "f8fafc" });
    }
    sc(ws, R_CAT, RC, "Criteria", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW, RC, "Criteria weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CRIT, RC, "Sub-criterion", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW2, RC, "Sub-criteria weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_EW, RC, "Effective weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });

    sc(ws, R_HDR, OC, "Option", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
    sc(ws, R_HDR, TC, "Total Score", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });
    sc(ws, R_HDR, RC, "Rank", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });

    ws.getColumn(OC).width = 20; ws.getColumn(TC).width = 12; ws.getColumn(RC).width = 8;

    // Category/criteria headers
    zCatColInfo.forEach(({ critStart, critEnd, cscoreCol, ci }) => {
      const cat = processed[ci];
      const color = CAT_COLORS[ci % CAT_COLORS.length];
      const [s_w] = catRowRanges[ci];
      const cwRef = catWeightCells[ci];
      const mergeEnd = cscoreCol || critEnd;

      if (mergeEnd > critStart) ws.mergeCells(R_CAT, critStart, R_CAT, mergeEnd);
      sc(ws, R_CAT, critStart, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: color });
      if (mergeEnd > critStart) ws.mergeCells(R_CW, critStart, R_CW, mergeEnd);
      sc(ws, R_CW, critStart, { formula: "Weights!" + cwRef }, { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });

      if (cat.hasSub) {
        cat.criteria.forEach((crit, j) => {
          const cc = critStart + j;
          sc(ws, R_CRIT, cc, crit.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc", align: { horizontal: "center", vertical: "middle", wrapText: true } });
          sc(ws, R_CW2, cc, { formula: "Weights!" + critWeightCells[ci + "-" + j] }, { font: { name: "Arial", size: 9, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });
          sc(ws, R_EW, cc, { formula: "Weights!E" + (s_w + j) }, { font: { name: "Arial", size: 9 }, fill: "f1f5f9", fmt: "0.0%" });
          sc(ws, R_HDR, cc, "", { fill: "e2e8f0" });
          ws.getColumn(cc).width = 16;
        });
        sc(ws, R_CRIT, cscoreCol, cat.name + " Score", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF" + color } }, fill: "f0fdf4", wrap: true });
        sc(ws, R_CW2, cscoreCol, "", { fill: "f0fdf4" });
        sc(ws, R_EW, cscoreCol, "", { fill: "f0fdf4" });
        sc(ws, R_HDR, cscoreCol, "", { fill: "f0fdf4" });
        ws.getColumn(cscoreCol).width = 14;
      } else {
        const cc = critStart;
        sc(ws, R_CRIT, cc, cat.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc", align: { horizontal: "center", vertical: "middle", wrapText: true } });
        sc(ws, R_CW2, cc, "", { fill: "f1f5f9" });
        sc(ws, R_EW, cc, { formula: "Weights!E" + s_w }, { font: { name: "Arial", size: 9 }, fill: "f1f5f9", fmt: "0.0%" });
        sc(ws, R_HDR, cc, "", { fill: "e2e8f0" });
        ws.getColumn(cc).width = 16;
      }
    });

    // Option rows
    for (let opt = 0; opt < numOptions; opt++) {
      const r = R_OPT + opt;
      const optName = (optionNames[opt] && optionNames[opt].trim()) ? optionNames[opt].trim() : "Option " + (opt + 1);
      sc(ws, r, OC, optName, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });

      const catScoreCells = [];
      zCatColInfo.forEach(({ critStart, cscoreCol, ci }, catIdx) => {
        const cat = processed[ci];
        const [s_w, e_w] = catRowRanges[ci];
        const iCritStart = iCatColInfo[catIdx].critStart;
        const numCrit = cat.hasSub ? cat.criteria.length : 1;

        for (let j = 0; j < numCrit; j++) {
          const cc = critStart + j;
          const icc = iCritStart + j;
          const iCl = colLetter(icc);
          const iRange = "'Input'!" + iCl + R_OPT + ":" + iCl + (R_OPT + numOptions - 1);
          const raw = "'Input'!" + iCl + r;
          sc(ws, r, cc, { formula: "IFERROR((" + raw + "-AVERAGE(" + iRange + "))/STDEV(" + iRange + "),0)" }, { font: { name: "Arial", size: 10 }, fmt: "0.00" });
        }

        if (cat.hasSub) {
          const parts = cat.criteria.map((_, j) => colLetter(critStart + j) + r + "*Weights!" + critWeightCells[ci + "-" + j]);
          const critWtSum = "SUM(Weights!D" + s_w + ":D" + e_w + ")";
          sc(ws, r, cscoreCol, { formula: "IFERROR((" + parts.join("+") + ")/" + critWtSum + ",0)" }, { font: { name: "Arial", size: 10, bold: true }, fill: "f0fdf4", fmt: "0.00" });
          catScoreCells.push({ col: cscoreCol, ci });
        } else {
          catScoreCells.push({ col: critStart, ci });
        }
      });

      // Total score
      const tParts = catScoreCells.map(({ col: csc, ci: cci }) => colLetter(csc) + r + "*Weights!" + catWeightCells[cci]);
      const catWtSum = processed.map((_, cci) => "Weights!" + catWeightCells[cci]).join("+");
      sc(ws, r, TC, { formula: "IFERROR((" + tParts.join("+") + ")/(" + catWtSum + "),0)" }, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "dcfce7", fmt: "0.00" });

      // Rank
      const tl = colLetter(TC);
      const rankRange = tl + "$" + R_OPT + ":" + tl + "$" + (R_OPT + numOptions - 1);
      sc(ws, r, RC, { formula: 'IFERROR(RANK(' + tl + r + ',' + rankRange + ',0),"")' }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0" });
    }

    // Conditional formatting on total and rank
    const totalRef = colLetter(TC) + R_OPT + ":" + colLetter(TC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({ ref: totalRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 1 }] });
    const rankRef = colLetter(RC) + R_OPT + ":" + colLetter(RC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({ ref: rankRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFBBF7D0" }, { argb: "FFFFFBEB" }, { argb: "FFFECACA" }], priority: 2 }] });

    // Conditional formatting: single color scale across ALL z-score data cells
    const allDataCols = [];
    zCatColInfo.forEach(({ critStart, ci }) => {
      const cat = processed[ci];
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      for (let j = 0; j < numCrit; j++) allDataCols.push(critStart + j);
    });
    const allDataRef = allDataCols.map(cc => colLetter(cc) + R_OPT + ":" + colLetter(cc) + (R_OPT + numOptions - 1)).join(" ");
    ws.addConditionalFormatting({ ref: allDataRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 3 }] });

    ws.views = [{ state: "frozen", xSplit: DC - 1, ySplit: R_OPT - 1 }];

    return { TC, R_OPT }; // Return for Results sheet references
  }

  // ============ RESULTS SHEET ============
  function buildResultsSheet(zTC, zOptStart) {
    const ws = wb.addWorksheet("Results", { properties: { tabColor: { argb: "FF059669" } } });

    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, name, { font: { name: "Arial", size: 18, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });

    let nextRow = 2;
    if (desc) {
      ws.mergeCells("A2:D2");
      sc(ws, 2, 1, desc, { font: { name: "Arial", size: 11, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" } });
      nextRow = 3;
    }
    nextRow++;

    const hdrRow = nextRow;
    sc(ws, hdrRow, 1, "Rank", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });
    sc(ws, hdrRow, 2, "Option", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
    sc(ws, hdrRow, 3, "Score", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });

    ws.getColumn(1).width = 8;
    ws.getColumn(2).width = 30;
    ws.getColumn(3).width = 14;

    // Auto-sorted by rank using INDEX/MATCH
    // Z-Score sheet: col A = option name, col B = total score, col C = rank
    const zS = "'Z-Score Normalisation'";
    const zOptEnd = zOptStart + numOptions - 1;
    const rankRange = zS + "!$C$" + zOptStart + ":$C$" + zOptEnd;
    const nameRange = zS + "!$A$" + zOptStart + ":$A$" + zOptEnd;
    const scoreRange = zS + "!$" + colLetter(zTC) + "$" + zOptStart + ":$" + colLetter(zTC) + "$" + zOptEnd;

    const dataStart = hdrRow + 1;
    for (let i = 0; i < numOptions; i++) {
      const r = dataStart + i;
      const rank = i + 1;

      // Rank: just the number 1, 2, 3...
      sc(ws, r, 1, rank, { font: { name: "Arial", size: 12, bold: true, color: { argb: "FF1a1a2e" } }, fmt: "0" });

      // Option name: INDEX(names, MATCH(rank, ranks, 0))
      sc(ws, r, 2, { formula: "IFERROR(INDEX(" + nameRange + ",MATCH(" + rank + "," + rankRange + ",0)),\"\")" }, { font: { name: "Arial", size: 11, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });

      // Score: INDEX(scores, MATCH(rank, ranks, 0))
      sc(ws, r, 3, { formula: "IFERROR(INDEX(" + scoreRange + ",MATCH(" + rank + "," + rankRange + ",0)),\"\")" }, { font: { name: "Arial", size: 11, bold: true, color: { argb: "FF1a1a2e" } }, fmt: "0.00" });
    }

    const scoreRef = "C" + dataStart + ":C" + (dataStart + numOptions - 1);
    ws.addConditionalFormatting({ ref: scoreRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 1 }] });

    const rankRef2 = "A" + dataStart + ":A" + (dataStart + numOptions - 1);
    ws.addConditionalFormatting({ ref: rankRef2, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFBBF7D0" }, { argb: "FFFFFBEB" }, { argb: "FFFECACA" }], priority: 2 }] });

    const noteRow = dataStart + numOptions + 1;
    ws.mergeCells("A" + noteRow + ":C" + noteRow);
    sc(ws, noteRow, 1, "Ranked by z-score normalised weighted total from the Z-Score Normalisation tab.", { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });
  }

  // ============ DEFINITIONS SHEET ============
  function buildDefinitionsSheet() {
    const ws = wb.addWorksheet("Definitions", { properties: { tabColor: { argb: "FF64748b" } } });

    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, name + " - Criteria Definitions", { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
    let instrRow = 2;
    if (desc) {
      sc(ws, 2, 1, desc, { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });
      instrRow = 3;
    }
    sc(ws, instrRow, 1, "Definitions for each criterion used in the model. Edit as needed.", { font: { name: "Arial", size: 9, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });

    const hdrRow = instrRow + 1;
    const headers = ["Criteria", "Sub-criterion", "Definition", "Source"];
    headers.forEach((h, i) => sc(ws, hdrRow, i + 1, h, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "64748b" }));

    ws.getColumn(1).width = 22;
    ws.getColumn(2).width = 26;
    ws.getColumn(3).width = 50;
    ws.getColumn(4).width = 30;

    let dr = hdrRow + 1;
    processed.forEach((cat, ci) => {
      if (cat.hasSub) {
        const catStart = dr;
        cat.criteria.forEach((crit, cri) => {
          if (cri === 0) {
            sc(ws, dr, 1, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "f1f5f9", align: { horizontal: "left", vertical: "top" } });
          } else {
            sc(ws, dr, 1, "", { fill: "f1f5f9" });
          }
          sc(ws, dr, 2, crit.name, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "top" } });
          sc(ws, dr, 3, crit.description || "", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "top", wrapText: true } });
          sc(ws, dr, 4, "", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "top", wrapText: true } });
          dr++;
        });
        if (cat.criteria.length > 1) {
          ws.mergeCells(catStart, 1, dr - 1, 1);
        }
      } else {
        sc(ws, dr, 1, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "f1f5f9", align: { horizontal: "left", vertical: "top" } });
        sc(ws, dr, 2, cat.name, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "top" } });
        sc(ws, dr, 3, (cat.criteria[0] && cat.criteria[0].description) || "", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "top", wrapText: true } });
        sc(ws, dr, 4, "", { font: { name: "Arial", size: 10, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "top", wrapText: true } });
        dr++;
      }
    });
  }

  // Build all sheets
  buildInputSheet();
  const { TC: zTC, R_OPT: zOptStart } = buildZScoreSheet();
  buildResultsSheet(zTC, zOptStart);
  buildDefinitionsSheet();

  // Reorder: move Weights to last tab (Overview, Input, Z-Score, Results, Definitions, Weights)
  const wsIdx = wb.worksheets.indexOf(wsW);
  if (wsIdx >= 0) {
    wb.worksheets.splice(wsIdx, 1);
    wb.worksheets.push(wsW);
    wb.worksheets.forEach((ws, i) => { ws.orderNo = i; });
  }

  // Set Overview as the active tab
  wb.views = [{ activeTab: 0 }];

  // Generate and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const filename = (name).replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "-") + ".xlsx";
  saveAs(blob, filename);
}
