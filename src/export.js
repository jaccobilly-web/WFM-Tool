import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const CAT_COLORS = ["3b82f6", "8b5cf6", "10b981", "f59e0b", "ec4899", "6366f1"];
const thin = { style: "thin", color: { argb: "FFcbd5e1" } };
const bdr = { top: thin, bottom: thin, left: thin, right: thin };

function sc(ws, r, c, value, { font, fill, align, fmt, wrap } = {}) {
  const cell = ws.getCell(r, c);
  if (value !== undefined && value !== null) cell.value = value;
  if (font) cell.font = font;
  if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + fill } };
  if (fmt) cell.numFmt = fmt;
  cell.alignment = align || { horizontal: "center", vertical: "middle", wrapText: !!wrap };
  cell.border = bdr;
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
  const wHeaders = ["Category", "Category Weight", "Criterion", "Criterion Weight", "Effective Weight", "Check"];
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
    sc(wsW, s, 6, { formula: 'IF(SUM(' + critRange + ')=0,"Empty",IF(ABS(SUM(' + critRange + ')-1)<0.001,"OK","Sum: "&TEXT(SUM(' + critRange + '),"0%")))' }, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF16a34a" } }, fill: "fefce8" });
    for (let row = s + 1; row <= e; row++) sc(wsW, row, 6, "", { fill: "fefce8" });
  });

  wr += 1;
  const catSumParts = processed.map((_, ci) => catWeightCells[ci]).join("+");
  sc(wsW, wr, 1, "TOTAL CHECK", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FF1a1a2e" } }, fill: "e2e8f0", align: { horizontal: "left", vertical: "middle" } });
  sc(wsW, wr, 2, { formula: catSumParts }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0%" });
  sc(wsW, wr, 3, "", { fill: "e2e8f0" }); sc(wsW, wr, 4, "", { fill: "e2e8f0" });
  sc(wsW, wr, 5, { formula: "SUM(E" + firstDataRow + ":E" + (wr - 2) + ")" }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0.0%" });
  sc(wsW, wr, 6, { formula: 'IF(ABS(' + catSumParts + '-1)<0.001,"All weights balanced","Category weights sum to "&TEXT(' + catSumParts + ',"0%")&" (need 100%)")' }, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF16a34a" } }, fill: "fefce8" });

  // ============ SHARED CONSTANTS ============
  // Input tab has no total/rank; Z-Score tab has them
  const R_CAT = 4, R_CW = 5, R_CRIT = 6, R_CW2 = 7, R_EW = 8, R_HDR = 9, R_OPT = 10;

  // ============ INPUT SHEET (data entry only, no scores) ============
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

    // No Total/Rank columns. Col A = Option, Col B onwards = data
    const OC = 1;
    const DC = 2;

    // Column map
    let col = DC;
    const catColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = col;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      col += numCrit;
      catColInfo.push({ critStart, critEnd: critStart + numCrit - 1, ci });
    });

    // Label col
    for (let rr = R_CAT; rr <= R_EW; rr++) sc(ws, rr, OC, "", { fill: "f8fafc" });
    sc(ws, R_CAT, OC, "Category", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW, OC, "Category weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CRIT, OC, "Criterion", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW2, OC, "Criterion weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_EW, OC, "Effective weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_HDR, OC, "Option", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
    ws.getColumn(OC).width = 20;

    catColInfo.forEach(({ critStart, critEnd, ci }) => {
      const cat = processed[ci];
      const color = CAT_COLORS[ci % CAT_COLORS.length];
      const [s_w] = catRowRanges[ci];
      const cwRef = catWeightCells[ci];

      if (critEnd > critStart) ws.mergeCells(R_CAT, critStart, R_CAT, critEnd);
      sc(ws, R_CAT, critStart, cat.name, { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: color });
      if (critEnd > critStart) ws.mergeCells(R_CW, critStart, R_CW, critEnd);
      sc(ws, R_CW, critStart, { formula: "Weights!" + cwRef }, { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });

      if (cat.hasSub) {
        cat.criteria.forEach((crit, j) => {
          const cc = critStart + j;
          sc(ws, R_CRIT, cc, crit.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc" });
          sc(ws, R_CW2, cc, { formula: "Weights!" + critWeightCells[ci + "-" + j] }, { font: { name: "Arial", size: 9, color: { argb: "FF0000FF" } }, fill: "f1f5f9", fmt: "0%" });
          sc(ws, R_EW, cc, { formula: "Weights!E" + (s_w + j) }, { font: { name: "Arial", size: 9 }, fill: "f1f5f9", fmt: "0.0%" });
          sc(ws, R_HDR, cc, "", { fill: "e2e8f0" });
          ws.getColumn(cc).width = 16;
        });
      } else {
        const cc = critStart;
        sc(ws, R_CRIT, cc, cat.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc" });
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
      catColInfo.forEach(({ critStart, ci }) => {
        const cat = processed[ci];
        const numCrit = cat.hasSub ? cat.criteria.length : 1;
        for (let j = 0; j < numCrit; j++) {
          sc(ws, r, critStart + j, null, { font: { name: "Arial", size: 10, color: { argb: "FF0000FF" } }, fill: "eff6ff" });
        }
      });
    }

    ws.views = [{ state: "frozen", xSplit: DC - 1, ySplit: R_OPT - 1 }];
  }

  // ============ Z-SCORE NORMALISATION SHEET (with total + rank) ============
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

    // Input sheet column map (no total/rank, starts at col 2)
    let iCol = 2;
    const iCatColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = iCol;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      iCol += numCrit;
      iCatColInfo.push({ critStart, ci });
    });

    // Label columns A-C
    for (let rr = R_CAT; rr <= R_EW; rr++) {
      sc(ws, rr, OC, "", { fill: "f8fafc" });
      sc(ws, rr, TC, "", { fill: "f8fafc" });
    }
    sc(ws, R_CAT, RC, "Category", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW, RC, "Category weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CRIT, RC, "Criterion", { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
    sc(ws, R_CW2, RC, "Criterion weight", { font: { name: "Arial", size: 9, color: { argb: "FF64748b" } }, fill: "f8fafc", align: { horizontal: "right", vertical: "middle" } });
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
          sc(ws, R_CRIT, cc, crit.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc" });
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
        sc(ws, R_CRIT, cc, cat.name, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF334155" } }, fill: "f8fafc" });
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

    // Conditional formatting
    const totalRef = colLetter(TC) + R_OPT + ":" + colLetter(TC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({ ref: totalRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 1 }] });
    const rankRef = colLetter(RC) + R_OPT + ":" + colLetter(RC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({ ref: rankRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFBBF7D0" }, { argb: "FFFFFBEB" }, { argb: "FFFECACA" }], priority: 2 }] });

    ws.views = [{ state: "frozen", xSplit: DC - 1, ySplit: R_OPT - 1 }];

    return { TC, R_OPT }; // Return for Results sheet references
  }

  // ============ RESULTS SHEET ============
  function buildResultsSheet(zTC, zOptStart) {
    const ws = wb.addWorksheet("Results", { properties: { tabColor: { argb: "FF059669" } } });

    // Title
    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, name, { font: { name: "Arial", size: 18, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });

    // Description
    let nextRow = 2;
    if (desc) {
      ws.mergeCells("A2:D2");
      sc(ws, 2, 1, desc, { font: { name: "Arial", size: 11, color: { argb: "FF475569" } }, align: { horizontal: "left", vertical: "middle" } });
      nextRow = 3;
    }

    // Spacer
    nextRow++;

    // Table header
    const hdrRow = nextRow;
    sc(ws, hdrRow, 1, "Rank", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });
    sc(ws, hdrRow, 2, "Option", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e", align: { horizontal: "left", vertical: "middle" } });
    sc(ws, hdrRow, 3, "Score", { font: { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: "1a1a2e" });

    ws.getColumn(1).width = 8;
    ws.getColumn(2).width = 30;
    ws.getColumn(3).width = 14;

    // One row per option, referencing z-score sheet
    const zSheetName = "'Z-Score Normalisation'";
    const zTL = colLetter(zTC);
    const zRankRange = zSheetName + "!" + colLetter(zTC + 1) + "$" + zOptStart + ":" + colLetter(zTC + 1) + "$" + (zOptStart + numOptions - 1);
    const zTotalRange = zSheetName + "!" + zTL + "$" + zOptStart + ":" + zTL + "$" + (zOptStart + numOptions - 1);

    const dataStart = hdrRow + 1;
    for (let opt = 0; opt < numOptions; opt++) {
      const r = dataStart + opt;
      const zRow = zOptStart + opt;

      // Rank (from z-score sheet col C = rank)
      sc(ws, r, 1, { formula: zSheetName + "!C" + zRow }, { font: { name: "Arial", size: 12, bold: true, color: { argb: "FF1a1a2e" } }, fmt: "0" });

      // Option name (from z-score sheet col A)
      sc(ws, r, 2, { formula: zSheetName + "!A" + zRow }, { font: { name: "Arial", size: 11, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });

      // Total score (from z-score sheet col B = total)
      sc(ws, r, 3, { formula: zSheetName + "!" + zTL + zRow }, { font: { name: "Arial", size: 11, bold: true, color: { argb: "FF1a1a2e" } }, fmt: "0.00" });
    }

    // Conditional formatting on score column
    const scoreRef = "C" + dataStart + ":C" + (dataStart + numOptions - 1);
    ws.addConditionalFormatting({ ref: scoreRef, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }], priority: 1 }] });

    // Conditional formatting on rank column
    const rankRef2 = "A" + dataStart + ":A" + (dataStart + numOptions - 1);
    ws.addConditionalFormatting({ ref: rankRef2, rules: [{ type: "colorScale", cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }], color: [{ argb: "FFBBF7D0" }, { argb: "FFFFFBEB" }, { argb: "FFFECACA" }], priority: 2 }] });

    // Note at bottom
    const noteRow = dataStart + numOptions + 1;
    ws.mergeCells("A" + noteRow + ":C" + noteRow);
    sc(ws, noteRow, 1, "Scores are based on z-score normalised data from the Z-Score Normalisation tab. Options are not auto-sorted; rank column indicates position.", { font: { name: "Arial", size: 8, italic: true, color: { argb: "FF94a3b8" } }, align: { horizontal: "left", vertical: "middle" } });
  }

  // Build all sheets
  buildInputSheet();
  const { TC: zTC, R_OPT: zOptStart } = buildZScoreSheet();
  buildResultsSheet(zTC, zOptStart);

  // Generate and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const filename = (name).replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "-") + ".xlsx";
  saveAs(blob, filename);
}
