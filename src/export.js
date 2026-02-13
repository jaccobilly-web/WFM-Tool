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

export async function exportToExcel(categories, modelName, numOptions, optionNames) {
  const wb = new ExcelJS.Workbook();

  // Determine which categories have subcriteria vs standalone
  const processed = categories.map(cat => {
    const hasSub = cat.criteria.length > 1 ||
      (cat.criteria.length === 1 && cat.criteria[0].name.trim() !== "" && cat.criteria[0].name.trim() !== cat.name.trim());
    return { ...cat, hasSub };
  });

  // ============ WEIGHTS SHEET ============
  const wsW = wb.addWorksheet("Weights", { properties: { tabColor: { argb: "FF10b981" } } });

  wsW.mergeCells("A1:F1");
  sc(wsW, 1, 1, "Weight Structure", { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
  sc(wsW, 2, 1, "Edit weights here (blue cells). All other sheets reference this tab.", { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });

  const wHeaders = ["Category", "Category Weight", "Criterion", "Criterion Weight", "Effective Weight", "Check"];
  wHeaders.forEach((h, i) => sc(wsW, 4, i + 1, h, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FFFFFFFF" } }, fill: "10b981" }));
  [22, 16, 26, 16, 18, 30].forEach((w, i) => { wsW.getColumn(i + 1).width = w; });

  const catWeightCells = {};
  const critWeightCells = {};
  const catRowRanges = {};
  let wr = 5;

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

  // Effective weight formulas + checks
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
  sc(wsW, wr, 3, "", { fill: "e2e8f0" });
  sc(wsW, wr, 4, "", { fill: "e2e8f0" });
  sc(wsW, wr, 5, { formula: "SUM(E5:E" + (wr - 2) + ")" }, { font: { name: "Arial", size: 10, bold: true }, fill: "e2e8f0", fmt: "0.0%" });
  sc(wsW, wr, 6, { formula: 'IF(ABS(' + catSumParts + '-1)<0.001,"All weights balanced","Category weights sum to "&TEXT(' + catSumParts + ',"0%")&" (need 100%)")' }, { font: { name: "Arial", size: 9, bold: true, color: { argb: "FF16a34a" } }, fill: "fefce8" });

  // ============ DATA SHEETS ============
  function buildDataSheet(sheetName, isZscore) {
    const ws = wb.addWorksheet(sheetName, { properties: { tabColor: { argb: isZscore ? "FF8b5cf6" : "FF3b82f6" } } });

    ws.mergeCells("A1:D1");
    sc(ws, 1, 1, (modelName || "Weighted Factor Model") + " - " + (isZscore ? "Z-Score Normalised" : "Data Input"), { font: { name: "Arial", size: 14, bold: true, color: { argb: "FF1a1a2e" } }, align: { horizontal: "left", vertical: "middle" } });
    sc(ws, 2, 1, isZscore ? "Scores standardised per criterion (mean=0, std=1), then weighted" : "Enter data in the blue cells. Weights are pulled from the Weights tab.", { font: { name: "Arial", size: 10, color: { argb: "FF666666" } }, align: { horizontal: "left", vertical: "middle" } });

    var OC = 1, TC = 2, RC = 3, DC = 4;
    var R_CAT = 4, R_CW = 5, R_CRIT = 6, R_CW2 = 7, R_EW = 8, R_HDR = 9, R_OPT = 10;

    // Column map
    let col = DC;
    const catColInfo = [];
    processed.forEach((cat, ci) => {
      const critStart = col;
      const numCrit = cat.hasSub ? cat.criteria.length : 1;
      col += numCrit;
      const cscoreCol = cat.hasSub ? col : null;
      if (cat.hasSub) col++;
      catColInfo.push({ critStart, critEnd: critStart + numCrit - 1, cscoreCol, ci, numCrit });
    });

    // Style rows 4-8 cols A-C
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

    ws.getColumn(OC).width = 20;
    ws.getColumn(TC).width = 12;
    ws.getColumn(RC).width = 8;

    // Category/criteria headers
    catColInfo.forEach(({ critStart, critEnd, cscoreCol, ci }) => {
      const cat = processed[ci];
      const color = CAT_COLORS[ci % CAT_COLORS.length];
      const [s_w, e_w] = catRowRanges[ci];
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
      const name = (optionNames[opt] && optionNames[opt].trim()) ? optionNames[opt].trim() : "Option " + (opt + 1);
      sc(ws, r, OC, name, { font: { name: "Arial", size: 10, color: { argb: "FF334155" } }, align: { horizontal: "left", vertical: "middle" } });

      const catScoreCells = [];
      catColInfo.forEach(({ critStart, cscoreCol, ci }) => {
        const cat = processed[ci];
        const [s_w, e_w] = catRowRanges[ci];

        if (cat.hasSub) {
          cat.criteria.forEach((_, j) => {
            const cc = critStart + j;
            const cl = colLetter(cc);
            if (isZscore) {
              const iRange = "'Input'!" + cl + R_OPT + ":" + cl + (R_OPT + numOptions - 1);
              const raw = "'Input'!" + cl + r;
              sc(ws, r, cc, { formula: "IFERROR((" + raw + "-AVERAGE(" + iRange + "))/STDEV(" + iRange + "),0)" }, { font: { name: "Arial", size: 10 }, fmt: "0.00" });
            } else {
              sc(ws, r, cc, null, { font: { name: "Arial", size: 10, color: { argb: "FF0000FF" } }, fill: "eff6ff" });
            }
          });
          const parts = cat.criteria.map((_, j) => colLetter(critStart + j) + r + "*Weights!" + critWeightCells[ci + "-" + j]);
          const critWtSum = "SUM(Weights!D" + s_w + ":D" + e_w + ")";
          sc(ws, r, cscoreCol, { formula: "IFERROR((" + parts.join("+") + ")/" + critWtSum + ",0)" }, { font: { name: "Arial", size: 10, bold: true }, fill: "f0fdf4", fmt: "0.00" });
          catScoreCells.push({ col: cscoreCol, ci: ci });
        } else {
          const cc = critStart;
          const cl = colLetter(cc);
          if (isZscore) {
            const iRange = "'Input'!" + cl + R_OPT + ":" + cl + (R_OPT + numOptions - 1);
            const raw = "'Input'!" + cl + r;
            sc(ws, r, cc, { formula: "IFERROR((" + raw + "-AVERAGE(" + iRange + "))/STDEV(" + iRange + "),0)" }, { font: { name: "Arial", size: 10 }, fmt: "0.00" });
          } else {
            sc(ws, r, cc, null, { font: { name: "Arial", size: 10, color: { argb: "FF0000FF" } }, fill: "eff6ff" });
          }
          catScoreCells.push({ col: cc, ci: ci });
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

    // Conditional formatting: Total Score (red-yellow-green)
    const totalRef = colLetter(TC) + R_OPT + ":" + colLetter(TC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({
      ref: totalRef,
      rules: [{
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
        color: [{ argb: "FFFECACA" }, { argb: "FFFFFBEB" }, { argb: "FFBBF7D0" }],
        priority: 1,
      }],
    });

    // Conditional formatting: Rank (green=1, red=last, reversed)
    const rankRef = colLetter(RC) + R_OPT + ":" + colLetter(RC) + (R_OPT + numOptions - 1);
    ws.addConditionalFormatting({
      ref: rankRef,
      rules: [{
        type: "colorScale",
        cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
        color: [{ argb: "FFBBF7D0" }, { argb: "FFFFFBEB" }, { argb: "FFFECACA" }],
        priority: 2,
      }],
    });

    ws.views = [{ state: "frozen", xSplit: DC - 1, ySplit: R_OPT - 1 }];
  }

  buildDataSheet("Input", false);
  buildDataSheet("Z-Score Normalised", true);

  // Generate and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const filename = (modelName || "weighted-factor-model").replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "-") + ".xlsx";
  saveAs(blob, filename);
}
