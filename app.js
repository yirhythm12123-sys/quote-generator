/* global Vue, ExcelJS, saveAs */

const PROFIT_MODE_OPTIONS = {
  "部": { initial: ["sales", "none"], monthly: ["sales", "none"], spot: ["sales", "none"] },
  "SE": { initial: ["both"],          monthly: ["both", "dept", "none"], spot: ["both"] },
};

const PROFIT_LABELS = {
  both:  "部+営業 (/0.8/0.9)",
  sales: "営業のみ (/0.9)",
  dept:  "部のみ (/0.8)",
  none:  "そのまま",
};

const FEE_LABELS = { initial: "初期費用", monthly: "月額費用", spot: "都度費用" };
const FEE_ORDER = ["initial", "monthly", "spot"];

function todayISO() {
  const d = new Date();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}

function newLineItem() {
  return {
    name: "",
    source: "部",
    feeType: "initial",
    profitMode: "sales",
    quantity: 1,
    unit: "式",
    unitOfferPrice: 0,
  };
}

function newCase(name) {
  return { name, lineItems: [newLineItem()] };
}

const { createApp, reactive, computed, ref } = Vue;

createApp({
  setup() {
    const quote = reactive({
      quoteNumber: "",
      issueDate: todayISO(),
      validUntil: "",
      customer: { company: "", department: "", contact: "", honorific: "御中" },
      ourCompany: { company: "", department: "", address: "", staff: "", manager: "" },
      notes: "",
      cases: [newCase("案A")],
    });

    const activeCaseIndex = ref(0);
    const activeCase = computed(() => quote.cases[activeCaseIndex.value]);

    function profitOptions(item) {
      return PROFIT_MODE_OPTIONS[item.source][item.feeType];
    }
    function profitLabel(mode) { return PROFIT_LABELS[mode] || mode; }

    function onSourceFeeChange(item) {
      const allowed = profitOptions(item);
      if (!allowed.includes(item.profitMode)) {
        item.profitMode = allowed[0];
      }
    }

    function calcUnitPrice(item) {
      const p = Number(item.unitOfferPrice) || 0;
      switch (item.profitMode) {
        case "both":  return Math.round(p / 0.8 / 0.9);
        case "sales": return Math.round(p / 0.9);
        case "dept":  return Math.round(p / 0.8);
        case "none":  return p;
        default:      return p;
      }
    }
    function calcLineTotal(item) {
      return calcUnitPrice(item) * (Number(item.quantity) || 0);
    }
    function calcCustomerPrice(item) { return calcLineTotal(item); }
    function calcPlanRequest(item) {
      const cp = calcCustomerPrice(item);
      return (item.profitMode === "sales" || item.profitMode === "both")
        ? Math.round(cp * 0.9)
        : cp;
    }
    function calcSalesProfit(item) { return calcCustomerPrice(item) - calcPlanRequest(item); }
    function calcSEAmount(item) {
      if (item.source !== "SE") return 0;
      return (Number(item.unitOfferPrice) || 0) * (Number(item.quantity) || 0);
    }
    function calcGrossProfit(item) { return calcPlanRequest(item) - calcSEAmount(item); }

    function caseTotalsByFee(feeType) {
      return activeCase.value.lineItems
        .filter(i => i.feeType === feeType)
        .reduce((s, i) => s + calcLineTotal(i), 0);
    }
    function caseGrandTotal() {
      return activeCase.value.lineItems.reduce((s, i) => s + calcLineTotal(i), 0);
    }
    function summaryTotals(c = activeCase.value) {
      return c.lineItems.reduce((acc, i) => {
        acc.customer    += calcCustomerPrice(i);
        acc.planRequest += calcPlanRequest(i);
        acc.salesProfit += calcSalesProfit(i);
        acc.seAmount    += calcSEAmount(i);
        acc.grossProfit += calcGrossProfit(i);
        return acc;
      }, { customer: 0, planRequest: 0, salesProfit: 0, seAmount: 0, grossProfit: 0 });
    }
    function grossMarginPct(c = activeCase.value) {
      const t = summaryTotals(c);
      if (t.customer === 0) return "0.0";
      return (t.grossProfit / t.customer * 100).toFixed(1);
    }

    function fmt(n) {
      const v = Number(n) || 0;
      return v.toLocaleString("ja-JP");
    }

    function addLine() { activeCase.value.lineItems.push(newLineItem()); }
    function removeLine(i) {
      if (activeCase.value.lineItems.length <= 1) {
        activeCase.value.lineItems.splice(i, 1, newLineItem());
      } else {
        activeCase.value.lineItems.splice(i, 1);
      }
    }
    function addCase() {
      const nextLetter = String.fromCharCode(65 + quote.cases.length);
      quote.cases.push(newCase(`案${nextLetter}`));
      activeCaseIndex.value = quote.cases.length - 1;
    }
    function removeCase() {
      if (quote.cases.length <= 1) return;
      if (!confirm(`「${activeCase.value.name}」を削除しますか？`)) return;
      quote.cases.splice(activeCaseIndex.value, 1);
      activeCaseIndex.value = Math.max(0, activeCaseIndex.value - 1);
    }
    function renameCase() {
      const next = prompt("案名を入力", activeCase.value.name);
      if (next && next.trim()) activeCase.value.name = next.trim();
    }

    async function exportExcel() {
      const wb = new ExcelJS.Workbook();
      wb.creator = quote.ourCompany.staff || "見積作成ツール";
      wb.created = new Date();

      const usedNames = new Set();
      for (const c of quote.cases) {
        const sheetName = sanitizeSheetName(c.name, usedNames);
        const ws = wb.addWorksheet(sheetName, {
          properties: { defaultRowHeight: 18 },
          pageSetup: { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
        });
        renderCaseSheet(ws, c);
      }

      const buf = await wb.xlsx.writeBuffer();
      const filename = `${quote.quoteNumber || "見積書"}_${todayISO()}.xlsx`;
      saveAs(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
    }

    function sanitizeSheetName(name, used) {
      let n = (name || "案").replace(/[\\/?*\[\]:]/g, "_").slice(0, 28);
      let candidate = n, i = 2;
      while (used.has(candidate)) candidate = `${n}_${i++}`;
      used.add(candidate);
      return candidate;
    }

    function renderCaseSheet(ws, c) {
      ws.columns = [
        { width: 6 },  { width: 28 }, { width: 10 }, { width: 8 },
        { width: 14 }, { width: 14 }, { width: 14 }, { width: 8 },
      ];
      const COLS = 8;
      let row = 1;

      // タイトル
      ws.mergeCells(row, 1, row, COLS);
      const titleCell = ws.getCell(row, 1);
      titleCell.value = "見 積 書";
      titleCell.font = { name: "游ゴシック", size: 20, bold: true };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      ws.getRow(row).height = 32;
      row += 2;

      // 顧客（左）と見積情報・自社（右）と印枠
      const headerStart = row;
      // 左: 顧客
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = `${quote.customer.company || ""}　${quote.customer.honorific || ""}`;
      ws.getCell(row, 1).font = { name: "游ゴシック", size: 14, bold: true };
      ws.getCell(row, 1).border = { bottom: { style: "thin" } };
      // 右: 見積番号
      ws.getCell(row, 6).value = "見積番号";
      ws.getCell(row, 7).value = quote.quoteNumber;
      // 印枠タイトル
      ws.getCell(row, 8).value = "担当";
      ws.getCell(row, 8).alignment = { horizontal: "center" };
      ws.getCell(row, 8).font = { size: 9 };
      row++;
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = quote.customer.department || "";
      ws.getCell(row, 6).value = "発行日";
      ws.getCell(row, 7).value = quote.issueDate;
      // 印枠（担当）正方形
      ws.getCell(row, 8).value = "";
      ws.getCell(row, 8).border = boxBorder();
      ws.getRow(row).height = 36;
      row++;
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = quote.customer.contact ? `ご担当 ${quote.customer.contact} 様` : "";
      ws.getCell(row, 6).value = "有効期限";
      ws.getCell(row, 7).value = quote.validUntil;
      ws.getCell(row, 8).value = "上長";
      ws.getCell(row, 8).alignment = { horizontal: "center" };
      ws.getCell(row, 8).font = { size: 9 };
      row++;
      // 自社情報行 + 印枠（上長）
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = quote.ourCompany.company || "";
      ws.getCell(row, 1).font = { bold: true };
      ws.getCell(row, 8).value = "";
      ws.getCell(row, 8).border = boxBorder();
      ws.getRow(row).height = 36;
      row++;
      ws.mergeCells(row, 1, row, 7);
      ws.getCell(row, 1).value = [quote.ourCompany.department, quote.ourCompany.address].filter(Boolean).join("　");
      row++;
      ws.mergeCells(row, 1, row, 7);
      ws.getCell(row, 1).value = [
        quote.ourCompany.staff ? `担当 ${quote.ourCompany.staff}` : "",
        quote.ourCompany.manager ? `上長 ${quote.ourCompany.manager}` : "",
      ].filter(Boolean).join("　");
      row += 2;

      // リード文
      ws.mergeCells(row, 1, row, COLS);
      ws.getCell(row, 1).value = "下記の通りお見積申し上げます。";
      ws.getCell(row, 1).font = { size: 11 };
      row += 2;

      // 費用区分セクション（存在するもののみ）
      const grandTotal = c.lineItems.reduce((s, i) => s + calcLineTotal(i), 0);
      for (const ft of FEE_ORDER) {
        const items = c.lineItems.filter(i => i.feeType === ft);
        if (items.length === 0) continue;
        // 見出し
        ws.mergeCells(row, 1, row, COLS);
        const h = ws.getCell(row, 1);
        h.value = `【${FEE_LABELS[ft]}】`;
        h.font = { bold: true, size: 12 };
        h.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8EDF2" } };
        row++;
        // テーブルヘッダ
        const hdrs = ["No.", "品目", "数量", "単位", "単価", "金額", "備考", ""];
        hdrs.forEach((v, idx) => {
          const cell = ws.getCell(row, idx + 1);
          cell.value = v;
          cell.font = { bold: true };
          cell.alignment = { horizontal: "center" };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF2F6" } };
          cell.border = thinAll();
        });
        row++;
        // 明細
        items.forEach((item, idx) => {
          const vals = [
            idx + 1,
            item.name || "",
            Number(item.quantity) || 0,
            item.unit || "",
            calcUnitPrice(item),
            calcLineTotal(item),
            "",
            "",
          ];
          vals.forEach((v, i) => {
            const cell = ws.getCell(row, i + 1);
            cell.value = v;
            cell.border = thinAll();
            if (i === 0 || i === 2 || i === 3) cell.alignment = { horizontal: "center" };
            if (i === 4 || i === 5) {
              cell.alignment = { horizontal: "right" };
              cell.numFmt = "#,##0";
            }
          });
          row++;
        });
        // 小計
        const subtotal = items.reduce((s, i) => s + calcLineTotal(i), 0);
        ws.mergeCells(row, 1, row, 4);
        const stLabel = ws.getCell(row, 1);
        stLabel.value = `${FEE_LABELS[ft]} 小計`;
        stLabel.alignment = { horizontal: "right" };
        stLabel.font = { bold: true };
        stLabel.border = thinAll();
        ws.getCell(row, 5).border = thinAll();
        const stVal = ws.getCell(row, 6);
        stVal.value = subtotal;
        stVal.numFmt = "#,##0";
        stVal.alignment = { horizontal: "right" };
        stVal.font = { bold: true };
        stVal.border = thinAll();
        ws.getCell(row, 7).border = thinAll();
        row += 2;
      }

      // 総合計
      ws.mergeCells(row, 1, row, 4);
      const gtLabel = ws.getCell(row, 1);
      gtLabel.value = "総合計（税抜）";
      gtLabel.alignment = { horizontal: "right" };
      gtLabel.font = { bold: true, size: 12 };
      gtLabel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
      gtLabel.border = thinAll();
      ws.getCell(row, 5).fill = gtLabel.fill;
      ws.getCell(row, 5).border = thinAll();
      const gtVal = ws.getCell(row, 6);
      gtVal.value = grandTotal;
      gtVal.numFmt = "#,##0";
      gtVal.alignment = { horizontal: "right" };
      gtVal.font = { bold: true, size: 12 };
      gtVal.fill = gtLabel.fill;
      gtVal.border = thinAll();
      ws.getCell(row, 7).fill = gtLabel.fill;
      ws.getCell(row, 7).border = thinAll();
      row += 2;

      // 備考
      if (quote.notes) {
        ws.mergeCells(row, 1, row, COLS);
        ws.getCell(row, 1).value = "【備考】";
        ws.getCell(row, 1).font = { bold: true };
        row++;
        const noteRows = quote.notes.split(/\r?\n/);
        for (const line of noteRows) {
          ws.mergeCells(row, 1, row, COLS);
          ws.getCell(row, 1).value = line;
          row++;
        }
        row++;
      }

      // 内部用サマリ
      row += 1;
      ws.mergeCells(row, 1, row, COLS);
      const sumTitle = ws.getCell(row, 1);
      sumTitle.value = "【内部用サマリ】 ※社外秘";
      sumTitle.font = { bold: true, size: 12, color: { argb: "FFCC4444" } };
      sumTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF6E8" } };
      row++;
      const sumHeaders = ["品目", "顧客出値", "企画依頼金額", "営業利益", "SE出値", "粗利", "", ""];
      sumHeaders.forEach((v, idx) => {
        const cell = ws.getCell(row, idx + 1);
        cell.value = v;
        if (v) {
          cell.font = { bold: true };
          cell.alignment = { horizontal: "center" };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7E6D4" } };
          cell.border = thinAll();
        }
      });
      row++;
      const sumFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFAF2" } };
      for (const item of c.lineItems) {
        const vals = [
          item.name || "(無題)",
          calcCustomerPrice(item),
          calcPlanRequest(item),
          calcSalesProfit(item),
          calcSEAmount(item),
          calcGrossProfit(item),
        ];
        vals.forEach((v, i) => {
          const cell = ws.getCell(row, i + 1);
          cell.value = v;
          cell.border = thinAll();
          cell.fill = sumFill;
          if (i >= 1) {
            cell.alignment = { horizontal: "right" };
            cell.numFmt = "#,##0";
          }
        });
        row++;
      }
      // 案合計
      const tot = c.lineItems.reduce((acc, i) => {
        acc.customer    += calcCustomerPrice(i);
        acc.planRequest += calcPlanRequest(i);
        acc.salesProfit += calcSalesProfit(i);
        acc.seAmount    += calcSEAmount(i);
        acc.grossProfit += calcGrossProfit(i);
        return acc;
      }, { customer: 0, planRequest: 0, salesProfit: 0, seAmount: 0, grossProfit: 0 });
      const totVals = ["案合計", tot.customer, tot.planRequest, tot.salesProfit, tot.seAmount, tot.grossProfit];
      totVals.forEach((v, i) => {
        const cell = ws.getCell(row, i + 1);
        cell.value = v;
        cell.border = thinAll();
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7E6D4" } };
        if (i >= 1) {
          cell.alignment = { horizontal: "right" };
          cell.numFmt = "#,##0";
        }
      });
      row++;
      const margin = tot.customer === 0 ? "0.0" : (tot.grossProfit / tot.customer * 100).toFixed(1);
      ws.mergeCells(row, 1, row, 5);
      const marginLabel = ws.getCell(row, 1);
      marginLabel.value = "粗利率（粗利 ÷ 顧客出値）";
      marginLabel.alignment = { horizontal: "right" };
      marginLabel.font = { bold: true };
      marginLabel.border = thinAll();
      const marginVal = ws.getCell(row, 6);
      marginVal.value = `${margin} %`;
      marginVal.alignment = { horizontal: "right" };
      marginVal.font = { bold: true };
      marginVal.border = thinAll();
    }

    function thinAll() {
      const s = { style: "thin", color: { argb: "FF888888" } };
      return { top: s, bottom: s, left: s, right: s };
    }
    function boxBorder() {
      const s = { style: "medium", color: { argb: "FF333333" } };
      return { top: s, bottom: s, left: s, right: s };
    }

    return {
      quote, activeCaseIndex, activeCase,
      profitOptions, profitLabel, onSourceFeeChange,
      calcUnitPrice, calcLineTotal,
      calcCustomerPrice, calcPlanRequest, calcSalesProfit, calcSEAmount, calcGrossProfit,
      caseTotalsByFee, caseGrandTotal, summaryTotals, grossMarginPct,
      fmt,
      addLine, removeLine, addCase, removeCase, renameCase,
      exportExcel,
    };
  },
}).mount("#app");
