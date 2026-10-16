// =====================================================================
// GBS Billings — branded exports (Excel workbooks and client PDF statements)
// Needs ExcelJS (for Excel) and jsPDF + jsPDF-AutoTable (for PDFs) loaded on the page.
// =====================================================================
(function () {
  const BLUE = "FF2050A0", BLUE_SOFT = "FFEEF3FA", INK = "FF111111", MUTED = "FF6E6E73", LINE = "FFE5E5EA", GREY = "FFF5F5F7";
  const RGB_BLUE = [32, 80, 160], RGB_SOFT = [238, 243, 250], RGB_INK = [17, 17, 17], RGB_MUTED = [110, 110, 115], RGB_LINE = [229, 229, 234];

  let logoCache;
  async function logoDataUrl() {
    if (logoCache !== undefined) return logoCache;
    try {
      const res = await fetch("assets/gbs-logo.png", { cache: "force-cache" });
      if (!res.ok) throw new Error("no logo");
      const blob = await res.blob();
      logoCache = await new Promise((resolve) => {
        const fr = new FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = () => resolve(null);
        fr.readAsDataURL(blob);
      });
    } catch (e) { logoCache = null; }      // e.g. opened from a local file — exports still work without the logo
    return logoCache;
  }

  function download(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const nowSA = () => new Date().toLocaleString("en-GB", { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Africa/Johannesburg" });
  const money = (n) => "R " + Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // -------------------------------------------------------------------
  // Excel: branded table
  //   columns: [{ header, key, width, type: 'text'|'money'|'number'|'date'|'pct', wrap }]
  //   totalKeys: keys to SUM in the totals row
  // -------------------------------------------------------------------
  async function xlsx({ filename, sheetName = "Billings", title, subtitle = "", columns, rows, totalKeys = [] }) {
    if (!window.ExcelJS) throw new Error("Excel library didn't load — check your internet connection and try again.");
    const wb = new ExcelJS.Workbook();
    wb.creator = "GBS Billings"; wb.created = new Date();
    const ws = wb.addWorksheet(sheetName.replace(/[\\/?*[\]:]/g, " ").slice(0, 31), {
      views: [{ state: "frozen", ySplit: 5, showGridLines: false }],
      pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
                   margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
      headerFooter: { oddFooter: "&L&8Global Business Solutions · GBS Billings&R&8Page &P of &N" },
    });
    ws.columns = columns.map((c) => ({ key: c.key, width: c.width || 14 }));
    const lastCol = columns.length, span = Math.max(2, Math.min(lastCol, 8));

    // Header block
    ws.getRow(1).height = 44;
    const logo = await logoDataUrl();
    if (logo) {
      const id = wb.addImage({ base64: logo, extension: "png" });
      ws.addImage(id, { tl: { col: 0.15, row: 0.15 }, ext: { width: 154, height: 42 } });
    }
    ws.mergeCells(2, 1, 2, span);
    ws.mergeCells(3, 1, 3, span);
    Object.assign(ws.getCell("A2"), { value: title });
    ws.getCell("A2").font = { name: "Calibri", size: 16, bold: true, color: { argb: BLUE } };
    ws.getRow(2).height = 24;
    ws.getCell("A3").value = subtitle;
    ws.getCell("A3").font = { name: "Calibri", size: 10, color: { argb: MUTED } };
    for (let i = 1; i <= lastCol; i++) ws.getRow(4).getCell(i).border = { top: { style: "medium", color: { argb: BLUE } } };
    ws.getRow(4).height = 6;

    // Column headers
    const hdr = ws.getRow(5);
    hdr.height = 26;
    columns.forEach((c, i) => {
      const cell = hdr.getCell(i + 1);
      cell.value = c.header;
      cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE } };
      cell.alignment = { vertical: "middle", horizontal: ["money", "number", "pct"].includes(c.type) ? "right" : "left", wrapText: true, indent: 0 };
    });

    // Data rows
    rows.forEach((r, idx) => {
      const row = ws.getRow(6 + idx);
      columns.forEach((c, i) => {
        const cell = row.getCell(i + 1);
        let v = r[c.key];
        if (v === undefined || v === "") v = null;
        if (c.type === "date" && v) v = new Date(String(v).slice(0, 10) + "T00:00:00Z");
        if (["money", "number", "pct"].includes(c.type) && v != null) v = Number(v);
        cell.value = v;
        cell.font = { name: "Calibri", size: 10, color: { argb: INK } };
        cell.alignment = { vertical: "top", horizontal: ["money", "number", "pct"].includes(c.type) ? "right" : "left", wrapText: !!c.wrap };
        if (c.type === "money") cell.numFmt = '"R" #,##0.00;[Red]-"R" #,##0.00';
        if (c.type === "number") cell.numFmt = "#,##0.##";
        if (c.type === "pct") cell.numFmt = "0%";
        if (c.type === "date") cell.numFmt = "dd mmm yyyy";
        if (idx % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_SOFT } };
        cell.border = { bottom: { style: "hair", color: { argb: LINE } } };
      });
    });

    // Totals row
    const first = 6, last = 5 + rows.length, totalRow = ws.getRow(6 + rows.length);
    totalRow.height = 22;
    columns.forEach((c, i) => {
      const cell = totalRow.getCell(i + 1);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREY } };
      cell.border = { top: { style: "medium", color: { argb: BLUE } }, bottom: { style: "medium", color: { argb: BLUE } } };
      cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: INK } };
      cell.alignment = { vertical: "middle", horizontal: ["money", "number", "pct"].includes(c.type) ? "right" : "left" };
      if (i === 0) cell.value = `Total · ${rows.length} line${rows.length === 1 ? "" : "s"}`;
      if (totalKeys.includes(c.key) && rows.length) {
        const L = colLetter(i + 1);
        const result = rows.reduce((a, r) => a + (Number(r[c.key]) || 0), 0);
        cell.value = { formula: `SUM(${L}${first}:${L}${last})`, result };
        cell.numFmt = c.type === "money" ? '"R" #,##0.00' : "#,##0.##";
      }
    });

    if (rows.length) ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: last, column: lastCol } };
    const note = ws.getCell(`A${8 + rows.length}`);
    note.value = `Generated ${nowSA()} · GBS Billings · Amounts in South African Rand`;
    note.font = { name: "Calibri", size: 9, italic: true, color: { argb: MUTED } };

    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
  }

  // -------------------------------------------------------------------
  // PDF: client billing statement (to attach to the invoice)
  //   rows: [{ work_date, consultant, service, narrative, hours, fee, expenses, total }]
  // -------------------------------------------------------------------
  async function clientStatement({ filename, client, contact, period, rows, vatRate = 0.15, invoiceNos = [] }) {
    if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF library didn't load — check your internet connection and try again.");
    const doc = new window.jspdf.jsPDF({ unit: "mm", format: "a4" });
    const W = doc.internal.pageSize.getWidth(), M = 16;
    const logo = await logoDataUrl();

    // Header
    if (logo) doc.addImage(logo, "PNG", M, 12, 55, 15);
    else { doc.setFont("helvetica", "bold"); doc.setFontSize(13); doc.setTextColor(...RGB_BLUE); doc.text("Global Business Solutions", M, 21); }
    doc.setFont("helvetica", "bold"); doc.setFontSize(17); doc.setTextColor(...RGB_BLUE);
    doc.text("Billing statement", W - M, 19, { align: "right" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(...RGB_MUTED);
    doc.text(`Prepared ${nowSA()}`, W - M, 25, { align: "right" });
    doc.setDrawColor(...RGB_BLUE); doc.setLineWidth(0.8); doc.line(M, 32, W - M, 32);

    // Client + period
    doc.setFontSize(8.5); doc.setTextColor(...RGB_MUTED);
    doc.text("CLIENT", M, 40);
    doc.text("PERIOD", W / 2 + 10, 40);
    doc.setFont("helvetica", "bold"); doc.setFontSize(12); doc.setTextColor(...RGB_INK);
    doc.text(doc.splitTextToSize(client || "", W / 2 - M), M, 46);
    doc.text(period || "", W / 2 + 10, 46);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.setTextColor(...RGB_MUTED);
    if (contact) doc.text(doc.splitTextToSize("Attention: " + contact, W / 2 - M), M, 52);
    if (invoiceNos.length) doc.text(doc.splitTextToSize("Invoice " + invoiceNos.join(", "), W / 2 - M - 10), W / 2 + 10, 52);

    // Table
    const fees = rows.reduce((a, r) => a + Number(r.fee || 0), 0);
    const exp = rows.reduce((a, r) => a + Number(r.expenses || 0), 0);
    const excl = fees + exp, vat = Math.round(excl * vatRate * 100) / 100;
    doc.autoTable({
      startY: 60,
      margin: { left: M, right: M, bottom: 20 },
      head: [["Date", "Consultant", "Service", "Description of work", "Hrs", "Fees", "Expenses", "Total excl. VAT"]],
      body: rows.map((r) => [
        new Date(r.work_date + "T00:00:00Z").toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" }),
        r.consultant || "", r.service || "", r.narrative || "",
        r.hours != null && r.hours !== "" ? String(Math.round(Number(r.hours) * 100) / 100) : "",
        money(r.fee), Number(r.expenses) ? money(r.expenses) : "—", money(r.total),
      ]),
      styles: { font: "helvetica", fontSize: 8.5, cellPadding: 2.2, textColor: RGB_INK, lineColor: RGB_LINE, lineWidth: 0, valign: "top" },
      headStyles: { fillColor: RGB_BLUE, textColor: 255, fontStyle: "bold", fontSize: 8.5 },
      alternateRowStyles: { fillColor: RGB_SOFT },
      columnStyles: {
        0: { cellWidth: 20 }, 1: { cellWidth: 25 }, 2: { cellWidth: 22 }, 3: { cellWidth: "auto" },
        4: { cellWidth: 11, halign: "right" }, 5: { cellWidth: 22, halign: "right" }, 6: { cellWidth: 20, halign: "right" }, 7: { cellWidth: 24, halign: "right", fontStyle: "bold" },
      },
      didParseCell: (d) => { if (d.section === "head" && d.column.index >= 4) d.cell.styles.halign = "right"; },
    });

    // Totals
    let y = doc.lastAutoTable.finalY + 8;
    if (y > doc.internal.pageSize.getHeight() - 50) { doc.addPage(); y = 24; }
    const lines = [["Professional fees", money(fees)], ["Expenses (travel, car hire, accommodation)", money(exp)],
                   ["Total excl. VAT", money(excl)], [`VAT ${Math.round(vatRate * 100)}%`, money(vat)]];
    doc.setFontSize(9.5);
    lines.forEach(([k, v], i) => {
      doc.setFont("helvetica", i === 2 ? "bold" : "normal"); doc.setTextColor(...(i === 2 ? RGB_INK : RGB_MUTED));
      doc.text(k, W - M - 60, y, { align: "right" }); doc.setTextColor(...RGB_INK); doc.text(v, W - M, y, { align: "right" });
      y += 6;
    });
    doc.setFillColor(...RGB_BLUE); doc.roundedRect(W - M - 100, y - 3, 100, 10, 2, 2, "F");
    doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(255);
    doc.text("Total incl. VAT", W - M - 60, y + 3.6, { align: "right" });
    doc.text(money(excl + vat), W - M - 3, y + 3.6, { align: "right" });

    // Footer on every page
    const pages = doc.internal.getNumberOfPages(), H = doc.internal.pageSize.getHeight();
    for (let i = 1; i <= pages; i++) {
      doc.setPage(i);
      doc.setDrawColor(...RGB_LINE); doc.setLineWidth(0.3); doc.line(M, H - 14, W - M, H - 14);
      doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(...RGB_MUTED);
      doc.text("Global Business Solutions · Billing statement — supporting detail, not a tax invoice", M, H - 9);
      doc.text(`Page ${i} of ${pages}`, W - M, H - 9, { align: "right" });
    }
    doc.save(filename);
  }

  window.GBSExport = { xlsx, clientStatement };
})();
