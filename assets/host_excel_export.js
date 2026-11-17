let lastHostExport = null;

async function downloadHostExcel() {
  const button = document.getElementById("downloadHostExcel");
  if (!lastHostExport || button.disabled) return;
  button.disabled = true;
  try {
    if (typeof JSZip === "undefined") throw new Error("Không tải được thư viện xuất Excel.");
    const mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const doc = document.implementation.createDocument(mainNs, "worksheet", null);
    const root = doc.documentElement;
    function child(parent, name, attributes = {}, text = null) {
      const element = doc.createElementNS(mainNs, name);
      Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, String(value)));
      if (text !== null) element.textContent = String(text);
      parent.appendChild(element);
      return element;
    }
    const views = child(root, "sheetViews");
    const view = child(views, "sheetView", { workbookViewId: 0 });
    child(view, "pane", { ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" });
    const columns = child(root, "cols");
    lastHostExport.headers.forEach((_, index) => {
      child(columns, "col", { min: index + 1, max: index + 1, width: index === 0 ? 16 : index === 1 ? 16 : 20, customWidth: 1 });
    });
    const sheetData = child(root, "sheetData");
    const values = [lastHostExport.headers, ...lastHostExport.rows, lastHostExport.total];
    values.forEach((row, rowIndex) => {
      const sheetRow = child(sheetData, "row", { r: rowIndex + 1 });
      row.forEach((value, columnIndex) => {
        const reference = String.fromCharCode(65 + columnIndex) + (rowIndex + 1);
        const isHeader = rowIndex === 0;
        const isTotal = rowIndex === values.length - 1;
        const isPercent = columnIndex === 9 || columnIndex === 13;
        const style = isHeader ? 1 : isTotal ? (isPercent ? 4 : 3) : isPercent ? 2 : 0;
        if (typeof value === "number" && Number.isFinite(value)) {
          const cell = child(sheetRow, "c", { r: reference, s: style, t: "n" });
          child(cell, "v", {}, value);
        } else {
          const cell = child(sheetRow, "c", { r: reference, s: style, t: "inlineStr" });
          child(child(cell, "is"), "t", {}, value ?? "");
        }
      });
    });
    const zip = new JSZip();
    zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + relNs + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + mainNs + '" xmlns:r="' + relNs + '"><sheets><sheet name="Host Performance" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + relNs + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + relNs + '/styles" Target="styles.xml"/></Relationships>');
    zip.file("xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="' + mainNs + '"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="10" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
    zip.file("xl/worksheets/sheet1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + new XMLSerializer().serializeToString(doc));
    const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const range = selectedRange();
    const dateText = (date) => [date.getFullYear(), String(date.getMonth() + 1).padStart(2, "0"), String(date.getDate()).padStart(2, "0")].join("-");
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `host-performance-${dateText(range.start)}-${dateText(range.end)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    console.error("Không thể xuất Excel:", error);
    alert("Không thể xuất file Excel. Vui lòng thử lại.");
  } finally {
    button.disabled = false;
  }
}
