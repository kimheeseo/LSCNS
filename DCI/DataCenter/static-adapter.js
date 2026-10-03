(() => {
  'use strict';
  const nativeFetch = window.fetch.bind(window);
  const jsonResponse = (value, status = 200) => new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json; charset=utf-8' } });
  const text = value => value == null ? '' : typeof value === 'string' ? value : JSON.stringify(value);
  const flatten = (value, prefix = '', rows = [], depth = 0) => {
    if (rows.length >= 6000) return rows;
    if (depth > 8 || value == null || typeof value !== 'object') { rows.push([prefix || 'value', text(value)]); return rows; }
    if (Array.isArray(value)) value.slice(0, 1200).forEach((v, i) => flatten(v, `${prefix}[${i}]`, rows, depth + 1));
    else Object.entries(value).forEach(([k, v]) => flatten(v, prefix ? `${prefix}.${k}` : k, rows, depth + 1));
    return rows;
  };
  const addRows = (sheet, rows) => {
    for (const row of rows || []) sheet.addRow(row.map(cell => typeof cell === 'object' && cell !== null ? text(cell.text ?? cell) : text(cell)));
  };
  async function workbook(payload) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'AI Data Center BOM Design Tool';
    const req = wb.addWorksheet('1_Requirements');
    req.addRow(['Data Center BOM Design', `v${payload.version || '7.3.2'}`]);
    req.addRow(['Language', payload.language || 'ko']);
    req.addRow([]); req.addRow(['Field', 'Value']);
    addRows(req, (payload.requirements || []).map(r => [r.field, r.value]));
    const result = wb.addWorksheet('2_Design Result');
    result.addRow(['Field', 'Value']);
    addRows(result, (payload.designResult || []).map(r => [r.field, r.value]));
    result.addRow([]); addRows(result, flatten(payload.design || {}));
    const generic = wb.addWorksheet('3_Generic BOM'); generic.addRow(['Category','Item','Purchase qty','Installed qty','Spare qty','Unit','Basis']);addRows(generic,(payload.design?.bom||[]).map(x=>[x.category,x.item,x.qty,x.installedQty??x.qty,x.spareQty??0,x.unit,x.basis]));
    const receipt = wb.addWorksheet('4_Product Match');receipt.addRow(['Category','Vendor','Product','Quantity','Evidence','Source']);addRows(receipt,(payload.design?.products||[]).map(x=>[x.category,x.vendor,x.product,x.qty,x.evidence,x.source]));const audit=wb.addWorksheet('5_Engineering Audit');audit.addRow(['Field','Value']);addRows(audit,flatten({input:payload.design?.input,warnings:payload.design?.warnings,facility:payload.design?.facility,optical:payload.design?.optical,racks:payload.design?.racks,portAudit:payload.design?.portAudit}));
    const nodes=wb.addWorksheet('6_Port Nodes');nodes.addRow(['Endpoint','Role','Logical capacity','Logical / cage','Server ports','Leaf-spine ports','Core ports']);addRows(nodes,(payload.design?.portAudit?.nodes||[]).map(x=>[x.id,x.role,x.capacity,x.logicalPerCage,x.used.server,x.used.leafSpine,x.used.core]));
    const routes=wb.addWorksheet('7_Link Routes');routes.addRow(['Segment','Endpoint A','Endpoint B','Installed logical links']);addRows(routes,(payload.design?.portAudit?.routes||[]).map(x=>[x.segment,x.a,x.b,x.links]));
    for (const ws of [req, result, generic, receipt, audit, nodes, routes]) {
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.columns = Array.from({ length: Math.max(2, ws.columnCount) }, () => ({ width: 36 }));
      ws.getRow(1).font = { bold: true };
      ws.eachRow(row => row.eachCell(cell => { cell.alignment = { vertical: 'top', wrapText: true }; }));
    }
    return wb.xlsx.writeBuffer();
  }
  window.fetch = async (input, init = {}) => {
    const raw = typeof input === 'string' ? input : input.url;
    const url = new URL(raw, location.href);
    if (url.pathname.endsWith('/api/design')) {
      try { return jsonResponse(DCBOMEngine.design(JSON.parse(init.body || '{}'))); }
      catch (error) { return jsonResponse({ status: 'REVIEW', error: error.message }, 400); }
    }
    if (url.pathname.endsWith('/api/export-xlsx')) {
      try {
        const data = await workbook(JSON.parse(init.body || '{}'));
        return new Response(data, { status: 200, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
      } catch (error) { return jsonResponse({ error: error.message }, 400); }
    }
    return nativeFetch(input, init);
  };
})();

