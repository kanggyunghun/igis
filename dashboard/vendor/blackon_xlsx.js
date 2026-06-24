/* black_on 대시보드 - 편집분을 통합 단일시트(Sheet1) 워크북에 패치하는 순수 로직.
 * 브라우저(인라인)와 node 테스트가 동일 코드를 공유한다. DOM 의존 없음.
 *
 * rows: 편집 가능한 4개 탭(국내주식/국내선물/해외주식/해외선물)의 레코드 평면 배열.
 *   각 row: { sourceRow:int|null, deleted:bool, isNew:bool, category, group, kind, sub,
 *             stockName, code, ticker, qty, price, value, targetPct(소수|null), memo }
 *   - 주식: qty/price/value 편집, targetPct/memo 신규열.
 *   - 선물: qty(계약수, 부호=방향)/value(부호 보존) 편집, memo 신규열 (목표 없음).
 *   소계/총계/현금 행은 sourceRow 로 지정되지 않으므로 건드리지 않는다.
 */
(function (root) {
  "use strict";

  // 통합 단일시트 0-indexed 컬럼 (Python COL 과 동일 레이아웃)
  var C = {
    no: 0, group: 1, kind: 2, sub: 3, fundCode: 4, fundName: 5, code: 6, name: 7,
    prevQty: 9, dQty: 10, price: 11, value: 29, ticker: 27,
    target: 32, memo: 33,  // 신규 컬럼
  };
  var STOCK_CATS = { domesticStock: 1, overseasStock: 1 };

  function colToA1(c) {
    var s = "";
    c += 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function a1ToCol(a1) { var c = 0; for (var i = 0; i < a1.length; i++) c = c * 26 + (a1.charCodeAt(i) - 64); return c - 1; }
  function encRange(rng) { return colToA1(rng.s.c) + (rng.s.r + 1) + ":" + colToA1(rng.e.c) + (rng.e.r + 1); }
  function decRange(ref) {
    var p = ref.split(":");
    function pt(x) { var m = x.match(/([A-Z]+)(\d+)/); return { c: a1ToCol(m[1]), r: parseInt(m[2], 10) - 1 }; }
    var s = pt(p[0]); return { s: s, e: p[1] ? pt(p[1]) : s };
  }
  function maxRow(ws) { return ws["!ref"] ? decRange(ws["!ref"]).e.r : 0; }
  function bump(ws, r, c) {
    var rng = ws["!ref"] ? decRange(ws["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    if (r > rng.e.r) rng.e.r = r;
    if (c > rng.e.c) rng.e.c = c;
    ws["!ref"] = encRange(rng);
  }
  function setCell(XLSX, ws, r, c, v, t) {
    var addr = XLSX.utils.encode_cell({ r: r, c: c });
    if (v === null || v === undefined || v === "") delete ws[addr];
    else ws[addr] = { t: t || (typeof v === "number" ? "n" : "s"), v: v };
    bump(ws, r, c);
  }
  function delRow(XLSX, ws, r, maxCol) {
    for (var c = 0; c <= maxCol; c++) delete ws[XLSX.utils.encode_cell({ r: r, c: c })];
  }
  function pct(v) { return v === null || v === undefined ? null : Math.round(v * 1000000) / 10000; }

  function applyEditsToWorkbook(XLSX, wb, data, rows) {
    var ws = wb.Sheets[data.sheetName];
    if (!ws) throw new Error("시트를 찾을 수 없음: " + data.sheetName);
    setCell(XLSX, ws, 0, C.target, "목표비중(%)");
    setCell(XLSX, ws, 0, C.memo, "메모");
    var appendAt = maxRow(ws) + 1;

    (rows || []).forEach(function (row) {
      if (row.deleted && row.sourceRow) { delRow(XLSX, ws, row.sourceRow - 1, C.memo); return; }
      var isNew = row.isNew || !row.sourceRow;
      var r = isNew ? appendAt++ : row.sourceRow - 1;
      var isStock = !!STOCK_CATS[row.category];

      if (isNew) {
        if (row.group) setCell(XLSX, ws, r, C.group, row.group);
        if (row.kind) setCell(XLSX, ws, r, C.kind, row.kind);
        if (row.sub) setCell(XLSX, ws, r, C.sub, row.sub);
        if (row.code) setCell(XLSX, ws, r, C.code, row.code);
        if (row.ticker) setCell(XLSX, ws, r, C.ticker, row.ticker);
        if (row.stockName) setCell(XLSX, ws, r, C.name, row.stockName);
      }
      // 보유수량/계약수 -> 전일보유수량 칸에 쓰고 당일증감은 0 (당일보유 = 신규값)
      if (row.qty !== undefined && row.qty !== null) { setCell(XLSX, ws, r, C.prevQty, row.qty); setCell(XLSX, ws, r, C.dQty, 0); }
      if (isStock && row.price !== undefined && row.price !== null) setCell(XLSX, ws, r, C.price, row.price);
      if (row.value !== undefined && row.value !== null) setCell(XLSX, ws, r, C.value, row.value);
      if (isStock) setCell(XLSX, ws, r, C.target, pct(row.targetPct));
      setCell(XLSX, ws, r, C.memo, row.memo || "");
    });
    return wb;
  }

  var api = { applyEditsToWorkbook: applyEditsToWorkbook };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.BlackOnXlsx = api;
})(typeof window !== "undefined" ? window : this);
