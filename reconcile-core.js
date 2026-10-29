/* 居家服務月底對帳 — 判斷核心（純函式，無 DOM，瀏覽器與 Node 共用）
 *
 * A＝巨安系統服務紀錄清冊；B＝服務員繳回紙本服務紀錄表。
 * 比對鍵＝日期｜個案｜服務員｜碼別，兩邊各自加總次數後比較。
 * 簽名只看紙本（B）；系統有、紙本無的列沒有紙本可核對，簽名狀態為「無紙本」，不算簽名異常。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.LTC = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // 長期照顧給付及支付基準 BA 碼（113.05 民眾參閱版；價格僅供估算金額，正式以最新公告為準）
  const BA_CODES = {
    BA01: ['基本身體清潔', 260], BA02: ['基本日常照顧', 195], BA03: ['測量生命徵象', 35],
    BA04: ['協助進食或管灌餵食', 130], BA05: ['餐食照顧', 310], BA07: ['協助沐浴及洗頭', 325],
    BA08: ['足部照護', 500], BA09: ['到宅沐浴車服務-第1型', 2200], BA09a: ['到宅沐浴車服務-第2型', 2500],
    BA10: ['翻身拍背', 155], BA11: ['肢體關節活動', 195], BA12: ['協助上(下)樓梯', 130],
    BA13: ['陪同外出', 195], BA14: ['陪同就醫', 685], BA15: ['家務協助', 195],
    BA16: ['代購或代領或代送服務', 130], BA17a: ['人工氣道管內分泌物抽吸', 75],
    BA17b: ['口腔內分泌物抽吸', 65], BA17c: ['尿管及鼻胃管之清潔與固定', 50],
    BA17d1: ['血糖機驗血糖', 50], BA17d2: ['甘油球通便', 50], BA17e: ['BA17e（以最新公告為準）', null],
    BA18: ['安全看視', 200], BA20: ['陪伴服務', 175], BA22: ['巡視服務', 130],
    BA23: ['協助洗頭', 200], BA24: ['協助排泄', 220],
  };

  /* ---------- 正規化 ---------- */

  function toHalf(s) {
    return String(s == null ? '' : s)
      .replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/　/g, ' ');
  }

  function normName(s) {
    return toHalf(s).replace(/\s+/g, '').trim();
  }

  // BA4 / ba04 / ＢＡ０７ / BA17D1 → BA04 / BA07 / BA17d1；無法辨識回傳 null
  function normCode(s) {
    const t = toHalf(s).replace(/\s+/g, '').toUpperCase();
    const m = t.match(/^BA0*(\d{1,2})([A-E]\d?)?$/);
    if (!m) return null;
    return 'BA' + m[1].padStart(2, '0') + (m[2] ? m[2].toLowerCase() : '');
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function validYmd(y, m, d) {
    if (!(y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1) return null; // 2/30 之類
    return y + '-' + pad2(m) + '-' + pad2(d);
  }

  // 支援 2026/10/17、2026-10-17、115/10/17（民國）、20261017、1151017、10/17（用 defaultYear）、Excel 序號、Date 物件
  function normDate(v, defaultYear) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) {
      // Excel 讀出的日期在台灣時區偶爾會落在前一天 23:5x（1900 年代的歷史時差），貼近午夜就進位到隔天
      if (v.getHours() === 23 && v.getMinutes() >= 50) v = new Date(v.getTime() + 15 * 60000);
      return validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const dt = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
      return validYmd(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
    const s = toHalf(v).trim().replace(/\s.*$/, '').replace(/[年月]/g, '/').replace(/日/g, '');
    let m = s.match(/^(\d{2,4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
    if (m) {
      let y = +m[1];
      if (y < 1911) y += 1911; // 民國年
      return validYmd(y, +m[2], +m[3]);
    }
    m = s.match(/^(\d{8})$/);
    if (m) return validYmd(+s.slice(0, 4), +s.slice(4, 6), +s.slice(6, 8));
    m = s.match(/^(\d{7})$/);
    if (m) return validYmd(+s.slice(0, 3) + 1911, +s.slice(3, 5), +s.slice(5, 7));
    m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})$/);
    if (m && defaultYear) return validYmd(defaultYear, +m[1], +m[2]);
    return null;
  }

  // 回傳 { n, note }；空白以 1 計並註記；非正數或非數字回傳 n=null
  function parseCount(v) {
    if (v == null || String(v).trim() === '') return { n: 1, note: '次數空白，以 1 計' };
    const t = toHalf(v).replace(/[次趟回\s]/g, '');
    if (!/^\d+(\.\d+)?$/.test(t)) return { n: null, note: '次數無法辨識：' + v };
    const n = Number(t);
    if (!(n > 0)) return { n: null, note: '次數須大於 0：' + v };
    return { n, note: '' };
  }

  // 簽名：signed / proxy（家屬代簽，視同有效） / unsigned / unknown（沒有簽名欄位）
  function parseSignature(v, hasColumn) {
    if (!hasColumn) return { s: 'unknown', note: '' };
    const t = toHalf(v == null ? '' : v).trim();
    if (t === '') return { s: 'unsigned', note: '' };
    if (/未簽|沒簽|無簽|漏簽|^無$|^否$|^N$|^X$|^✗$|^×$/i.test(t)) return { s: 'unsigned', note: '' };
    const who = (t.match(/[（(]\s*([^）)]+?)\s*[）)]/) || [])[1] || '';
    if (/代簽|家屬|代理/.test(t)) return { s: 'proxy', note: who };
    if (/已簽|簽名|簽章|有|^是$|^Y$|^V$|^✓$|^✔$|^O$|^○$/i.test(t) || who) return { s: 'signed', note: who };
    // 欄位內只寫了人名，也視為已簽
    if (/^[一-鿿]{2,4}$/.test(t)) return { s: 'signed', note: t };
    return { s: 'unsigned', note: '簽名欄無法辨識：' + t };
  }

  /* ---------- 表格解析 ---------- */

  const HEADER_SYNONYMS = {
    date: ['服務日期', '日期', '服務日', '執行日期'],
    client: ['個案姓名', '個案', '案主姓名', '案主', '個案名稱', '姓名'],
    worker: ['服務員', '服務人員', '照服員', '居服員', '照顧服務員', '服務員姓名'],
    code: ['服務項目', '碼別', '服務代碼', '項目代碼', '照顧組合', '組合代碼', '項目'],
    count: ['次數', '數量', '單位', '服務次數', '組數'],
    sig: ['簽名', '案主簽名', '簽章', '個案簽名', '案家簽名', '簽收'],
  };

  function matchHeader(cell) {
    const t = normName(cell);
    if (!t) return null;
    for (const key of Object.keys(HEADER_SYNONYMS)) {
      if (HEADER_SYNONYMS[key].some(w => t === w)) return key;
    }
    for (const key of Object.keys(HEADER_SYNONYMS)) {
      if (HEADER_SYNONYMS[key].some(w => w.length >= 2 && t.includes(w))) return key;
    }
    return null;
  }

  // 純文字 → 2D 陣列。有 tab 用 tab；有逗號（CSV）用逗號；否則以空白切，多出的欄位併入最後一欄
  function textToRows(text) {
    const lines = String(text || '').replace(/\r/g, '').split('\n').filter(l => l.trim() !== '');
    if (!lines.length) return [];
    if (lines.some(l => l.includes('\t'))) return lines.map(l => l.split('\t'));
    if (lines.every(l => l.includes(','))) return lines.map(parseCsvLine);
    const headerCols = lines[0].trim().split(/\s+/).length;
    return lines.map(l => {
      const parts = l.trim().split(/\s+/);
      if (parts.length > headerCols) {
        return parts.slice(0, headerCols - 1).concat(parts.slice(headerCols - 1).join(' '));
      }
      return parts;
    });
  }

  function parseCsvLine(line) {
    const out = []; let cur = ''; let q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (q) {
        if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (c === '"') q = false;
        else cur += c;
      } else if (c === '"') q = true;
      else if (c === ',') { out.push(cur); cur = ''; }
      else cur += c;
    }
    out.push(cur);
    return out;
  }

  // 一格可能寫多個碼：「BA01×1、BA07*2」
  function splitCodes(cell) {
    const t = toHalf(cell);
    const re = /BA\s*0*\d{1,2}(?:[A-E]\d?)?(?![A-Z0-9])\s*(?:[x×X*＊]\s*(\d+(?:\.\d+)?))?/gi;
    const out = []; let m;
    while ((m = re.exec(t))) {
      const raw = m[0].replace(/\s*[x×X*＊]\s*\d+(\.\d+)?$/, '');
      out.push({ code: normCode(raw), raw: raw.trim(), count: m[1] != null ? Number(m[1]) : null });
    }
    return out;
  }

  /**
   * 解析一份清冊（2D 陣列）成紀錄。
   * 回傳 { records, problems, columns, hasSig }
   * problems：無法使用的列（缺日期／個案／碼別、次數錯誤），一律列出，不默默丟棄。
   */
  function parseRecords(rows, src, opts) {
    opts = opts || {};
    const problems = [];
    let headerIdx = -1; let colMap = null;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const map = {};
      rows[i].forEach((cell, j) => { const k = matchHeader(cell); if (k && map[k] == null) map[k] = j; });
      const found = ['date', 'client', 'code'].filter(k => map[k] != null).length;
      if (found === 3) { headerIdx = i; colMap = map; break; }
    }
    if (headerIdx < 0) {
      return { records: [], problems: [{ src, row: 0, msg: '找不到表頭：至少需要「日期」「個案姓名」「服務項目／碼別」三欄' }], columns: null, hasSig: false };
    }
    const hasSig = colMap.sig != null;
    const body = rows.slice(headerIdx + 1);

    // 推定年份：用資料中出現最多的完整年份，供「10/17」這種沒寫年的日期使用
    let defaultYear = opts.defaultYear;
    if (!defaultYear) {
      const years = {};
      body.forEach(r => { const d = normDate(r[colMap.date]); if (d) years[d.slice(0, 4)] = (years[d.slice(0, 4)] || 0) + 1; });
      const best = Object.entries(years).sort((a, b) => b[1] - a[1])[0];
      defaultYear = best ? +best[0] : new Date().getFullYear();
    }

    const records = [];
    body.forEach((r, i) => {
      const rowNo = headerIdx + i + 2; // 對應試算表列號（1 起算、含表頭）
      const cells = r.map(c => (c == null ? '' : c));
      if (cells.every(c => String(c).trim() === '')) return;
      const rawDate = cells[colMap.date];
      const date = normDate(rawDate, defaultYear);
      const client = normName(cells[colMap.client]);
      const worker = colMap.worker != null ? normName(cells[colMap.worker]) : '';
      const codeCell = cells[colMap.code];
      const errs = [];
      if (!date) errs.push('日期無法辨識：「' + (rawDate === '' ? '空白' : rawDate) + '」');
      if (!client) errs.push('個案姓名空白');
      const codes = splitCodes(codeCell);
      if (!codes.length) errs.push('服務項目無法辨識：「' + (String(codeCell).trim() || '空白') + '」');
      if (errs.length) { problems.push({ src, row: rowNo, msg: errs.join('；'), raw: cells.join(' | ') }); return; }

      const sig = parseSignature(hasSig ? cells[colMap.sig] : '', hasSig);
      const countCell = colMap.count != null ? cells[colMap.count] : '';
      codes.forEach(c => {
        let n; let note = '';
        if (c.count != null) n = c.count;
        else if (codes.length > 1) { n = 1; note = '同格多碼未註明次數，以 1 計'; }
        else { const pc = parseCount(countCell); n = pc.n; note = pc.note; }
        if (n == null) { problems.push({ src, row: rowNo, msg: note, raw: cells.join(' | ') }); return; }
        records.push({
          src, row: rowNo, date, client, worker, code: c.code, rawCode: c.raw, count: n,
          sig: sig.s, sigNote: sig.note, notes: note ? [note] : [],
        });
      });
    });
    return { records, problems, columns: colMap, hasSig };
  }

  /* ---------- 比對 ---------- */

  function keyOf(r) { return [r.date, r.client, r.worker, r.code].join('|'); }

  function aggregate(records) {
    const map = new Map();
    for (const r of records) {
      const k = keyOf(r);
      let g = map.get(k);
      if (!g) { g = { date: r.date, client: r.client, worker: r.worker, code: r.code, count: 0, rows: [], sigs: [], notes: [] }; map.set(k, g); }
      g.count += r.count; g.rows.push(r.row); g.sigs.push(r); g.notes.push(...r.notes);
    }
    return map;
  }

  // 同一列多筆紙本的簽名合併：全部有效才算有效；有效與未簽混雜 → mixed
  function combineSig(recs) {
    if (!recs || !recs.length) return { s: 'na', note: '' };
    const ss = recs.map(r => r.sig);
    if (ss.every(s => s === 'unknown')) return { s: 'unknown', note: '' };
    const ok = ss.filter(s => s === 'signed' || s === 'proxy').length;
    const notes = [...new Set(recs.map(r => r.sigNote).filter(Boolean))].join('、');
    if (ok === ss.length) return { s: ss.includes('proxy') ? 'proxy' : 'signed', note: notes };
    if (ok === 0) return { s: 'unsigned', note: notes };
    return { s: 'mixed', note: notes };
  }

  function levenshtein(a, b) {
    const m = a.length, n = b.length;
    const d = Array.from({ length: m + 1 }, (_, i) => [i].concat(Array(n).fill(0)));
    for (let j = 1; j <= n; j++) d[0][j] = j;
    for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    return d[m][n];
  }

  function price(code) { return BA_CODES[code] ? BA_CODES[code][1] : null; }
  function codeName(code) { return BA_CODES[code] ? BA_CODES[code][0] : null; }

  /**
   * 主比對。回傳 { rows, summary }
   * row.status：
   *   match         兩邊項目與次數一致
   *   count         兩邊都有，次數不同
   *   a_only        系統有、紙本無（可能溢報）
   *   b_only        紙本有、系統無（可能漏報）
   *   code_mismatch 同日同個案同服務員，系統與紙本的碼別不同（疑似填錯碼）
   *   client_mismatch 同日同服務員同碼別，個案姓名差一字（疑似寫錯字）
   *   worker_mismatch 同日同個案同碼別，服務員不同（代班或登錯人）
   * row.sig：signed / proxy / unsigned / mixed / unknown（紙本無簽名欄） / na（無紙本）
   */
  function reconcile(aRecords, bRecords, opts) {
    // 任一邊沒有服務員欄位時，比對鍵不含服務員，否則每一列都會被誤判成服務員不一致
    if (opts && opts.ignoreWorker) {
      aRecords = aRecords.map(r => Object.assign({}, r, { worker: '' }));
      bRecords = bRecords.map(r => Object.assign({}, r, { worker: '' }));
    }
    const A = aggregate(aRecords);
    const B = aggregate(bRecords);
    const rows = [];
    const aOnly = []; const bOnly = [];

    for (const [k, a] of A) {
      const b = B.get(k);
      if (!b) { aOnly.push(a); continue; }
      rows.push(makeRow(b.count === a.count ? 'match' : 'count', a, b));
    }
    for (const [k, b] of B) if (!A.has(k)) bOnly.push(b);

    // 配對單邊列：先找「只差一個欄位」的，越具體越先配；每組只在一對一時配對，避免猜錯
    const usedA = new Set(); const usedB = new Set();
    function pairBy(groupKeyFn, accept, status) {
      const ga = new Map(); const gb = new Map();
      aOnly.forEach(x => { if (!usedA.has(x)) push(ga, groupKeyFn(x), x); });
      bOnly.forEach(x => { if (!usedB.has(x)) push(gb, groupKeyFn(x), x); });
      for (const [gk, as] of ga) {
        const bs = gb.get(gk);
        if (!bs) continue;
        const cands = [];
        as.forEach(a => bs.forEach(b => { if (accept(a, b)) cands.push([a, b]); }));
        // 每一邊在候選中只出現一次才採用
        const ca = new Map(); const cb = new Map();
        cands.forEach(([a, b]) => { ca.set(a, (ca.get(a) || 0) + 1); cb.set(b, (cb.get(b) || 0) + 1); });
        cands.forEach(([a, b]) => {
          if (ca.get(a) === 1 && cb.get(b) === 1 && !usedA.has(a) && !usedB.has(b)) {
            usedA.add(a); usedB.add(b); rows.push(makeRow(status, a, b));
          }
        });
      }
    }
    pairBy(x => [x.date, x.worker, x.code].join('|'),
      (a, b) => a.client.length === b.client.length && levenshtein(a.client, b.client) === 1, 'client_mismatch');
    pairBy(x => [x.date, x.client, x.code].join('|'), (a, b) => a.worker !== b.worker, 'worker_mismatch');
    pairBy(x => [x.date, x.client, x.worker].join('|'), (a, b) => a.code !== b.code && a.count === b.count, 'code_mismatch');

    aOnly.forEach(a => { if (!usedA.has(a)) rows.push(makeRow('a_only', a, null)); });
    bOnly.forEach(b => { if (!usedB.has(b)) rows.push(makeRow('b_only', null, b)); });

    rows.sort((x, y) => (x.date + x.client + x.worker + x.code).localeCompare(y.date + y.client + y.worker + y.code));

    // 月份：以 A（系統）最多的年月為對帳月，其他月份標註
    const months = {};
    rows.forEach(r => { const ym = r.date.slice(0, 7); months[ym] = (months[ym] || 0) + 1; });
    const month = (Object.entries(months).sort((a, b) => b[1] - a[1])[0] || [''])[0];
    rows.forEach(r => {
      if (month && r.date.slice(0, 7) !== month) r.flags.push('非對帳月份（' + month + '）資料');
      const codes = [r.aCode, r.bCode].filter(Boolean);
      codes.forEach(c => { if (!BA_CODES[c]) r.flags.push(c + ' 不在 BA 碼表內，請確認'); });
    });

    return { rows, summary: summarize(rows), month };
  }

  function push(map, k, v) { if (!map.has(k)) map.set(k, []); map.get(k).push(v); }

  function makeRow(status, a, b) {
    const base = a || b;
    const sig = b ? combineSig(b.sigs) : { s: 'na', note: '' };
    const row = {
      status,
      date: base.date,
      client: a ? a.client : b.client, bClient: b ? b.client : null,
      worker: a ? a.worker : b.worker, bWorker: b ? b.worker : null,
      code: a ? a.code : b.code, aCode: a ? a.code : null, bCode: b ? b.code : null,
      aCount: a ? a.count : 0, bCount: b ? b.count : 0,
      aRows: a ? a.rows : [], bRows: b ? b.rows : [],
      sig: sig.s, sigNote: sig.note,
      flags: [...new Set([...(a ? a.notes : []), ...(b ? b.notes : [])])],
    };
    if (a && a.rows.length > 1) row.flags.push('系統同鍵 ' + a.rows.length + ' 列已合併計次（列 ' + a.rows.join('、') + '），請確認非重複登打');
    if (b && b.rows.length > 1) row.flags.push('紙本同鍵 ' + b.rows.length + ' 列已合併計次');
    row.itemOk = status === 'match';
    row.sigOk = sig.s === 'signed' || sig.s === 'proxy' || sig.s === 'unknown';
    row.sigIssue = b != null && (sig.s === 'unsigned' || sig.s === 'mixed');
    row.ok = row.itemOk && row.sigOk;
    row.amount = estimateAmount(row);
    row.advice = advise(row);
    return row;
  }

  // 系統申報金額 − 紙本支持金額；正數＝系統多報（有被追扣風險），負數＝系統少報（漏收）
  function estimateAmount(r) {
    const pa = r.aCode ? price(r.aCode) : 0;
    const pb = r.bCode ? price(r.bCode) : 0;
    if ((r.aCode && pa == null) || (r.bCode && pb == null)) return null;
    return r.aCount * (pa || 0) - r.bCount * (pb || 0);
  }

  function advise(r) {
    const nm = c => c + (codeName(c) ? ' ' + codeName(c) : '');
    switch (r.status) {
      case 'match':
        if (r.sig === 'unsigned') return '項目次數一致，但紙本未簽名。請服務員補請案主或家屬簽名後再申報。';
        if (r.sig === 'mixed') return '項目次數一致，但同鍵多張紙本中有未簽名者，請逐張核對。';
        if (r.sig === 'unknown') return '項目次數一致。紙本資料沒有簽名欄，簽名未核對。';
        return '項目、次數一致，紙本已' + (r.sig === 'proxy' ? '由家屬代簽' : '簽名') + '。';
      case 'count': {
        const d = r.aCount - r.bCount;
        return d > 0
          ? '系統比紙本多 ' + d + ' 次。以紙本為準時，系統應刪減 ' + d + ' 次，避免溢報。'
          : '紙本比系統多 ' + (-d) + ' 次。確認服務確實提供後，於系統補登 ' + (-d) + ' 次。';
      }
      case 'a_only':
        return '系統已登 ' + nm(r.aCode) + '，紙本沒有這筆。請向服務員確認是否漏繳紙本；若未提供服務，應從系統刪除。';
      case 'b_only':
        return '紙本有 ' + nm(r.bCode) + '，系統沒有登錄。確認後於系統補登，否則無法申報。';
      case 'code_mismatch':
        return '系統登 ' + nm(r.aCode) + '，紙本寫 ' + nm(r.bCode) + '。請確認實際服務內容，修正錯的一邊。';
      case 'client_mismatch':
        return '個案姓名不一致：系統「' + r.client + '」、紙本「' + r.bClient + '」，差一個字，疑似誤寫。確認後更正。';
      case 'worker_mismatch':
        return '服務員不一致：系統「' + r.worker + '」、紙本「' + r.bWorker + '」。若為代班，系統應登實際服務者。';
    }
    return '';
  }

  function summarize(rows) {
    const s = {
      total: rows.length, ok: 0, itemIssue: 0, sigIssue: 0, both: 0, attention: 0,
      byStatus: {}, overClaim: 0, underClaim: 0, amountUnknown: 0,
    };
    rows.forEach(r => {
      s.byStatus[r.status] = (s.byStatus[r.status] || 0) + 1;
      if (r.ok) s.ok++;
      if (!r.itemOk) s.itemIssue++;
      if (r.sigIssue) s.sigIssue++;
      if (!r.itemOk && r.sigIssue) s.both++;
      if (!r.itemOk || r.sigIssue) s.attention++;
      if (!r.itemOk) {
        if (r.amount == null) s.amountUnknown++;
        else if (r.amount > 0) s.overClaim += r.amount;
        else s.underClaim += -r.amount;
      }
    });
    return s;
  }

  return {
    BA_CODES, normCode, normDate, normName, parseCount, parseSignature,
    textToRows, splitCodes, parseRecords, reconcile, codeName, price,
  };
});
