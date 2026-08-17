(function () {
  "use strict";

  var PASS = 'MR@123';
  var SCHEMA = 'bond.key.v3';
  var PAGES = ['p1', 'p2', 'p3', 'p4', 'p5'];
  var unlocked = false;
  var LOG = [];
  var SRC = null;

  function $(s, r) { return (r || document).querySelector(s); }
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  function u16(d, p) { return d[p] | (d[p + 1] << 8); }
  function u32(d, p) { return (d[p] | (d[p + 1] << 8) | (d[p + 2] << 16) | (d[p + 3] << 24)) >>> 0; }

  function zipEntries(buf) {
    var d = new Uint8Array(buf), eocd = -1;
    for (var i = d.length - 22; i >= 0 && i > d.length - 66000; i--) {
      if (u32(d, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Không đọc được cấu trúc .xlsx (thiếu EOCD).');
    var n = u16(d, eocd + 10), off = u32(d, eocd + 16), out = {}, p = off;
    for (var k = 0; k < n; k++) {
      if (u32(d, p) !== 0x02014b50) throw new Error('Central directory hỏng ở entry ' + k + '.');
      var csize = u32(d, p + 20), nlen = u16(d, p + 28), elen = u16(d, p + 30),
          clen = u16(d, p + 32), method = u16(d, p + 10), lho = u32(d, p + 42),
          name = new TextDecoder().decode(d.subarray(p + 46, p + 46 + nlen));
      var lnlen = u16(d, lho + 26), lelen = u16(d, lho + 28), start = lho + 30 + lnlen + lelen;
      out[name] = { method: method, data: d.subarray(start, start + csize) };
      p += 46 + nlen + elen + clen;
    }
    return out;
  }

  function inflate(e) {
    if (e.method === 0) return Promise.resolve(e.data);
    if (typeof DecompressionStream === 'undefined') {
      return Promise.reject(new Error('Trình duyệt không hỗ trợ DecompressionStream. Cần Chrome/Edge đời mới.'));
    }
    var s = new Blob([e.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(s).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }
  function text(e) { return inflate(e).then(function (u) { return new TextDecoder('utf-8').decode(u); }); }
  function xml(s) { return new DOMParser().parseFromString(s, 'application/xml'); }

  function colOf(ref) {
    var m = /^([A-Z]+)/.exec(ref);
    if (!m) return 0;
    var c = 0, s = m[1];
    for (var i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
    return c;
  }

  function sheetRows(doc, shared) {
    var rows = [];
    doc.querySelectorAll('sheetData > row').forEach(function (r) {
      var arr = [];
      r.querySelectorAll('c').forEach(function (c) {
        var t = c.getAttribute('t'), v = c.querySelector('v'), val = null;
        if (t === 'inlineStr') { var is = c.querySelector('is'); val = is ? is.textContent : null; }
        else if (v) {
          if (t === 's') val = shared[parseInt(v.textContent, 10)];
          else if (t === 'str' || t === 'e') val = v.textContent;
          else { var f = parseFloat(v.textContent); val = isNaN(f) ? v.textContent : f; }
        }
        arr[colOf(c.getAttribute('r') || '') - 1] = (val === '' ? null : val);
      });
      rows.push(arr);
    });
    return rows;
  }

  function readWorkbook(buf) {
    var z = zipEntries(buf);
    ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels'].forEach(function (n) {
      if (!z[n]) throw new Error('Không phải file .xlsx hợp lệ (thiếu ' + n + ').');
    });
    var shared = [];
    var pShared = z['xl/sharedStrings.xml']
      ? text(z['xl/sharedStrings.xml']).then(function (s) {
          xml(s).querySelectorAll('si').forEach(function (si) {
            var o = '';
            si.querySelectorAll('t').forEach(function (t) { o += t.textContent; });
            shared.push(o);
          });
        })
      : Promise.resolve();
    return pShared
      .then(function () { return Promise.all([text(z['xl/workbook.xml']), text(z['xl/_rels/workbook.xml.rels'])]); })
      .then(function (r) {
        var wbx = xml(r[0]), rels = xml(r[1]), map = {};
        rels.querySelectorAll('Relationship').forEach(function (x) {
          map[x.getAttribute('Id')] = x.getAttribute('Target').replace(/^\/?xl\//, '');
        });
        var jobs = [], names = [];
        wbx.querySelectorAll('sheets > sheet').forEach(function (s) {
          var rid = s.getAttribute('r:id') ||
            s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
          if (!map[rid]) return;
          names.push(s.getAttribute('name'));
          jobs.push(text(z['xl/' + map[rid]]));
        });
        return Promise.all(jobs).then(function (docs) {
          var out = {};
          docs.forEach(function (d, i) { out[names[i]] = sheetRows(xml(d), shared); });
          return out;
        });
      });
  }

  function toMap(rows) {
    var head = (rows && rows[0]) || [], idx = {}, out = [];
    head.forEach(function (h, i) { if (h != null) idx[String(h).trim()] = i; });
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || row.every(function (v) { return v == null; })) continue;
      var o = {};
      Object.keys(idx).forEach(function (k) { o[k] = row[idx[k]]; });
      out.push(o);
    }
    return out;
  }

  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = typeof v === 'number' ? v : parseFloat(v);
    return isFinite(n) ? n : null;
  }

  function grid(rows, block) {
    return rows.filter(function (r) { return String(r.Block || '').trim() === block; });
  }

  function build(sheets) {
    var problems = [];
    ['META', 'DATA_S2', 'POS', 'CURVE', 'GRID', 'SCEN', 'TEXT', 'TS'].forEach(function (s) {
      if (!sheets[s]) problems.push('Thiếu sheet ' + s);
    });
    if (problems.length) throw new Error(problems.join(' · '));

    var meta = {};
    toMap(sheets.META).forEach(function (r) { if (r.Key != null) meta[String(r.Key).trim()] = r.Value; });
    if (String(meta.schema || '') !== SCHEMA) {
      problems.push('schema "' + (meta.schema || '(trống)') + '" khác "' + SCHEMA + '"');
    }

    var texts = {};
    toMap(sheets.TEXT).forEach(function (r) { if (r.KeyID != null) texts[String(r.KeyID).trim()] = r.Value; });

    var s2 = toMap(sheets.DATA_S2), byKey = {}, dup = 0;
    s2.forEach(function (r) {
      var k = String(r.KeyID || '').trim();
      if (!k) return;
      if (byKey[k]) dup++;
      byKey[k] = r;
    });
    if (dup) problems.push(dup + ' KeyID trùng trong DATA_S2');

    var ts = {}, tdup = 0, tseen = {};
    toMap(sheets.TS).forEach(function (r) {
      var s = String(r.Series || '').trim(), f = String(r.Field || '').trim(), d = String(r.Date || '').trim();
      if (!s || !f || !d) return;
      var sig = s + '|' + f + '|' + d;
      if (tseen[sig]) { tdup++; return; }
      tseen[sig] = 1;
      if (!ts[s]) ts[s] = { date: [], _f: {} };
      if (!ts[s]._f[f]) ts[s]._f[f] = {};
      ts[s]._f[f][d] = num(r.Value);
      if (ts[s].date.indexOf(d) < 0) ts[s].date.push(d);
    });
    if (tdup) problems.push(tdup + ' dòng TS trùng bộ ba Series+Field+Date');

    Object.keys(ts).forEach(function (s) {
      var o = ts[s], dates = o.date;
      Object.keys(o._f).forEach(function (f) {
        o[f] = dates.map(function (d) { return o._f[f][d]; });
      });
      delete o._f;
    });

    var pos = toMap(sheets.POS), curve = toMap(sheets.CURVE),
        g = toMap(sheets.GRID), scen = toMap(sheets.SCEN),
        vira = sheets.VIRA ? toMap(sheets.VIRA) : [],
        rating = sheets.RATING ? toMap(sheets.RATING) : [],
        vol = sheets.VOL ? toMap(sheets.VOL) : [],
        vira4 = sheets.VIRA4 ? toMap(sheets.VIRA4) : [];

    return {
      meta: meta, texts: texts, problems: problems,
      s2: s2, byKey: byKey, pos: pos, curve: curve, grid: g, scen: scen,
      vira: vira, vira4: vira4, rating: rating, vol: vol, ts: ts,
      counts: {
        'Khối 2': s2.length, 'Position': pos.length, 'Đường cong': curve.length,
        'Khối 3.x/7.x': g.length, 'Kịch bản': scen.length, 'VIRA': vira.length,
        'Xếp hạng': rating.length, 'Delta yield': vol.length,
        'Nhận định': Object.keys(texts).length, 'Điểm biểu đồ': Object.keys(ts).length + ' chuỗi'
      }
    };
  }

  function posRows(K, book) {
    return K.pos.filter(function (r) { return String(r.Book || '').trim() === book; })
      .map(function (r) {
        return {
          t: String(r.Tenor || '').replace('<=', '≤'),
          face: num(r.Face), faceYest: num(r.FaceYest), faceLM: num(r.FaceLM),
          pv01: num(r.PV01), itd: num(r.Itd), itdYest: num(r.ItdYest), itdLM: num(r.ItdLM),
          itdMtD: num(r.ItdMtD), itdYtD: num(r.ItdYtD), daily: num(r.Daily)
        };
      });
  }

  function markup(v) {
    if (v === null || v === undefined) return v;
    var s = String(v);
    if (!s) return s;
    s = s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    s = s.replace(/\[b\]/g, '<b>').replace(/\[\/b\]/g, '</b>');
    s = s.replace(/\[i\]/g, '<i>').replace(/\[\/i\]/g, '</i>');
    s = s.replace(/\[r\]/g, '<span style="color:#9B2C2C">').replace(/\[\/r\]/g, '</span>');
    s = s.replace(/\[h\]/g, '<mark style="background:#ffe27a;color:inherit">').replace(/\[\/h\]/g, '</mark>');
    return s;
  }

  function apply(K) {
    var R = window.RPT || {};
    var m = K.meta;

    R.asOf = m.asOf || R.asOf;
    R.dates = {
      today: m.asOf, yest: m.dateYest, lm: m.dateLastMonth,
      lq: m.dateLastQuarter, ly: m.dateLastYear
    };
    R.colDates = {
      today: String(m.asOf || '').slice(0, 5), yest: String(m.dateYest || '').slice(0, 5),
      lm: m.dateLastMonth, lq: m.dateLastQuarter, ly: m.dateLastYear
    };

    var t = {};
    Object.keys(K.texts).forEach(function (k) { t[k] = markup(K.texts[k]); });
    R.highlight = R.highlight || {};
    if (t['txt.assessment']) R.highlight.compliance = t['txt.assessment'];
    if (t['txt.itdTB']) R.highlight.tb = t['txt.itdTB'];
    if (t['txt.itdBB']) R.highlight.bb = t['txt.itdBB'];
    if (t['txt.noteFI']) R.highlight.fibond = t['txt.noteFI'];
    if (t['txt.note10d']) R.highlight.riskPrice = t['txt.note10d'];
    if (t['txt.noteRealized']) R.highlight.riskNim = t['txt.noteRealized'];
    if (t['txt.market']) R.reportMarket = t['txt.market'];
    if (t['txt.noteFV']) R.fvNote = t['txt.noteFV'];
    if (t['txt.noteRating']) R.fiNote = t['txt.noteRating'];
    if (t['txt.note10d']) R.note10d = t['txt.note10d'];
    if (t['txt.noteVar']) R.noteVar = t['txt.noteVar'];

    var BOOKMAP = {};
    K.s2.forEach(function (r) {
      var b = String(r.Book || '').trim();
      if (!b || BOOKMAP[b]) return;
      if (!/^(TB|BB)_/i.test(b)) return;
      var scope = /SBV/i.test(b) ? 'SBV' : 'MSB';
      BOOKMAP[b] = [scope, /^BB/i.test(b) ? 1 : 0];
    });
    R.books = R.books || { MSB: [], SBV: [] };
    Object.keys(BOOKMAP).forEach(function (bk) {
      var scope = BOOKMAP[bk][0], i = BOOKMAP[bk][1];
      var rows = K.s2.filter(function (r) { return String(r.Book || '').trim() === bk; });
      if (!rows.length || !R.books[scope] || !R.books[scope][i]) return;
      R.books[scope][i].rows = rows.map(function (r) {
        var o = {
          cat: r.Nhom || '', label: String(r.ChiTieu || '').replace(/\s+/g, ' ').trim(),
          today: num(r.Today), dtd: num(r.DtD), yest: num(r.Yesterday),
          lm: num(r.LastMonth), lq: num(r.LastQuarter), ly: num(r.LastYear)
        };
        if (String(r.Sub) === '1') o.sub = true;
        if (r.Limit != null && r.Limit !== '') o.limit = String(r.Limit);
        if (num(r.Used) !== null) o.used = num(r.Used);
        if (/YTM|Coupon rate|Repo rate|Concentration|Tỷ lệ|Cơ cấu/i.test(o.label)) o.pct = true;
        return o;
      });
    });

    var other = K.s2.filter(function (r) { return String(r.Book || '').trim() === 'OTHER'; });
    if (other.length && R.other) {
      R.other = other.map(function (r) {
        return {
          cat: 'Khác', label: String(r.ChiTieu || '').trim(), today: num(r.Today),
          dtd: num(r.DtD), yest: num(r.Yesterday), lm: num(r.LastMonth), lq: num(r.LastQuarter),
          limit: r.Limit ? String(r.Limit) : undefined, used: num(r.Used), pct: true
        };
      });
    }
    var fib = K.s2.filter(function (r) { return String(r.Book || '').trim() === 'FIBOND'; });
    if (fib.length && R.fibond) {
      R.fibond = fib.map(function (r) {
        var label = String(r.ChiTieu || '').replace(/\s+/g, ' ').trim();
        return {
          label: label, today: num(r.Today), dtd: num(r.DtD), yest: num(r.Yesterday),
          lm: num(r.LastMonth), lq: num(r.LastQuarter),
          limit: r.Limit ? String(r.Limit) : undefined, used: num(r.Used),
          pct: /Cơ cấu|Tỷ lệ/i.test(label)
        };
      });
    }

    var tb = posRows(K, 'TB'), bb = posRows(K, 'BB');
    if (tb.length) {
      R.posTrading = tb.filter(function (r) { return !/Tổng|Total/i.test(r.t); });
      var tt = tb.filter(function (r) { return /Tổng|Total/i.test(r.t); })[0];
      if (tt) { tt.stdPv01 = tt.pv01; R.posTradingTotal = tt; }
    }
    if (bb.length) {
      R.posBanking = bb.filter(function (r) { return !/Tổng|Total/i.test(r.t); });
      var bt = bb.filter(function (r) { return /Tổng|Total/i.test(r.t); })[0];
      if (bt) { bt.stdPv01 = bt.pv01; R.posBankingTotal = bt; }
    }

    var yc = K.curve.filter(function (r) { return String(r.Type || '') === 'YIELD'; });
    if (yc.length) {
      R.yieldCurve = yc.map(function (r) {
        return { t: String(r.Tenor || ''), bid: num(r.V1), ask: num(r.V2), mid: num(r.V3),
          spread: num(r.V4), dtd: num(r.DtD), mtd: num(r.MtD), ytd: num(r.YtD) };
      });
    }
    var rc = K.curve.filter(function (r) { return String(r.Type || '') === 'REPO'; });
    if (rc.length) {
      R.repoCurve = rc.map(function (r) {
        return { t: String(r.Tenor || ''), today: num(r.V1), yest: num(r.V2), lm: num(r.V3),
          dtd: num(r.DtD), mtd: num(r.MtD) };
      });
    }

    var mix = grid(K.grid, 'MIX');
    if (mix.length) {
      var body = mix.filter(function (r) { return !/Tổng|%/.test(String(r.Label)); });
      R.issuerMix = R.issuerMix || {};
      R.issuerMix.tenors = body.map(function (r) { return String(r.Label); });
      R.issuerMix.TB = {
        TPCP: body.map(function (r) { return num(r.C1); }),
        TPCPBL: body.map(function (r) { return num(r.C2); }),
        TPCQDP: body.map(function (r) { return num(r.C3); })
      };
      R.issuerMix.BB = {
        TPCP: body.map(function (r) { return num(r.C4); }),
        TPCPBL: body.map(function (r) { return num(r.C5); }),
        TPCQDP: body.map(function (r) { return num(r.C6); })
      };
      var tot = mix.filter(function (r) { return /Tổng/.test(String(r.Label)); })[0];
      if (tot) {
        R.issuerMix.TB.total = { TPCP: num(tot.C1), TPCPBL: num(tot.C2), TPCQDP: num(tot.C3) };
        R.issuerMix.BB.total = { TPCP: num(tot.C4), TPCPBL: num(tot.C5), TPCQDP: num(tot.C6) };
      }
    }

    var hold = grid(K.grid, 'HOLD');
    if (hold.length) {
      R.holdTime = hold.filter(function (r) { return !/Tổng/.test(String(r.Label)); })
        .map(function (r) {
          return { t: String(r.Label), tb: num(r.C1), tbItd: num(r.C2), bb: num(r.C3), bbItd: num(r.C4) };
        });
      var ht = hold.filter(function (r) { return /Tổng/.test(String(r.Label)); })[0];
      if (ht) R.holdTimeTotal = { tb: num(ht.C1), tbItd: num(ht.C2), bb: num(ht.C3), bbItd: num(ht.C4) };
    }

    var pnl = grid(K.grid, 'PNL');
    if (pnl.length) {
      R.pnlBreakdown = R.pnlBreakdown || {};
      R.pnlBreakdown.rows = pnl.filter(function (r) { return !/^Total/i.test(String(r.Label)); })
        .map(function (r) {
          var L = String(r.Label).trim();
          return { k: L, ytd: num(r.C1), mtd: num(r.C2), dtd: num(r.C3), sub: /^\d\.\d/.test(L) };
        });
      var pt = pnl.filter(function (r) { return /^Total/i.test(String(r.Label)); })[0];
      if (pt) R.pnlBreakdown.total = { ytd: num(pt.C1), mtd: num(pt.C2), dtd: num(pt.C3) };
    }

    var cap = grid(K.grid, 'CAPITAL');
    if (cap.length) {
      R.capital = R.capital || {};
      R.capital.rows = cap.filter(function (r) { return !/^Tổng/i.test(String(r.Label)); })
        .map(function (r) {
          var L = String(r.Label).trim();
          return { k: L, today: num(r.C1), yest: num(r.C2), dtd: num(r.C3), lm: num(r.C4),
            mtd: num(r.C5), sub: /^[ab]\./.test(L) };
        });
      var ct = cap.filter(function (r) { return /^Tổng/i.test(String(r.Label)); })[0];
      if (ct) R.capital.total = { k: String(ct.Label), today: num(ct.C1), yest: num(ct.C2),
        dtd: num(ct.C3), lm: num(ct.C4), mtd: num(ct.C5) };
    }

    var bs = grid(K.grid, 'BS');
    if (bs.length) {
      var pick = function (r) {
        return {
          on: { mtm: num(r.C1), face: num(r.C2), ptck: num(r.C3), diff: num(r.C4) },
          off: { mtm: num(r.C5), face: num(r.C6), ptck: num(r.C7), diff: num(r.C8) },
          out: { mtm: num(r.C9), face: num(r.C10), ptck: num(r.C11), diff: num(r.C12) }
        };
      };
      R.bsBook = R.bsBook || {};
      R.bsBook.rows = bs.filter(function (r) { return !/Total/i.test(String(r.Label)); })
        .map(function (r) { var o = pick(r); o.k = String(r.Label).trim(); return o; });
      var bt2 = bs.filter(function (r) { return /Total/i.test(String(r.Label)); })[0];
      if (bt2) R.bsBook.total = pick(bt2);
      var tbrow = R.bsBook.rows.filter(function (r) { return /Trading/i.test(r.k); })[0];
      if (tbrow) {
        R.bsLayers = { onBS: tbrow.on, offBS: tbrow.off, outBS: tbrow.out,
          netFace: (R.bsLayers && R.bsLayers.netFace) };
      }
    }

    var fio = grid(K.grid, 'FIONBS');
    if (fio.length) {
      R.fiOnBS = R.fiOnBS || {};
      R.fiOnBS.rows = fio.filter(function (r) { return /^\d\./.test(String(r.Label).trim()); })
        .map(function (r) {
          return { k: String(r.Label).trim(), on: num(r.C1), off: num(r.C2), out: num(r.C3), total: num(r.C4) };
        });
      fio.forEach(function (r) {
        var L = String(r.Label).trim(), o = { k: L, on: num(r.C1), off: num(r.C2), out: num(r.C3), total: num(r.C4) };
        if (/chưa điều chỉnh|Tổng chưa/i.test(L)) R.fiOnBS.gross = o;
        else if (/Giảm trừ/i.test(L)) R.fiOnBS.offset = o;
        else if (/sau giảm trừ/i.test(L)) R.fiOnBS.net = o;
      });
    }

    var ps = grid(K.grid, 'PNLSCEN');
    if (ps.length && R.pnlScenario) {
      R.pnlScenario.rows = ps.filter(function (r) { return !/^Tổng/i.test(String(r.Label)); })
        .map(function (r) {
          var L = String(r.Label).trim();
          return { k: L, a: [num(r.C1), num(r.C2), num(r.C3)], b: [num(r.C4), num(r.C5), num(r.C6)],
            sub: /^\d\.\d/.test(L) };
        });
      var pst = ps.filter(function (r) { return /^Tổng/i.test(String(r.Label)); })[0];
      if (pst) R.pnlScenario.total = { k: String(pst.Label), a: [num(pst.C1), num(pst.C2), num(pst.C3)],
        b: [num(pst.C4), num(pst.C5), num(pst.C6)] };
    }

    ['TB', 'BB'].forEach(function (bk) {
      var target = bk === 'TB' ? 'scenTB' : 'scenBB';
      var rows = K.scen.filter(function (r) { return String(r.Book || '') === bk; });
      if (!rows.length || !R[target]) return;
      var S = R[target];
      var tenors = [];
      rows.forEach(function (r) {
        var t = String(r.Tenor || '');
        if (!/Total|Tổng/i.test(t) && tenors.indexOf(t) < 0) tenors.push(t);
      });
      S.tenors = tenors;
      S.pv01 = tenors.map(function (t) {
        var x = rows.filter(function (r) { return String(r.Tenor) === t; })[0];
        return x ? num(x.PV01) : null;
      });
      S.itd = tenors.map(function (t) {
        var x = rows.filter(function (r) { return String(r.Tenor) === t; })[0];
        return x ? num(x.Itd) : null;
      });
      var totalRow = rows.filter(function (r) { return /Total|Tổng/i.test(String(r.Tenor)); })[0];
      if (totalRow) { S.pv01Total = num(totalRow.PV01); S.itdTotal = num(totalRow.Itd); }

      ['recent', 'var'].forEach(function (kind) {
        var kr = rows.filter(function (r) { return String(r.Kind || '') === kind; });
        var names = [];
        kr.forEach(function (r) {
          var nm = String(r.Scenario || '');
          if (nm && names.indexOf(nm) < 0) names.push(nm);
        });
        var list = names.map(function (nm) {
          var mine = kr.filter(function (r) { return String(r.Scenario) === nm; });
          var sub = mine.length ? String(mine[0].Sub || '') : '';
          var tot = mine.filter(function (r) { return /Total|Tổng/i.test(String(r.Tenor)); })[0];
          var o = {
            y: tenors.map(function (t) {
              var x = mine.filter(function (r) { return String(r.Tenor) === t; })[0];
              return x ? num(x.YieldBps) : null;
            }),
            p: tenors.map(function (t) {
              var x = mine.filter(function (r) { return String(r.Tenor) === t; })[0];
              return x ? num(x.ItdChange) : null;
            })
          };
          if (kind === 'recent') { o.name = nm; o.sub = sub; o.total = tot ? num(tot.ItdChange) : null; }
          else { o.h = nm; o.day = sub; o.loss = tot ? num(tot.ItdChange) : null; }
          return o;
        });
        if (kind === 'recent') S.recent = list; else S.varScen = list;
      });
    });

    if (K.vira.length && R.vira) {
      var vb = K.vira.filter(function (r) { return String(r.Scope || '') === 'BOOK'; });
      var names2 = [];
      vb.forEach(function (r) {
        var nm = String(r.Scenario || '');
        if (nm && names2.indexOf(nm) < 0) names2.push(nm);
      });
      if (names2.length) {
        R.vira.scenarios = names2.map(function (nm) {
          var x = vb.filter(function (r) { return String(r.Scenario) === nm; })[0];
          return { k: nm, bps: x ? num(x.YieldBps) : null };
        });
        var labels = [];
        vb.forEach(function (r) {
          var L = String(r.Label || '');
          if (L && labels.indexOf(L) < 0) labels.push(L);
        });
        R.vira.books = labels.map(function (L) {
          var mine = vb.filter(function (r) { return String(r.Label) === L; });
          return { k: L, pv01: num(mine[0].PV01), base: num(mine[0].Itd),
            v: names2.map(function (nm) {
              var x = mine.filter(function (r) { return String(r.Scenario) === nm; })[0];
              return x ? num(x.ItdAfter) : null;
            }) };
        });
      }
    }

    if (K.vira4 && K.vira4.length && R.vira4) {
      R.vira4 = {
        months: K.vira4.map(function (r) { return String(r.Month || ''); }),
        VIRA: K.vira4.map(function (r) { return num(r.VIRA); }),
        Big4: K.vira4.map(function (r) { return num(r.Big4); }),
        MarketMaker: K.vira4.map(function (r) { return num(r.MarketMaker); }),
        Top3: K.vira4.map(function (r) { return num(r.Top3); }),
        Actual: K.vira4.map(function (r) { return num(r.Actual); })
      };
    }

    if (K.rating.length) {
      var body2 = K.rating.filter(function (r) { return !/Total/i.test(String(r.Issuer)); });
      R.fiRating = body2.map(function (r) {
        var o = { issuer: String(r.Issuer), amt: num(r.Amount), rating: r.Rating || '',
          pct: num(r.Pct), cum: num(r.CumPct) };
        if (r.Fitch) o.fitch = String(r.Fitch);
        if (r.Moody) o.moody = String(r.Moody);
        if (r.SP) o.sp = String(r.SP);
        return o;
      });
      var rt2 = K.rating.filter(function (r) { return /Total/i.test(String(r.Issuer)); })[0];
      if (rt2) R.fiRatingTotal = num(rt2.Amount);
    }

    if (K.vol.length && R.volatility) {
      var tenorsV = Object.keys(K.vol[0]).filter(function (k) { return k !== 'Date'; });
      var series = tenorsV.map(function (t) {
        return K.vol.map(function (r) { return num(r[t]); }).filter(function (v) { return v !== null; });
      });
      var sma = function (a, w) {
        var s = a.slice(0, w).filter(function (v) { return v !== null; });
        if (!s.length) return null;
        var mean = s.reduce(function (x, y) { return x + y; }, 0) / s.length;
        var varr = s.reduce(function (x, y) { return x + (y - mean) * (y - mean); }, 0) / s.length;
        return Math.sqrt(varr);
      };
      var ewma = function (a, lam) {
        var v = null;
        for (var i = a.length - 1; i >= 0; i--) {
          var x = a[i] * a[i];
          v = (v === null) ? x : lam * v + (1 - lam) * x;
        }
        return v === null ? null : Math.sqrt(v);
      };
      var rnd = function (x, d) { return x === null ? null : Math.round(x * Math.pow(10, d)) / Math.pow(10, d); };
      var daily = [
        { k: 'Simple MA · 252 ngày', v: series.map(function (a) { return rnd(sma(a, 252), 3); }) },
        { k: 'Simple MA · 504 ngày', v: series.map(function (a) { return rnd(sma(a, 504), 3); }) },
        { k: 'Simple MA · 756 ngày', v: series.map(function (a) { return rnd(sma(a, 756), 3); }) },
        { k: 'EWMA · λ=0.98', v: series.map(function (a) { return rnd(ewma(a, 0.98), 3); }) }
      ];
      R.volatility = {
        tenors: tenorsV,
        daily: daily,
        annual: daily.map(function (d) {
          return { k: d.k, v: d.v.map(function (x) { return rnd(x === null ? null : x * Math.sqrt(252), 2); }) };
        })
      };
    }

    if (Object.keys(K.ts).length) {
      var TS = window.TS || {};
      var MAP = {
        tdPv01: 'tdpv01', tdPnl: 'tdpnl', bkPnl: 'bkpnl', bbPv01: 'bkpv01',
        klgdThuCap: 'klgd', yield: 'yieldTs', lsRepo: 'repoTs', lsttSoCap: 'lsttPrimary',
        bidAskSpread: 'spreadTs', tuongQuanLs10y: 'corr', gdTpdn: 'gdTPDN',
        posiFibondCd: 'fiPos', couponBqFibondCd: 'fiCoupon', yieldFiBond: 'fiYield',
        mucDoSuDungVon: 'vonSD', qlhsFibondCd: 'qlhsFi', qlhsGovbond: 'qlhsGov',
        tuongQuanYtmVsRepo: 'ytmRepo'
      };
      Object.keys(K.ts).forEach(function (s) {
        var target = MAP[s];
        if (!target || !TS[target]) return;
        var src = K.ts[s], dst = {};
        Object.keys(src).forEach(function (f) { dst[f] = src[f]; });
        dst.date = src.date.map(function (d) { return String(d).slice(0, 5); });
        Object.keys(TS[target]).forEach(function (f) {
          if (dst[f] === undefined) dst[f] = TS[target][f];
        });
        TS[target] = dst;
      });
      window.TS = TS;
    }

    window.RPT = R;
  }

  function ageDays(d) {
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(d || ''));
    if (!m) return null;
    return Math.floor((new Date().setHours(0, 0, 0, 0) - new Date(+m[3], +m[2] - 1, +m[1]).getTime()) / 86400000);
  }

  function badge() {
    document.querySelectorAll('[data-asof-badge]').forEach(function (b) { b.remove(); });
  }

  function lockNav(btn) {
    if (btn.getAttribute('data-locked')) return;
    btn.setAttribute('data-locked', '1');
    btn.addEventListener('click', function (e) {
      if (unlocked) return;
      e.stopImmediatePropagation();
      e.preventDefault();
      var v = window.prompt('Nhap mat khau de mo tab Quan tri du lieu:');
      if (v === null) return;
      if (v !== PASS) { window.alert('Mat khau khong dung.'); return; }
      unlocked = true;
      if (window.RRTT && window.RRTT.goPage) window.RRTT.goPage('p6');
    }, true);
  }

  function ui() {
    var nav = document.querySelector('[data-nav="p5"]');
    if (!nav) return false;

    var existing = document.querySelector('[data-nav="p6"]');
    if (existing) {
      var sec0 = document.querySelector('[data-page="p6"]');
      if (!sec0) return false;
      if (!document.getElementById('rrttFile')) sec0.innerHTML = panelHTML();
      lockNav(existing);
      wire();
      return true;
    }

    var btn = nav.cloneNode(true);
    btn.setAttribute('data-nav', 'p6');
    btn.removeAttribute('data-active');
    var lbl = btn.querySelector('[data-rail-text]');
    if (lbl) { lbl.textContent = 'Quản trị dữ liệu'; lbl.removeAttribute('data-k'); }
    nav.parentNode.appendChild(btn);

    var host = document.querySelector('[data-page="p5"]');
    if (!host) return false;
    var sec = document.createElement('section');
    sec.setAttribute('data-page', 'p6');
    sec.setAttribute('data-screen-label', 'Quản trị dữ liệu');
    sec.style.cssText = 'display:grid;gap:18px;';
    sec.innerHTML = panelHTML();
    host.parentNode.appendChild(sec);

    lockNav(btn);
    wire();
    return true;
  }


  var EDITS = {}, ORIG = {}, editMode = false, skipped = 0, TMP = null;

  var INLINE = { B: 1, I: 1, EM: 1, STRONG: 1, SPAN: 1, MARK: 1, BR: 1, SMALL: 1,
    SUP: 1, SUB: 1, A: 1, U: 1, FONT: 1 };

  function isLeafText(el) {
    if (!(el.textContent || '').trim()) return false;
    if (el.querySelector('table,canvas,svg,input,button,select')) return false;
    for (var i = 0; i < el.children.length; i++) {
      var c = el.children[i];
      if (!INLINE[c.tagName]) return false;
    }
    return true;
  }

  function editable(root) {
    var out = [];
    (root || document).querySelectorAll('[data-page]').forEach(function (pg) {
      if (!/^p[1-5]$/.test(pg.dataset.page || '')) return;
      pg.querySelectorAll('td,th,p,li,h1,h2,h3,h4,h5,div,span').forEach(function (el) {
        if (el.hasAttribute('data-asof-badge')) return;
        if (el.closest('[data-asof-badge]')) return;
        if (isLeafText(el)) out.push(el);
      });
    });
    return out;
  }

  function hash(str) {
    var h = 5381, i = str.length;
    while (i) h = (h * 33 ^ str.charCodeAt(--i)) >>> 0;
    return h.toString(36);
  }

  function tagEditables() {
    var seen = {};
    editable().forEach(function (el) {
      var host = el.closest('[data-page]');
      if (!host) return;
      var pg = host.dataset.page;
      var txt = (el.textContent || '').trim();
      var prev = el.getAttribute('data-eid');
      var base = (prev && EDITS[prev]) ? prev.split('#')[0] : pg + ':' + hash(txt);
      seen[base] = (seen[base] || 0) + 1;
      var id = base + '#' + seen[base];
      el.setAttribute('data-eid', id);
      if (!EDITS[id]) ORIG[id] = txt;
    });
  }

  function applyEdits() {
    tagEditables();
    skipped = 0;
    Object.keys(EDITS).forEach(function (id) {
      var e = EDITS[id];
      var el = document.querySelector('[data-eid="' + id + '"]');
      if (!el) { skipped++; return; }
      var now = (el.textContent || '').trim();
      if (TMP == null) TMP = document.createElement('div');
      TMP.innerHTML = e.h;
      var edited = (TMP.textContent || '').trim();
      var ok = (e.o == null) || now === e.o || now === edited || el.innerHTML === e.h;
      if (!ok) { skipped++; return; }
      if (el.innerHTML !== e.h) el.innerHTML = e.h;
    });
    var n = document.getElementById('rrttEditCount');
    if (n) n.textContent = String(Object.keys(EDITS).length) + (skipped ? ' (' + skipped + ' khong ap duoc)' : '');
    if (editMode) setEditable(true);
  }

  function setEditable(on) {
    editable().forEach(function (el) {
      if (on) {
        el.setAttribute('contenteditable', 'true');
        el.style.outline = '1px dashed rgba(123,45,59,.35)';
        el.style.outlineOffset = '1px';
      } else {
        el.removeAttribute('contenteditable');
        el.style.outline = '';
        el.style.outlineOffset = '';
      }
    });
  }

  function captureEdit(node) {
    var el = node && node.nodeType === 1 ? node : (node && node.parentElement);
    var host = el && el.closest ? el.closest('[data-eid]') : null;
    if (!host) return;
    var id = host.getAttribute('data-eid');
    EDITS[id] = { h: host.innerHTML, o: ORIG[id] };
    var n = document.getElementById('rrttEditCount');
    if (n) n.textContent = String(Object.keys(EDITS).length);
  }

  function exec(cmd, val) {
    try { document.execCommand('styleWithCSS', false, true); } catch (e) {}
    document.execCommand(cmd, false, val === undefined ? null : val);
    var sel = window.getSelection();
    if (sel && sel.anchorNode) captureEdit(sel.anchorNode);
  }

  function toggleEdit(on) {
    editMode = on;
    var bar = ensureFloatBar();
    setEditable(on);
    bar.style.display = on ? 'flex' : 'none';
    var hint = document.getElementById('rrttEditHint');
    if (hint) hint.style.display = on ? 'block' : 'none';
    var btn = document.getElementById('rrttEditBtn');
    if (btn) {
      btn.textContent = on ? 'Tat che do sua' : 'Sua bao cao';
      btn.style.background = on ? '#5E1F2B' : '#fdfcfa';
      btn.style.color = on ? '#fff' : '#7B2D3B';
    }
    if (on && window.RRTT && window.RRTT.goPage) window.RRTT.goPage('p1');
  }

  document.addEventListener('input', function (e) {
    if (editMode) captureEdit(e.target);
  }, true);

  document.addEventListener('keydown', function (e) {
    if (!editMode || !(e.ctrlKey || e.metaKey)) return;
    var k = String(e.key || '').toLowerCase();
    if (k === 'b') { e.preventDefault(); exec('bold'); }
    else if (k === 'i') { e.preventDefault(); exec('italic'); }
    else if (k === 'z' && !e.shiftKey) { e.preventDefault(); exec('undo'); }
    else if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); exec('redo'); }
  }, true);

  var EBTN = 'font-family:inherit;font-size:12.5px;font-weight:600;border:1px solid #e6e2db;' +
    'background:#fff;color:#33312e;border-radius:8px;padding:7px 12px;cursor:pointer;';

  function editBarHTML() { return ''; }

  function ensureFloatBar() {
    var bar = document.getElementById('rrttEditBar');
    if (bar) return bar;
    bar = document.createElement('div');
    bar.id = 'rrttEditBar';
    bar.style.cssText = 'position:fixed;left:50%;transform:translateX(-50%);bottom:18px;z-index:9999;' +
      'display:none;gap:7px;flex-wrap:wrap;align-items:center;padding:9px 12px;background:#fdfcfa;' +
      'border:1px solid #e6e2db;border-radius:12px;box-shadow:0 6px 24px rgba(32,31,29,.18);' +
      'font-family:inherit;';
    bar.innerHTML =
      '<span style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#8d877c;margin-right:2px;">Sua</span>' +
      '<button type="button" data-cmd="bold" style="' + EBTN + 'font-weight:800;min-width:34px;">B</button>' +
      '<button type="button" data-cmd="italic" style="' + EBTN + 'font-style:italic;min-width:34px;">I</button>' +
      '<button type="button" data-cmd="hilite" style="' + EBTN + 'background:#ffe27a;">Boi vang</button>' +
      '<button type="button" data-cmd="hiliteOff" style="' + EBTN + '">Bo boi</button>' +
      '<button type="button" data-cmd="black" style="' + EBTN + 'color:#201f1d;">Chu den</button>' +
      '<button type="button" data-cmd="red" style="' + EBTN + 'color:#9B2C2C;">Chu do</button>' +
      '<span style="width:1px;height:22px;background:#e6e2db;margin:0 3px;"></span>' +
      '<button type="button" data-cmd="undo" title="Ctrl+Z" style="' + EBTN + 'min-width:34px;">&#8630;</button>' +
      '<button type="button" data-cmd="redo" title="Ctrl+Y" style="' + EBTN + 'min-width:34px;">&#8631;</button>' +
      '<span style="width:1px;height:22px;background:#e6e2db;margin:0 3px;"></span>' +
      '<span style="font-size:11.5px;color:#6b665e;white-space:nowrap;">Da sua <b id="rrttEditCount">0</b> o</span>' +
      '<button type="button" data-cmd="reset" style="' + EBTN + 'color:#9B2C2C;">Hoan tac tat ca</button>' +
      '<button type="button" data-cmd="off" style="' + EBTN + 'background:#7B2D3B;color:#fff;border-color:#7B2D3B;">Tat sua</button>';
    document.body.appendChild(bar);
    bindBar(bar);
    return bar;
  }

  function bindBar(bar) {
    bar.addEventListener('mousedown', function (e) {
      if (e.target && e.target.tagName === 'BUTTON') e.preventDefault();
    });
    bar.addEventListener('click', function (e) {
      var b = e.target.closest ? e.target.closest('button') : null;
      if (!b) return;
      var c = b.getAttribute('data-cmd');
      if (c === 'bold') exec('bold');
      else if (c === 'italic') exec('italic');
      else if (c === 'hilite') exec('hiliteColor', '#ffe27a');
      else if (c === 'hiliteOff') exec('hiliteColor', 'transparent');
      else if (c === 'black') exec('foreColor', '#201f1d');
      else if (c === 'red') exec('foreColor', '#9B2C2C');
      else if (c === 'undo') exec('undo');
      else if (c === 'redo') exec('redo');
      else if (c === 'off') toggleEdit(false);
      else if (c === 'reset') {
        if (!window.confirm('Bo toan bo sua doi thu cong, tra ve dung so lieu tu file key?')) return;
        EDITS = {}; ORIG = {};
        var n = document.getElementById('rrttEditCount');
        if (n) n.textContent = '0';
        if (window.RRTT && window.RRTT.rerender) window.RRTT.rerender();
        setTimeout(function () { if (editMode) setEditable(true); }, 60);
      }
    });
  }

  function wireEdit() {
    var btn = document.getElementById('rrttEditBtn');
    if (!btn || btn.getAttribute('data-wired')) return;
    btn.setAttribute('data-wired', '1');
    btn.addEventListener('click', function () { toggleEdit(!editMode); });
  }

  var CARD = 'background:#fdfcfa;border:1px solid #e6e2db;border-radius:12px;overflow:hidden;';
  var HEAD = 'background:#faf7f4;border-bottom:1px solid #e6e2db;padding:11px 20px;font-size:14.5px;' +
    'font-weight:600;color:#7B2D3B;letter-spacing:.02em;';
  var BODY = 'padding:16px 20px;';

  function panelHTML() {
    return '' +
      '<div style="' + CARD + '"><div style="' + HEAD + '">Nạp dữ liệu kỳ mới</div><div style="' + BODY + '">' +
        '<input type="file" id="rrttFile" accept=".xlsx,.xlsm" style="position:absolute;width:1px;height:1px;opacity:0;clip:rect(0 0 0 0)">' +
        '<label for="rrttFile" id="rrttDrop" style="display:block;border:2px dashed #e6e2db;border-radius:12px;' +
          'padding:30px 20px;text-align:center;background:#faf8f6;cursor:pointer;transition:.15s;">' +
          '<div style="font-size:14.5px;font-weight:600;color:#7B2D3B;">Bấm vào đây để chọn file Key, hoặc kéo thả vào ô này</div>' +
          '<div style="font-size:12px;color:#6b665e;margin-top:5px;">Key_YYYYMMDD.xlsx — sinh từ File 02 bằng macro XuatFileKey</div>' +
        '</label>' +
        '<div id="rrttStatus" style="margin-top:12px;border-radius:10px;padding:12px 16px;font-size:13px;' +
          'border:1px solid #e6d3a8;background:#fff8e8;color:#6b5320;">Chưa nạp file nào trong phiên này. Báo cáo đang hiển thị dữ liệu nhúng sẵn.</div>' +
        '<div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;">' +
          '<button type="button" id="rrttEditBtn" style="font-family:inherit;font-size:13px;font-weight:600;border:1px solid #7B2D3B;background:#fdfcfa;color:#7B2D3B;border-radius:9px;padding:10px 18px;cursor:pointer;">Sua bao cao</button>' +
          '<button type="button" id="rrttExport" style="font-family:inherit;font-size:13px;font-weight:600;' +
            'border:1px solid #7B2D3B;background:#7B2D3B;color:#fff;border-radius:9px;padding:10px 18px;cursor:pointer;">Xuất bản gửi đi</button>' +
        '</div>' +
        editBarHTML() +
        '<p style="font-size:12.5px;color:#6b665e;margin-top:10px;line-height:1.7;">Toàn bộ xử lý chạy trong trình duyệt trên máy này. Không có kết nối mạng nào được tạo ra, không dữ liệu nào rời khỏi máy.</p>' +
      '</div></div>' +
      '<div style="' + CARD + '"><div style="' + HEAD + '">Trạng thái dữ liệu</div>' +
        '<div style="' + BODY + '" id="rrttStat"></div></div>' +
      '<div style="' + CARD + '"><div style="' + HEAD + '">Kiểm tra trước khi gửi</div>' +
        '<div style="' + BODY + '" id="rrttChecks"></div></div>' +
      '<div style="' + CARD + '"><div style="' + HEAD + '">Nhật ký nạp file</div>' +
        '<div style="' + BODY + '" id="rrttLog"></div></div>' +
      '<div style="' + CARD + '"><div style="' + HEAD + '">Môi trường</div>' +
        '<div style="' + BODY + '" id="rrttEnv"></div></div>';
  }

  function setStatus(kind, title, sub) {
    var el = $('#rrttStatus');
    if (!el) return;
    var C = { ok: ['#cfe3d4', '#f2f8f3', '#255c3a'], warn: ['#e6d3a8', '#fff8e8', '#6b5320'],
      err: ['#e8c9c9', '#fdf2f2', '#9B2C2C'] }[kind] || ['#e6d3a8', '#fff8e8', '#6b5320'];
    el.style.borderColor = C[0]; el.style.background = C[1]; el.style.color = C[2];
    el.innerHTML = '<b>' + esc(title) + '</b>' +
      (sub ? '<span style="display:block;font-size:11.5px;opacity:.85;margin-top:3px;">' + esc(sub) + '</span>' : '');
  }

  function renderPanel() {
    var st = $('#rrttStat'), ck = $('#rrttChecks'), lg = $('#rrttLog'), ev = $('#rrttEnv');
    if (!st) return;
    var box = 'display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;';
    var cell = function (k, v, color) {
      return '<div style="background:#faf8f6;border:1px solid #e6e2db;border-radius:9px;padding:11px 14px;">' +
        '<div style="font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#8d877c;">' + esc(k) + '</div>' +
        '<div style="font-size:15px;font-weight:600;margin-top:3px;' + (color ? 'color:' + color : '') + '">' + esc(v) + '</div></div>';
    };

    var d = (window.RPT && window.RPT.asOf) || '—', age = ageDays(d);
    st.innerHTML = '<div style="' + box + '">' +
      cell('Ngày chốt dữ liệu', d + ' · ' + (age === null ? 'khong ro' :
        (age <= 0 ? 'du lieu hom nay' : 'da cu ' + age + ' ngay')),
        age === null || age > 3 ? '#a56a1b' : '#276749') +
      cell('Nguồn', SRC ? SRC.file : 'dữ liệu nhúng sẵn') +
      cell('Sinh lúc', SRC ? (SRC.generatedAt || '—') : '—') +
      '</div>' +
      (SRC ? '<div style="' + box + 'margin-top:12px;">' + Object.keys(SRC.counts).map(function (k) {
        return cell(k, String(SRC.counts[k]));
      }).join('') + '</div>' : '');

    var items = [];
    if (age === null) items.push(['no', 'Không đọc được ngày chốt', String(d)]);
    else if (age > 3) items.push(['wn', 'Dữ liệu đã cũ ' + age + ' ngày', 'Ngày chốt ' + d]);
    else items.push(['ok', 'Ngày chốt ' + d, age <= 0 ? 'Dữ liệu hôm nay' : 'Cách hôm nay ' + age + ' ngày']);
    (SRC ? SRC.problems : []).forEach(function (p) { items.push(['no', 'Cấu trúc file', p]); });

    var R = window.RPT || {};
    [['Trading nội bộ · Face = AFS + HTM', 'MSB', 0, 'Face Value', ['Book AFS', 'Book HTM']],
     ['Banking nội bộ · Face = DCM + Pool', 'MSB', 1, 'Face Value', ['DCM tự fund', 'fund từ Pool']]]
      .forEach(function (spec) {
        var bk = R.books && R.books[spec[1]] && R.books[spec[1]][spec[2]];
        if (!bk || !bk.rows) return;
        var tot = bk.rows.filter(function (r) { return r.label === spec[3] && !r.sub; })[0];
        var parts = spec[4].map(function (nm) {
          return bk.rows.filter(function (r) { return r.sub && r.label.indexOf(nm) >= 0; })[0];
        });
        if (!tot || parts.some(function (p) { return !p; })) return;
        var a = tot.today, b = parts.reduce(function (s, p) { return s + (p.today || 0); }, 0);
        if (a === null) return;
        var diff = a - b;
        items.push([Math.abs(diff) < 0.5 ? 'ok' : 'no', spec[0],
          fm(a) + ' vs ' + fm(b) + (Math.abs(diff) < 0.5 ? '' : ' · lệch ' + fm(diff))]);
      });

    var breach = [];
    ['MSB', 'SBV'].forEach(function (sc) {
      ((R.books && R.books[sc]) || []).forEach(function (bk) {
        (bk.rows || []).forEach(function (r) {
          if (typeof r.used === 'number' && r.used >= 0.8) breach.push([bk.title || sc, r.label, r.used]);
        });
      });
    });
    (R.fibond || []).forEach(function (r) {
      if (typeof r.used === 'number' && r.used >= 0.8) breach.push(['FI Bond & CD', r.label, r.used]);
    });
    breach.sort(function (x, y) { return y[2] - x[2]; });
    if (breach.length) {
      breach.forEach(function (x) {
        items.push([x[2] >= 1 ? 'no' : 'wn', (x[2] >= 1 ? 'Vượt hạn mức: ' : 'Sát hạn mức: ') + x[1],
          x[0] + ' · ' + (x[2] * 100).toFixed(1) + '%']);
      });
    } else {
      items.push(['ok', 'Không chỉ tiêu nào chạm ngưỡng cảnh báo', 'XANH < 80% · HỔ PHÁCH 80–100% · ĐỎ ≥ 100%']);
    }

    ck.innerHTML = '<ul style="list-style:none;font-size:13px;margin:0;padding:0;">' + items.map(function (it) {
      var col = it[0] === 'ok' ? '#276749' : it[0] === 'wn' ? '#a56a1b' : '#9B2C2C';
      var ic = it[0] === 'ok' ? '✓' : it[0] === 'wn' ? '!' : '✕';
      return '<li style="padding:8px 0;border-bottom:1px solid rgba(32,31,29,.06);display:flex;gap:10px;">' +
        '<span style="color:' + col + ';font-weight:700;flex:none;width:14px;">' + ic + '</span>' +
        '<span>' + esc(it[1]) + '<span style="display:block;color:#6b665e;font-size:12px;font-variant-numeric:tabular-nums;">' +
        esc(it[2]) + '</span></span></li>';
    }).join('') + '</ul>';

    lg.innerHTML = LOG.length
      ? '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr>' +
        ['Lúc', 'File', 'Ngày chốt', 'Cảnh báo'].map(function (h) {
          return '<th style="text-align:left;font-size:10px;letter-spacing:.08em;text-transform:uppercase;' +
            'color:#8d877c;padding:6px 8px;border-bottom:1px solid #e6e2db;">' + h + '</th>';
        }).join('') + '</tr></thead><tbody>' +
        LOG.slice().reverse().map(function (r) {
          return '<tr>' + [r.at, r.file, r.asOf, String(r.w)].map(function (c) {
            return '<td style="padding:6px 8px;border-bottom:1px solid rgba(32,31,29,.06);">' + esc(c) + '</td>';
          }).join('') + '</tr>';
        }).join('') + '</tbody></table>'
      : '<p style="color:#8d877c;font-size:12.5px;font-style:italic;">Chưa có lần nạp nào trong file này.</p>';

    var deco = typeof DecompressionStream !== 'undefined';
    ev.innerHTML = '<div style="' + box + '">' +
      cell('Đọc file .xlsx', deco ? 'Hỗ trợ' : 'KHÔNG hỗ trợ', deco ? '#276749' : '#9B2C2C') +
      cell('Trình duyệt', /Edg\//.test(navigator.userAgent) ? 'Edge' :
        /Chrome\//.test(navigator.userAgent) ? 'Chrome' : 'Khác') +
      cell('Hôm nay', new Date().toLocaleDateString('en-GB')) + '</div>';
  }

  function fm(v) {
    if (v === null || v === undefined) return '—';
    var a = Math.abs(v), dp = a >= 1000 ? 0 : (a >= 10 ? 1 : 2);
    return v.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
  }

  function countIssues() {
    var n = 0, el = $('#rrttChecks');
    if (!el) return 0;
    el.querySelectorAll('li').forEach(function (li) {
      var t = li.firstChild && li.firstChild.textContent;
      if (t === '✕' || t === '!') n++;
    });
    return n;
  }

  function viaReader(f) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () { res(fr.result); };
      fr.onerror = function () { rej(fr.error || new Error('FileReader thất bại')); };
      fr.readAsArrayBuffer(f);
    });
  }

  function describe(f) {
    var d = f.lastModified ? new Date(f.lastModified).toLocaleString('en-GB', { hour12: false }) : '?';
    return f.name + ' · ' + (f.size != null ? f.size.toLocaleString('en-US') + ' byte' : 'không rõ cỡ') +
      ' · sửa lần cuối ' + d;
  }

  function readBytes(f) {
    var tried = [];
    var step = function (label, fn) {
      return function () {
        return Promise.resolve().then(fn).then(function (buf) {
          if (!buf || !buf.byteLength) throw new Error('đọc ra 0 byte');
          return buf;
        }).catch(function (e) {
          tried.push(label + ': ' + ((e && e.name ? e.name + ' — ' : '') + ((e && e.message) || e)));
          throw e;
        });
      };
    };
    var a = step('arrayBuffer', function () {
      if (!f.arrayBuffer) throw new Error('trình duyệt không có File.arrayBuffer');
      return f.arrayBuffer();
    });
    var b = step('FileReader', function () { return viaReader(f); });
    var c = step('slice+arrayBuffer', function () {
      var s = f.slice(0, f.size);
      if (s.arrayBuffer) return s.arrayBuffer();
      return viaReader(s);
    });

    return a().catch(b).catch(c).catch(function () {
      var hint = (f.size === 0)
        ? 'File có kích thước 0 byte — bản tải về bị hỏng hoặc chưa tải xong. Tải lại file.'
        : 'Thường do: file đang mở trong Excel · file nằm trên OneDrive/ổ mạng chưa tải hẳn về máy · ' +
          'file đã bị di chuyển/xoá sau khi chọn · phần mềm bảo mật khoá file.';
      throw new Error(hint + ' | ' + describe(f) + ' | Chi tiết: ' + tried.join(' ; '));
    });
  }

  function handleFile(f, done) {
    done = done || function () {};
    if (!f) { done(); return; }
    if (!/\.xls[xm]$/i.test(f.name)) {
      setStatus('err', 'Sai định dạng: ' + f.name, 'Cần file .xlsx sinh từ macro XuatFileKey.');
      done();
      return;
    }
    setStatus('warn', 'Đang đọc ' + f.name + '…', '');
    readBytes(f).then(readWorkbook).then(function (sheets) {
      var K = build(sheets);
      apply(K);
      SRC = { file: f.name, generatedAt: K.meta.generatedAt, counts: K.counts, problems: K.problems };
      if (window.RRTT && window.RRTT.rerender) window.RRTT.rerender();
      badge();
      applyEdits();
      renderPanel();
      var n = countIssues();
      LOG.push({ at: new Date().toLocaleString('en-GB', { hour12: false }).replace(',', ''),
        file: f.name, asOf: String(K.meta.asOf || '—'), w: n });
      renderPanel();
      if (n) setStatus('warn', 'Đã nạp — ' + n + ' mục cần xem', 'Ngày chốt ' + (K.meta.asOf || 'KHÔNG RÕ'));
      else setStatus('ok', 'Đã nạp dữ liệu chốt ngày ' + (K.meta.asOf || 'KHÔNG RÕ'), 'nguồn ' + f.name);
      done();
    }).catch(function (e) {
      setStatus('err', 'Không nạp được — dữ liệu cũ giữ nguyên', e && e.message ? e.message : String(e));
      done();
    });
  }

  function wire() {
    var inp = $('#rrttFile'), drop = $('#rrttDrop');
    if (!inp) return;
    inp.addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      handleFile(f, function () { try { inp.value = ''; } catch (x) {} });
    });
    ['dragenter', 'dragover'].forEach(function (t) {
      drop.addEventListener(t, function (e) { e.preventDefault(); drop.style.borderColor = '#7B2D3B'; });
    });
    ['dragleave', 'drop'].forEach(function (t) {
      drop.addEventListener(t, function (e) { e.preventDefault(); drop.style.borderColor = '#e6e2db'; });
    });
    drop.addEventListener('drop', function (e) {
      handleFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); });

    $('#rrttExport').addEventListener('click', function () {
      if (window.RRTT && window.RRTT.goPage) window.RRTT.goPage('p1');
      var clone = document.documentElement.cloneNode(true);
      var inject = document.createElement('script');
      inject.textContent = 'window.RPT = ' + JSON.stringify(window.RPT).replace(/</g, '\\u003c') + ';\n' +
        'window.TS = ' + JSON.stringify(window.TS).replace(/</g, '\\u003c') + ';\n' +
        'window.RRTT_EDITS = ' + JSON.stringify(EDITS).replace(/</g, '\\u003c') + ';';
      clone.querySelector('head').appendChild(inject);
      var blob = new Blob(['<!DOCTYPE html>\n' + clone.outerHTML], { type: 'text/html;charset=utf-8' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'BaoCao_RRTT_Bond_' + String((window.RPT && window.RPT.asOf) || '').replace(/\//g, '') + '.html';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    });

    wireEdit();
    renderPanel();
    badge();
    applyEdits();
    [300, 1200, 2500].forEach(function (ms) { setTimeout(applyEdits, ms); });
    watchDom();
  }

  setInterval(function () {
    if (unlocked) return;
    var sec = document.querySelector('[data-page="p6"][data-active]');
    if (!sec) return;
    var v = window.prompt('Nhap mat khau de mo tab Quan tri du lieu:');
    if (v === PASS) { unlocked = true; return; }
    if (v !== null) window.alert('Mat khau khong dung.');
    if (window.RRTT && window.RRTT.goPage) window.RRTT.goPage('p1');
  }, 400);

  var moTimer = null, mo = null;
  function watchDom() {
    if (mo || !window.MutationObserver) return;
    var host = document.body;
    if (!host) return;
    mo = new MutationObserver(function () {
      if (moTimer) clearTimeout(moTimer);
      moTimer = setTimeout(function () {
        if (Object.keys(EDITS).length) applyEdits(); else tagEditables();
      }, 120);
    });
    mo.observe(host, { childList: true, subtree: true });
  }

  if (window.RRTT_EDITS) EDITS = window.RRTT_EDITS;

  var tries = 0;
  (function boot() {
    if (ui()) return;
    if (++tries > 60) return;
    setTimeout(boot, 100);
  })();

  window.RRTT_ADMIN = { apply: apply, edits: function () { return EDITS; }, applyEdits: applyEdits, log: function () { return LOG; }, source: function () { return SRC; } };
})();
