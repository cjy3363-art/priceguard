/*
 * clean.js — 정제 규칙(JSON) 실행 엔진
 *
 * 규칙의 조건·순서·패턴·임계값·사유 문구는 모두 rules/*.json 에서 읽는다.
 * 이 파일에는 "조건 유형"별 실행기만 있다. (예: parse_price, column_compare)
 * 브라우저(window.PriceCleaner)와 Node(module.exports)에서 같은 코드로 동작한다.
 * 네트워크 요청을 하지 않는다.
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PriceCleaner = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- 셀 값 모델 ----------
   * { t: 'blank' | 'text' | 'number' | 'date' | 'bool', v }
   * date 는 엑셀 날짜값 셀이며 v 에 'YYYY-MM-DD' 를 담는다.
   */
  var BLANK = { t: 'blank', v: null };

  /** 정리기가 정리본에 덧붙이는 컬럼 이름 (재업로드 시 원본 행 번호를 이어받는 데 쓴다) */
  var ID_COLS = { row: '_원본_엑셀행번호', seq: '_데이터순번' };
  var OUTPUT_COLS = ['_검토상태', '_검토사유', ID_COLS.row, ID_COLS.seq, '_유지된_행(원본_엑셀행번호)'];
  function positiveInt(c) {
    if (!c) return null;
    if (c.t === 'number' && c.v > 0 && Math.floor(c.v) === c.v) return c.v;
    if (c.t === 'text' && /^\s*\d+\s*$/.test(c.v) && parseInt(c.v, 10) > 0) return parseInt(c.v, 10);
    return null;
  }

  function pad(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }
  function ymd(y, m, d) { return pad(y, 4) + '-' + pad(m, 2) + '-' + pad(d, 2); }
  function validDate(y, m, d) {
    if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }
  function isBlank(c) { return !c || c.t === 'blank' || (c.t === 'text' && c.v.trim() === ''); }
  function sameCell(a, b) { return a.t === b.t && a.v === b.v; }
  function copyCell(c) { return { t: c.t, v: c.v }; }

  function cellFromSheet(XLSX, c) {
    if (!c || c.t === 'z' || c.v === undefined || c.v === null) return copyCell(BLANK);
    if (c.t === 's') return { t: 'text', v: String(c.v) };
    if (c.t === 'n') {
      if (c.z && XLSX.SSF.is_date(c.z)) {
        var p = XLSX.SSF.parse_date_code(c.v);
        if (p) return { t: 'date', v: ymd(p.y, p.m, p.d) };
      }
      return { t: 'number', v: c.v };
    }
    if (c.t === 'd') {
      var dt = c.v instanceof Date ? c.v : new Date(c.v);
      return { t: 'date', v: ymd(dt.getFullYear(), dt.getMonth() + 1, dt.getDate()) };
    }
    if (c.t === 'b') return { t: 'bool', v: !!c.v };
    if (c.t === 'e') return { t: 'text', v: c.w || '#ERROR' };
    return { t: 'text', v: String(c.v) };
  }

  /** 화면·변경 기록용 표시값 */
  function display(c, opts) {
    if (!c || c.t === 'blank') return '';
    if (c.t === 'date') return opts && opts.markExcelDate ? c.v + ' (엑셀 날짜값)' : c.v;
    if (c.t === 'bool') return c.v ? 'TRUE' : 'FALSE';
    return c.v;
  }
  /** 엑셀 출력용 값 */
  function outValue(c) {
    if (!c || c.t === 'blank') return null;
    return c.v;
  }

  function fill(template, params) {
    return String(template || '').replace(/\{(\w+)\}/g, function (m, k) {
      return params && params[k] !== undefined && params[k] !== null ? String(params[k]) : m;
    });
  }
  function resolvePath(obj, path) {
    return String(path).split('.').reduce(function (o, k) { return o == null ? undefined : o[k]; }, obj);
  }
  function asArray(x) { return x == null ? [] : Array.isArray(x) ? x : [x]; }

  /* ---------- 1. 엑셀 읽기 ---------- */

  function readWorkbook(XLSX, data, rules) {
    var isNodeBuf = typeof Buffer !== 'undefined' && Buffer.isBuffer && Buffer.isBuffer(data);
    var wb = XLSX.read(data, { type: isNodeBuf ? 'buffer' : 'array', cellNF: true, cellDates: false });
    var ds = rules.dataset;
    var sheetName = wb.SheetNames[ds.sheet_index || 0];
    if (!sheetName) throw new Error('읽을 시트가 없습니다.');
    var ws = wb.Sheets[sheetName];
    if (!ws['!ref']) throw new Error('시트 "' + sheetName + '"이(가) 비어 있습니다.');
    var range = XLSX.utils.decode_range(ws['!ref']);
    var h0 = ds.header_row - 1;

    var headerIdx = {};
    var headerCount = 0;
    var extra = [];
    var known = ds.columns.map(function (c) { return c.name; });
    for (var c = range.s.c; c <= range.e.c; c++) {
      var hc = cellFromSheet(XLSX, ws[XLSX.utils.encode_cell({ r: h0, c: c })]);
      if (isBlank(hc)) continue;
      var name = String(hc.v).trim();
      headerCount++;
      if (headerIdx[name] === undefined) headerIdx[name] = c;
      if (known.indexOf(name) < 0 && OUTPUT_COLS.indexOf(name) < 0) extra.push(name);
    }
    var cleanerCols = OUTPUT_COLS.filter(function (n) { return headerIdx[n] !== undefined; });
    var rowIdIdx = headerIdx[ID_COLS.row];
    var seqIdIdx = headerIdx[ID_COLS.seq];
    var missing = known.filter(function (n) { return headerIdx[n] === undefined; });
    if (missing.length) {
      var err = new Error('필수 컬럼이 없습니다: ' + missing.join(', ') + ' (머리글 ' + ds.header_row + '행 기준)');
      err.code = 'MISSING_COLUMNS';
      throw err;
    }

    var rows = [];
    var skippedEmpty = 0;
    for (var r = ds.first_data_row - 1; r <= range.e.r; r++) {
      var cells = {};
      var allBlank = true;
      known.forEach(function (n) {
        var cell = cellFromSheet(XLSX, ws[XLSX.utils.encode_cell({ r: r, c: headerIdx[n] })]);
        if (!isBlank(cell)) allBlank = false;
        cells[n] = cell;
      });
      if (allBlank) { skippedEmpty++; continue; }
      var item = { excelRow: r + 1, cells: cells };
      // 정리본을 다시 올린 경우: 처음 업로드한 원본의 행 번호·데이터 순번을 이어받는다
      if (rowIdIdx !== undefined) item.origExcelRow = positiveInt(cellFromSheet(XLSX, ws[XLSX.utils.encode_cell({ r: r, c: rowIdIdx })]));
      if (seqIdIdx !== undefined) item.origSeq = positiveInt(cellFromSheet(XLSX, ws[XLSX.utils.encode_cell({ r: r, c: seqIdIdx })]));
      rows.push(item);
    }

    var identity = { source: rowIdIdx !== undefined ? 'inherited' : 'file', column: rowIdIdx !== undefined ? ID_COLS.row : null, inherited: 0, missing: 0, duplicated: [] };
    if (rowIdIdx !== undefined) {
      var seen = {};
      rows.forEach(function (it) {
        if (it.origExcelRow) {
          identity.inherited++;
          if (seen[it.origExcelRow] && identity.duplicated.indexOf(it.origExcelRow) < 0) identity.duplicated.push(it.origExcelRow);
          seen[it.origExcelRow] = true;
        } else identity.missing++;
      });
    }

    return {
      rows: rows,
      meta: {
        sheetNames: wb.SheetNames.slice(),
        sheetCount: wb.SheetNames.length,
        sheetName: sheetName,
        dataRowCount: rows.length,
        columnCount: headerCount,
        extraColumns: extra,
        cleanerColumns: cleanerCols,
        rowIdentity: identity,
        skippedEmptyRows: skippedEmpty
      }
    };
  }

  /* ---------- 2. 형식 정규화기 (auto_fix) ----------
   * 각 함수는 (rule, cell, ctx) → { matched, out } 을 돌려준다.
   * matched 이고 out 이 cell 과 다르면 자동 수정 대상이다.
   */
  var NORMALIZERS = {
    parse_price: function (rule, cell) {
      if (cell.t === 'number') return { matched: true, out: cell };
      if (cell.t !== 'text') return { matched: false };
      if (!new RegExp(rule.condition.pattern).test(cell.v)) return { matched: false };
      var s = cell.v;
      asArray(rule.transform.remove).forEach(function (ch) { s = s.split(ch).join(''); });
      var n = rule.transform.to === 'integer' ? parseInt(s, 10) : Number(s);
      if (!isFinite(n)) return { matched: false };
      return { matched: true, out: { t: 'number', v: n } };
    },

    trim: function (rule, cell) {
      if (cell.t !== 'text') return { matched: false };
      var out = cell.v;
      if (rule.transform.trim === 'both' || rule.transform.trim === 'start') out = out.replace(/^\s+/, '');
      if (rule.transform.trim === 'both' || rule.transform.trim === 'end') out = out.replace(/\s+$/, '');
      return { matched: true, out: { t: 'text', v: out } };
    },

    parse_date: function (rule, cell) {
      var cond = rule.condition;
      var order = String(cond.group_order || 'year-month-day').split('-');
      var inputs = cond.accepted_inputs || [];
      for (var i = 0; i < inputs.length; i++) {
        var inp = inputs[i];
        if (inp.kind === 'excel_date' && cell.t === 'date') {
          return { matched: true, out: { t: 'text', v: cell.v } };
        }
        if (inp.kind === 'text' && cell.t === 'text') {
          var m = new RegExp(inp.pattern).exec(cell.v);
          if (!m) continue;
          var parts = {};
          order.forEach(function (name, idx) { parts[name] = parseInt(m[idx + 1], 10); });
          if (cond.require_valid_calendar_date && !validDate(parts.year, parts.month, parts.day)) return { matched: false };
          var fmt = (rule.transform && rule.transform.format) || 'YYYY-MM-DD';
          var v = fmt.replace('YYYY', pad(parts.year, 4)).replace('MM', pad(parts.month, 2)).replace('DD', pad(parts.day, 2));
          return { matched: true, out: { t: 'text', v: v } };
        }
      }
      return { matched: false };
    },

    normalize_count: function (rule, cell) {
      var s;
      if (cell.t === 'text') s = cell.v;
      else if (cell.t === 'number' && rule.condition.type === 'text_or_number_matches') s = String(cell.v);
      else return { matched: false };
      var m = new RegExp(rule.condition.pattern).exec(s);
      if (!m) return { matched: false };
      var out = String(rule.transform.template).replace(/\{(\d+)\}/g, function (x, i) { return m[+i] !== undefined ? m[+i] : ''; });
      return { matched: true, out: { t: 'text', v: out } };
    },

    map_to_standard: function (rule, cell, ctx) {
      if (cell.t !== 'text' || !cell.v.trim()) return { matched: false };
      var list = asArray(resolvePath(ctx.rules, rule.condition.standard_list));
      var key = cell.v.replace(/\s+/g, '');
      for (var i = 0; i < list.length; i++) {
        if (String(list[i]).replace(/\s+/g, '') === key) return { matched: true, out: { t: 'text', v: list[i] } };
      }
      return { matched: false };
    }
  };

  /* ---------- 3. 검사기 (review / keep) ---------- */

  function cmp(a, op, b) {
    switch (op) {
      case '<': return a < b;
      case '<=': return a <= b;
      case '>': return a > b;
      case '>=': return a >= b;
      case '==': return a === b;
      case '!=': return a !== b;
    }
    throw new Error('알 수 없는 비교 연산자: ' + op);
  }
  var DAY_MS = 86400000;
  function comparable(cell) {
    if (cell.t === 'number') return { k: 'n', v: cell.v };
    var s = cell.t === 'text' || cell.t === 'date' ? String(cell.v).trim() : null;
    var m = s && /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) {
      var y = +m[1], mo = +m[2], d = +m[3];
      if (validDate(y, mo, d)) return { k: 'd', v: Date.UTC(y, mo - 1, d) / DAY_MS };
    }
    return null;
  }
  function activeRows(ctx) { return ctx.rows.filter(function (r) { return !r.removed; }); }
  function groupByKey(ctx, key) {
    var groups = new Map();
    activeRows(ctx).forEach(function (row) {
      var c = row.cur[key];
      if (isBlank(c)) return;
      var k = String(c.v).trim();
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(row);
    });
    return groups;
  }
  function allSame(rows, col) {
    for (var i = 1; i < rows.length; i++) if (!sameCell(rows[0].cur[col], rows[i].cur[col])) return false;
    return true;
  }

  var CHECKS = {
    text_matches: function (ctx, rule) {
      var re = new RegExp(rule.condition.pattern);
      activeRows(ctx).forEach(function (row) {
        rule.columns.forEach(function (col) {
          var c = row.cur[col];
          if (c.t === 'text' && re.test(c.v)) addReview(ctx, row, rule, [col], { column: col, value: c.v });
        });
      });
    },

    not_in_standard_after_mapping: function (ctx, rule) {
      var list = asArray(resolvePath(ctx.rules, rule.condition.standard_list));
      activeRows(ctx).forEach(function (row) {
        rule.columns.forEach(function (col) {
          var c = row.cur[col];
          if (isBlank(c)) { if (rule.condition.skip_blank) return; }
          else if (list.indexOf(String(c.v)) >= 0) return;
          addReview(ctx, row, rule, [col], { column: col, value: display(c) });
        });
      });
    },

    blank: function (ctx, rule) {
      activeRows(ctx).forEach(function (row) {
        rule.columns.forEach(function (col) {
          if (isBlank(row.cur[col])) addReview(ctx, row, rule, [col], { column: col });
        });
      });
    },

    number_compare: function (ctx, rule) {
      var cond = rule.condition;
      activeRows(ctx).forEach(function (row) {
        rule.columns.forEach(function (col) {
          var c = row.cur[col];
          if (c.t === 'number' && cmp(c.v, cond.operator, cond.value)) addReview(ctx, row, rule, [col], { column: col, value: c.v });
        });
      });
    },

    column_compare: function (ctx, rule) {
      var cond = rule.condition;
      activeRows(ctx).forEach(function (row) {
        var L = comparable(row.cur[cond.left]);
        var R = comparable(row.cur[cond.right]);
        if (!L || !R || L.k !== R.k) return; // 비교할 수 없으면 판정하지 않는다 (빈칸 등은 다른 규칙이 표시)
        if (cmp(L.v, cond.operator, R.v)) {
          addReview(ctx, row, rule, [cond.left, cond.right], {
            left: display(row.cur[cond.left]), right: display(row.cur[cond.right]), days: L.v - R.v, diff: L.v - R.v
          });
        }
      });
    },

    price_ratio_outlier: function (ctx, rule) {
      var cond = rule.condition;
      var parts = String(cond.ratio).split('/').map(function (s) { return s.trim(); });
      var f = cond.flag_if || {};
      activeRows(ctx).forEach(function (row) {
        var a = row.cur[parts[0]], b = row.cur[parts[1]];
        if (a.t !== 'number' || b.t !== 'number' || !(a.v > 0) || !(b.v > 0)) return;
        var ratio = a.v / b.v;
        var hit = (f.ratio_lt != null && ratio < f.ratio_lt) || (f.ratio_lte != null && ratio <= f.ratio_lte) ||
                  (f.ratio_gt != null && ratio > f.ratio_gt) || (f.ratio_gte != null && ratio >= f.ratio_gte);
        if (hit) addReview(ctx, row, rule, rule.columns.slice(), { ratio: Number(ratio.toFixed(3)) });
      });
    },

    count_out_of_range: function (ctx, rule) {
      var cond = rule.condition;
      activeRows(ctx).forEach(function (row) {
        rule.columns.forEach(function (col) {
          var c = row.cur[col];
          if (c.t !== 'text' && c.t !== 'number') return;
          var m = /^\s*(\d+)/.exec(String(c.v));
          if (!m) return;
          var n = parseInt(m[1], 10);
          if (n < cond.observed_min || n > cond.observed_max) addReview(ctx, row, rule, [col], { column: col, value: c.v });
        });
      });
    },

    minority_spacing_variant: function (ctx, rule) {
      rule.columns.forEach(function (col) {
        var items = [];
        activeRows(ctx).forEach(function (row) {
          var c = row.cur[col];
          if (c.t !== 'text' || !c.v.trim()) return;
          items.push({ row: row, tokens: c.v.trim().split(/\s+/) });
        });
        var pairs = new Map();
        items.forEach(function (it) {
          for (var i = 0; i + 1 < it.tokens.length; i++) {
            var key = it.tokens[i] + '\u0000' + it.tokens[i + 1];
            if (!pairs.has(key)) pairs.set(key, [it.tokens[i], it.tokens[i + 1]]);
          }
        });
        var flagged = new Map(); // row -> [{variant, majority}]
        function add(list, variant, majority) {
          list.forEach(function (it) {
            if (!flagged.has(it.row)) flagged.set(it.row, []);
            var arr = flagged.get(it.row);
            if (!arr.some(function (x) { return x.variant === variant; })) arr.push({ variant: variant, majority: majority });
          });
        }
        pairs.forEach(function (ab) {
          var A = ab[0], B = ab[1], joined = A + B;
          var sep = items.filter(function (it) {
            for (var i = 0; i + 1 < it.tokens.length; i++) if (it.tokens[i] === A && it.tokens[i + 1] === B) return true;
            return false;
          });
          var join = items.filter(function (it) { return it.tokens.some(function (t) { return t.indexOf(joined) === 0; }); });
          if (!sep.length || !join.length) return;
          var sepForm = A + ' ' + B;
          if (sep.length < join.length) add(sep, sepForm, joined);
          else if (join.length < sep.length) add(join, joined, sepForm);
          else { add(sep, sepForm, joined); add(join, joined, sepForm); }
        });
        flagged.forEach(function (list, row) {
          addReview(ctx, row, rule, [col], {
            column: col,
            variant: list.map(function (x) { return x.variant; }).join(', '),
            majority: list.map(function (x) { return x.majority; }).join(', ')
          });
        });
      });
    },

    same_key_different_values: function (ctx, rule) {
      var cond = rule.condition;
      groupByKey(ctx, cond.key).forEach(function (rows, code) {
        if (rows.length < 2) return;
        var diff = cond.compare_columns.filter(function (c) { return !allSame(rows, c); });
        if (!diff.length) return;
        rows.forEach(function (row) {
          addReview(ctx, row, rule, diff, {
            code: code,
            columns: diff.join(', '),
            other_rows: rows.filter(function (r) { return r !== row; }).map(function (r) { return r.excelRow; }).join(', ')
          });
        });
      });
    },

    same_key_same_product_other_diff: function (ctx, rule) {
      var cond = rule.condition;
      var others = ctx.colNames.filter(function (c) { return c !== cond.key && cond.same_columns.indexOf(c) < 0; });
      groupByKey(ctx, cond.key).forEach(function (rows, code) {
        if (rows.length < 2) return;
        if (!cond.same_columns.every(function (c) { return allSame(rows, c); })) return;
        var diff = others.filter(function (c) { return !allSame(rows, c); });
        if (!diff.length) return;
        rows.forEach(function (row) {
          addReview(ctx, row, rule, diff, {
            code: code,
            columns: diff.join(', '),
            other_rows: rows.filter(function (r) { return r !== row; }).map(function (r) { return r.excelRow; }).join(', ')
          });
        });
      });
    },

    // KEEP: 시술코드가 다르면 병합하지 않는다 — 아무것도 바꾸지 않는다.
    different_key: function () {}
  };

  /* ---------- 4. 기록 ---------- */

  function actionLabel(ctx, action) {
    var s = (ctx.rules.actions || {})[action] || action;
    return String(s).split(':')[0].trim();
  }
  function stat(ctx, rule) {
    if (!ctx.stats[rule.id]) ctx.stats[rule.id] = { id: rule.id, executed: false, entries: 0, rows: new Set() };
    return ctx.stats[rule.id];
  }
  function logEntry(ctx, row, column, before, after, method, reason, ruleId) {
    var e = {};
    e['원본_엑셀행번호'] = row.excelRow;
    e['데이터순번'] = row.seq;
    e['시술코드'] = display(row.raw[ctx.keyCol]);
    e['변경컬럼'] = column;
    e['변경전값'] = before;
    e['변경후값'] = after;
    e['처리방식'] = method;
    e['처리이유'] = reason;
    e['규칙ID'] = ruleId;
    e._order = ctx.log.length;
    ctx.log.push(e);
  }
  function addFix(ctx, row, rule, col, before, after) {
    row.cur[col] = after;
    row.fixes.push({ ruleId: rule.id, col: col, before: copyCell(before), after: copyCell(after) });
    var s = stat(ctx, rule); s.entries++; s.rows.add(row.excelRow);
    logEntry(ctx, row, col, display(before, { markExcelDate: true }), display(after),
      actionLabel(ctx, rule.action), rule.name + ' — ' + rule.why, rule.id);
  }
  function addReview(ctx, row, rule, cols, params) {
    var reason = fill(rule.review_reason, params);
    row.reviews.push({ ruleId: rule.id, name: rule.name, reason: reason, cols: cols.slice(), flagType: rule.flag_type || null, why: rule.why });
    var s = stat(ctx, rule); s.entries++; s.rows.add(row.excelRow);
    logEntry(ctx, row, cols.join(', '),
      cols.map(function (c) { return display(row.cur[c]); }).join(' / '),
      '(변경 없음)', actionLabel(ctx, rule.action), reason, rule.id);
  }

  /* ---------- 5. 실행 ---------- */

  function runDedupe(ctx, rule) {
    var cond = rule.condition;
    var cols = rule.columns;
    var seen = new Map();
    var list = cond.keep === 'last_occurrence' ? ctx.rows.slice().reverse() : ctx.rows;
    list.forEach(function (row) {
      var sig = JSON.stringify(cols.map(function (c) {
        var cell = cond.compare === 'raw_values' ? row.raw[c] : row.cur[c];
        return [cell.t, cell.v];
      }));
      if (!seen.has(sig)) { seen.set(sig, row); return; }
      var kept = seen.get(sig);
      row.removed = true;
      row.keptRow = kept.excelRow;
      var s = stat(ctx, rule); s.entries++; s.rows.add(row.excelRow);
      var reason = fill(rule.log && rule.log.reason_template, { kept_row: kept.excelRow });
      logEntry(ctx, row, '(행 전체)',
        ctx.colNames.map(function (c) { return display(row.raw[c], { markExcelDate: true }); }).join(' | '),
        '(삭제)', actionLabel(ctx, rule.action) + '(행 삭제)', reason, rule.id);
    });
  }

  function runNormalizer(ctx, rule) {
    var fn = NORMALIZERS[rule.operation];
    if (!fn) throw new Error('규칙 ' + rule.id + '의 처리(' + rule.operation + ')를 실행할 수 없습니다.');
    activeRows(ctx).forEach(function (row) {
      rule.columns.forEach(function (col) {
        var cell = row.cur[col];
        if (isBlank(cell)) return;
        var res = fn(rule, cell, ctx);
        if (res.matched && !sameCell(res.out, cell)) addFix(ctx, row, rule, col, cell, res.out);
      });
    });
  }

  function runRule(ctx, rule) {
    if (rule.action === 'auto_fix') {
      if (rule.operation === 'remove_row') runDedupe(ctx, rule);
      else runNormalizer(ctx, rule);
    } else {
      var fn = CHECKS[rule.condition && rule.condition.type];
      if (!fn) throw new Error('규칙 ' + rule.id + '의 조건 유형(' + (rule.condition && rule.condition.type) + ')을 실행할 수 없습니다.');
      fn(ctx, rule);
    }
    stat(ctx, rule).executed = true;
  }

  /** format_fallback: 형식 통일 컬럼에 어떤 형식에도 맞지 않는 값이 있으면 검토 표시 */
  function runFallback(ctx) {
    var fb = ctx.rules.format_fallback;
    if (!fb || fb.policy !== 'review') return;
    var pseudo = { id: 'format_fallback', name: '알 수 없는 형식', action: 'review', why: fb.why, review_reason: '{column} 값을 정해진 형식으로 해석할 수 없음' };
    var byCol = {};
    asArray(fb.applies_to).forEach(function (id) {
      var r = ctx.ruleById[id];
      if (!r) return;
      r.columns.forEach(function (c) { (byCol[c] = byCol[c] || []).push(r); });
    });
    activeRows(ctx).forEach(function (row) {
      Object.keys(byCol).forEach(function (col) {
        var cell = row.cur[col];
        if (isBlank(cell)) return;
        if (row.reviews.some(function (rv) { return rv.cols.indexOf(col) >= 0; })) return;
        var ok = byCol[col].some(function (r) {
          var res = NORMALIZERS[r.operation](r, cell, ctx);
          return res.matched && sameCell(res.out, cell);
        });
        if (!ok) addReview(ctx, row, pseudo, [col], { column: col });
      });
    });
    stat(ctx, pseudo).executed = true;
  }

  function checkOrder(rules) {
    var pos = {};
    var i = 0;
    rules.execution_order.forEach(function (step) { step.rules.forEach(function (id) { pos[id] = i++; }); });
    var warnings = [];
    rules.rules.forEach(function (r) {
      if (pos[r.id] === undefined) warnings.push('규칙 ' + r.id + '이(가) 실행 순서(execution_order)에 없습니다.');
      asArray(r.condition && r.condition.evaluate_after).forEach(function (dep) {
        if (pos[dep] === undefined || pos[dep] >= pos[r.id]) warnings.push('규칙 ' + r.id + '은(는) ' + dep + ' 다음에 실행되어야 합니다.');
      });
    });
    Object.keys(pos).forEach(function (id) {
      if (!rules.rules.some(function (r) { return r.id === id; })) warnings.push('실행 순서에 있는 ' + id + ' 규칙이 rules 목록에 없습니다.');
    });
    return warnings;
  }

  function clean(table, rules) {
    var ds = rules.dataset;
    var colNames = ds.columns.map(function (c) { return c.name; });
    var keyDef = ds.columns.filter(function (c) { return c.role === 'key'; })[0] || ds.columns[0];
    var rows = table.rows.map(function (r) {
      var raw = {}, cur = {};
      colNames.forEach(function (n) { raw[n] = copyCell(r.cells[n] || BLANK); cur[n] = copyCell(r.cells[n] || BLANK); });
      // excelRow 는 항상 "처음 업로드한 원본 엑셀의 행 번호"다. 정리본을 다시 올리면 _원본_엑셀행번호 를 이어받는다.
      var orig = r.origExcelRow || r.excelRow;
      var seq = r.origSeq || (orig - ds.header_row);
      return { excelRow: orig, seq: seq, fileRow: r.excelRow, inherited: !!r.origExcelRow,
        raw: raw, cur: cur, removed: false, keptRow: null, reviews: [], fixes: [] };
    });
    var ctx = { rules: rules, colNames: colNames, keyCol: keyDef.name, rows: rows, log: [], stats: {}, ruleById: {} };
    rules.rules.forEach(function (r) { ctx.ruleById[r.id] = r; stat(ctx, r); });
    var warnings = checkOrder(rules);

    rules.execution_order.forEach(function (step) {
      step.rules.forEach(function (id) {
        var rule = ctx.ruleById[id];
        if (!rule) throw new Error('실행 순서의 규칙 ' + id + '을(를) 찾을 수 없습니다.');
        runRule(ctx, rule);
      });
    });
    runFallback(ctx);

    rows.forEach(function (row) {
      var conflict = row.reviews.some(function (r) { return r.flagType === 'conflict'; });
      row.status = row.removed ? '삭제됨' : conflict ? '충돌' : row.reviews.length ? '검토 필요' : '정상';
      row.reasonText = row.reviews.map(function (r) { return '[' + r.ruleId + '] ' + r.reason; }).join('; ');
    });

    var log = ctx.log.slice().sort(function (a, b) { return a['원본_엑셀행번호'] - b['원본_엑셀행번호'] || a._order - b._order; });
    log.forEach(function (e) { delete e._order; });

    var active = rows.filter(function (r) { return !r.removed; });
    var ruleStats = rules.rules.map(function (r) {
      var s = ctx.stats[r.id];
      return { id: r.id, name: r.name, action: r.action, executed: s.executed, entries: s.entries, rowCount: s.rows.size };
    });
    var fb = ctx.stats.format_fallback;

    return {
      colNames: colNames,
      keyCol: keyDef.name,
      rows: rows,
      active: active,
      removed: rows.filter(function (r) { return r.removed; }),
      log: log,
      ruleStats: ruleStats,
      fallbackStats: fb ? { entries: fb.entries, rowCount: fb.rows.size } : { entries: 0, rowCount: 0 },
      warnings: warnings,
      summary: {
        rowsBefore: rows.length,
        rowsAfter: active.length,
        removedRows: rows.length - active.length,
        autoFixCells: rows.reduce(function (n, r) { return n + r.fixes.length; }, 0),
        autoFixRows: rows.filter(function (r) { return r.fixes.length; }).length,
        reviewRows: active.filter(function (r) { return r.reviews.length; }).length,
        reviewEntries: active.reduce(function (n, r) { return n + r.reviews.length; }, 0),
        conflictRows: active.filter(function (r) { return r.status === '충돌'; }).length,
        rulesTotal: rules.rules.length,
        rulesExecuted: ruleStats.filter(function (s) { return s.executed; }).length,
        rulesWithHits: ruleStats.filter(function (s) { return s.entries > 0; }).length
      }
    };
  }

  /* ---------- 6. 결과 엑셀 ---------- */

  function sheetFrom(XLSX, aoa, widths) {
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    if (widths) ws['!cols'] = widths.map(function (w) { return { wch: w }; });
    return ws;
  }

  function buildSheets(XLSX, result, rules) {
    var cols = result.colNames;
    var fields = (rules.change_log && rules.change_log.fields) || [];

    var cleanAoa = [cols.concat(['_검토상태', '_검토사유', ID_COLS.row])];
    result.active.forEach(function (row) {
      cleanAoa.push(cols.map(function (c) { return outValue(row.cur[c]); }).concat([row.status, row.reasonText, row.excelRow]));
    });

    var logAoa = [fields];
    result.log.forEach(function (e) { logAoa.push(fields.map(function (f) { return e[f] === undefined ? null : e[f]; })); });

    var remAoa = [cols.concat([ID_COLS.row, ID_COLS.seq, '_유지된_행(원본_엑셀행번호)'])];
    result.removed.forEach(function (row) {
      remAoa.push(cols.map(function (c) { return outValue(row.raw[c]); }).concat([row.excelRow, row.seq, row.keptRow]));
    });

    var w10 = [10, 10, 14, 26, 12, 10, 12, 12, 13, 13];
    return {
      clean: sheetFrom(XLSX, cleanAoa, w10.concat([10, 60, 14])),
      log: sheetFrom(XLSX, logAoa, [14, 10, 10, 16, 30, 20, 14, 60, 10]),
      removed: sheetFrom(XLSX, remAoa, w10.concat([14, 10, 22]))
    };
  }

  var SHEET = { clean: '정리본', log: '변경기록', removed: '제거된_중복행' };

  function buildWorkbooks(XLSX, result, rules) {
    function book(parts) {
      var s = buildSheets(XLSX, result, rules);
      var wb = XLSX.utils.book_new();
      parts.forEach(function (p) { XLSX.utils.book_append_sheet(wb, s[p], SHEET[p]); });
      return wb;
    }
    return {
      cleaned: book(['clean', 'log', 'removed']),
      log: book(['log']),
      removed: book(['removed'])
    };
  }

  /* ---------- 7. 정답 CSV 대조 ---------- */

  function parseCsv(XLSX, text) {
    var wb = XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
    var ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false });
  }

  function cellText(c) {
    if (!c || c.t === 'blank') return '';
    return String(c.v).trim();
  }
  function sameValue(a, b) {
    if (a === b) return true;
    if (a !== '' && b !== '') {
      var na = Number(a), nb = Number(b);
      if (isFinite(na) && isFinite(nb)) return na === nb;
    }
    return false;
  }

  function compareWithAnswer(XLSX, text, result) {
    var aoa = parseCsv(XLSX, text);
    if (!aoa.length) throw new Error('정답 파일이 비어 있습니다.');
    var header = aoa[0].map(function (h) { return String(h).replace(/^﻿/, '').trim(); });
    // line = 정답 CSV의 데이터 순번(머리글 제외, 1부터)
    var body = aoa.slice(1).map(function (r, i) { return { cells: r, line: i + 1 }; })
      .filter(function (a) { return a.cells.some(function (v) { return String(v).trim() !== ''; }); });

    var colIdx = {};
    header.forEach(function (h, i) { if (colIdx[h] === undefined) colIdx[h] = i; });
    var compared = result.colNames.filter(function (c) { return colIdx[c] !== undefined; });
    var missing = result.colNames.filter(function (c) { return colIdx[c] === undefined; });
    if (!compared.length) throw new Error('정답 파일에서 비교할 컬럼을 찾지 못했습니다. 머리글을 확인해 주세요.');

    // 대응 기준: 정답에 원본 행 번호 컬럼이 있으면 그 번호, 없으면 시술코드.
    // 같은 키가 정리본·정답 중 어느 한쪽에라도 두 번 이상 있으면 행을 고르지 않고 '대응 보류'로 둔다.
    var rowCol = header.findIndex(function (h) { return /원본.*행/.test(h) && !/데이터/.test(h); });
    var keyCol = result.keyCol;
    var keyMode = rowCol >= 0 ? 'row' : 'code';
    if (keyMode === 'code' && colIdx[keyCol] === undefined) throw new Error('정답 파일에 "' + keyCol + '" 컬럼이나 원본 행 번호 컬럼이 없어 행을 맞출 수 없습니다.');

    function grouped(list, getKey) {
      var m = new Map();
      list.forEach(function (item) {
        var k = getKey(item);
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(item);
      });
      return m;
    }
    var ours = grouped(result.active, function (r) { return keyMode === 'row' ? String(r.excelRow) : cellText(r.cur[keyCol]); });
    var theirs = grouped(body, function (a) { return keyMode === 'row' ? String(a.cells[rowCol]).trim() : String(a.cells[colIdx[keyCol]]).trim(); });

    function answerValues(a) {
      var o = {};
      compared.forEach(function (c) { var v = a.cells[colIdx[c]]; o[c] = String(v == null ? '' : v).trim(); });
      return o;
    }
    function diffCols(row, ans) {
      return compared.filter(function (c) { return !sameValue(cellText(row.cur[c]), ans[c]); })
        .map(function (c) { return { col: c, cleaned: cellText(row.cur[c]), answer: ans[c] }; });
    }

    var matched = 0;
    var diffs = [];
    var pairs = [];
    var held = [];
    var onlyCleaned = [];
    var onlyAnswer = [];
    var keys = [];
    ours.forEach(function (v, k) { keys.push(k); });
    theirs.forEach(function (v, k) { if (!ours.has(k)) keys.push(k); });

    keys.forEach(function (k) {
      var o = ours.get(k) || [], t = theirs.get(k) || [];
      if (!t.length) { o.forEach(function (row) { onlyCleaned.push({ key: k, row: row, flagged: row.reviews.length > 0 }); }); return; }
      if (!o.length) { t.forEach(function (a) { onlyAnswer.push({ key: k, line: a.line, values: answerValues(a) }); }); return; }
      if (o.length === 1 && t.length === 1) {
        var ans = answerValues(t[0]);
        var cols = diffCols(o[0], ans);
        var pair = { key: k, row: o[0], line: t[0].line, flagged: o[0].reviews.length > 0, cols: cols, answer: ans };
        pairs.push(pair);
        if (!cols.length) matched++;
        else diffs.push(pair);
        return;
      }
      // 대응 보류: 어느 행이 정답과 대응하는지 정하지 않는다. 참고용 차이만 계산한다(정답이 1행일 때).
      var answers = t.map(function (a) { return { line: a.line, values: answerValues(a) }; });
      held.push({
        key: k,
        cleaned: o.map(function (row) {
          return { row: row, flagged: row.reviews.length > 0, cols: answers.length === 1 ? diffCols(row, answers[0].values) : null };
        }),
        answers: answers
      });
    });

    var heldCleanedRows = held.reduce(function (n, g) { return n + g.cleaned.length; }, 0);
    var heldAnswerRows = held.reduce(function (n, g) { return n + g.answers.length; }, 0);
    return {
      keyMode: keyMode,
      keyColumnName: keyMode === 'row' ? header[rowCol] : keyCol,
      comparedColumns: compared,
      missingColumns: missing,
      answerRows: body.length,
      cleanedRows: result.active.length,
      pairedRows: pairs.length,
      matched: matched,
      mismatched: diffs.length,
      diffCells: diffs.reduce(function (n, d) { return n + d.cols.length; }, 0),
      diffs: diffs,
      pairs: pairs,
      held: held,
      heldCodes: held.length,
      heldCleanedRows: heldCleanedRows,
      heldAnswerRows: heldAnswerRows,
      onlyCleaned: onlyCleaned,
      onlyAnswer: onlyAnswer,
      mismatchFlagged: diffs.filter(function (d) { return d.flagged; }).length,
      mismatchUnflagged: diffs.filter(function (d) { return !d.flagged; }).length
    };
  }

  /** 두 실행 결과가 같은지 비교하기 위한 서명 */
  function signature(result) {
    return JSON.stringify({
      rows: result.rows.map(function (r) {
        return [r.excelRow, r.removed, r.keptRow, result.colNames.map(function (c) { return [r.cur[c].t, r.cur[c].v]; }), r.status, r.reasonText];
      }),
      log: result.log
    });
  }

  return {
    readWorkbook: readWorkbook,
    clean: clean,
    buildWorkbooks: buildWorkbooks,
    compareWithAnswer: compareWithAnswer,
    signature: signature,
    display: display,
    isBlank: isBlank,
    checkOrder: checkOrder
  };
});
