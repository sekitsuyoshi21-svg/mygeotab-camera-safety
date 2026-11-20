/*
 * 安全運転ダッシュボード MyGeotab Add-In
 * GO Focus Plus（カメラ）の警報と、急ブレーキ・急加速などの運転ルールの例外イベント（ExceptionEvent）を
 * 集計し、種類別件数・日別推移・運転手/車両別の安全スコア・警報一覧（映像・地図）を表示する。
 * MyGeotab 外（ローカルでファイルを開いた場合）はデモデータで動作する。
 */
/* global geotab */
var geotab = window.geotab || { addin: {} };
geotab.addin = geotab.addin || {};

geotab.addin.cameraSafety = function () {
  'use strict';

  var VERSION = '0.1.0';

  // 警報の種類。ルール名（日本語・英語）から判定する。上から順に判定。
  // sev: crit=重大 / warn=注意 / info=その他、w: 安全スコアの重み（0はスコアに影響しない）
  var TYPES = [
    { key: 'near', label: '衝突警告', sev: 'crit', w: 5, re: /near\s*collision|forward\s*collision|ニア.*衝突|衝突.*(警告|寸前|回避|注意)|前方衝突/i },
    { key: 'collision', label: '衝突検知', sev: 'crit', w: 8, re: /collision|crash|accident|衝突|事故|衝撃/i },
    { key: 'phone', label: 'スマホ使用', sev: 'crit', w: 5, re: /phone|cell|携帯|スマホ|通話/i },
    { key: 'distract', label: '脇見運転', sev: 'crit', w: 4, re: /distract|inattent|脇見|わき見|よそ見/i },
    { key: 'fatigue', label: '居眠り・疲労', sev: 'crit', w: 5, re: /fatigue|drows|yawn|居眠り|眠気|疲労|あくび/i },
    { key: 'tailgate', label: '車間距離不足', sev: 'warn', w: 3, re: /tailgat|close\s*follow|following\s*distance|車間/i },
    { key: 'seatbelt', label: 'シートベルト', sev: 'warn', w: 3, re: /seat\s*belt|シートベルト/i },
    { key: 'stop', label: '一時停止違反', sev: 'warn', w: 3, re: /rolling\s*stop|stop\s*sign|一時停止/i },
    { key: 'lane', label: '車線はみ出し', sev: 'warn', w: 2, re: /lane|straddl|車線|はみ出/i },
    { key: 'brake', label: '急ブレーキ', sev: 'warn', w: 2, re: /harsh\s*brak|hard\s*brak|急ブレーキ|急制動|急減速/i },
    { key: 'accel', label: '急加速', sev: 'warn', w: 2, re: /harsh\s*accel|hard\s*accel|jackrabbit|急加速|急発進/i },
    { key: 'corner', label: '急ハンドル', sev: 'warn', w: 2, re: /corner|急ハンドル|急旋回|急カーブ/i },
    { key: 'speed', label: '速度超過', sev: 'warn', w: 2, re: /speed|速度超過|スピード/i },
    { key: 'smoking', label: '喫煙', sev: 'info', w: 1, re: /smok|喫煙|たばこ|タバコ/i },
    { key: 'drinking', label: '飲食', sev: 'info', w: 1, re: /drink|eating|飲食|飲み物/i },
    { key: 'button', label: '手動録画', sev: 'info', w: 0, re: /button|manual|ボタン|手動/i }
  ];
  var OTHER = { key: 'other', label: 'その他', sev: 'info', w: 1 };
  var TYPE_BY_KEY = {}; TYPES.concat([OTHER]).forEach(function (t) { TYPE_BY_KEY[t.key] = t; });
  // Geotab標準ルールの固定ID（言語によらない）
  var KNOWN_RULES = { RuleHarshBrakingId: 'brake', RuleJackrabbitStartsId: 'accel', RuleHarshCorneringId: 'corner', RuleSeatbeltId: 'seatbelt', RulePostedSpeedingId: 'speed' };
  // 自動選択の対象：上の種類に当てはまるか、カメラ関連の言葉を含むルール
  var CAMERA_RE = /camera|カメラ|go\s*focus|video|ビデオ|映像|adas|dms|ドラレコ/i;
  var SEV_ORDER = { crit: 0, warn: 1, info: 2 };
  var UNKNOWN_DRIVER = 'UnknownDriverId';

  var api = null, state = null, demo = false, storageKey = 'cameraSafety:demo', server = '', session = null;
  var settings = { rules: null, videoHash: '' };   // rules: [{id, name}] / null=未設定（自動選択）
  var marks = {};                                   // 確認済み・メモ { eventId: {ack, memo} }
  var data = null, loading = false;
  var view = { type: null, entity: null, rankBy: 'driver', onlyTodo: false };
  var listRows = [], detailIdx = -1;
  var $ = function (id) { return document.getElementById(id); };

  // ---------- ユーティリティ ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  var JST = 9 * 3600e3, DAY = 864e5;
  function fmtT(d) { return d ? d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '-'; }
  function fmtShort(d) { return d ? d.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', month: 'numeric', day: 'numeric', weekday: 'short', hour: '2-digit', minute: '2-digit' }) : '-'; }
  function jstMidnight(t) { return Math.floor((t + JST) / DAY) * DAY - JST; }
  function call(method, params) { return new Promise(function (res, rej) { api.call(method, params, res, rej); }); }
  function multi(calls) { return new Promise(function (res, rej) { api.multiCall(calls, res, rej); }); }
  function showMsg(text, info) { var el = $('cs-message'); el.textContent = text; el.className = 'cs-message' + (info ? ' info' : ''); el.hidden = !text; }
  function typeOfRule(rule) {
    if (KNOWN_RULES[rule.id]) return TYPE_BY_KEY[KNOWN_RULES[rule.id]];
    var n = rule.name || '';
    for (var i = 0; i < TYPES.length; i++) if (TYPES[i].re.test(n)) return TYPES[i];
    return OTHER;
  }
  function isTarget(rule) { return !!KNOWN_RULES[rule.id] || typeOfRule(rule) !== OTHER || CAMERA_RE.test(rule.name || ''); }
  function load(key, def) { try { var v = JSON.parse(localStorage.getItem(storageKey + ':' + key)); return v == null ? def : v; } catch (e) { return def; } }
  function save(key, v) { try { localStorage.setItem(storageKey + ':' + key, JSON.stringify(v)); } catch (e) { /* ignore */ } }

  // ---------- 期間 ----------
  function period() {
    var d = parseInt($('cs-days').value, 10), now = Date.now(), from, to = now;
    if (d === 1) from = jstMidnight(now);
    else if (d === -1) { to = jstMidnight(now); from = to - DAY; }
    else from = now - d * DAY;
    return { from: new Date(from), to: new Date(to), prevFrom: new Date(from - (to - from)), hourly: to - from <= DAY + 1 };
  }

  // ---------- データ取得（ライブ） ----------
  function loadRules() {
    return call('Get', { typeName: 'Rule' }).then(function (list) {
      return (list || []).map(function (r) { return { id: r.id, name: r.name || r.id, baseType: r.baseType }; })
        .sort(function (a, b) { return a.name.localeCompare(b.name, 'ja'); });
    });
  }

  function loadLive() {
    var p = period();
    return loadRules().then(function (rules) {
      var ruleById = {}; rules.forEach(function (r) { ruleById[r.id] = r; });
      if (!settings.rules) {                     // 初回は自動選択
        settings.rules = rules.filter(isTarget).map(function (r) { return { id: r.id, name: r.name }; });
        save('settings', settings);
      }
      var sel = settings.rules.filter(function (r) { return ruleById[r.id]; });
      if (!sel.length) { var e = new Error('対象の警報ルールが見つかりませんでした。「設定」から集計するルールを選んでください。'); e.info = true; throw e; }
      var groups = state && state.getGroupFilter ? state.getGroupFilter() : null;
      var calls = [
        ['Get', { typeName: 'Device', search: groups && groups.length ? { groups: groups } : {}, resultsLimit: 5000 }],
        ['Get', { typeName: 'User', search: { isDriver: true }, resultsLimit: 5000 }],
        ['Get', { typeName: 'Trip', search: { fromDate: p.from.toISOString(), toDate: p.to.toISOString() }, resultsLimit: 50000 }]
      ];
      sel.forEach(function (r) {
        calls.push(['Get', { typeName: 'ExceptionEvent', search: { ruleSearch: { id: r.id }, fromDate: p.prevFrom.toISOString(), toDate: p.to.toISOString() }, resultsLimit: 50000 }]);
      });
      return Promise.all([multi(calls), loadMedia(p)]).then(function (res) {
        var r = res[0], media = res[1];
        var devices = {}, drivers = {};
        r[0].forEach(function (d) { devices[d.id] = d.name || d.id; });
        r[1].forEach(function (u) { drivers[u.id] = ((u.lastName || '') + ' ' + (u.firstName || '')).trim() || u.name || u.id; });
        var dist = { driver: {}, device: {} };
        r[2].forEach(function (t) {
          var dv = t.device && t.device.id; if (!devices[dv]) return;
          var dr = driverId(t.driver);
          dist.device[dv] = (dist.device[dv] || 0) + (t.distance || 0);
          dist.driver[dr] = (dist.driver[dr] || 0) + (t.distance || 0);
        });
        var events = [];
        sel.forEach(function (rule, i) {
          var type = typeOfRule(ruleById[rule.id]);
          (r[3 + i] || []).forEach(function (x) {
            var dv = x.device && x.device.id; if (!devices[dv]) return;
            events.push({ id: x.id, t: new Date(x.activeFrom), end: x.activeTo ? new Date(x.activeTo) : null, devId: dv,
              drvId: driverId(x.driver), type: type, ruleName: ruleById[rule.id].name, pos: undefined });
          });
        });
        attachMedia(events, media);
        return build(events, devices, drivers, dist, p);
      });
    });
  }
  function driverId(d) { return d && d.id && d.id !== 'NoDriverId' ? d.id : UNKNOWN_DRIVER; }

  // MediaFile（映像ファイル）。GO Focus Plusでは使われないことが多いため、取れなくても続行する。
  function loadMedia(p) {
    return call('Get', { typeName: 'MediaFile', search: { fromDate: p.from.toISOString(), toDate: p.to.toISOString() }, resultsLimit: 5000 })
      .then(function (l) { return (l || []).filter(function (m) { return m.device && (!m.status || m.status === 'Ready'); }); }, function () { return []; });
  }
  function attachMedia(events, media) {
    if (!media.length) return;
    events.forEach(function (e) {
      var t0 = +e.t - 60000, t1 = +(e.end || e.t) + 60000;
      var hit = media.filter(function (m) { return m.device.id === e.devId && +new Date(m.toDate || m.fromDate) >= t0 && +new Date(m.fromDate) <= t1; })[0];
      if (hit) e.media = { id: hit.id, type: hit.mediaType };
    });
  }
  function mediaUrl(m) {
    if (!session) return '';
    var cred = { userName: session.userName, database: session.database, sessionId: session.sessionId };
    return 'https://' + server + '/apiv1/DownloadMediaFile?mediaFile=' + encodeURIComponent(JSON.stringify({ id: m.id })) +
      '&credentials=' + encodeURIComponent(JSON.stringify(cred));
  }

  // 位置・速度（LogRecord）は一覧に表示する分だけ、まとめて取得
  var posLoading = false;
  function ensurePositions(rows) {
    if (demo || posLoading) return;
    var need = rows.filter(function (e) { return e.pos === undefined; }).slice(0, 100);
    if (!need.length) return;
    posLoading = true;
    var d = data;
    multi(need.map(function (e) {
      return ['Get', { typeName: 'LogRecord', search: { deviceSearch: { id: e.devId }, fromDate: new Date(+e.t - 30000).toISOString(), toDate: new Date(+e.t + 30000).toISOString() }, resultsLimit: 20 }];
    })).then(function (res) {
      need.forEach(function (e, i) {
        var best = null;
        (res[i] || []).forEach(function (l) {
          if (!l.latitude && !l.longitude) return;
          if (!best || Math.abs(new Date(l.dateTime) - e.t) < Math.abs(new Date(best.dateTime) - e.t)) best = l;
        });
        e.pos = best ? { lat: best.latitude, lon: best.longitude, spd: best.speed || 0 } : null;
      });
    }, function () { need.forEach(function (e) { e.pos = null; }); })
      .then(function () { posLoading = false; if (d === data) { renderList(); if (detailIdx >= 0) renderDetail(); } });
  }

  // ---------- デモデータ ----------
  function loadDemo() {
    var p = period(), now = Date.now();
    var seed = 11; function rnd() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
    var drvNames = ['佐藤 健一', '鈴木 大輔', '高橋 誠', '田中 美咲', '伊藤 翔太', '渡辺 浩二', '山本 拓也', '中村 優子', '小林 剛', '加藤 直樹', '吉田 亮', '山田 由美', '松本 隆', '井上 和也'];
    var devNames = ['品川 100 あ 12-34', '品川 100 あ 12-35', '練馬 400 か 56-01', '練馬 400 か 56-02', '横浜 130 さ 77-10', '横浜 130 さ 77-11', '名古屋 500 た 30-08', '名古屋 500 た 30-09', '大宮 100 な 41-22', '川崎 400 は 18-03', '多摩 300 ま 92-15', '千葉 100 や 65-40'];
    var devices = {}, drivers = {}, dist = { driver: {}, device: {} };
    devNames.forEach(function (n, i) { devices['b' + i] = n; });
    drvNames.forEach(function (n, i) { drivers['u' + i] = n; });
    // 運転手ごとの「くせ」（警報の出やすさ）
    var risk = [0.5, 2.4, 0.8, 0.3, 3.2, 1.0, 1.4, 0.4, 1.9, 0.7, 1.1, 0.6, 2.0, 0.9];
    var favor = [['brake'], ['phone', 'distract'], ['tailgate'], [], ['speed', 'distract', 'brake'], ['lane'], ['accel', 'brake'], [], ['tailgate', 'smoking'], ['corner'], ['seatbelt'], [], ['fatigue', 'lane'], ['stop']];
    var base = ['brake', 'brake', 'accel', 'corner', 'speed', 'speed', 'tailgate', 'lane', 'distract', 'stop', 'seatbelt', 'smoking', 'button', 'drinking'];
    var span = p.to - p.prevFrom, days = span / DAY, events = [], n = 0;
    var cities = [[35.6284, 139.7387], [35.7356, 139.6517], [35.4658, 139.6223], [35.1709, 136.8815], [35.9064, 139.6239], [35.5309, 139.7029], [35.6510, 139.4426], [35.6074, 140.1065]];
    drvNames.forEach(function (_, di) {
      var dev = 'b' + (di % devNames.length), city = cities[di % cities.length];
      var kmPerDay = 60 + rnd() * 90;
      var cur = Math.min(p.to - p.from, span) / DAY;
      dist.driver['u' + di] = kmPerDay * cur;
      dist.device[dev] = (dist.device[dev] || 0) + kmPerDay * cur;
      var count = Math.round(risk[di] * days * 0.45 * (0.8 + rnd() * 0.4));
      for (var k = 0; k < count; k++) {
        var t = +p.prevFrom + rnd() * span;
        var hour = 7 + Math.floor(rnd() * 12);                     // 7時〜19時に発生
        t = jstMidnight(t) + hour * 3600e3 + Math.floor(rnd() * 3600e3);
        if (t > now || t < +p.prevFrom || t > +p.to) continue;
        var pool = favor[di].length && rnd() < 0.55 ? favor[di] : base;
        var key = pool[Math.floor(rnd() * pool.length)];
        if (rnd() < 0.01) key = 'near';
        events.push({ id: 'e' + (n++), t: new Date(t), end: new Date(t + 4000 + rnd() * 8000), devId: dev, drvId: 'u' + di,
          type: TYPE_BY_KEY[key], ruleName: TYPE_BY_KEY[key].label + '（デモ）',
          pos: { lat: city[0] + (rnd() - 0.5) * 0.08, lon: city[1] + (rnd() - 0.5) * 0.1, spd: key === 'button' ? 0 : Math.round(15 + rnd() * 55) },
          media: { demo: true } });
      }
    });
    // 運転手未登録の走行
    dist.driver[UNKNOWN_DRIVER] = 140;
    events.push({ id: 'e' + (n++), t: new Date(now - 2 * DAY), devId: 'b11', drvId: UNKNOWN_DRIVER, type: TYPE_BY_KEY.brake, ruleName: '急ブレーキ（デモ）', pos: { lat: 35.61, lon: 140.11, spd: 42 }, media: { demo: true } });
    // デモでは確認済みの印を一部に付けておく
    if (!load('demoMarked', false)) {
      events.forEach(function (e, i) { if (+e.t < now - 5 * DAY && i % 3) marks[e.id] = { ack: true }; });
      save('marks', marks); save('demoMarked', true);
    }
    return Promise.resolve(build(events, devices, drivers, dist, p));
  }

  function build(events, devices, drivers, dist, p) {
    var cur = events.filter(function (e) { return e.t >= p.from && e.t <= p.to; });
    var prev = events.filter(function (e) { return e.t >= p.prevFrom && e.t < p.from; });
    cur.sort(function (a, b) { return b.t - a.t; });
    return { events: cur, prevEvents: prev, devices: devices, drivers: drivers, dist: dist, p: p };
  }

  // ---------- 集計 ----------
  function nameOf(by, id) {
    if (by === 'driver') return id === UNKNOWN_DRIVER ? '（運転手未登録）' : data.drivers[id] || id;
    return data.devices[id] || id;
  }
  function points(list) { return list.reduce(function (a, e) { return a + e.type.w; }, 0); }
  // 安全スコア：100kmあたりの「重み付き警報点数」から算出（距離が短すぎる場合は50km として扱う）
  function scoreOf(pts, km) { var r = pts / Math.max(km, 50) * 100; return Math.max(0, Math.round(100 - r * 8)); }
  function scoreCls(s) { return s >= 80 ? 's-good' : s >= 60 ? 's-mid' : 's-bad'; }

  function filtered(opts) {
    opts = opts || {};
    var q = ($('cs-filter').value || '').trim();
    return data.events.filter(function (e) {
      if (!opts.noType && view.type && e.type.key !== view.type) return false;
      if (!opts.noEntity && view.entity && (view.entity.by === 'driver' ? e.drvId : e.devId) !== view.entity.id) return false;
      if (!opts.noTodo && view.onlyTodo && (marks[e.id] || {}).ack) return false;
      if (q && (nameOf('driver', e.drvId) + ' ' + nameOf('device', e.devId)).indexOf(q) < 0) return false;
      return true;
    });
  }

  // ---------- 描画 ----------
  function render() {
    renderKpis(); renderTypes(); renderTrend(); renderRank(); renderList();
  }

  function renderKpis() {
    var ev = data.events, totalKm = 0;
    Object.keys(data.dist.device).forEach(function (k) { totalKm += data.dist.device[k]; });
    $('kpi-total').textContent = ev.length + '件';
    var pv = data.prevEvents.length, sub = $('kpi-total-sub');
    if (pv) {
      var ch = Math.round((ev.length - pv) / pv * 100);
      sub.textContent = '前の期間より ' + (ch > 0 ? '+' : '') + ch + '%（' + pv + '件）';
      sub.className = ch > 0 ? 'up' : ch < 0 ? 'down' : '';
    } else { sub.textContent = '前の期間：0件'; sub.className = ''; }
    $('kpi-rate').textContent = totalKm ? (ev.length / totalKm * 100).toFixed(2) + '件' : '-';
    $('kpi-rate-sub').textContent = '走行 ' + Math.round(totalKm).toLocaleString() + 'km';
    $('kpi-crit').textContent = ev.filter(function (e) { return e.type.sev === 'crit'; }).length + '件';
    var rows = rankRows('driver').filter(function (r) { return r.km > 0 || r.n > 0; });
    $('kpi-score').textContent = rows.length ? Math.round(rows.reduce(function (a, r) { return a + r.score; }, 0) / rows.length) + '点' : '-';
    $('kpi-todo').textContent = ev.filter(function (e) { return !(marks[e.id] || {}).ack; }).length + '件';
  }

  function renderTypes() {
    var list = filtered({ noType: true, noTodo: true }), cnt = {};
    list.forEach(function (e) { cnt[e.type.key] = (cnt[e.type.key] || 0) + 1; });
    var keys = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; });
    var max = keys.length ? cnt[keys[0]] : 1;
    $('cs-types').innerHTML = keys.map(function (k) {
      var t = TYPE_BY_KEY[k];
      return '<button type="button" class="cs-type' + (view.type === k ? ' on' : '') + '" data-type="' + k + '" title="' + esc(t.label) + 'の警報だけを一覧に表示">' +
        '<span class="name">' + esc(t.label) + '</span><span><span class="bar ' + t.sev + '" style="width:' + (cnt[k] / max * 100).toFixed(1) + '%;display:block"></span></span><b>' + cnt[k] + '</b></button>';
    }).join('') || '<div class="cs-empty">期間内の警報はありません</div>';
  }

  function renderTrend() {
    var p = data.p, list = filtered({ noTodo: true }), step = p.hourly ? 3600e3 : DAY;
    var t0 = p.hourly ? +p.from : jstMidnight(+p.from), t1 = +p.to, buckets = [];
    for (var t = t0; t < t1; t += step) buckets.push({ t: t, crit: 0, warn: 0, info: 0 });
    list.forEach(function (e) { var i = Math.floor((e.t - t0) / step); if (buckets[i]) buckets[i][e.type.sev]++; });
    var W = Math.max(320, Math.round($('cs-trend').clientWidth || 640)), H = 190, L = 30, R = 8, T = 10, B = 26, PW = W - L - R, PH = H - T - B;
    var max = Math.max(1, Math.max.apply(null, buckets.map(function (b) { return b.crit + b.warn + b.info; })));
    var stepY = max <= 5 ? 1 : max <= 10 ? 2 : max <= 25 ? 5 : max <= 50 ? 10 : Math.ceil(max / 50) * 10;
    max = Math.ceil(max / stepY) * stepY;
    var bw = PW / buckets.length, svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="警報件数の推移">';
    for (var g = 0; g <= max; g += stepY) {
      var y = T + PH - g / max * PH;
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#eef0f3"/><text x="' + (L - 4) + '" y="' + (y + 4) + '" font-size="10" text-anchor="end" fill="#6b7280">' + g + '</text>';
    }
    var labelEvery = Math.ceil(buckets.length / 10);
    buckets.forEach(function (b, i) {
      var x = L + i * bw + Math.max(1, bw * 0.15), w = Math.max(1, bw * 0.7), y = T + PH;
      ['info', 'warn', 'crit'].forEach(function (s) {
        if (!b[s]) return; var h = b[s] / max * PH; y -= h;
        svg += '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + h.toFixed(1) + '" fill="' + { crit: '#dc2626', warn: '#f59e0b', info: '#94a3b8' }[s] + '"><title>' + label(b.t) + '：' + ({ crit: '重大', warn: '注意', info: 'その他' })[s] + ' ' + b[s] + '件</title></rect>';
      });
      if (i % labelEvery === 0) svg += '<text x="' + (L + i * bw + bw / 2).toFixed(1) + '" y="' + (H - 8) + '" font-size="10" text-anchor="middle" fill="#6b7280">' + label(b.t) + '</text>';
    });
    function label(t) { var d = new Date(t + JST); return p.hourly ? d.getUTCHours() + '時' : (d.getUTCMonth() + 1) + '/' + d.getUTCDate(); }
    $('cs-trend').innerHTML = svg + '</svg>';
    $('cs-trend-note').textContent = (p.hourly ? '1時間ごと' : '1日ごと') + (view.type || view.entity ? '・絞り込み中' : '');
  }

  function rankRows(by) {
    var groups = {};
    data.events.forEach(function (e) { var id = by === 'driver' ? e.drvId : e.devId; (groups[id] = groups[id] || []).push(e); });
    var ids = Object.keys(groups);
    Object.keys(data.dist[by]).forEach(function (id) { if (!groups[id] && data.dist[by][id] > 1) ids.push(id); });
    return ids.map(function (id) {
      var list = groups[id] || [], km = data.dist[by][id] || 0, cnt = {};
      list.forEach(function (e) { if (e.type.w) cnt[e.type.key] = (cnt[e.type.key] || 0) + 1; });
      var top = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; }).slice(0, 2);
      return { id: id, name: nameOf(by, id), km: km, n: list.length, rate: km ? list.length / km * 100 : null, score: scoreOf(points(list), km),
        top: top.map(function (k) { return TYPE_BY_KEY[k].label + ' ' + cnt[k]; }).join('、'), short: km < 50 };
    }).sort(function (a, b) { return a.score - b.score || b.n - a.n; });
  }

  function renderRank() {
    var by = view.rankBy, rows = rankRows(by);
    $('cs-rank-name').textContent = by === 'driver' ? '運転手' : '車両';
    $('cs-rank').innerHTML = rows.map(function (r, i) {
      var on = view.entity && view.entity.by === by && view.entity.id === r.id;
      return '<tr data-id="' + esc(r.id) + '"' + (on ? ' class="on"' : '') + '><td><span class="cs-rankno">' + (rows.length - i) + '</span></td>' +
        '<td>' + esc(r.name) + '</td>' +
        '<td><span class="cs-score"><b class="' + scoreCls(r.score) + '">' + r.score + '</b><span class="track"><span class="fill ' + scoreCls(r.score) + '" style="width:' + r.score + '%;display:block"></span></span></span></td>' +
        '<td class="r">' + Math.round(r.km).toLocaleString() + 'km' + (r.short ? ' <span class="cs-muted" title="走行距離が短いため、50km として計算しています">※</span>' : '') + '</td>' +
        '<td class="r">' + r.n + '件</td><td class="r">' + (r.rate == null ? '-' : r.rate.toFixed(2)) + '</td><td>' + esc(r.top || '-') + '</td></tr>';
    }).join('') || '<tr><td colspan="7" class="cs-empty">データがありません</td></tr>';
  }

  function renderList() {
    listRows = filtered();
    var chips = [];
    if (view.type) chips.push('<button type="button" class="cs-chip" data-clear="type">' + esc(TYPE_BY_KEY[view.type].label) + '</button>');
    if (view.entity) chips.push('<button type="button" class="cs-chip" data-clear="entity">' + esc(nameOf(view.entity.by, view.entity.id)) + '</button>');
    $('cs-chips').innerHTML = chips.join('');
    $('cs-list-count').textContent = listRows.length + '件' + (listRows.length > 500 ? '（新しい500件を表示）' : '');
    var shown = listRows.slice(0, 500);
    $('cs-events').innerHTML = shown.map(function (e, i) {
      var m = marks[e.id] || {};
      return '<tr data-id="' + esc(e.id) + '" data-idx="' + i + '"' + (m.ack ? ' class="acked"' : '') + '>' +
        '<td><input type="checkbox" class="cs-ack" data-ack="' + esc(e.id) + '"' + (m.ack ? ' checked' : '') + ' title="確認済みにする"></td>' +
        '<td>' + fmtShort(e.t) + '</td><td><span class="cs-tag ' + e.type.sev + '" title="' + esc(e.ruleName) + '">' + esc(e.type.label) + '</span>' + (m.memo ? ' <span class="cs-muted" title="' + esc(m.memo) + '">📝</span>' : '') + '</td>' +
        '<td>' + esc(nameOf('driver', e.drvId)) + '</td><td>' + esc(nameOf('device', e.devId)) + '</td>' +
        '<td class="r">' + (e.pos === undefined ? '<span class="cs-muted">…</span>' : e.pos ? Math.round(e.pos.spd) + 'km/h' : '-') + '</td>' +
        '<td><button type="button" class="cs-video-btn' + (e.media ? ' has' : '') + '" data-open="' + i + '">' + (e.media ? '▶ 映像' : '詳細') + '</button></td></tr>';
    }).join('') || '<tr><td colspan="7" class="cs-empty">' + (view.onlyTodo ? '未確認の警報はありません' : '該当する警報はありません') + '</td></tr>';
    ensurePositions(shown);
  }

  // ---------- 詳細（映像・地図） ----------
  // MyGeotabの左メニュー・上部バーに隠れないよう、Add-Inが見えている範囲に合わせて配置
  function scrollParent(el) {
    for (var p = el.parentElement; p && p !== document.body && p !== document.documentElement; p = p.parentElement) {
      var oy = getComputedStyle(p).overflowY;
      if ((oy === 'auto' || oy === 'scroll') && p.clientHeight > 0) return p;
    }
    return null;
  }
  function placeModal(el) {
    var vw = window.innerWidth, vh = window.innerHeight, root = $('cameraSafety');
    var rr = root.getBoundingClientRect(), sp = scrollParent(root), ar = sp ? sp.getBoundingClientRect() : { left: 0, top: 0, right: vw, bottom: vh };
    var left = Math.max(0, ar.left, rr.left), right = Math.min(vw, ar.right, rr.right), top = Math.max(0, ar.top), bottom = Math.min(vh, ar.bottom);
    if (right - left < 300 || bottom - top < 200) { left = 0; top = 0; right = vw; bottom = vh; }
    el.style.left = left + 'px'; el.style.top = top + 'px'; el.style.width = (right - left) + 'px'; el.style.height = (bottom - top) + 'px';
    el.style.right = 'auto'; el.style.bottom = 'auto';
  }
  function openDetail(i) { detailIdx = i; var m = $('cs-detail'); m.hidden = false; placeModal(m); renderDetail(); }
  function closeDetail() { detailIdx = -1; $('cs-detail').hidden = true; $('cs-d-video').innerHTML = ''; $('cs-d-map').innerHTML = ''; }

  var addrCache = {};
  function renderDetail() {
    var e = listRows[detailIdx]; if (!e) { closeDetail(); return; }
    var m = marks[e.id] || {};
    $('cs-d-title').innerHTML = '<span class="cs-tag ' + e.type.sev + '">' + esc(e.type.label) + '</span> ' + esc(nameOf('driver', e.drvId)) + ' <span class="cs-muted">' + fmtT(e.t) + '</span>';
    $('cs-d-prev').disabled = detailIdx <= 0; $('cs-d-next').disabled = detailIdx >= Math.min(listRows.length, 500) - 1;
    // 映像
    var v = $('cs-d-video');
    if (e.media && e.media.demo) v.innerHTML = demoVideo(e);
    else if (e.media && e.media.id) v.innerHTML = e.media.type === 'Image' ? '<img alt="" style="width:100%;height:100%;object-fit:contain" src="' + esc(mediaUrl(e.media)) + '">'
      : '<video controls autoplay muted src="' + esc(mediaUrl(e.media)) + '"></video>';
    else v.innerHTML = '<div class="cs-d-novideo"><p>この警報の映像は、MyGeotabの「ビデオ」画面で確認できます。</p>' +
      (settings.videoHash ? '<button type="button" class="cs-btn" id="cs-d-goto">MyGeotabのビデオ画面を開く</button><p>' + esc(nameOf('device', e.devId)) + '・' + fmtT(e.t) + ' の映像を探してください</p>'
        : '<p>「設定 → 映像を開くページ」にビデオ画面のアドレスを登録すると、ここから開けます。</p>') + '</div>';
    // 地図
    var mp = $('cs-d-map');
    if (e.pos) {
      var la = e.pos.lat, lo = e.pos.lon, d = 0.006;
      var key = la.toFixed(5) + ',' + lo.toFixed(5);
      if (mp.dataset.key !== key) {
        mp.dataset.key = key;
        mp.innerHTML = '<iframe title="発生場所の地図" loading="lazy" src="https://www.openstreetmap.org/export/embed.html?bbox=' + (lo - d) + ',' + (la - d * 0.7) + ',' + (lo + d) + ',' + (la + d * 0.7) + '&layer=mapnik&marker=' + la + ',' + lo + '"></iframe>' +
          '<a class="cs-d-maplink" target="_blank" rel="noopener" href="https://www.google.com/maps?q=' + la + ',' + lo + '">Googleマップで開く</a>';
      }
    } else { mp.dataset.key = ''; mp.innerHTML = '<div class="cs-d-novideo" style="color:#6b7280"><p style="color:#6b7280">' + (e.pos === undefined ? '位置を取得中…' : '位置データがありません') + '</p></div>'; }
    // 情報
    var addr = e.pos ? addressOf(e.pos) : null;
    $('cs-d-info').innerHTML = [
      ['日時', fmtT(e.t)], ['運転手', esc(nameOf('driver', e.drvId))], ['車両', esc(nameOf('device', e.devId))],
      ['速度', e.pos ? Math.round(e.pos.spd) + ' km/h' : '-'], ['場所', e.pos ? esc(addr == null ? '住所を取得中…' : addr || '-') : '-'],
      ['ルール名', esc(e.ruleName)], ['続いた時間', e.end ? Math.max(1, Math.round((e.end - e.t) / 1000)) + '秒' : '-']
    ].map(function (x) { return '<div><span>' + x[0] + '</span><strong>' + x[1] + '</strong></div>'; }).join('');
    $('cs-d-ack').checked = !!m.ack; $('cs-d-memo').value = m.memo || '';
    var g = $('cs-d-goto'); if (g) g.onclick = function () { gotoVideo(e); };
  }
  function addressOf(pos) {
    var key = pos.lat.toFixed(4) + ',' + pos.lon.toFixed(4);
    if (addrCache[key] !== undefined) return addrCache[key];
    if (demo) return (addrCache[key] = ['東京都品川区', '東京都練馬区', '神奈川県横浜市', '愛知県名古屋市', '埼玉県さいたま市', '神奈川県川崎市', '東京都多摩市', '千葉県千葉市'][Math.abs(Math.round(pos.lat * 100 + pos.lon * 10)) % 8] + '（デモ）');
    addrCache[key] = null;
    call('GetAddresses', { coordinates: [{ x: pos.lon, y: pos.lat }] }).then(function (r) {
      var a = r && r[0]; addrCache[key] = a ? String(a.formattedAddress || '').replace(/,?\s*(Japan|日本)\s*$/i, '').replace(/〒?\s*\d{3}-\d{4}\s*/, '') : '';
    }, function () { addrCache[key] = ''; }).then(function () { if (detailIdx >= 0) renderDetail(); });
    return null;
  }
  function gotoVideo(e) {
    var h = String(settings.videoHash || '').replace(/^.*#/, '');
    if (!h) return;
    var page = h.split(',')[0];
    if (state && state.gotoPage && h.indexOf(',') < 0) { closeDetail(); state.gotoPage(page); return; }
    window.open('https://' + server + '/' + (session ? session.database : '') + '/#' + h, '_blank');
  }

  // デモ用の映像の代わり（ドラレコ風のイラスト）
  function demoVideo(e) {
    return '<svg viewBox="0 0 320 200" role="img" aria-label="デモ映像">' +
      '<rect width="320" height="200" fill="#1e293b"/><rect y="0" width="320" height="90" fill="#334155"/>' +
      '<polygon points="120,90 200,90 300,200 20,200" fill="#475569"/>' +
      '<line x1="160" y1="92" x2="160" y2="200" stroke="#fde68a" stroke-width="3" stroke-dasharray="14 12"><animate attributeName="stroke-dashoffset" from="0" to="-26" dur=".6s" repeatCount="indefinite"/></line>' +
      (e.type.key === 'tailgate' || e.type.key === 'near' || e.type.key === 'brake' ? '<rect x="135" y="98" width="50" height="34" rx="4" fill="#94a3b8"><animate attributeName="width" values="50;66;50" dur="2s" repeatCount="indefinite"/></rect>' : '') +
      '<text x="8" y="16" font-size="11" fill="#fff">' + esc(fmtT(e.t)) + '</text>' +
      '<text x="312" y="16" font-size="11" fill="#fff" text-anchor="end">' + (e.pos ? Math.round(e.pos.spd) : 0) + ' km/h</text>' +
      '<rect x="8" y="170" width="' + (e.type.label.length * 12 + 16) + '" height="22" rx="4" fill="' + (e.type.sev === 'crit' ? '#dc2626' : e.type.sev === 'warn' ? '#d97706' : '#64748b') + '"/>' +
      '<text x="16" y="185" font-size="12" fill="#fff">' + esc(e.type.label) + '</text>' +
      '<text x="160" y="40" font-size="10" fill="#cbd5e1" text-anchor="middle">デモ映像（実際の映像ではありません）</text></svg>';
  }

  // ---------- 確認済み・メモ ----------
  function setMark(id, patch) {
    var m = Object.assign({}, marks[id] || {}, patch);
    if (!m.ack && !m.memo) delete marks[id]; else marks[id] = m;
    save('marks', marks);
  }

  // ---------- CSV ----------
  function exportCsv() {
    var rows = [['日時', '種類', 'ルール名', '運転手', '車両', '速度(km/h)', '緯度', '経度', '確認済み', 'メモ']];
    listRows.forEach(function (e) {
      var m = marks[e.id] || {};
      rows.push([fmtT(e.t), e.type.label, e.ruleName, nameOf('driver', e.drvId), nameOf('device', e.devId),
        e.pos ? Math.round(e.pos.spd) : '', e.pos ? e.pos.lat : '', e.pos ? e.pos.lon : '', m.ack ? '済' : '', m.memo || '']);
    });
    var csv = '﻿' + rows.map(function (r) { return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    a.download = '安全運転_警報一覧.csv'; document.body.appendChild(a); a.click(); a.remove();
  }

  // ---------- 更新 ----------
  function refresh() {
    if (loading) return;
    loading = true; showMsg(''); $('cs-refresh').disabled = true; $('cs-refresh').textContent = '読み込み中…';
    (demo ? loadDemo() : loadLive()).then(function (d) {
      data = d;
      if (view.entity && !rankRows(view.entity.by).some(function (r) { return r.id === view.entity.id; })) view.entity = null;
      render();
      if (!demo && !d.events.length) showMsg('期間内に該当する警報はありません。対象のルールは「設定」で確認できます。', true);
    }, function (e) { showMsg((e && e.info ? '' : 'データの取得に失敗しました：') + (e && e.message || e), e && e.info); })
      .then(function () { loading = false; $('cs-refresh').disabled = false; $('cs-refresh').textContent = '更新'; });
  }

  // ---------- 設定 ----------
  var draftRules = [], allRules = null;
  function renderRuleList(list) {
    var sel = {}; draftRules.forEach(function (r) { sel[r.id] = 1; });
    var shown = list.slice().sort(function (a, b) { return (sel[b.id] || 0) - (sel[a.id] || 0); });
    $('set-rule-list').innerHTML = shown.map(function (r) {
      var t = typeOfRule(r);
      return '<label><input type="checkbox" data-rule="' + esc(r.id) + '"' + (sel[r.id] ? ' checked' : '') + '> ' + esc(r.name) +
        ' <small>→ ' + esc(t.label) + '</small></label>';
    }).join('') || '<div class="cs-muted">該当するルールがありません</div>';
  }
  function demoRules() { return TYPES.map(function (t) { return { id: 'r-' + t.key, name: t.label + '（デモ）' }; }); }
  function openSettings() {
    draftRules = (settings.rules || []).slice();
    $('set-video-hash').value = settings.videoHash || '';
    $('set-check-out').hidden = true;
    $('cs-settings').hidden = false; placeModal($('cs-settings'));
    var ready = allRules ? Promise.resolve(allRules) : demo ? Promise.resolve(demoRules()) : loadRules();
    $('set-rule-list').innerHTML = '<div class="cs-muted">読み込み中…</div>';
    ready.then(function (l) {
      allRules = l;
      if (demo && !draftRules.length) draftRules = l.slice();
      renderRuleList(l.filter(function (r) { return isTarget(r) || draftRules.some(function (d) { return d.id === r.id; }); }));
    }, function (e) { $('set-rule-list').innerHTML = '<div class="cs-muted">ルールを読み込めませんでした：' + esc(e && e.message || e) + '</div>'; });
  }
  function syncDraft() {
    $('set-rule-list').querySelectorAll('input[data-rule]').forEach(function (cb) {
      var id = cb.dataset.rule, i = -1;
      draftRules.forEach(function (r, k) { if (r.id === id) i = k; });
      if (cb.checked && i < 0) { var r = (allRules || []).filter(function (x) { return x.id === id; })[0]; if (r) draftRules.push({ id: r.id, name: r.name }); }
      if (!cb.checked && i >= 0) draftRules.splice(i, 1);
    });
  }
  function searchRules() {
    syncDraft();
    var q = ($('set-rule-query').value || '').trim().toLowerCase();
    renderRuleList((allRules || []).filter(function (r) { return !q || r.name.toLowerCase().indexOf(q) >= 0; }));
  }
  function saveSettingsForm() {
    syncDraft();
    if (!demo) settings.rules = draftRules;
    settings.videoHash = ($('set-video-hash').value || '').trim();
    save('settings', settings); $('cs-settings').hidden = true; refresh();
  }
  // 接続チェック：このデータベースで何が取れるかを表示（導入時の確認用）
  function runCheck() {
    var out = $('set-check-out'); out.hidden = false; out.textContent = '確認中…';
    if (demo) { out.textContent = 'デモモードのため確認できません。MyGeotabに登録してから実行してください。'; return; }
    var to = new Date(), from = new Date(to - 7 * DAY), lines = [];
    function line(s) { lines.push(s); out.textContent = lines.join('\n'); }
    loadRules().then(function (rules) {
      var tgt = rules.filter(isTarget);
      line('ルール：全' + rules.length + '件、うち警報らしいもの ' + tgt.length + '件');
      tgt.slice(0, 40).forEach(function (r) { line('  ・' + r.name + '  → ' + typeOfRule(r).label); });
      var cams = rules.filter(function (r) { return CAMERA_RE.test(r.name); });
      line('カメラ関連の名前のルール：' + cams.length + '件');
    }).then(function () {
      return call('Get', { typeName: 'MediaFile', search: { fromDate: from.toISOString(), toDate: to.toISOString() }, resultsLimit: 100 })
        .then(function (l) { line('MediaFile（映像ファイル・直近7日）：' + l.length + '件'); }, function (e) { line('MediaFile：取得できません（' + (e && e.message || e) + '）'); });
    }).then(function () {
      return call('Get', { typeName: 'CameraEvent', search: { fromDate: from.toISOString(), toDate: to.toISOString() }, resultsLimit: 5 })
        .then(function (l) { line('CameraEvent（直近7日）：' + l.length + '件' + (l[0] ? '\n  例：' + JSON.stringify(l[0]).slice(0, 400) : '')); },
          function (e) { line('CameraEvent：この方法では取得できません（' + (e && e.message || e).toString().slice(0, 120) + '）'); });
    }).then(function () { line('\n確認が終わりました。この内容をコピーして開発担当に送ってください。'); });
  }

  // ---------- 操作 ----------
  function bind() {
    $('cs-refresh').onclick = refresh;
    $('cs-days').onchange = function () { save('days', this.value); refresh(); };
    $('cs-filter').oninput = function () { if (data) { renderTypes(); renderTrend(); renderList(); } };
    $('cs-only-todo').onchange = function () { view.onlyTodo = this.checked; if (data) renderList(); };
    $('cs-csv').onclick = exportCsv;
    $('cs-types').onclick = function (e) {
      var b = e.target.closest('[data-type]'); if (!b) return;
      view.type = view.type === b.dataset.type ? null : b.dataset.type; render();
    };
    $('cs-rank-by').onclick = function (e) {
      var b = e.target.closest('button[data-by]'); if (!b) return;
      view.rankBy = b.dataset.by;
      this.querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
      renderRank();
    };
    $('cs-rank').onclick = function (e) {
      var tr = e.target.closest('tr[data-id]'); if (!tr) return;
      var id = tr.dataset.id, by = view.rankBy;
      view.entity = view.entity && view.entity.by === by && view.entity.id === id ? null : { by: by, id: id };
      render();
    };
    $('cs-chips').onclick = function (e) {
      var b = e.target.closest('[data-clear]'); if (!b) return;
      view[b.dataset.clear] = null; render();
    };
    $('cs-events').onclick = function (e) {
      var cb = e.target.closest('[data-ack]');
      if (cb) { setMark(cb.dataset.ack, { ack: cb.checked }); renderKpis(); cb.closest('tr').classList.toggle('acked', cb.checked); return; }
      var tr = e.target.closest('tr[data-idx]'); if (tr) openDetail(+tr.dataset.idx);
    };
    $('cs-d-close').onclick = closeDetail;
    $('cs-d-prev').onclick = function () { if (detailIdx > 0) { detailIdx--; renderDetail(); } };
    $('cs-d-next').onclick = function () { if (detailIdx < Math.min(listRows.length, 500) - 1) { detailIdx++; renderDetail(); } };
    $('cs-detail').onclick = function (e) { if (e.target === this) closeDetail(); };
    $('cs-d-ack').onchange = function () {
      var ev = listRows[detailIdx]; if (!ev) return;
      setMark(ev.id, { ack: this.checked }); renderKpis(); renderList();
    };
    $('cs-d-memo').onchange = function () {
      var ev = listRows[detailIdx]; if (!ev) return;
      setMark(ev.id, { memo: this.value.trim() }); renderList();
    };
    document.addEventListener('keydown', function (e) {
      if ($('cs-detail').hidden || e.target.tagName === 'INPUT') { if (e.key === 'Escape') { closeDetail(); $('cs-settings').hidden = true; } return; }
      if (e.key === 'Escape') closeDetail();
      else if (e.key === 'ArrowLeft') $('cs-d-prev').click();
      else if (e.key === 'ArrowRight') $('cs-d-next').click();
    });
    window.addEventListener('resize', function () {
      if (!$('cs-detail').hidden) placeModal($('cs-detail'));
      if (!$('cs-settings').hidden) placeModal($('cs-settings'));
      clearTimeout(rzTimer); rzTimer = setTimeout(function () { if (data) renderTrend(); }, 150);
    });
    $('cs-open-settings').onclick = openSettings;
    $('set-cancel').onclick = function () { $('cs-settings').hidden = true; };
    $('set-save').onclick = saveSettingsForm;
    $('set-rule-search').onclick = searchRules;
    $('set-rule-query').onkeydown = function (e) { if (e.key === 'Enter') searchRules(); };
    $('set-rule-auto').onclick = function () {
      draftRules = (allRules || []).filter(isTarget).map(function (r) { return { id: r.id, name: r.name }; });
      renderRuleList((allRules || []).filter(isTarget));
    };
    $('set-check').onclick = runCheck;
  }

  var bound = false, rzTimer = null;
  function start(a, isDemo) {
    api = a; demo = isDemo;
    var badge = $('cs-source'); badge.textContent = (demo ? 'デモ' : 'ライブ') + ' v' + VERSION; badge.className = 'cs-badge' + (demo ? ' demo' : '');
    var ready = demo ? Promise.resolve() : new Promise(function (res) {
      api.getSession(function (s, srv) {
        session = s; server = srv || window.location.hostname || 'my.geotab.com';
        storageKey = 'cameraSafety:' + s.database; res();
      });
    });
    return ready.then(function () {
      settings = Object.assign({ rules: null, videoHash: '' }, load('settings', {}));
      marks = load('marks', {});
      var d = load('days', null); if (d) $('cs-days').value = d;
      if (!bound) { bind(); bound = true; }
    });
  }

  return {
    initialize: function (a, st, callback) { state = st; start(a, false).then(function () { callback(); }, function () { callback(); }); },
    focus: function (a, st) { api = a; if (st) state = st; refresh(); },
    blur: function () { closeDetail(); },
    _standalone: function () { start(null, true).then(refresh); }
  };
};

// MyGeotab外（ローカルでHTMLを開いた場合）はデモモードで起動。
// MyGeotab内では initialize が呼ばれるので、2秒待っても呼ばれなければデモとみなす。
(function () {
  var orig = geotab.addin.cameraSafety;
  geotab.addin.cameraSafety = function () {
    var inst = orig();
    var init = inst.initialize;
    inst.initialize = function () { window.__cameraSafetyInit = true; return init.apply(inst, arguments); };
    return inst;
  };
  window.addEventListener('load', function () {
    setTimeout(function () { if (!window.__cameraSafetyInit) orig()._standalone(); }, 2000);
  });
})();
