/*
 * 安全運転ダッシュボード MyGeotab Add-In
 * GO Focus Plus（カメラ）の警報と、急ブレーキ・急加速などの運転ルールの例外イベント（ExceptionEvent）を集計し、
 * 地図（車両全体の走行軌跡＋警報のヒートマップ）をメインに、左の列で絞り込み・件数・ランキング・警報一覧を表示する。
 * MyGeotab 外（ローカルでファイルを開いた場合）はデモデータで動作する。
 */
/* global geotab */
var geotab = window.geotab || { addin: {} };
geotab.addin = geotab.addin || {};

geotab.addin.cameraSafety = function () {
  'use strict';

  var VERSION = '0.3.3';

  // 警報の種類。ルール名（日本語・英語）から判定する。上から順に判定。
  // sev: crit=重大 / warn=注意 / info=その他、w: 安全スコアの重み（0はスコアに影響しない）
  var TYPES = [
    // カメラ機器の状態（遮られている・カバーが外れた・向きの調整が必要など）。運転の危険ではないため初期設定では集計しない
    { key: 'camera', label: 'カメラの状態', sev: 'info', w: 0, re: /obstruct|cover\s*removed|adjustment\s*required|camera\s*(blocked|covered|tamper)|遮|調整が必要/i },
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
  var view = { rankBy: 'driver', onlyTodo: false };
  var listRows = [], detailList = [], detailIdx = -1;
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
  function isTarget(rule) {
    var t = typeOfRule(rule);
    if (t.key === 'camera') return false;
    return !!KNOWN_RULES[rule.id] || t !== OTHER || CAMERA_RE.test(rule.name || '');
  }
  var AUTO_VER = 2;   // 自動選択の基準を変えたら上げる（利用者が設定で選んでいない場合は選び直す）
  function load(key, def) { try { var v = JSON.parse(localStorage.getItem(storageKey + ':' + key)); return v == null ? def : v; } catch (e) { return def; } }
  function save(key, v) { try { localStorage.setItem(storageKey + ':' + key, JSON.stringify(v)); } catch (e) { /* ignore */ } }
  // 地図の座標（Webメルカトル、0〜1に正規化）
  function mercX(lon) { return (lon + 180) / 360; }
  function mercY(lat) { var s = Math.sin(lat * Math.PI / 180); return 0.5 - Math.log((1 + s) / (1 - s)) / (4 * Math.PI); }

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
  function driverId(d) { return d && d.id && d.id !== 'NoDriverId' ? d.id : UNKNOWN_DRIVER; }

  function loadLive() {
    var p = period();
    return loadRules().then(function (rules) {
      var ruleById = {}; rules.forEach(function (r) { ruleById[r.id] = r; });
      if (!settings.rules || (!settings.manual && settings.autoVer !== AUTO_VER)) {   // 初回（または自動選択の基準が変わったとき）は自動選択
        settings.rules = rules.filter(isTarget).map(function (r) { return { id: r.id, name: r.name }; });
        settings.autoVer = AUTO_VER;
        save('settings', settings);
      }
      var sel = settings.rules.filter(function (r) { return ruleById[r.id]; });
      if (!sel.length) { var e = new Error('対象の警報ルールが見つかりませんでした。「設定」から集計するルールを選んでください。'); e.info = true; throw e; }
      var groups = state && state.getGroupFilter ? state.getGroupFilter() : null;
      var calls = [
        ['Get', { typeName: 'Device', search: groups && groups.length ? { groups: groups } : {}, resultsLimit: 5000 }],
        ['Get', { typeName: 'User', search: { isDriver: true }, resultsLimit: 5000 }],
        ['Get', { typeName: 'Trip', search: { fromDate: p.from.toISOString(), toDate: p.to.toISOString() }, resultsLimit: 50000 }],
        ['Get', { typeName: 'Group', resultsLimit: 10000 }]
      ];
      sel.forEach(function (r) {
        calls.push(['Get', { typeName: 'ExceptionEvent', search: { ruleSearch: { id: r.id }, fromDate: p.prevFrom.toISOString(), toDate: p.to.toISOString() }, resultsLimit: 50000 }]);
      });
      return Promise.all([multi(calls), loadMedia(p)]).then(function (res) {
        var r = res[0], media = res[1];
        var devices = {}, drivers = {}, trips = [], devRaw = {};
        r[0].forEach(function (d) { devices[d.id] = d.name || d.id; devRaw[d.id] = (d.groups || []).map(function (g) { return g.id; }); });
        r[1].forEach(function (u) { drivers[u.id] = ((u.lastName || '') + ' ' + (u.firstName || '')).trim() || u.name || u.id; });
        r[2].forEach(function (t) {
          var dv = t.device && t.device.id; if (!devices[dv]) return;
          trips.push({ devId: dv, drvId: driverId(t.driver), km: t.distance || 0, start: +new Date(t.start), stop: +new Date(t.stop || t.start) });
        });
        var events = [];
        sel.forEach(function (rule, i) {
          var type = typeOfRule(ruleById[rule.id]);
          (r[4 + i] || []).forEach(function (x) {
            var dv = x.device && x.device.id; if (!devices[dv]) return;
            events.push({ id: x.id, t: new Date(x.activeFrom), end: x.activeTo ? new Date(x.activeTo) : null, devId: dv,
              drvId: driverId(x.driver), type: type, ruleName: ruleById[rule.id].name, pos: undefined });
          });
        });
        events = dedupe(events);
        attachMedia(events, media);
        var d = build(events, devices, drivers, trips, p);
        setGroups(d, (r[3] || []).map(function (g) { return { id: g.id, name: g.name, children: (g.children || []).map(function (c) { return c.id; }) }; }), devRaw);
        return d;
      });
    });
  }

  // 同じ出来事が新旧・複数のルールで重ねて記録される場合（例：Harsh Braking と Harsh Braking (New)）、
  // 同じ車両・同じ種類で15秒以内の警報は1件にまとめる（ルール名は「／」でつなぐ）
  var DEDUPE_MS = 15000;
  function dedupe(events) {
    events.sort(function (a, b) { return a.devId < b.devId ? -1 : a.devId > b.devId ? 1 : a.type.key < b.type.key ? -1 : a.type.key > b.type.key ? 1 : a.t - b.t; });
    var out = [], last = null;
    events.forEach(function (e) {
      if (last && last.devId === e.devId && last.type === e.type && e.t - last.t <= DEDUPE_MS) {
        if (last.ruleName.indexOf(e.ruleName) < 0) last.ruleName += '／' + e.ruleName;
        if (e.end && (!last.end || e.end > last.end)) last.end = e.end;
        if (last.drvId === UNKNOWN_DRIVER) last.drvId = e.drvId;
        return;
      }
      out.push(e); last = e;
    });
    return out;
  }

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

  // ---------- 走行軌跡（車両ごとに位置記録 LogRecord をまとめて取得） ----------
  // 軌跡は運転記録（Trip）ごとの線に分け、運転手で絞り込めるようにする。約40m以上離れた点だけ残して軽くする。
  var THIN_M = 40;
  function fetchLogs(devId, from, to, acc) {
    acc = acc || [];
    return call('Get', { typeName: 'LogRecord', search: { deviceSearch: { id: devId }, fromDate: new Date(from).toISOString(), toDate: new Date(to).toISOString() }, resultsLimit: 50000 })
      .then(function (list) {
        list = list || [];
        list.forEach(function (l) { if (l.latitude || l.longitude) acc.push({ t: +new Date(l.dateTime), lat: l.latitude, lon: l.longitude, spd: l.speed || 0 }); });
        if (list.length >= 50000) {                // 上限に達したら続きを取得
          var last = +new Date(list[list.length - 1].dateTime);
          if (last > from && last < to) return fetchLogs(devId, last + 1, to, acc);
        }
        return acc;
      });
  }
  function loadTracksLive(d) {
    var devs = {}; d.trips.forEach(function (t) { devs[t.devId] = 1; });
    var ids = Object.keys(devs), i = 0;
    d.trk = { done: 0, total: ids.length, loading: ids.length > 0 };
    hmStatus();
    function next() {
      if (d !== data) return null;
      if (i >= ids.length) return null;
      var dv = ids[i++];
      return fetchLogs(dv, +d.p.from, +d.p.to).then(function (pts) {
        pts.sort(function (a, b) { return a.t - b.t; });
        addTrackSegments(d, dv, pts);
        positionsFromTrack(d, dv, pts);
      }, function () { /* この車両は読めなかった */ }).then(function () {
        d.trk.done++;
        if (d === data) { trackChanged(); }
        return next();
      });
    }
    Promise.all([next(), next(), next()]).then(function () {
      if (d !== data) return;
      d.trk.loading = false;
      // 軌跡で位置がわからなかった警報は、個別に位置を取得
      ensurePositions(filtered().slice(0, 3000));
      trackChanged(true);
    });
  }
  function addTrackSegments(d, dv, pts) {
    var trips = d.trips.filter(function (t) { return t.devId === dv; }).sort(function (a, b) { return a.start - b.start; });
    var k = 0;
    trips.forEach(function (tr) {
      while (k < pts.length && pts[k].t < tr.start - 60000) k++;
      var xs = [], ys = [], lastLat = null, lastLon = null;
      for (var j = k; j < pts.length && pts[j].t <= tr.stop + 60000; j++) {
        var p = pts[j];
        if (lastLat != null) {
          var dy = (p.lat - lastLat) * 111000, dx = (p.lon - lastLon) * 91000;
          if (dx * dx + dy * dy < THIN_M * THIN_M) continue;
        }
        xs.push(mercX(p.lon)); ys.push(mercY(p.lat)); lastLat = p.lat; lastLon = p.lon;
      }
      if (xs.length > 1) d.tracks.push({ devId: dv, drvId: tr.drvId, x: new Float64Array(xs), y: new Float64Array(ys), n: 1 });
    });
  }
  // 位置記録から、その車両の警報が起きた場所・速度を求める（前後2分以内で一番近い記録）
  function positionsFromTrack(d, dv, pts) {
    if (!pts.length) return;
    d.events.concat(d.prevEvents).forEach(function (e) {
      if (e.devId !== dv || e.pos) return;
      var t = +e.t, lo = 0, hi = pts.length - 1;
      while (hi - lo > 1) { var m = (lo + hi) >> 1; if (pts[m].t <= t) lo = m; else hi = m; }
      var best = Math.abs(pts[lo].t - t) <= Math.abs(pts[hi].t - t) ? pts[lo] : pts[hi];
      if (Math.abs(best.t - t) <= 120000) e.pos = { lat: best.lat, lon: best.lon, spd: best.spd };
    });
  }
  var trackTimer = null;
  function trackChanged(final) {
    clearTimeout(trackTimer);
    trackTimer = setTimeout(function () {
      if (!hm.userMoved) fitMap();
      renderMapSide(); renderList(); requestDraw(); hmStatus();
    }, final ? 0 : 400);
  }

  // 位置・速度（LogRecord）を警報ごとに取得（一覧に表示する分を先に、100件ずつ順番に）
  var posBusy = false, posQueue = [];
  function ensurePositions(rows) {
    if (demo) return;
    rows.forEach(function (e) { if (e.pos === undefined && !e.queued) { e.queued = true; posQueue.push(e); } });
    pumpPositions();
  }
  function pumpPositions() {
    hmStatus();
    if (posBusy || !posQueue.length) return;
    posBusy = true;
    var d = data, need = posQueue.splice(0, 100).filter(function (e) { return e.pos === undefined; });
    if (!need.length) { posBusy = false; pumpPositions(); return; }
    multi(need.map(function (e) {
      return ['Get', { typeName: 'LogRecord', search: { deviceSearch: { id: e.devId }, fromDate: new Date(+e.t - 30000).toISOString(), toDate: new Date(+e.t + 30000).toISOString() }, resultsLimit: 20 }];
    })).then(function (res) {
      need.forEach(function (e, i) {
        if (e.pos) return;
        var best = null;
        (res[i] || []).forEach(function (l) {
          if (!l.latitude && !l.longitude) return;
          if (!best || Math.abs(new Date(l.dateTime) - e.t) < Math.abs(new Date(best.dateTime) - e.t)) best = l;
        });
        e.pos = best ? { lat: best.latitude, lon: best.longitude, spd: best.speed || 0 } : null;
      });
    }, function () { need.forEach(function (e) { if (!e.pos) e.pos = null; }); })
      .then(function () {
        posBusy = false;
        if (d === data) { renderList(); renderMapSide(); requestDraw(); if (detailIdx >= 0) renderDetail(); }
        pumpPositions();
      });
  }

  // ---------- デモデータ ----------
  function loadDemo() {
    var p = period(), now = Date.now();
    var seed = 11; function rnd() { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; }
    var drvNames = ['佐藤 健一', '鈴木 大輔', '高橋 誠', '田中 美咲', '伊藤 翔太', '渡辺 浩二', '山本 拓也', '中村 優子', '小林 剛', '加藤 直樹', '吉田 亮', '山田 由美', '松本 隆', '井上 和也'];
    var devNames = ['品川 100 あ 12-34', '品川 100 あ 12-35', '練馬 400 か 56-01', '練馬 400 か 56-02', '横浜 130 さ 77-10', '横浜 130 さ 77-11', '名古屋 500 た 30-08', '名古屋 500 た 30-09', '大宮 100 な 41-22', '川崎 400 は 18-03', '多摩 300 ま 92-15', '千葉 100 や 65-40'];
    var devices = {}, drivers = {}, trips = [], tracks = [];
    devNames.forEach(function (n, i) { devices['b' + i] = n; });
    drvNames.forEach(function (n, i) { drivers['u' + i] = n; });
    var cities = [[35.6284, 139.7387], [35.7356, 139.6517], [35.4658, 139.6223], [35.1709, 136.8815], [35.9064, 139.6239], [35.5309, 139.7029], [35.6510, 139.4426], [35.6074, 140.1065]];
    // 警報が集中しやすい交差点など（デモ用）。各都市に3か所
    var hot = cities.map(function (c) { return [0, 1, 2].map(function () { return [c[0] + (rnd() - 0.5) * 0.05, c[1] + (rnd() - 0.5) * 0.06]; }); });
    var span = p.to - p.prevFrom, days = span / DAY, curDays = (p.to - p.from) / DAY;
    // 車両ごとの決まったルート（デモ用の軌跡）。碁盤の目のように縦横に曲がりながら、交差点（警報が多い場所）を通る
    var routes = devNames.map(function (_, vi) {
      var c = cities[vi % cities.length], h = hot[vi % cities.length];
      var depot = [c[0] + (rnd() - 0.5) * 0.02, c[1] + (rnd() - 0.5) * 0.02];
      return [0, 1, 2, 3].map(function (k) {
        var way = [depot, h[k % 3], [c[0] + (rnd() - 0.5) * 0.12, c[1] + (rnd() - 0.5) * 0.14], h[(k + 1) % 3], depot], pts = [];
        for (var i = 0; i < way.length - 1; i++) {
          var a = way[i], b = way[i + 1], corner = rnd() < 0.5 ? [a[0], b[1]] : [b[0], a[1]];
          [[a, corner], [corner, b]].forEach(function (sg) {
            var len = Math.max(Math.abs(sg[1][0] - sg[0][0]), Math.abs(sg[1][1] - sg[0][1])), steps = Math.max(1, Math.round(len / 0.0015));
            for (var s = 0; s < steps; s++) {
              var f = s / steps;
              pts.push([sg[0][0] + (sg[1][0] - sg[0][0]) * f + (rnd() - 0.5) * 0.0002, sg[0][1] + (sg[1][1] - sg[0][1]) * f + (rnd() - 0.5) * 0.0002]);
            }
          });
        }
        pts.push(depot);
        return pts;
      });
    });
    // 運転手ごとの「くせ」（警報の出やすさ）
    var risk = [0.5, 2.4, 0.8, 0.3, 3.2, 1.0, 1.4, 0.4, 1.9, 0.7, 1.1, 0.6, 2.0, 0.9];
    var favor = [['brake'], ['phone', 'distract'], ['tailgate'], [], ['speed', 'distract', 'brake'], ['lane'], ['accel', 'brake'], [], ['tailgate', 'smoking'], ['corner'], ['seatbelt'], [], ['fatigue', 'lane'], ['stop']];
    var base = ['brake', 'brake', 'accel', 'corner', 'speed', 'speed', 'tailgate', 'lane', 'distract', 'stop', 'seatbelt', 'smoking', 'button', 'drinking'];
    var events = [], n = 0;
    drvNames.forEach(function (_, di) {
      var vi = di % devNames.length, dev = 'b' + vi, ci = vi % cities.length, myRoutes = routes[vi];
      var kmPerDay = 60 + rnd() * 90;
      trips.push({ devId: dev, drvId: 'u' + di, km: kmPerDay * curDays });
      // 軌跡：このルートを何回走ったか（重ね描きの回数）
      myRoutes.forEach(function (rt, k) {
        if (vi + 12 < drvNames.length && k % 2 !== (di < 12 ? 0 : 1)) return;   // 車両を2人で使う場合はルートを分担
        var times = Math.max(1, Math.round(curDays * (0.3 + rnd() * 0.9)));
        tracks.push({ devId: dev, drvId: 'u' + di, x: new Float64Array(rt.map(function (q) { return mercX(q[1]); })), y: new Float64Array(rt.map(function (q) { return mercY(q[0]); })), n: times });
      });
      var count = Math.round(risk[di] * days * 0.45 * (0.8 + rnd() * 0.4));
      for (var k = 0; k < count; k++) {
        var t = +p.prevFrom + rnd() * span;
        var hour = 7 + Math.floor(rnd() * 12);                     // 7時〜19時に発生
        t = jstMidnight(t) + hour * 3600e3 + Math.floor(rnd() * 3600e3);
        if (t > now || t < +p.prevFrom || t > +p.to) continue;
        var pool = favor[di].length && rnd() < 0.55 ? favor[di] : base;
        var key = pool[Math.floor(rnd() * pool.length)];
        if (rnd() < 0.01) key = 'near';
        var ll;
        if (rnd() < 0.55) { var h = hot[ci][Math.floor(rnd() * 3)]; ll = [h[0] + (rnd() - 0.5) * 0.0016, h[1] + (rnd() - 0.5) * 0.002]; }
        else { var rt2 = myRoutes[Math.floor(rnd() * myRoutes.length)]; ll = rt2[Math.floor(rnd() * rt2.length)]; }
        events.push({ id: 'e' + (n++), t: new Date(t), end: new Date(t + 4000 + rnd() * 8000), devId: dev, drvId: 'u' + di,
          type: TYPE_BY_KEY[key], ruleName: TYPE_BY_KEY[key].label + '（デモ）',
          pos: { lat: ll[0], lon: ll[1], spd: key === 'button' ? 0 : Math.round(15 + rnd() * 55) },
          media: { demo: true } });
      }
    });
    // 運転手未登録の走行
    trips.push({ devId: 'b11', drvId: UNKNOWN_DRIVER, km: 140 });
    events.push({ id: 'e' + (n++), t: new Date(now - 2 * DAY), devId: 'b11', drvId: UNKNOWN_DRIVER, type: TYPE_BY_KEY.brake, ruleName: '急ブレーキ（デモ）', pos: { lat: 35.61, lon: 140.11, spd: 42 }, media: { demo: true } });
    // デモでは確認済みの印を一部に付けておく
    if (!load('demoMarked', false)) {
      events.forEach(function (e, i) { if (+e.t < now - 5 * DAY && i % 3) marks[e.id] = { ack: true }; });
      save('marks', marks); save('demoMarked', true);
    }
    var d = build(events, devices, drivers, trips, p);
    d.tracks = tracks; d.trk = { done: 1, total: 1, loading: false };
    // デモ用のグループ（関東エリアの下に営業所）
    setGroups(d, [
      { id: 'g-kanto', name: '関東エリア', children: ['g-tokyo', 'g-kanagawa', 'g-saitama'] },
      { id: 'g-tokyo', name: '東京営業所', children: [] },
      { id: 'g-kanagawa', name: '神奈川営業所', children: [] },
      { id: 'g-saitama', name: '埼玉・千葉営業所', children: [] },
      { id: 'g-nagoya', name: '名古屋営業所', children: [] }
    ], { b0: ['g-tokyo'], b1: ['g-tokyo'], b2: ['g-tokyo'], b3: ['g-tokyo'], b9: ['g-kanagawa'], b10: ['g-tokyo'],
      b4: ['g-kanagawa'], b5: ['g-kanagawa'], b6: ['g-nagoya'], b7: ['g-nagoya'], b8: ['g-saitama'], b11: ['g-saitama'] });
    return Promise.resolve(d);
  }

  // グループ：MyGeotabのグループ（営業所など）。上位のグループを選んだ場合は、その下のグループの車両も含める。
  // 「GroupCompanyId」などGeotab標準の仕組み用グループは選択肢に出さない。
  function setGroups(d, list, devRaw) {
    var byId = {}, parent = {};
    list.forEach(function (g) { byId[g.id] = g; });
    list.forEach(function (g) { g.children.forEach(function (c) { parent[c] = g.id; }); });
    var devGroups = {}, count = {};
    Object.keys(devRaw).forEach(function (dv) {
      var set = {};
      devRaw[dv].forEach(function (gid) { for (var g = gid, k = 0; g && k < 30; g = parent[g], k++) set[g] = 1; });
      devGroups[dv] = set;
      Object.keys(set).forEach(function (g) { count[g] = (count[g] || 0) + 1; });
    });
    var builtin = /^Group\w*Id$/;
    var opts = [];
    function walk(gid, depth) {
      var g = byId[gid]; if (!g) return;
      var show = !builtin.test(gid) && count[gid];
      if (show) opts.push({ id: gid, name: g.name || gid, depth: depth, n: count[gid] });
      g.children.map(function (c) { return byId[c]; }).filter(Boolean)
        .sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'ja'); })
        .forEach(function (c) { walk(c.id, show ? depth + 1 : depth); });
    }
    list.filter(function (g) { return !parent[g.id]; }).forEach(function (g) { walk(g.id, 0); });
    d.groups = opts; d.devGroups = devGroups;
  }
  var savedGroup = null;
  function inGroup(devId, gid) { return !gid || !!((data.devGroups || {})[devId] || {})[gid]; }

  function build(events, devices, drivers, trips, p) {
    var cur = events.filter(function (e) { return e.t >= p.from && e.t <= p.to; });
    var prev = events.filter(function (e) { return e.t >= p.prevFrom && e.t < p.from; });
    cur.sort(function (a, b) { return b.t - a.t; });
    return { events: cur, prevEvents: prev, devices: devices, drivers: drivers, trips: trips, tracks: [], trk: null, p: p, groups: [], devGroups: {} };
  }

  // ---------- 絞り込み ----------
  function nameOf(by, id) {
    if (by === 'driver') return id === UNKNOWN_DRIVER ? '（運転手未登録）' : data.drivers[id] || id;
    return data.devices[id] || id;
  }
  function f() { return { grp: $('f-group').value, drv: $('f-driver').value, dev: $('f-device').value, type: $('f-type').value }; }
  function typeOk(e, ty) {
    if (ty === '_crit') return e.type.sev === 'crit';
    if (ty === '_noinfo') return e.type.sev !== 'info';
    return !ty || e.type.key === ty;
  }
  // opts.noType: 種類の条件を外す / opts.noWho: ドライバー・車両の条件を外す（グループの条件は残す） / opts.list: 一覧の条件（未確認だけ・検索）も使う
  function filtered(opts, src) {
    opts = opts || {}; var c = f(), q = opts.list ? ($('cs-filter').value || '').trim() : '';
    return (src || data.events).filter(function (e) {
      if (c.grp && !inGroup(e.devId, c.grp)) return false;
      if (!opts.noWho && c.drv && e.drvId !== c.drv) return false;
      if (!opts.noWho && c.dev && e.devId !== c.dev) return false;
      if (!opts.noType && !typeOk(e, c.type)) return false;
      if (opts.list && view.onlyTodo && (marks[e.id] || {}).ack) return false;
      if (q && (nameOf('driver', e.drvId) + ' ' + nameOf('device', e.devId)).indexOf(q) < 0) return false;
      return true;
    });
  }
  function kmOf(c) {
    return data.trips.reduce(function (a, t) { return a + (inGroup(t.devId, c.grp) && (!c.drv || t.drvId === c.drv) && (!c.dev || t.devId === c.dev) ? t.km : 0); }, 0);
  }
  function points(list) { return list.reduce(function (a, e) { return a + e.type.w; }, 0); }
  // 安全スコア：100kmあたりの「重み付き警報点数」r から算出。100 × e^(−r/20)
  // （r=1 → 95点、r=5 → 78点、r=10 → 61点、r=20 → 37点。0点にはならず、警報が多い人どうしでも差が見える）
  // 距離が短すぎる場合は50km として扱う
  function scoreOf(pts, km) { var r = pts / Math.max(km, 50) * 100; return Math.round(100 * Math.exp(-r / 20)); }
  function scoreCls(s) { return s >= 80 ? 's-good' : s >= 60 ? 's-mid' : 's-bad'; }

  function renderFilters() {
    var gSel = $('f-group'), dSel = $('f-driver'), vSel = $('f-device'), tSel = $('f-type'), c = f(), drvs = {}, devs = {}, cnt = {};
    // グループ（階層は字下げで表示）。前回選んだグループを最初に復元
    var gids = {};
    if (savedGroup != null) { c.grp = savedGroup; savedGroup = null; }
    gSel.innerHTML = '<option value="">すべてのグループ</option>' + (data.groups || []).map(function (g) {
      gids[g.id] = 1;
      return '<option value="' + esc(g.id) + '">' + new Array(g.depth + 1).join('　') + esc(g.name) + '（' + g.n + '台）</option>';
    }).join('');
    gSel.value = gids[c.grp] ? c.grp : ''; c.grp = gSel.value;
    gSel.disabled = !(data.groups || []).length;
    // ドライバー・車両の選択肢は、選んだグループの車両に関係するものだけ
    data.trips.forEach(function (t) { if (!inGroup(t.devId, c.grp)) return; drvs[t.drvId] = drvs[t.drvId] || 0; devs[t.devId] = devs[t.devId] || 0; });
    data.events.forEach(function (e) { if (!inGroup(e.devId, c.grp)) return; drvs[e.drvId] = (drvs[e.drvId] || 0) + 1; devs[e.devId] = (devs[e.devId] || 0) + 1; cnt[e.type.key] = (cnt[e.type.key] || 0) + 1; });
    function opts(map, by, all) {
      return '<option value="">' + all + '</option>' + Object.keys(map).sort(function (a, b) { return nameOf(by, a).localeCompare(nameOf(by, b), 'ja'); })
        .map(function (id) { return '<option value="' + esc(id) + '">' + esc(nameOf(by, id)) + '（' + map[id] + '件）</option>'; }).join('');
    }
    dSel.innerHTML = opts(drvs, 'driver', '全員'); dSel.value = drvs[c.drv] !== undefined ? c.drv : '';
    vSel.innerHTML = opts(devs, 'device', '全車両'); vSel.value = devs[c.dev] !== undefined ? c.dev : '';
    var keys = Object.keys(cnt).sort(function (a, b) { return SEV_ORDER[TYPE_BY_KEY[a].sev] - SEV_ORDER[TYPE_BY_KEY[b].sev] || cnt[b] - cnt[a]; });
    tSel.innerHTML = '<option value="">すべて</option><option value="_noinfo">危険なもの（重大＋注意）</option><option value="_crit">重大のみ</option>' +
      keys.map(function (k) { return '<option value="' + k + '">' + esc(TYPE_BY_KEY[k].label) + '（' + cnt[k] + '件）</option>'; }).join('');
    tSel.value = c.type === '' || c.type.charAt(0) === '_' || cnt[c.type] ? c.type : '';
    markFilters();
  }
  function markFilters() { ['f-group', 'f-driver', 'f-device', 'f-type'].forEach(function (id) { $(id).classList.toggle('on', !!$(id).value); }); }
  function setFilter(id, value) { var el = $(id); el.value = el.value === value ? '' : value; if (el.value !== value && el.value !== '') el.value = ''; markFilters(); filterChanged(); }
  function filterChanged() { hm.userMoved = false; render(); fitMap(); requestDraw(); }

  // ---------- 描画（左の列） ----------
  function render() {
    markFilters(); renderKpis(); renderTypes(); renderTrend(); renderMapSide(); renderRank(); renderList(); requestDraw();
  }

  function renderKpis() {
    var c = f(), ev = filtered(), km = kmOf(c);
    $('kpi-total').textContent = ev.length + '件';
    var pv = filtered({}, data.prevEvents).length, sub = $('kpi-total-sub');
    if (pv) {
      var ch = Math.round((ev.length - pv) / pv * 100);
      sub.textContent = '前の期間より ' + (ch > 0 ? '+' : '') + ch + '%（' + pv + '件）';
      sub.className = ch > 0 ? 'up' : ch < 0 ? 'down' : '';
    } else { sub.textContent = '前の期間：0件'; sub.className = ''; }
    $('kpi-rate').textContent = km ? (ev.length / km * 100).toFixed(2) + '件' : '-';
    $('kpi-rate-sub').textContent = '警報 ÷ 走行距離';
    $('kpi-crit').textContent = ev.filter(function (e) { return e.type.sev === 'crit'; }).length + '件';
    var rows = c.drv ? rankRows('driver').filter(function (r) { return r.id === c.drv; }) : rankRows('driver').filter(function (r) { return r.km > 0 || r.n > 0; });
    $('kpi-score').textContent = rows.length ? Math.round(rows.reduce(function (a, r) { return a + r.score; }, 0) / rows.length) + '点' : '-';
    $('kpi-todo').textContent = ev.filter(function (e) { return !(marks[e.id] || {}).ack; }).length + '件';
    $('kpi-km').textContent = Math.round(km).toLocaleString() + 'km';
    var nd = {}; data.trips.forEach(function (t) { if (inGroup(t.devId, c.grp) && (!c.drv || t.drvId === c.drv) && (!c.dev || t.devId === c.dev)) nd[t.devId] = 1; });
    $('kpi-km-sub').textContent = Object.keys(nd).length + '台';
  }

  function renderTypes() {
    var list = filtered({ noType: true }), cnt = {}, cur = f().type;
    list.forEach(function (e) { cnt[e.type.key] = (cnt[e.type.key] || 0) + 1; });
    var keys = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; });
    var max = keys.length ? cnt[keys[0]] : 1;
    $('cs-types').innerHTML = keys.map(function (k) {
      var t = TYPE_BY_KEY[k];
      return '<button type="button" class="cs-type' + (cur === k ? ' on' : '') + '" data-type="' + k + '" title="' + esc(t.label) + 'だけを表示">' +
        '<span class="name">' + esc(t.label) + '</span><span><span class="bar ' + t.sev + '" style="width:' + (cnt[k] / max * 100).toFixed(1) + '%"></span></span><b>' + cnt[k] + '</b></button>';
    }).join('') || '<div class="cs-empty">期間内の警報はありません</div>';
  }

  function renderTrend() {
    var p = data.p, list = filtered(), step = p.hourly ? 3600e3 : DAY;
    var t0 = p.hourly ? +p.from : jstMidnight(+p.from), t1 = +p.to, buckets = [];
    for (var t = t0; t < t1; t += step) buckets.push({ t: t, crit: 0, warn: 0, info: 0 });
    list.forEach(function (e) { var i = Math.floor((e.t - t0) / step); if (buckets[i]) buckets[i][e.type.sev]++; });
    var W = Math.max(280, Math.round($('cs-trend').clientWidth || 360)), H = 130, L = 24, R = 4, T = 6, B = 20, PW = W - L - R, PH = H - T - B;
    var max = Math.max(1, Math.max.apply(null, buckets.map(function (b) { return b.crit + b.warn + b.info; })));
    var stepY = max <= 5 ? 1 : max <= 10 ? 2 : max <= 25 ? 5 : max <= 50 ? 10 : Math.ceil(max / 50) * 10;
    max = Math.ceil(max / stepY) * stepY;
    var bw = PW / buckets.length, svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="警報件数の推移">';
    for (var g = 0; g <= max; g += stepY) {
      var y = T + PH - g / max * PH;
      svg += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y + '" y2="' + y + '" stroke="#eef0f3"/><text x="' + (L - 4) + '" y="' + (y + 4) + '" font-size="10" text-anchor="end" fill="#6b7280">' + g + '</text>';
    }
    var labelEvery = Math.ceil(buckets.length / 6);
    function label(t) { var d = new Date(t + JST); return p.hourly ? d.getUTCHours() + '時' : (d.getUTCMonth() + 1) + '/' + d.getUTCDate(); }
    buckets.forEach(function (b, i) {
      var x = L + i * bw + Math.max(0.5, bw * 0.15), w = Math.max(1, bw * 0.7), y = T + PH;
      ['info', 'warn', 'crit'].forEach(function (s) {
        if (!b[s]) return; var h = b[s] / max * PH; y -= h;
        svg += '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + w.toFixed(1) + '" height="' + h.toFixed(1) + '" fill="' + { crit: '#dc2626', warn: '#f59e0b', info: '#94a3b8' }[s] + '"><title>' + label(b.t) + '：' + ({ crit: '重大', warn: '注意', info: 'その他' })[s] + ' ' + b[s] + '件</title></rect>';
      });
      if (i % labelEvery === 0) svg += '<text x="' + (L + i * bw + bw / 2).toFixed(1) + '" y="' + (H - 6) + '" font-size="10" text-anchor="middle" fill="#6b7280">' + label(b.t) + '</text>';
    });
    $('cs-trend').innerHTML = svg + '</svg>';
    $('cs-trend-note').textContent = p.hourly ? '1時間ごと' : '1日ごと';
  }

  function rankRows(by) {
    var groups = {}, km = {};
    filtered({ noWho: true }).forEach(function (e) { var id = by === 'driver' ? e.drvId : e.devId; (groups[id] = groups[id] || []).push(e); });
    var g = f().grp;
    data.trips.forEach(function (t) { if (!inGroup(t.devId, g)) return; var id = by === 'driver' ? t.drvId : t.devId; km[id] = (km[id] || 0) + t.km; });
    var ids = Object.keys(groups);
    Object.keys(km).forEach(function (id) { if (!groups[id] && km[id] > 1) ids.push(id); });
    return ids.map(function (id) {
      var list = groups[id] || [], k = km[id] || 0;
      return { id: id, name: nameOf(by, id), km: k, n: list.length, score: scoreOf(points(list), k), short: k < 50 };
    }).sort(function (a, b) { return a.score - b.score || b.n - a.n; });
  }

  function renderRank() {
    var by = view.rankBy, rows = rankRows(by), c = f(), cur = by === 'driver' ? c.drv : c.dev;
    $('cs-rank-name').textContent = by === 'driver' ? '運転手' : '車両';
    $('cs-rank').innerHTML = rows.map(function (r, i) {
      return '<tr data-id="' + esc(r.id) + '"' + (cur === r.id ? ' class="on"' : '') + '><td><span class="cs-rankno">' + (rows.length - i) + '</span></td>' +
        '<td>' + esc(r.name) + '</td>' +
        '<td><span class="cs-score"><b class="' + scoreCls(r.score) + '">' + r.score + '</b><span class="track"><span class="fill ' + scoreCls(r.score) + '" style="width:' + r.score + '%"></span></span></span></td>' +
        '<td class="r">' + r.n + '件</td><td class="r">' + Math.round(r.km).toLocaleString() + 'km' + (r.short ? '<span class="cs-muted" title="走行距離が短いため、50km として計算しています">※</span>' : '') + '</td></tr>';
    }).join('') || '<tr><td colspan="5" class="cs-empty">データがありません</td></tr>';
  }

  function renderList() {
    listRows = filtered({ list: true });
    $('cs-list-count').textContent = listRows.length + '件' + (listRows.length > 500 ? '（新しい500件を表示）' : '');
    var shown = listRows.slice(0, 500);
    $('cs-events').innerHTML = shown.map(function (e, i) {
      var m = marks[e.id] || {};
      return '<tr data-id="' + esc(e.id) + '" data-idx="' + i + '" class="' + (m.ack ? 'acked' : '') + (hm.focus === e ? ' on' : '') + '">' +
        '<td><input type="checkbox" class="cs-ack" data-ack="' + esc(e.id) + '"' + (m.ack ? ' checked' : '') + ' title="確認済みにする"></td>' +
        '<td><span class="cs-tag ' + e.type.sev + '" title="' + esc(e.ruleName) + '">' + esc(e.type.label) + '</span>' + (m.memo ? ' <span title="' + esc(m.memo) + '">📝</span>' : '') +
        '<small>' + fmtShort(e.t) + (e.pos ? '・' + Math.round(e.pos.spd) + 'km/h' : '') + '</small></td>' +
        '<td>' + esc(nameOf('driver', e.drvId)) + '<small>' + esc(nameOf('device', e.devId)) + '</small></td>' +
        '<td><button type="button" class="cs-video-btn" data-open="' + i + '" title="映像・詳細を開く">▶</button></td></tr>';
    }).join('') || '<tr><td colspan="4" class="cs-empty">' + (view.onlyTodo ? '未確認の警報はありません' : '該当する警報はありません') + '</td></tr>';
    ensurePositions(shown.slice(0, 100));
  }

  // ---------- 地図（OpenStreetMapの地図画像の上に、走行軌跡とヒートマップを重ねる） ----------
  var TILE = 256, hm = { lat: 36.5, lon: 138, z: 5, userMoved: false, pts: [], screen: [], tiles: {}, raf: 0, focus: null };
  function wsize(z) { return TILE * Math.pow(2, z); }
  function lon2x(lon, z) { return mercX(lon) * wsize(z); }
  function lat2y(lat, z) { return mercY(lat) * wsize(z); }
  function x2lon(x, z) { return x / wsize(z) * 360 - 180; }
  function y2lat(y, z) { var n = Math.PI - 2 * Math.PI * y / wsize(z); return 180 / Math.PI * Math.atan(0.5 * (Math.exp(n) - Math.exp(-n))); }
  var SEV_W = { crit: 1, warn: 0.6, info: 0.3 }, SEV_C = { crit: '#dc2626', warn: '#f59e0b', info: '#64748b' };

  function mapTracks() {
    if (!data || !data.tracks) return [];
    var c = f();
    return data.tracks.filter(function (t) { return inGroup(t.devId, c.grp) && (!c.drv || t.drvId === c.drv) && (!c.dev || t.devId === c.dev); });
  }
  function renderMapSide() {
    if (!data) return;
    var list = filtered();
    hm.pts = list.filter(function (e) { return e.pos; });
    var trk = data.trk, tr = mapTracks().length;
    $('hm-count').textContent = '警報 ' + list.length + '件' + (hm.pts.length < list.length ? '（地図に ' + hm.pts.length + '件）' : '') + '・軌跡 ' + tr + '本';
    renderHotspots(); hmStatus();
  }
  function hmStatus() {
    var el = $('hm-status'); if (!el || !data) return;
    var msg = [], trk = data.trk;
    if (trk && trk.loading) msg.push('走行軌跡を読み込み中… ' + trk.done + ' / ' + trk.total + '台');
    var n = posQueue.length + (posBusy ? 100 : 0);
    if (n) msg.push('警報の位置を取得中…（残り約' + n + '件）');
    el.hidden = !msg.length; el.textContent = msg.join('　');
  }
  function fitMap() {
    var el = $('hm-map'), W = el.clientWidth || 800, H = el.clientHeight || 500;
    var x0 = 1, x1 = 0, y0 = 1, y1 = 0, any = false;
    function add(x, y) { any = true; if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    mapTracks().forEach(function (t) { for (var i = 0; i < t.x.length; i += 5) add(t.x[i], t.y[i]); });
    hm.pts.forEach(function (e) { add(mercX(e.pos.lon), mercY(e.pos.lat)); });
    if (!any) return;
    var z = 16;
    while (z > 3 && ((x1 - x0) * wsize(z) > W - 80 || (y1 - y0) * wsize(z) > H - 80)) z--;
    hm.z = z;
    hm.lon = x2lon((x0 + x1) / 2 * wsize(z), z);
    hm.lat = y2lat((y0 + y1) / 2 * wsize(z), z);
  }
  function requestDraw() { if (!hm.raf) hm.raf = requestAnimationFrame(function () { hm.raf = 0; drawMap(); }); }

  var palette = null, stamp = null, STAMP_R = 22;
  function heatAssets() {
    if (palette) return;
    var c = document.createElement('canvas'); c.width = 256; c.height = 1;
    var g = c.getContext('2d'), gr = g.createLinearGradient(0, 0, 256, 0);
    gr.addColorStop(0, '#22c55e'); gr.addColorStop(0.4, '#facc15'); gr.addColorStop(0.7, '#f97316'); gr.addColorStop(1, '#dc2626');
    g.fillStyle = gr; g.fillRect(0, 0, 256, 1); palette = g.getImageData(0, 0, 256, 1).data;
    stamp = document.createElement('canvas'); stamp.width = stamp.height = STAMP_R * 2;
    var s = stamp.getContext('2d'), rg = s.createRadialGradient(STAMP_R, STAMP_R, 0, STAMP_R, STAMP_R, STAMP_R);
    rg.addColorStop(0, 'rgba(0,0,0,1)'); rg.addColorStop(1, 'rgba(0,0,0,0)');
    s.fillStyle = rg; s.fillRect(0, 0, STAMP_R * 2, STAMP_R * 2);
  }
  function viewBox() {
    var el = $('hm-map'), W = el.clientWidth, H = el.clientHeight;
    return { W: W, H: H, ox: lon2x(hm.lon, hm.z) - W / 2, oy: lat2y(hm.lat, hm.z) - H / 2, ws: wsize(hm.z) };
  }
  function drawTiles(v) {
    var z = hm.z, n = Math.pow(2, z), box = $('hm-tiles'), want = {};
    for (var tx = Math.floor(v.ox / TILE); tx <= Math.floor((v.ox + v.W) / TILE); tx++) {
      for (var ty = Math.floor(v.oy / TILE); ty <= Math.floor((v.oy + v.H) / TILE); ty++) {
        if (ty < 0 || ty >= n) continue;
        var wx = ((tx % n) + n) % n, key = z + '/' + wx + '/' + ty + '/' + tx, img = hm.tiles[key];
        if (!img) { img = new Image(); img.alt = ''; img.draggable = false; img.src = 'https://tile.openstreetmap.org/' + z + '/' + wx + '/' + ty + '.png'; hm.tiles[key] = img; box.appendChild(img); }
        img.style.left = Math.round(tx * TILE - v.ox) + 'px'; img.style.top = Math.round(ty * TILE - v.oy) + 'px';
        want[key] = 1;
      }
    }
    Object.keys(hm.tiles).forEach(function (k) { if (!want[k]) { hm.tiles[k].remove(); delete hm.tiles[k]; } });
  }
  function drawMap() {
    if (!data) return;
    var v = viewBox(); if (!v.W || !v.H) return;
    drawTiles(v);
    var cv = $('hm-canvas'), dpr = window.devicePixelRatio || 1;
    cv.style.transform = '';
    cv.width = Math.round(v.W * dpr); cv.height = Math.round(v.H * dpr);
    var ctx = cv.getContext('2d'); ctx.setTransform(dpr, 0, 0, dpr, 0, 0); ctx.clearRect(0, 0, v.W, v.H);
    // 走行軌跡：薄い青の線を重ねる（よく通る道ほど濃くなる）
    if ($('l-tracks').checked) {
      var trs = mapTracks(), many = trs.reduce(function (a, t) { return a + t.n; }, 0);
      var a1 = many > 2000 ? 0.06 : many > 500 ? 0.1 : 0.16;
      ctx.lineJoin = 'round'; ctx.lineCap = 'round';
      ctx.lineWidth = hm.z >= 14 ? 3.5 : hm.z >= 11 ? 2.5 : 2;
      ctx.strokeStyle = '#1d4ed8';
      var minA = hm.z < 11 ? 0.35 : 0.18;   // 縮小しているときは見えやすいよう濃いめに
      trs.forEach(function (t) {
        ctx.globalAlpha = Math.min(0.85, Math.max(minA, 1 - Math.pow(1 - a1, t.n)));
        ctx.beginPath();
        var lx = -1e9, ly = -1e9, started = false;
        for (var i = 0; i < t.x.length; i++) {
          var x = t.x[i] * v.ws - v.ox, y = t.y[i] * v.ws - v.oy;
          if (!started) { ctx.moveTo(x, y); started = true; lx = x; ly = y; continue; }
          if (Math.abs(x - lx) + Math.abs(y - ly) < 1.5 && i < t.x.length - 1) continue;   // 画面上で近すぎる点は省略
          ctx.lineTo(x, y); lx = x; ly = y;
        }
        ctx.stroke();
      });
      ctx.globalAlpha = 1;
    }
    // ヒートマップ：画面を小さなマス目に分けて警報の数（重大ほど重く）を数え、ぼかしてから、
    // 表示範囲の中で一番多い場所を赤として比べて色を付ける（警報が多くても全部が赤にならないように）
    var PAD = 48;
    hm.screen = [];
    hm.pts.forEach(function (e) {
      var x = mercX(e.pos.lon) * v.ws - v.ox, y = mercY(e.pos.lat) * v.ws - v.oy;
      if (x < -PAD || y < -PAD || x > v.W + PAD || y > v.H + PAD) return;
      hm.screen.push({ x: x, y: y, e: e });
    });
    if ($('l-heat').checked && hm.screen.length) {
      heatAssets();
      var CELL = 5, BL = 3, pc = Math.ceil(PAD / CELL);
      var gw = Math.ceil(v.W / CELL) + pc * 2, gh = Math.ceil(v.H / CELL) + pc * 2, grid = new Float32Array(gw * gh), tmp = new Float32Array(gw * gh);
      hm.screen.forEach(function (s) {
        var i = Math.floor(s.x / CELL) + pc, j = Math.floor(s.y / CELL) + pc;
        if (i >= 0 && j >= 0 && i < gw && j < gh) grid[j * gw + i] += SEV_W[s.e.type.sev];
      });
      // ぼかし（横・縦の平均を2回ずつ）
      function blurH(a, b) { for (var y = 0; y < gh; y++) { var row = y * gw, acc = 0; for (var x = -BL; x < gw + BL; x++) { if (x + BL < gw) acc += a[row + x + BL]; if (x - BL - 1 >= 0) acc -= a[row + x - BL - 1]; if (x >= 0 && x < gw) b[row + x] = acc / (BL * 2 + 1); } } }
      function blurV(a, b) { for (var x = 0; x < gw; x++) { var acc = 0; for (var y = -BL; y < gh + BL; y++) { if (y + BL < gh) acc += a[(y + BL) * gw + x]; if (y - BL - 1 >= 0) acc -= a[(y - BL - 1) * gw + x]; if (y >= 0 && y < gh) b[y * gw + x] = acc / (BL * 2 + 1); } } }
      blurH(grid, tmp); blurV(tmp, grid); blurH(grid, tmp); blurV(tmp, grid);
      var max = 0, k;
      for (k = 0; k < grid.length; k++) if (grid[k] > max) max = grid[k];
      if (max > 0) {
        var small = document.createElement('canvas'); small.width = gw; small.height = gh;
        var sc = small.getContext('2d'), img2 = sc.createImageData(gw, gh), d = img2.data;
        for (k = 0; k < grid.length; k++) {
          var val = grid[k] / max; if (val < 0.06) continue;      // ごく少ない所は色を付けない
          var vv = Math.pow(val, 0.7), pi = Math.min(255, Math.round(vv * 255)) * 4, o4 = k * 4;
          d[o4] = palette[pi]; d[o4 + 1] = palette[pi + 1]; d[o4 + 2] = palette[pi + 2]; d[o4 + 3] = Math.round(70 + vv * 160);
        }
        sc.putImageData(img2, 0, 0);
        ctx.imageSmoothingEnabled = true;
        ctx.drawImage(small, -pc * CELL, -pc * CELL, gw * CELL, gh * CELL);
      }
    }
    if ($('l-points').checked) {
      hm.screen.forEach(function (s) {
        ctx.beginPath(); ctx.arc(s.x, s.y, 4.5, 0, Math.PI * 2);
        ctx.fillStyle = SEV_C[s.e.type.sev]; ctx.fill(); ctx.lineWidth = 1.5; ctx.strokeStyle = '#fff'; ctx.stroke();
      });
    }
    // 一覧で選んだ警報の位置
    if (hm.focus && hm.focus.pos) {
      var fx = mercX(hm.focus.pos.lon) * v.ws - v.ox, fy = mercY(hm.focus.pos.lat) * v.ws - v.oy;
      ctx.beginPath(); ctx.arc(fx, fy, 11, 0, Math.PI * 2); ctx.lineWidth = 3; ctx.strokeStyle = '#111827'; ctx.stroke();
      ctx.beginPath(); ctx.arc(fx, fy, 5, 0, Math.PI * 2); ctx.fillStyle = SEV_C[hm.focus.type.sev]; ctx.fill();
    }
  }
  // 警報が多い場所：約300m四方ごとに集計
  var hotspots = [];
  function renderHotspots() {
    var cells = {};
    hm.pts.forEach(function (e) {
      var k = Math.round(e.pos.lat / 0.003) + ':' + Math.round(e.pos.lon / 0.0036);
      var c = cells[k] || (cells[k] = { n: 0, lat: 0, lon: 0, types: {} });
      c.n++; c.lat += e.pos.lat; c.lon += e.pos.lon; c.types[e.type.label] = (c.types[e.type.label] || 0) + 1;
    });
    hotspots = Object.keys(cells).map(function (k) { var c = cells[k]; c.lat /= c.n; c.lon /= c.n; return c; })
      .filter(function (c) { return c.n >= 2; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 5);
    $('hm-top').innerHTML = hotspots.map(function (c, i) {
      var ad = addressOf({ lat: c.lat, lon: c.lon });
      var tys = Object.keys(c.types).sort(function (a, b) { return c.types[b] - c.types[a]; }).slice(0, 2).map(function (t) { return t + ' ' + c.types[t]; }).join('・');
      return '<li data-hot="' + i + '"><b class="no">' + (i + 1) + '</b><span class="addr">' + esc(ad == null ? '住所を取得中…' : ad || (c.lat.toFixed(4) + ', ' + c.lon.toFixed(4))) +
        '<small>' + esc(tys) + '</small></span><span class="n">' + c.n + '件</span></li>';
    }).join('') || '<li style="cursor:default;display:block" class="cs-muted">同じ場所で2件以上の警報はありません</li>';
  }
  function focusEvent(e) {
    hm.focus = e;
    if (e.pos) { hm.lat = e.pos.lat; hm.lon = e.pos.lon; hm.z = Math.max(hm.z, 15); hm.userMoved = true; }
    else ensurePositions([e]);
    requestDraw();
    var on = document.querySelector('#cs-events tr.on'); if (on) on.classList.remove('on');
    var tr = document.querySelector('#cs-events tr[data-id="' + e.id.replace(/"/g, '') + '"]'); if (tr) tr.classList.add('on');
  }
  function bindMap() {
    var map = $('hm-map'), drag = null;
    function rel(ev) { var r = map.getBoundingClientRect(); return { x: ev.clientX - r.left, y: ev.clientY - r.top }; }
    function hit(p) {
      if (!$('l-points').checked && !$('l-heat').checked) return null;
      var best = null, bd = $('l-points').checked ? 64 : 0;
      hm.screen.forEach(function (s) { var d = (s.x - p.x) * (s.x - p.x) + (s.y - p.y) * (s.y - p.y); if (d < bd) { bd = d; best = s; } });
      return best;
    }
    map.addEventListener('pointerdown', function (ev) {
      if (ev.target.closest('.cs-hm-tools, .cs-hm-attr, .cs-hm-legend') || (ev.pointerType === 'mouse' && ev.button !== 0)) return;
      var p = rel(ev);
      drag = { x: p.x, y: p.y, cx: lon2x(hm.lon, hm.z), cy: lat2y(hm.lat, hm.z), moved: false };
      try { map.setPointerCapture(ev.pointerId); } catch (e) { /* ignore */ }
    });
    map.addEventListener('pointermove', function (ev) {
      var p = rel(ev);
      if (!drag) { map.classList.toggle('hover', !!hit(p)); return; }
      if (Math.abs(p.x - drag.x) + Math.abs(p.y - drag.y) > 4) { drag.moved = true; map.classList.add('drag'); }
      if (!drag.moved) return;
      hm.lon = x2lon(drag.cx - (p.x - drag.x), hm.z);
      hm.lat = y2lat(Math.max(0, Math.min(wsize(hm.z), drag.cy - (p.y - drag.y))), hm.z);
      hm.userMoved = true;
      // ドラッグ中は描き直さず、描画済みの画像をずらすだけ（軽くするため）。離したら描き直す
      drawTiles(viewBox());
      var nx = lon2x(hm.lon, hm.z), ny = lat2y(hm.lat, hm.z);
      $('hm-canvas').style.transform = 'translate(' + Math.round(drag.cx - nx) + 'px,' + Math.round(drag.cy - ny) + 'px)';
    });
    function end(ev) {
      if (!drag) return;
      var d = drag; drag = null; map.classList.remove('drag');
      if (d.moved) { drawMap(); return; }
      if (ev.type !== 'pointerup') return;
      var s = hit(rel(ev)); if (!s) return;
      var list = hm.screen.map(function (x) { return x.e; }).sort(function (a, b) { return b.t - a.t; });
      openDetail(list.indexOf(s.e), list);
    }
    map.addEventListener('pointerup', end);
    map.addEventListener('pointercancel', end);
    // ページのスクロールを邪魔しないよう、拡大・縮小は Ctrl＋ホイールのときだけ
    var hintTimer = null;
    map.addEventListener('wheel', function (ev) {
      if (!ev.ctrlKey && !ev.metaKey) {
        var st = $('hm-status'); st.hidden = false; st.textContent = 'Ctrl キーを押しながらホイールで拡大・縮小できます';
        clearTimeout(hintTimer); hintTimer = setTimeout(hmStatus, 1500);
        return;
      }
      ev.preventDefault();
      zoomAt(rel(ev), ev.deltaY < 0 ? 1 : -1);
    }, { passive: false });
    function zoomAt(p, dz) {
      var W = map.clientWidth, H = map.clientHeight, z2 = Math.max(3, Math.min(18, hm.z + dz)); if (z2 === hm.z) return;
      var ox = lon2x(hm.lon, hm.z) - W / 2, oy = lat2y(hm.lat, hm.z) - H / 2;
      var lon = x2lon(ox + p.x, hm.z), lat = y2lat(oy + p.y, hm.z);
      hm.z = z2; hm.userMoved = true;
      hm.lon = x2lon(lon2x(lon, z2) - (p.x - W / 2), z2);
      hm.lat = y2lat(lat2y(lat, z2) - (p.y - H / 2), z2);
      requestDraw();
    }
    $('hm-zin').onclick = function () { zoomAt({ x: map.clientWidth / 2, y: map.clientHeight / 2 }, 1); };
    $('hm-zout').onclick = function () { zoomAt({ x: map.clientWidth / 2, y: map.clientHeight / 2 }, -1); };
    $('hm-fit').onclick = function () { hm.userMoved = false; fitMap(); requestDraw(); };
    ['l-tracks', 'l-heat', 'l-points'].forEach(function (id) { $(id).onchange = function () { save('layers', layerState()); requestDraw(); }; });
    $('hm-top').onclick = function (ev) {
      var li = ev.target.closest('[data-hot]'); if (!li) return;
      var c = hotspots[+li.dataset.hot]; hm.lat = c.lat; hm.lon = c.lon; hm.z = 16; hm.userMoved = true; requestDraw();
    };
  }
  function layerState() { return { tracks: $('l-tracks').checked, heat: $('l-heat').checked, points: $('l-points').checked }; }

  // ---------- 画面の高さ：MyGeotabで見えている範囲いっぱいに合わせる ----------
  function scrollParent(el) {
    for (var p = el.parentElement; p && p !== document.body && p !== document.documentElement; p = p.parentElement) {
      var oy = getComputedStyle(p).overflowY;
      if ((oy === 'auto' || oy === 'scroll') && p.clientHeight > 0) return p;
    }
    return null;
  }
  function fitHeight() {
    var root = $('cameraSafety');
    if (window.innerWidth <= 900) { root.style.height = ''; return; }
    var sp = scrollParent(root), vh = window.innerHeight;
    var bottom = sp ? Math.min(vh, sp.getBoundingClientRect().bottom) : vh;
    var top = Math.max(sp ? Math.max(0, sp.getBoundingClientRect().top) : 0, root.getBoundingClientRect().top);
    root.style.height = Math.max(560, Math.floor(bottom - top - 2)) + 'px';
  }

  // ---------- 詳細（映像・地図） ----------
  // MyGeotabの左メニュー・上部バーに隠れないよう、Add-Inが見えている範囲に合わせて配置
  function placeModal(el) {
    var vw = window.innerWidth, vh = window.innerHeight, root = $('cameraSafety');
    var rr = root.getBoundingClientRect(), sp = scrollParent(root), ar = sp ? sp.getBoundingClientRect() : { left: 0, top: 0, right: vw, bottom: vh };
    var left = Math.max(0, ar.left, rr.left), right = Math.min(vw, ar.right, rr.right), top = Math.max(0, ar.top), bottom = Math.min(vh, ar.bottom);
    if (right - left < 300 || bottom - top < 200) { left = 0; top = 0; right = vw; bottom = vh; }
    el.style.left = left + 'px'; el.style.top = top + 'px'; el.style.width = (right - left) + 'px'; el.style.height = (bottom - top) + 'px';
    el.style.right = 'auto'; el.style.bottom = 'auto';
  }
  function openDetail(i, list) { detailList = list; detailIdx = i; var m = $('cs-detail'); m.hidden = false; placeModal(m); renderDetail(); }
  function closeDetail() { detailIdx = -1; $('cs-detail').hidden = true; $('cs-d-video').innerHTML = ''; $('cs-d-map').innerHTML = ''; $('cs-d-map').dataset.key = ''; }

  var addrCache = {}, hotTimer = null;
  function renderDetail() {
    var e = detailList[detailIdx]; if (!e) { closeDetail(); return; }
    var m = marks[e.id] || {};
    $('cs-d-title').innerHTML = '<span class="cs-tag ' + e.type.sev + '">' + esc(e.type.label) + '</span> ' + esc(nameOf('driver', e.drvId)) + ' <span class="cs-muted">' + fmtT(e.t) + '</span>';
    $('cs-d-prev').disabled = detailIdx <= 0; $('cs-d-next').disabled = detailIdx >= detailList.length - 1;
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
      var la = e.pos.lat, lo = e.pos.lon, d = 0.006, key = la.toFixed(5) + ',' + lo.toFixed(5);
      if (mp.dataset.key !== key) {
        mp.dataset.key = key;
        mp.innerHTML = '<iframe title="発生場所の地図" loading="lazy" src="https://www.openstreetmap.org/export/embed.html?bbox=' + (lo - d) + ',' + (la - d * 0.7) + ',' + (lo + d) + ',' + (la + d * 0.7) + '&layer=mapnik&marker=' + la + ',' + lo + '"></iframe>' +
          '<a class="cs-d-maplink" target="_blank" rel="noopener" href="https://www.google.com/maps?q=' + la + ',' + lo + '">Googleマップで開く</a>';
      }
    } else { mp.dataset.key = ''; mp.innerHTML = '<div class="cs-d-novideo"><p style="color:#6b7280">' + (e.pos === undefined ? '位置を取得中…' : '位置データがありません') + '</p></div>'; ensurePositions([e]); }
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
    }, function () { addrCache[key] = ''; }).then(function () {
      if (detailIdx >= 0) renderDetail();
      clearTimeout(hotTimer); hotTimer = setTimeout(renderHotspots, 100);
    });
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
    loading = true; showMsg(''); posQueue = []; hm.userMoved = false; hm.focus = null;
    $('cs-refresh').disabled = true; $('cs-refresh').textContent = '読み込み中…';
    fitHeight();
    (demo ? loadDemo() : loadLive()).then(function (d) {
      data = d;
      renderFilters(); render(); fitMap(); requestDraw();
      var now = new Date();
      $('cs-loaded').textContent = '読み込み ' + now.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
      if (!demo) {
        if (!d.events.length) showMsg('期間内に該当する警報はありません。対象のルールは「設定」で確認できます。', true);
        loadTracksLive(d);
      }
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
    if (!demo) { settings.rules = draftRules; settings.manual = true; }
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
    ['f-driver', 'f-device', 'f-type'].forEach(function (id) { $(id).onchange = function () { markFilters(); filterChanged(); }; });
    // グループを変えたら、ドライバー・車両の選択肢をそのグループのものに作り直す
    $('f-group').onchange = function () { save('group', this.value); renderFilters(); filterChanged(); };
    $('f-reset').onclick = function () { $('f-group').value = ''; save('group', ''); $('f-driver').value = ''; $('f-device').value = ''; $('f-type').value = ''; renderFilters(); filterChanged(); };
    $('cs-filter').oninput = function () { if (data) renderList(); };
    $('cs-only-todo').onchange = function () { view.onlyTodo = this.checked; if (data) renderList(); };
    $('cs-csv').onclick = exportCsv;
    $('cs-types').onclick = function (e) { var b = e.target.closest('[data-type]'); if (b) setFilter('f-type', b.dataset.type); };
    $('cs-rank-by').onclick = function (e) {
      var b = e.target.closest('button[data-by]'); if (!b) return;
      view.rankBy = b.dataset.by;
      this.querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === b); });
      renderRank();
    };
    $('cs-rank').onclick = function (e) {
      var tr = e.target.closest('tr[data-id]'); if (!tr) return;
      setFilter(view.rankBy === 'driver' ? 'f-driver' : 'f-device', tr.dataset.id);
    };
    $('cs-events').onclick = function (e) {
      var cb = e.target.closest('[data-ack]');
      if (cb) { setMark(cb.dataset.ack, { ack: cb.checked }); renderKpis(); cb.closest('tr').classList.toggle('acked', cb.checked); return; }
      var b = e.target.closest('[data-open]');
      if (b) { openDetail(+b.dataset.open, listRows.slice(0, 500)); return; }
      var tr = e.target.closest('tr[data-idx]'); if (tr) focusEvent(listRows[+tr.dataset.idx]);
    };
    $('cs-d-close').onclick = closeDetail;
    $('cs-d-prev').onclick = function () { if (detailIdx > 0) { detailIdx--; renderDetail(); } };
    $('cs-d-next').onclick = function () { if (detailIdx < detailList.length - 1) { detailIdx++; renderDetail(); } };
    $('cs-detail').onclick = function (e) { if (e.target === this) closeDetail(); };
    $('cs-d-ack').onchange = function () {
      var ev = detailList[detailIdx]; if (!ev) return;
      setMark(ev.id, { ack: this.checked }); renderKpis(); renderList();
    };
    $('cs-d-memo').onchange = function () {
      var ev = detailList[detailIdx]; if (!ev) return;
      setMark(ev.id, { memo: this.value.trim() }); renderList();
    };
    document.addEventListener('keydown', function (e) {
      if ($('cs-detail').hidden || e.target.tagName === 'INPUT') { if (e.key === 'Escape') { closeDetail(); $('cs-settings').hidden = true; } return; }
      if (e.key === 'Escape') closeDetail();
      else if (e.key === 'ArrowLeft') $('cs-d-prev').click();
      else if (e.key === 'ArrowRight') $('cs-d-next').click();
    });
    var rz = null;
    window.addEventListener('resize', function () {
      if (!$('cs-detail').hidden) placeModal($('cs-detail'));
      if (!$('cs-settings').hidden) placeModal($('cs-settings'));
      clearTimeout(rz); rz = setTimeout(function () { fitHeight(); if (data) { renderTrend(); drawMap(); } }, 150);
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
    bindMap();
  }

  var bound = false;
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
      savedGroup = load('group', '') || null;
      var ly = load('layers', null); if (ly) { $('l-tracks').checked = ly.tracks !== false; $('l-heat').checked = ly.heat !== false; $('l-points').checked = !!ly.points; }
      if (!bound) { bind(); bound = true; }
      fitHeight();
    });
  }

  return {
    initialize: function (a, st, callback) { state = st; start(a, false).then(function () { callback(); }, function () { callback(); }); },
    focus: function (a, st) { api = a; if (st) state = st; refresh(); },
    blur: function () { closeDetail(); },
    _standalone: function () { document.body.style.margin = '0'; start(null, true).then(refresh); }
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
