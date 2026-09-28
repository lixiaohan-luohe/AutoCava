(function () {
  'use strict';

  var STORAGE_KEY = 'autocava.ai-report.browser-refresh.v1';
  var SNAPSHOT_VERSION = 2;
  var BRAND_CN = {
    Nissan: '日产（Nissan）', Kia: '起亚（Kia）', Chevrolet: '雪佛兰（Chevrolet）',
    Volkswagen: '大众（Volkswagen）', MG: '名爵（MG）', Mazda: '马自达（Mazda）',
    Geely: '吉利（Geely）', Toyota: '丰田（Toyota）', Changan: '长安（Changan）',
    Honda: '本田（Honda）', Hyundai: '现代（Hyundai）', Dodge: '道奇（Dodge）'
  };
  var KEYWORDS = [
    ['金融/贷款/融资', /financ|cr[eé]dito|loan|pr[eé]stamo|enganche|mensualidad|tasa|inter[eé]s|banco|bancos|plazo|pago|pagar|bur[oó]|金融|贷款|融资|首付|月供|利率|银行|征信/i],
    ['车型推荐/买什么', /conviene|recomienda|mejor|comprar|compra|auto|carro|coche|veh[ií]culo|sed[aá]n|suv|pickup|camioneta|versa|aveo|k3|mg5|yaris|sentra|virtus|mazda|nissan|kia|chevrolet|toyota|volkswagen|买什么|推荐|车型|车/i],
    ['新能源/政策/新车/降价', /el[eé]ctrico|h[ií]brido|hybrid|ev|新能源|电动|混动|政策|新车|降价/i],
    ['月供', /mensualidad|mensual|mes|monthly|月供/i],
    ['首付', /enganche|anticipo|down\s*payment|首付/i],
    ['车型对比/哪个好', /compar|versus| vs |cu[aá]l.*mejor|mejor.*cu[aá]l|哪个好|对比|比较/i],
    ['预算/收入/贷款额度', /presupuesto|ingreso|salario|sueldo|gano|ganar|capacidad|budget|income|预算|收入|工资|额度|能贷/i],
    ['银行/方案对比', /bbva|mstar|santander|banorte|banco|bancos|方案|银行/i],
    ['价格/报价/库存/优惠', /precio|cotiza|descuento|promoci[oó]n|oferta|inventario|stock|价格|报价|优惠|库存/i]
  ];

  function esc(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function text(value) { return value == null ? '' : String(value).trim(); }
  function nonempty(value) { var v = text(value); return v !== '' && v.toLowerCase() !== 'nan'; }
  function fmt(value) { return Math.round(Number(value) || 0).toLocaleString('en-US'); }
  function pct(num, den, digits) { return den ? Number((num / den * 100).toFixed(digits == null ? 1 : digits)) : 0; }
  function signed(value) { return (value >= 0 ? '+' : '') + value.toFixed(1); }
  function unique(values) { return new Set(values.filter(nonempty).map(text)); }

  function utcDay(year, month, day) { return Date.UTC(year, month - 1, day); }
  function parseTimestamp(value) {
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes(), value.getSeconds());
    }
    if (typeof value === 'number' && window.XLSX && XLSX.SSF) {
      var decoded = XLSX.SSF.parse_date_code(value);
      if (decoded) return Date.UTC(decoded.y, decoded.m - 1, decoded.d, decoded.H || 0, decoded.M || 0, Math.floor(decoded.S || 0));
    }
    var match = text(value).match(/(20\d{2})[-\/]?(\d{1,2})[-\/]?(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (!match) return NaN;
    return Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
  }

  function parsePeriod(name) {
    var match = name.match(/(20\d{2})-(\d{2})-(\d{2})\s*(?:至|~|_to_)\s*(20\d{2})-(\d{2})-(\d{2})/i);
    if (!match) match = name.match(/(20\d{2})(\d{2})(\d{2})-(20\d{2})(\d{2})(\d{2})/);
    if (!match) return null;
    var start = utcDay(+match[1], +match[2], +match[3]);
    var end = utcDay(+match[4], +match[5], +match[6]);
    if (!isFinite(start) || !isFinite(end) || end < start) return null;
    return { start: start, end: end };
  }

  function iso(day) { return new Date(day).toISOString().slice(0, 10); }
  function short(day) { return iso(day).slice(5); }
  function rowLabel(period) { return short(period.start) + '~' + short(period.end); }
  function fullLabel(period) { return iso(period.start) + ' 至 ' + iso(period.end); }
  function mergeOtherIntoFinanceForPeriod(period) {
    return iso(period.start) === '2026-09-14' && iso(period.end) === '2026-09-20';
  }

  async function rowsFromFile(file) {
    var workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    if (!workbook.SheetNames.length) throw new Error(file.name + ' 没有工作表');
    var rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: null, raw: true });
    return rows.map(function (row) {
      var clean = {};
      Object.keys(row).forEach(function (key) { clean[text(key)] = row[key]; });
      return clean;
    });
  }

  function requireColumns(rows, required, fileName) {
    var columns = rows.length ? Object.keys(rows[0]) : [];
    var missing = required.filter(function (key) { return columns.indexOf(key) < 0; });
    if (!rows.length) throw new Error(fileName + ' 没有数据行');
    if (missing.length) throw new Error(fileName + ' 缺少字段：' + missing.join('、'));
  }

  function inPeriod(value, period) {
    var timestamp = parseTimestamp(value);
    return isFinite(timestamp) && timestamp >= period.start && timestamp < period.end + 86400000;
  }

  function truthyMetric(value) {
    var v = text(value).toLowerCase();
    return v !== '' && v !== '0' && v !== '0.0' && v !== 'nan' && v !== 'false';
  }

  function cardTouched(row) { return nonempty(row['卡片消息内容']) || truthyMetric(row['卡片推荐类型']); }
  function hookTouched(row) { return truthyMetric(row['是否出现留资钩子列']); }

  function addSet(map, key, value) {
    if (!key || !value) return;
    if (!map[key]) map[key] = new Set();
    map[key].add(value);
  }

  function topSets(map, limit) {
    return Object.keys(map).map(function (key) { return [key, map[key].size]; })
      .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); }).slice(0, limit || 10);
  }

  function parseCards(rows) {
    var brands = {}, series = {}, models = {};
    rows.forEach(function (row) {
      var sid = text(row['会话ID']);
      var source = text(row['卡片消息内容']) + '\n' + text(row['纯文本消息内容']);
      var regex = /\[\[(FinanceCard|SeriesCard|ModelCard|OemBrandCard)::([\s\S]*?)\]\]/g;
      var match;
      while ((match = regex.exec(source))) {
        var payload;
        try { payload = JSON.parse(match[2]); } catch (error) { continue; }
        var kind = match[1];
        if (kind === 'OemBrandCard') {
          (Array.isArray(payload) ? payload : [payload]).forEach(function (item) { addSet(brands, text(item && item.brandName), sid); });
          continue;
        }
        if (kind === 'SeriesCard' || kind === 'ModelCard') {
          addSet(brands, text(payload.brandName), sid);
          addSet(series, text(payload.brandName) && text(payload.seriesName) ? text(payload.brandName) + ' ' + text(payload.seriesName) : '', sid);
          if (kind === 'ModelCard') addSet(models, text(payload.brandName) && text(payload.seriesName) && text(payload.modelName) ? text(payload.brandName) + ' ' + text(payload.seriesName) + ' ' + text(payload.modelName) : '', sid);
          continue;
        }
        var options = payload.modelOptions || [];
        var byId = {};
        options.forEach(function (item) { byId[item.modelId] = item; });
        var ids = [payload.currModelId].concat((payload.cards || []).map(function (item) { return item.modelId; })).filter(function (id) { return id != null; });
        if (!ids.length && options.length) ids = [options[0].modelId];
        ids.forEach(function (id) {
          var item = byId[id] || options[0] || {};
          addSet(brands, text(item.brandName), sid);
          addSet(series, text(item.brandName) && text(item.seriesName) ? text(item.brandName) + ' ' + text(item.seriesName) : '', sid);
          addSet(models, text(item.brandName) && text(item.seriesName) && text(item.modelName) ? text(item.brandName) + ' ' + text(item.seriesName) + ' ' + text(item.modelName) : '', sid);
        });
      }
    });
    return { brands: topSets(brands), series: topSets(series), models: topSets(models) };
  }

  function analyzeSession(sourceRows, period, mergeOtherIntoFinance) {
    var rows = sourceRows.filter(function (row) { return inPeriod(row['用户消息时间(UTC-6)'], period); });
    if (!rows.length) throw new Error(fullLabel(period) + ' 的 AI 会话表按日期过滤后没有数据');
    var sessionCounts = {}, userCounts = {}, durationBySession = {}, dailyMap = {}, intentRaw = {}, questionCounts = {};
    var cardSessions = new Set(), hookSessions = new Set(), cardUsers = new Set();
    rows.forEach(function (row) {
      var sid = text(row['会话ID']);
      var uid = text(row['用户身份标识']);
      if (sid) sessionCounts[sid] = (sessionCounts[sid] || 0) + 1;
      if (uid) userCounts[uid] = (userCounts[uid] || 0) + 1;
      if (sid && cardTouched(row)) cardSessions.add(sid);
      if (sid && hookTouched(row)) hookSessions.add(sid);
      if (uid && cardTouched(row)) cardUsers.add(uid);
      var intent = text(row['用户意图']) || '其他/空';
      intentRaw[intent] = (intentRaw[intent] || 0) + 1;
      var question = text(row['用户消息内容']);
      if (question) questionCounts[question] = (questionCounts[question] || 0) + 1;
      var userTime = parseTimestamp(row['用户消息时间(UTC-6)']);
      var aiTime = parseTimestamp(row['AI消息时间(UTC-6)']);
      if (sid && isFinite(userTime)) {
        if (!durationBySession[sid]) durationBySession[sid] = { min: userTime, max: NaN };
        durationBySession[sid].min = Math.min(durationBySession[sid].min, userTime);
        if (isFinite(aiTime)) durationBySession[sid].max = isFinite(durationBySession[sid].max) ? Math.max(durationBySession[sid].max, aiTime) : aiTime;
      }
      if (isFinite(userTime)) {
        var day = iso(Date.UTC(new Date(userTime).getUTCFullYear(), new Date(userTime).getUTCMonth(), new Date(userTime).getUTCDate()));
        if (!dailyMap[day]) dailyMap[day] = { messages: 0, users: {}, counts: {} };
        dailyMap[day].messages += 1;
        if (uid) {
          dailyMap[day].users[uid] = true;
          dailyMap[day].counts[uid] = (dailyMap[day].counts[uid] || 0) + 1;
        }
      }
    });
    var durations = Object.keys(durationBySession).map(function (key) {
      var item = durationBySession[key];
      return isFinite(item.max) ? (item.max - item.min) / 1000 : NaN;
    }).filter(function (value) { return isFinite(value) && value >= 0; });
    var sessionValues = Object.keys(sessionCounts).map(function (key) { return sessionCounts[key]; });
    var depth = [
      ['1轮', sessionValues.filter(function (v) { return v === 1; }).length],
      ['2轮', sessionValues.filter(function (v) { return v === 2; }).length],
      ['3-5轮', sessionValues.filter(function (v) { return v >= 3 && v <= 5; }).length],
      ['6轮以上', sessionValues.filter(function (v) { return v >= 6; }).length]
    ];
    var other = (intentRaw['其他/空'] || 0) + (intentRaw['其他'] || 0);
    var intents = [['金融', intentRaw['金融'] || 0], ['选车', intentRaw['选车'] || 0], ['其他/空', other]];
    if (mergeOtherIntoFinance) { intents[0][1] += other; intents[2][1] = 0; }
    var questions = Object.keys(questionCounts).map(function (key) { return [key, questionCounts[key]]; })
      .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); }).slice(0, 15);
    var messagesText = rows.map(function (row) { return text(row['用户消息内容']); });
    var demand = KEYWORDS.map(function (item) { return [item[0], messagesText.filter(function (value) { return item[1].test(value); }).length]; });
    var cards = parseCards(rows);
    var daily = Object.keys(dailyMap).sort().map(function (day) {
      var item = dailyMap[day];
      var users = Object.keys(item.users).length;
      var deepUsers = Object.keys(item.counts).filter(function (uid) { return item.counts[uid] >= 2; }).length;
      return { day: day.slice(5), messages: item.messages, users: users, deepUsers: deepUsers, deepPct: pct(deepUsers, users) };
    });
    return {
      period: period, rows: rows, messages: rows.length, sessions: Object.keys(sessionCounts).length,
      users: Object.keys(userCounts).length, userIds: Object.keys(userCounts), cardUserIds: Array.from(cardUsers),
      avgTurns: Object.keys(userCounts).length ? rows.length / Object.keys(userCounts).length : 0,
      avgDuration: durations.length ? durations.reduce(function (a, b) { return a + b; }, 0) / durations.length : 0,
      reachSessions: cardSessions.size, leadSessions: hookSessions.size,
      multiSessions: depth[1][1] + depth[2][1] + depth[3][1], intents: intents, depth: depth,
      demand: demand, questions: questions, brands: cards.brands, series: cards.series, models: cards.models, daily: daily
    };
  }

  function analyzeLeads(sourceRows, period, session) {
    var hasDate = sourceRows.length && Object.prototype.hasOwnProperty.call(sourceRows[0], '创建时间(utc)');
    var rows = hasDate ? sourceRows.filter(function (row) { return inPeriod(row['创建时间(utc)'], period); }) : sourceRows.slice();
    if (!rows.length) throw new Error(fullLabel(period) + ' 的主线索表按日期过滤后没有数据');
    var leadUuids = unique(rows.map(function (row) { return row.UUID; }));
    var sessionUsers = new Set(session.userIds);
    var cardUsers = new Set(session.cardUserIds);
    var matched = Array.from(leadUuids).filter(function (uid) { return sessionUsers.has(uid); });
    var matchedSet = new Set(matched);
    var matchedTurns = session.rows.filter(function (row) { return matchedSet.has(text(row['用户身份标识'])); }).length;
    var matchedCardUsers = Array.from(leadUuids).filter(function (uid) { return cardUsers.has(uid); }).length;
    var aiSource = /AI|(^|[^0-9])(17|24|25|26|27|28|29|30|31)([^0-9]|$)/i;
    var aiTotal = rows.filter(function (row) { return aiSource.test(text(row['来源类型'])); }).length;
    function topColumn(column, limit) {
      var counts = {};
      rows.forEach(function (row) { var key = text(row[column]) || '空'; counts[key] = (counts[key] || 0) + 1; });
      return Object.keys(counts).map(function (key) { return [key, counts[key]]; })
        .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); }).slice(0, limit || 6);
    }
    return {
      period: period, total: rows.length, uniqueMobile: unique(rows.map(function (row) { return row['手机号']; })).size,
      uniqueUuid: leadUuids.size, matchedUsers: matched.length, matchedTurns: matchedTurns,
      matchedAvgTurns: matched.length ? matchedTurns / matched.length : 0,
      cardLeadUsers: matchedCardUsers, cardLeadRate: pct(matchedCardUsers, cardUsers.size, 2),
      aiTotal: aiTotal, sourceTypes: topColumn('来源类型'), dealers: topColumn('经销商名称')
    };
  }

  function stripRuntime(stats) {
    var copy = {};
    Object.keys(stats).forEach(function (key) {
      if (key !== 'rows' && key !== 'userIds' && key !== 'cardUserIds') copy[key] = stats[key];
    });
    return copy;
  }

  function buildSnapshot(files, firstSessionRows, secondSessionRows, firstLeadRows, secondLeadRows) {
    var firstPeriod = parsePeriod(files[2].name);
    var secondPeriod = parsePeriod(files[3].name);
    if (!firstPeriod || !secondPeriod) throw new Error('两份主线索表文件名必须包含 YYYY-MM-DD至YYYY-MM-DD');
    if (firstPeriod.start >= secondPeriod.start) throw new Error('第一周的日期必须早于第二周');
    if (firstPeriod.end >= secondPeriod.start) throw new Error('第一周和第二周日期不能重叠');
    var sessionRequired = ['会话ID', '用户身份标识', '用户意图', '用户消息内容', '卡片推荐类型', '是否出现留资钩子列', '卡片消息内容', '用户消息时间(UTC-6)', 'AI消息时间(UTC-6)'];
    var leadRequired = ['UUID', '手机号', '来源类型'];
    requireColumns(firstSessionRows, sessionRequired, files[0].name);
    requireColumns(secondSessionRows, sessionRequired, files[1].name);
    requireColumns(firstLeadRows, leadRequired, files[2].name);
    requireColumns(secondLeadRows, leadRequired, files[3].name);
    var first = analyzeSession(firstSessionRows, firstPeriod, mergeOtherIntoFinanceForPeriod(firstPeriod));
    var second = analyzeSession(secondSessionRows, secondPeriod, mergeOtherIntoFinanceForPeriod(secondPeriod));
    var firstLeads = analyzeLeads(firstLeadRows, firstPeriod, first);
    var secondLeads = analyzeLeads(secondLeadRows, secondPeriod, second);
    var combined = analyzeSession(first.rows.concat(second.rows), { start: firstPeriod.start, end: secondPeriod.end }, false);
    return {
      version: SNAPSHOT_VERSION, createdAt: new Date().toISOString(),
      files: files.map(function (file) { return file.name; }),
      first: stripRuntime(first), second: stripRuntime(second), combined: stripRuntime(combined),
      firstLeads: firstLeads, secondLeads: secondLeads
    };
  }

  function overviewPayload(stats, leads) {
    var finance = stats.intents.filter(function (item) { return item[0] === '金融'; })[0][1];
    return {
      users: [fmt(stats.users), '本地 AI Session 去重用户'],
      messages: [pct(stats.multiSessions, stats.sessions, 2).toFixed(2) + '%', fmt(stats.multiSessions) + ' / ' + fmt(stats.sessions) + ' 会话'],
      bounce: [pct(stats.depth[0][1], stats.sessions, 2).toFixed(2) + '%', '1轮会话 / 全部会话'],
      finance: [pct(finance, stats.messages).toFixed(1) + '%', fmt(finance) + ' / ' + fmt(stats.messages) + ' 消息（业务归类）'],
      leadHook: [fmt(leads.total), '主线索表记录 · ' + rowLabel(stats.period)]
    };
  }

  function rowMap(rows) { var out = {}; rows.forEach(function (item) { out[item[0]] = item[1]; }); return out; }
  function tableRows(rows, den, valueHeaderClass) {
    return rows.map(function (item) { return '<tr><td style="text-align:left">' + esc(item[0]) + '</td><td class="' + (valueHeaderClass || 'hl') + '">' + fmt(item[1]) + ' · ' + pct(item[1], den).toFixed(1) + '%</td></tr>'; }).join('');
  }

  function demandRows(stats) {
    var sorted = stats.demand.slice().sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); });
    var max = Math.max.apply(null, sorted.map(function (item) { return item[1]; }).concat([1]));
    return sorted.map(function (item, index) {
      return '<tr><td>' + (index + 1) + '</td><td style="text-align:left">' + esc(item[0]) + '</td><td class="hl">' + fmt(item[1]) + '</td><td>' + pct(item[1], stats.messages).toFixed(1) + '%</td><td><div class="mini-track"><div class="mini-fill" style="width:' + pct(item[1], max) + '%;background:var(--s1)"></div></div></td></tr>';
    }).join('');
  }

  function questionJudge(question) {
    if (/mensualidad|enganche|cr[eé]dito|financ|banco|pr[eé]stamo/i.test(question)) return '<span class="tag">已命中</span>金融场景';
    if (/plan|conviene|atributos|versi[oó]n|compar|auto|carro|suv|sed[aá]n/i.test(question)) return '<span class="tag">部分命中</span>选车/方案场景';
    return '<span class="tag gap">需核对</span>入口/手打待拆分';
  }

  function questionRows(stats) {
    return stats.questions.map(function (item, index) {
      return '<tr><td>' + (index + 1) + '</td><td style="text-align:left">' + esc(item[0]) + '</td><td>' + fmt(item[1]) + '</td><td><div class="match">' + questionJudge(item[0]) + '</div></td></tr>';
    }).join('');
  }

  function vehicleRows(rows, sessions, translateBrand) {
    return rows.map(function (item, index) {
      var name = translateBrand ? (BRAND_CN[item[0]] || item[0]) : item[0];
      return '<tr><td>' + (index + 1) + '</td><td style="text-align:left">' + esc(name) + '</td><td class="hl">' + fmt(item[1]) + '</td><td>' + pct(item[1], sessions).toFixed(1) + '%</td></tr>';
    }).join('');
  }

  function chartSvg(daily, metric, colorClass, gradientId, guideId, aria) {
    var width = 860, height = 320, left = 50, right = 16, top = 22, bottom = 36;
    var innerW = width - left - right, innerH = height - top - bottom;
    var values = daily.map(function (d) { return Number(d[metric]) || 0; });
    var max = Math.max.apply(null, values.concat([1]));
    var xs = values.map(function (_, i) { return left + (values.length <= 1 ? innerW / 2 : innerW * i / (values.length - 1)); });
    var ys = values.map(function (v) { return top + innerH - innerH * v / max; });
    var points = xs.map(function (x, i) { return x.toFixed(1) + ',' + ys[i].toFixed(1); }).join(' ');
    var area = points + ' ' + (xs[xs.length - 1] || left) + ',' + (top + innerH) + ' ' + (xs[0] || left) + ',' + (top + innerH);
    var labels = daily.map(function (d, i) { return '<text class="xlab" x="' + xs[i].toFixed(1) + '" y="304">' + esc(i === 0 || d.day.slice(3) === '01' ? d.day : d.day.slice(3)) + '</text>'; }).join('');
    var circles = values.map(function (v, i) { return '<circle class="mk" cx="' + xs[i].toFixed(1) + '" cy="' + ys[i].toFixed(1) + '" r="3.5" fill="var(--' + colorClass + ')"/>'; }).join('');
    return '<svg viewBox="0 0 860 320" role="img" aria-label="' + esc(aria) + '"><defs><linearGradient id="' + gradientId + '" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="var(--' + colorClass + ')"/><stop offset="1" stop-color="var(--' + colorClass + ')" stop-opacity="0"/></linearGradient></defs><polygon points="' + area + '" fill="url(#' + gradientId + ')" opacity=".18"/><polyline points="' + points + '" fill="none" stroke="var(--' + colorClass + ')" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round"/>' + circles + '<g>' + labels + '</g><line class="guide" id="' + guideId + '" x1="0" y1="14" x2="0" y2="284"/></svg>';
  }

  function bindChart(wrap, daily, lines) {
    if (!wrap) return;
    var svg = wrap.querySelector('svg'), guide = wrap.querySelector('.guide'), tip = wrap.querySelector('.dctip');
    if (!svg || !guide || !tip || !daily.length) return;
    wrap.addEventListener('mousemove', function (event) {
      var rect = svg.getBoundingClientRect();
      var virtualX = (event.clientX - rect.left) * 860 / rect.width;
      var left = 50, right = 16, inner = 860 - left - right;
      var index = Math.max(0, Math.min(daily.length - 1, Math.round((virtualX - left) / inner * Math.max(1, daily.length - 1))));
      var x = left + (daily.length <= 1 ? inner / 2 : inner * index / (daily.length - 1));
      guide.setAttribute('x1', x); guide.setAttribute('x2', x);
      tip.style.left = (x / 860 * 100) + '%';
      tip.innerHTML = '<b>' + esc(daily[index].day) + '</b>' + lines.map(function (line) { return '<br><span style="color:var(--' + line.color + ');font-weight:700">● ' + esc(line.label) + ' ' + esc(line.value(daily[index])) + '</span>'; }).join('');
      wrap.classList.add('hover');
    });
    wrap.addEventListener('mouseleave', function () { wrap.classList.remove('hover'); });
  }

  function setNodeHtml(id, html) { var node = document.getElementById(id); if (node) node.innerHTML = html; }
  function findSection(label) { return Array.from(document.querySelectorAll('section')).find(function (section) { return section.textContent.indexOf(label) >= 0; }); }

  function applySnapshot(snapshot, restored) {
    var first = snapshot.first, second = snapshot.second, combined = snapshot.combined;
    var firstLeads = snapshot.firstLeads, secondLeads = snapshot.secondLeads;
    var headerBadge = document.querySelector('.date-badge');
    if (headerBadge) headerBadge.textContent = iso(first.period.start) + ' — ' + iso(second.period.end);
    var overview = { current: overviewPayload(second, secondLeads), previous: overviewPayload(first, firstLeads), total: overviewPayload(combined, { total: firstLeads.total + secondLeads.total }) };
    var overviewTabs = document.getElementById('overviewTabs');
    if (overviewTabs) {
      var buttons = overviewTabs.querySelectorAll('button[data-view]');
      if (buttons[0]) buttons[0].textContent = '第二周 · ' + rowLabel(second.period);
      if (buttons[1]) buttons[1].textContent = '第一周 · ' + rowLabel(first.period);
      if (buttons[2]) buttons[2].textContent = '两周综合 · ' + rowLabel(combined.period);
      overviewTabs.onclick = function (event) {
        var button = event.target.closest('button[data-view]'); if (!button) return;
        setTimeout(function () { renderOverview(overview[button.getAttribute('data-view')]); }, 0);
      };
    }
    function renderOverview(data) {
      Object.keys(data).forEach(function (key) {
        var number = document.querySelector('[data-key="' + key + '"]');
        var note = document.querySelector('[data-note="' + key + '"]');
        if (number) number.textContent = data[key][0]; if (note) note.textContent = data[key][1];
      });
      var userCard = document.querySelector('[data-key="users"]');
      if (userCard && userCard.nextElementSibling) userCard.nextElementSibling.textContent = '本地去重AI用户数';
    }
    renderOverview(overview.current);

    var weekly = findSection('周趋势对比');
    function weeklyRow(label, stats, leads) {
      return '<tr><td>' + esc(label) + '</td><td>' + fmt(stats.messages) + '</td><td>' + fmt(stats.sessions) + '</td><td class="hl">' + stats.avgTurns.toFixed(1) + '</td><td>' + stats.avgDuration.toFixed(0) + '秒</td><td>' + fmt(stats.reachSessions) + '</td><td>' + pct(stats.reachSessions, stats.sessions).toFixed(1) + '%</td><td>' + fmt(stats.leadSessions) + '</td><td>' + fmt(leads.total) + '</td><td>' + fmt(leads.aiTotal) + '</td></tr>';
    }
    if (weekly) {
      var weeklyBody = weekly.querySelector('tbody');
      if (weeklyBody) weeklyBody.innerHTML = weeklyRow('第二周 · ' + rowLabel(second.period), second, secondLeads) + weeklyRow('第一周 · ' + rowLabel(first.period), first, firstLeads) + weeklyRow('两周综合 · ' + rowLabel(combined.period), combined, { total: firstLeads.total + secondLeads.total, aiTotal: firstLeads.aiTotal + secondLeads.aiTotal });
      var insight = weekly.querySelector('.insight');
      if (insight) insight.innerHTML = '<b>第二周表现：</b>' + rowLabel(second.period) + ' 本地 AI Session 有 ' + fmt(second.sessions) + ' 个去重会话、' + fmt(second.messages) + ' 条消息；主线索表 ' + fmt(secondLeads.total) + ' 条，较第一周变化 ' + signed(pct(secondLeads.total - firstLeads.total, firstLeads.total)) + '%。消息数较第一周变化 ' + signed(pct(second.messages - first.messages, first.messages)) + '%。';
    }

    var daily = first.daily.concat(second.daily);
    var dailySection = findSection('每日趋势');
    if (dailySection) {
      var date = dailySection.querySelector('.section-date'); if (date) date.textContent = iso(first.period.start) + ' — ' + iso(second.period.end);
      var firstCard = dailySection.querySelector('#daychart');
      if (firstCard) {
        var scope = firstCard.querySelector('.scope'); if (scope) scope.textContent = '每日消息数 vs 本地去重AI用户数';
        var legend = firstCard.querySelectorAll('.dcl .item'); if (legend[1]) legend[1].lastChild.textContent = '本地去重AI用户数';
        var oldWrap = firstCard.querySelector('#dcwrap');
        if (oldWrap) { var newWrap = document.createElement('div'); newWrap.className = 'dcwrap'; newWrap.id = 'dcwrap'; newWrap.innerHTML = chartSvg(daily, 'messages', 's1', 'localGradMessages', 'localGuideMessages', '每日消息数趋势') + '<div class="dctip"></div>'; oldWrap.replaceWith(newWrap); bindChart(newWrap, daily, [{ color: 's1', label: '消息数', value: function (d) { return fmt(d.messages); } }, { color: 's3', label: '本地AI用户', value: function (d) { return fmt(d.users); } }]); }
      }
      var deepCard = dailySection.querySelector('#dcwrapDeep');
      if (deepCard) { var newDeep = document.createElement('div'); newDeep.className = 'dcwrap'; newDeep.id = 'dcwrapDeep'; newDeep.innerHTML = chartSvg(daily, 'deepPct', 's4', 'localGradDeep', 'localGuideDeep', '深度会话用户占比趋势') + '<div class="dctip"></div>'; deepCard.replaceWith(newDeep); bindChart(newDeep, daily, [{ color: 's4', label: '深度用户占比', value: function (d) { return d.deepPct.toFixed(1) + '%'; } }]); }
      var peak = second.daily.slice().sort(function (a, b) { return b.messages - a.messages; })[0];
      var trendInsight = dailySection.querySelector('.insight'); if (peak && trendInsight) trendInsight.innerHTML = '<b>趋势解读：</b>第二周最高消息数出现在 ' + esc(peak.day) + '，为 ' + fmt(peak.messages) + ' 条；当日本地去重 AI 用户 ' + fmt(peak.users) + ' 人。';
    }

    var intentSection = findSection('用户意图与会话深度');
    if (intentSection) {
      var bodies = intentSection.querySelectorAll('tbody');
      if (bodies[0]) {
        var firstIntent = rowMap(first.intents);
        bodies[0].innerHTML = second.intents.map(function (item) { var previous = firstIntent[item[0]] || 0; return '<tr><td style="text-align:left">' + esc(item[0]) + '</td><td class="hl">' + fmt(item[1]) + ' · ' + pct(item[1], second.messages).toFixed(1) + '%</td><td>' + fmt(previous) + ' · ' + pct(previous, first.messages).toFixed(1) + '%</td><td>' + signed(pct(item[1], second.messages) - pct(previous, first.messages)) + ' pct</td></tr>'; }).join('');
      }
      if (bodies[1]) {
        var firstDepth = rowMap(first.depth);
        bodies[1].innerHTML = second.depth.map(function (item) { var previous = firstDepth[item[0]] || 0; return '<tr><td style="text-align:left">' + esc(item[0]) + '</td><td class="hl">' + fmt(item[1]) + ' · ' + pct(item[1], second.sessions).toFixed(1) + '%</td><td>' + fmt(previous) + ' · ' + pct(previous, first.sessions).toFixed(1) + '%</td><td>' + signed(pct(item[1], second.sessions) - pct(previous, first.sessions)) + ' pct</td></tr>'; }).join('');
      }
    }

    var dynamic = {
      demand: { current: demandRows(second), previous: demandRows(first) },
      questions: { current: questionRows(second), previous: questionRows(first) },
      vehicle: {
        current: { brand: vehicleRows(second.brands, second.sessions, true), series: vehicleRows(second.series, second.sessions), model: vehicleRows(second.models, second.sessions) },
        previous: { brand: vehicleRows(first.brands, first.sessions, true), series: vehicleRows(first.series, first.sessions), model: vehicleRows(first.models, first.sessions) }
      }
    };
    setNodeHtml('demandBody', dynamic.demand.current); setNodeHtml('questionBody', dynamic.questions.current);
    setNodeHtml('vehicleBrandBody', dynamic.vehicle.current.brand); setNodeHtml('vehicleSeriesBody', dynamic.vehicle.current.series); setNodeHtml('vehicleModelBody', dynamic.vehicle.current.model);
    function bindDynamicTabs(tabId, render) {
      var tabs = document.getElementById(tabId); if (!tabs) return;
      tabs.addEventListener('click', function (event) { var button = event.target.closest('button[data-view]'); if (!button) return; setTimeout(function () { render(button.getAttribute('data-view')); }, 0); });
    }
    bindDynamicTabs('demandTabs', function (view) { setNodeHtml('demandBody', dynamic.demand[view] || ''); });
    bindDynamicTabs('questionTabs', function (view) { setNodeHtml('questionBody', dynamic.questions[view] || ''); });
    bindDynamicTabs('vehicleTabs', function (view) { var data = dynamic.vehicle[view] || {}; setNodeHtml('vehicleBrandBody', data.brand || ''); setNodeHtml('vehicleSeriesBody', data.series || ''); setNodeHtml('vehicleModelBody', data.model || ''); });

    var coverage = findSection('指标覆盖情况');
    if (coverage) {
      var finance = rowMap(second.demand)['金融/贷款/融资'] || 0;
      var coverageBody = coverage.querySelector('tbody');
      if (coverageBody) coverageBody.innerHTML = [
        ['对话深度', '平均' + second.avgTurns.toFixed(1) + '轮；1轮占' + pct(second.depth[0][1], second.sessions).toFixed(1) + '%', '按会话ID聚合有效用户消息数。'],
        ['平均会话轮次', second.avgTurns.toFixed(1), '有效用户消息行数 / 去重AI用户数。'],
        ['平均会话时长', '约' + second.avgDuration.toFixed(0) + '秒', '同一会话内最后 AI 时间 - 最早用户消息时间估算。'],
        ['真实线索用户平均对话轮次', secondLeads.matchedAvgTurns.toFixed(1), '主线索表 UUID 匹配到 AI Session 用户身份标识后计算。'],
        ['金融需求命中', fmt(finance) + ' 条 · ' + pct(finance, second.messages).toFixed(1) + '%', fmt(finance) + ' / ' + fmt(second.messages) + ' 条用户消息。'],
        ['对话解决率（UUID真实线索）', fmt(secondLeads.matchedUsers) + ' / ' + fmt(second.users) + '（' + pct(secondLeads.matchedUsers, second.users, 2).toFixed(2) + '%）', '主线索 UUID 匹配用户 / AI 会话去重用户。'],
        ['卡片触达后解决率', secondLeads.cardLeadRate.toFixed(2) + '%', '卡片触达用户中，UUID 出现在主线索表的用户占比。'],
        ['多轮对话完成率', fmt(second.multiSessions) + ' / ' + fmt(second.sessions) + '（' + pct(second.multiSessions, second.sessions).toFixed(1) + '%）', '同一会话内用户有效提问 ≥2 次。'],
        ['AI跳出率', pct(second.depth[0][1], second.sessions, 2).toFixed(2) + '%', '1轮会话 / 全部会话。']
      ].map(function (row) { return '<tr><td style="text-align:left">' + esc(row[0]) + '</td><td class="hl">' + esc(row[1]) + '</td><td>' + esc(row[2]) + '</td><td><span class="badge good">已更新</span></td></tr>'; }).join('');
    }

    var footer = document.querySelector('footer.foot');
    if (footer) footer.innerHTML = '<p><b>数据口径：</b>AI Session 消息数按日期区间内有效行计；会话数按“会话ID”去重；用户按“用户身份标识”去重。真实线索数为主线索表日期区间内记录行数；用户级线索闭环按 <code>UUID = 用户身份标识</code> 匹配。</p><p><b>数据来源：</b>' + snapshot.files.map(function (name) { return '<code>' + esc(name) + '</code>'; }).join('、') + '。</p><p><b>统计周期：</b>' + fullLabel(first.period) + '；' + fullLabel(second.period) + '。AI 会话表已按主线索表日期边界过滤，避免两周重叠。</p><p>AI 购车顾问 · 墨西哥市场 · 浏览器本地刷新 ' + esc(new Date(snapshot.createdAt).toLocaleString('zh-CN', { hour12: false })) + '</p>';

    var status = document.getElementById('refreshStatus');
    if (status) { status.textContent = restored ? '已恢复本机上次刷新结果' : '报告已刷新并保存在本浏览器'; status.className = 'refresh-status success'; }
    document.documentElement.setAttribute('data-local-report-refreshed', 'true');
  }

  function installUi() {
    var style = document.createElement('style');
    style.textContent = '.wizard-actions{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:18px}.wizard-primary,.wizard-secondary{appearance:none;font:inherit;font-size:14px;font-weight:700;border-radius:8px;padding:9px 15px;cursor:pointer}.wizard-primary{border:1px solid var(--accent);background:var(--accent);color:#fff}.wizard-primary:disabled{opacity:.45;cursor:not-allowed}.wizard-secondary{border:1px solid var(--line2);background:var(--surface);color:var(--muted)}.refresh-status{display:block;margin-top:12px;font-size:13px;color:var(--muted)}.refresh-status.error{color:var(--crit)}.refresh-status.success{color:var(--good)}.upload-card.ready{border-style:solid;border-color:var(--good)}';
    document.head.appendChild(style);
    var inputs = Array.from(document.querySelectorAll('[data-upload-step]'));
    if (inputs.length !== 4) return;
    inputs.forEach(function (input, index) { input.setAttribute('data-upload-role', ['first-session', 'second-session', 'first-leads', 'second-leads'][index]); });
    var headCopy = document.querySelector('.wizard-head p');
    if (headCopy) headCopy.textContent = '选择 4 份 Excel 后，网页会在当前浏览器中本地解析和重算报告。文件不会上传到服务器。';
    var note = document.querySelector('.wizard-note');
    if (note) note.innerHTML = '第一周和第二周的日期以两份主线索表文件名为准。AI 会话数据会按对应日期过滤，因此文件含下一周首日也不会重复计算。';
    var actions = document.createElement('div'); actions.className = 'wizard-actions';
    actions.innerHTML = '<button class="wizard-primary" id="runBrowserRefresh" type="button" disabled>生成并刷新报告</button><button class="wizard-secondary" id="clearBrowserRefresh" type="button">恢复网页发布版</button><span id="refreshStatus" class="refresh-status">请先选择 4 份文件</span>';
    var main = document.querySelector('.wizard-main'); if (main) main.appendChild(actions);
    var run = document.getElementById('runBrowserRefresh'), status = document.getElementById('refreshStatus');
    function refreshInputState() {
      inputs.forEach(function (input) { var card = input.closest('.upload-card'); if (card) card.classList.toggle('ready', !!(input.files && input.files[0])); });
      run.disabled = !inputs.every(function (input) { return input.files && input.files[0]; });
      if (!run.disabled) { status.textContent = '四份文件已齐全，可以刷新'; status.className = 'refresh-status'; }
    }
    inputs.forEach(function (input) { input.addEventListener('change', refreshInputState); });
    run.addEventListener('click', async function () {
      run.disabled = true; status.textContent = '正在读取和计算…'; status.className = 'refresh-status';
      try {
        if (!window.XLSX) throw new Error('Excel 解析组件加载失败，请刷新页面后重试');
        var files = inputs.map(function (input) { return input.files[0]; });
        var data = await Promise.all(files.map(rowsFromFile));
        var snapshot = buildSnapshot(files, data[0], data[1], data[2], data[3]);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
        applySnapshot(snapshot, false);
        var mask = document.getElementById('wizardMask'); if (mask) { mask.classList.remove('open'); mask.setAttribute('aria-hidden', 'true'); }
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } catch (error) {
        console.error(error); status.textContent = '刷新失败：' + (error && error.message ? error.message : String(error)); status.className = 'refresh-status error';
      } finally { run.disabled = false; }
    });
    document.getElementById('clearBrowserRefresh').addEventListener('click', function () { localStorage.removeItem(STORAGE_KEY); location.reload(); });
  }

  function restoreSaved() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY); if (!raw) return;
      var snapshot = JSON.parse(raw); if (snapshot.version !== SNAPSHOT_VERSION) return;
      applySnapshot(snapshot, true);
    } catch (error) { console.warn('Unable to restore local report snapshot', error); }
  }

  function start() { installUi(); restoreSaved(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  window.__AIReportRefresh = { parsePeriod: parsePeriod, parseTimestamp: parseTimestamp, analyzeSession: analyzeSession, analyzeLeads: analyzeLeads, buildSnapshot: buildSnapshot, applySnapshot: applySnapshot };
})();
