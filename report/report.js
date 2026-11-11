(() => {
  'use strict';
  const data = JSON.parse(document.getElementById('report-data').textContent);
  const $ = id => document.getElementById(id);
  const number = value => Number(value).toLocaleString('en-US');
  const GROUP_PAGE_SIZE = 12, ROW_PAGE_SIZE = 100;
  const state = {sheet: 0, query: '', sort: 'count', page: 0};
  const cache = new Map();
  const monthNames = {'一月':1,'二月':2,'三月':3,'四月':4,'五月':5,'六月':6,'七月':7,'八月':8,'九月':9,'十月':10,'十一月':11,'十二月':12,
    jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,sep:9,sept:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12};
  function parseReportTime(value) {
    const raw = String(value ?? ''), text = raw.replace(/\s+/g, '').replace(/，/g, ',');
    let match, month, day, year, hour, minute, second, period;
    if ((match = text.match(/^([^\d,\/]+)(\d{1,2}),?(\d{4})(\d{1,2}):(\d{2})(?::(\d{2}))?(上午|下午|AM|PM)?$/i))) {
      month = monthNames[match[1].toLowerCase()]; day = +match[2]; year = +match[3]; hour = +match[4]; minute = +match[5]; second = +(match[6] || 0); period = match[7];
    } else if ((match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(\d{1,2}):(\d{2})(?::(\d{2}))?(上午|下午|AM|PM)?$/i))) {
      month = +match[1]; day = +match[2]; year = +match[3]; hour = +match[4]; minute = +match[5]; second = +(match[6] || 0); period = match[7];
    } else if ((match = text.match(/^(\d{4})-(\d{2})-(\d{2})T?(\d{1,2}):(\d{2})(?::(\d{2}))?$/))) {
      year = +match[1]; month = +match[2]; day = +match[3]; hour = +match[4]; minute = +match[5]; second = +(match[6] || 0);
    } else return {raw, display: raw || '（時間空白）', valid: false, key: -Infinity};
    const invalid = () => ({raw, display: raw || '（時間空白）', valid: false, key: -Infinity});
    if (period) {
      if (hour < 1 || hour > 12) return invalid();
      hour = hour % 12 + (/^(下午|PM)$/i.test(period) ? 12 : 0);
    }
    if (!month || month > 12 || day < 1 || year < 1000 || hour < 0 || hour > 23 || minute > 59 || second > 59) return invalid();
    const key = Date.UTC(year, month - 1, day, hour, minute, second), date = new Date(key);
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return invalid();
    const pad = x => String(x).padStart(2, '0');
    return {raw, display: `${pad(month)}/${pad(day)} ${pad(hour)}:${pad(minute)}:${pad(second)}`, valid: true, key};
  }
  function node(tag, className, text) {
    const element = document.createElement(tag);
    if (className) element.className = className;
    if (text !== undefined) element.textContent = text;
    return element;
  }
  function buildGroups(report) {
    const map = new Map(); let invalidTimes = 0;
    for (const row of report.rows) {
      const identifier = row.values[report.keyIndex] ?? '';
      if (!map.has(identifier)) map.set(identifier, {identifier, rows: [], latest: null});
      const group = map.get(identifier), time = parseReportTime(row.values[report.timeIndex]);
      if (!time.valid) invalidTimes++;
      if (time.valid && (!group.latest || time.key > group.latest.key)) group.latest = time;
      group.rows.push({...row, time});
    }
    const groups = [...map.values()];
    for (const group of groups) {
      group.rows.sort((a, b) => {
        if (a.time.valid !== b.time.valid) return a.time.valid ? -1 : 1;
        return a.time.valid ? a.time.key - b.time.key : 0;
      });
    }
    return {groups, invalidTimes};
  }
  function getGroups(index) {
    if (!cache.has(index)) cache.set(index, buildGroups(data.reports[index]));
    return cache.get(index);
  }
  function pager(container, page, total, size, change) {
    container.replaceChildren();
    if (total <= size) { container.hidden = true; return; }
    container.hidden = false;
    const pages = Math.ceil(total / size);
    const previous = node('button', 'page-button', '上一頁'); previous.type = 'button'; previous.disabled = page === 0;
    const next = node('button', 'page-button', '下一頁'); next.type = 'button'; next.disabled = page >= pages - 1;
    previous.addEventListener('click', () => change(page - 1)); next.addEventListener('click', () => change(page + 1));
    container.append(node('span', '', `${number(page * size + 1)}–${number(Math.min((page + 1) * size, total))} / ${number(total)}`), previous, node('span', '', `${page + 1} / ${pages}`), next);
  }
  function renderGroup(report, group, open) {
    const details = node('details', 'group-card'); details.dataset.identifier = group.identifier;
    const summary = node('summary');
    summary.append(node('span', 'group-chevron', '›'));
    const heading = node('div', 'group-heading');
    heading.append(node('div', 'group-name', group.identifier || '（識別值空白）'));
    summary.append(heading, node('span', `count-badge${group.rows.length === 1 ? ' single' : ''}`, `${number(group.rows.length)} 筆`));
    details.append(summary);
    let built = false;
    function buildContent() {
      if (built) return; built = true;
      const content = node('div', 'group-content');
      const scroll = node('div', 'table-scroll'), table = node('table'), thead = node('thead'), header = node('tr'), tbody = node('tbody');
      header.append(node('th', 'key-column', report.columns[report.keyIndex]));
      for (const c of report.detailIndices) header.append(node('th', '', report.columns[c]));
      for (const th of header.children) th.scope = 'col';
      thead.append(header); table.append(thead, tbody); scroll.append(table);
      const pagination = node('div', 'pagination'); pagination.setAttribute('aria-label', '觸發明細頁碼');
      function showRows(page) {
        tbody.replaceChildren();
        const rows = group.rows.slice(page * ROW_PAGE_SIZE, (page + 1) * ROW_PAGE_SIZE);
        for (const [index, row] of rows.entries()) {
          const tr = node('tr');
          if (index === 0) {
            const key = node('th', 'key-value', group.identifier || '（識別值空白）');
            key.scope = 'rowgroup'; key.rowSpan = rows.length;
            tr.append(key);
          }
          for (const c of report.detailIndices) {
            const isTime = c === report.timeIndex;
            const cell = node('td', isTime ? `time-cell${row.time.valid ? '' : ' time-invalid'}` : '', isTime ? row.time.display : (row.values[c] || '—'));
            if (isTime) cell.title = `${row.time.raw}${row.file ? `\n來源：${row.file}${row.page ? ` · 第 ${row.page} 頁` : ''}` : ''}`;
            tr.append(cell);
          }
          tbody.append(tr);
        }
        pager(pagination, page, group.rows.length, ROW_PAGE_SIZE, showRows); scroll.scrollTop = 0;
      }
      showRows(0); content.append(scroll, pagination); details.append(content);
    }
    details.addEventListener('toggle', () => { if (details.open) buildContent(); });
    if (open) { buildContent(); details.open = true; }
    return details;
  }
  function renderGroups() {
    const report = data.reports[state.sheet]; if (!report) return;
    const query = state.query.trim().toLocaleLowerCase();
    const groups = getGroups(state.sheet).groups.filter(group => {
      if (!query) return true;
      if (group.searchText === undefined) group.searchText = group.rows.map(row => `${row.values.join('\n')}\n${row.time.display}`).join('\n').toLocaleLowerCase();
      return group.searchText.includes(query);
    });
    groups.sort((a, b) => {
      if (state.sort === 'name') return a.identifier.localeCompare(b.identifier);
      if (state.sort === 'latest') {
        const ak = a.latest?.key ?? -Infinity, bk = b.latest?.key ?? -Infinity;
        if (ak !== bk) return bk > ak ? 1 : -1;
      }
      return b.rows.length - a.rows.length || a.identifier.localeCompare(b.identifier);
    });
    state.page = Math.min(state.page, Math.max(0, Math.ceil(groups.length / GROUP_PAGE_SIZE) - 1));
    $('groups').replaceChildren(...groups.slice(state.page * GROUP_PAGE_SIZE, (state.page + 1) * GROUP_PAGE_SIZE).map((group, i) => renderGroup(report, group, i === 0)));
    $('empty-state').hidden = groups.length > 0;
    pager($('group-pagination'), state.page, groups.length, GROUP_PAGE_SIZE, page => { state.page = page; renderGroups(); $('groups').scrollIntoView({block: 'start'}); });
  }
  function selectSheet(index) {
    state.sheet = index; state.page = 0; state.query = ''; $('search').value = '';
    const report = data.reports[index];
    if (!report) {
      $('report-title').textContent = '尚無報告資料';
      $('groups').replaceChildren(); return;
    }
    const {invalidTimes} = getGroups(index);
    for (const [i, button] of [...$('report-nav').children].entries()) { button.setAttribute('aria-selected', String(i === index)); button.tabIndex = i === index ? 0 : -1; }
    $('report-panel').setAttribute('aria-labelledby', `tab-${report.id}`);
    $('report-title').textContent = report.title;
    $('time-warning').hidden = invalidTimes === 0;
    $('time-warning').textContent = `${number(invalidTimes)} 筆時間無法可靠轉換，明細以黃色保留原文；這些紀錄仍包含在統計中。`;
    renderGroups();
  }
  $('sheet-count').textContent = data.reports.length;
  $('source-label').textContent = `資料來源：${data.workbook} · ${data.reports.length} 張報告工作表`;
  for (const [i, report] of data.reports.entries()) {
    const button = node('button', 'nav-item'); button.type = 'button'; button.id = `tab-${report.id}`; button.setAttribute('role', 'tab'); button.setAttribute('aria-controls', 'report-panel');
    button.append(node('span', 'nav-name', report.title), node('span', 'nav-count', number(report.rows.length)));
    button.addEventListener('click', () => selectSheet(i));
    button.addEventListener('keydown', event => {
      let next;
      if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (i + 1) % data.reports.length;
      if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (i + data.reports.length - 1) % data.reports.length;
      if (event.key === 'Home') next = 0; if (event.key === 'End') next = data.reports.length - 1;
      if (next !== undefined) { event.preventDefault(); selectSheet(next); $('report-nav').children[next].focus(); }
    });
    $('report-nav').append(button);
  }
  let searchTimer;
  $('search').addEventListener('input', event => { clearTimeout(searchTimer); state.query = event.target.value; state.page = 0; searchTimer = setTimeout(renderGroups, 120); });
  $('sort').addEventListener('change', event => { state.sort = event.target.value; state.page = 0; renderGroups(); });
  window.ReportViewer = {data, state, parseReportTime, buildGroups, getGroups, selectSheet, renderGroups};
  selectSheet(0);
})();
