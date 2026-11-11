(async () => {
  const checks = [];
  const assert = (value, message) => { if (!value) throw Error(message); checks.push(message); };
  const viewer = ReportViewer, parse = viewer.parseReportTime;
  assert(parse('十一月 01,2026 12:00:00 上午').display === '11/01 00:00:00', 'Chinese midnight');
  assert(parse('十一月 01,2026 12:00:00 下 午').display === '11/01 12:00:00', 'Chinese noon and wrapped period');
  assert(parse('十月 31,2026 11:59:46 下\n午').display === '10/31 23:59:46', 'Chinese evening and line break');
  assert(parse('October 20,2026 01:53:42 PM').display === '10/20 13:53:42', 'English month and PM');
  assert(parse('10/20/2026 13:53:42').display === '10/20 13:53:42', 'Numeric month-day-year');
  assert(parse('2026-10-20 13:53:42').display === '10/20 13:53:42', 'ISO date');
  for (const invalid of ['二月 30,2026 01:00:00 上午', '十月 31,2026 13:00:00 PM', '', 'invalid time']) assert(!parse(invalid).valid, `Reject invalid time ${invalid}`);
  assert(!window.injected && !document.querySelector('img'), 'PDF content cannot execute HTML or scripts');
  assert(document.title === 'Fixture <safe>', 'Page title safely preserved');
  assert(!document.querySelector('header.topbar,#report-subtitle,#generated-at,.metrics,.section-heading,#results-caption,h3.subsection-label'), 'Removed metadata and repeated detail total are absent');
  assert(viewer.data.reports.length === 14 && document.querySelectorAll('[role=tab]').length === 14, 'Each report has a tab');
  const groups = viewer.getGroups(0).groups;
  const alice = groups.find(group => group.identifier === 'alice');
  assert(alice.rows.length === 103 && groups.length === 14, 'Grouping counts preserve all records and case');
  assert(!document.querySelector('.common-section,.group-avatar,.group-latest'), 'Simple headings without shared-info cards or avatars');
  const first = document.querySelector('.group-card');
  assert(first.open && first.querySelectorAll('tbody tr').length === 100, 'Large groups start with 100 detail rows');
  assert([...first.querySelectorAll('thead th')].map(x => x.textContent).join('|') === 'User Name|Domain Name|SID|Time Generated|Message|Source', 'Unknown report keeps every non-key column in source order');
  assert(first.querySelector('thead th').textContent === 'User Name' && first.querySelectorAll('tbody .key-value').length === 1 && first.querySelector('tbody .key-value').textContent === 'alice' && first.querySelector('tbody .key-value').rowSpan === 100, 'Key column has one merged row header spanning the visible page');
  assert(first.querySelector('tbody tr .time-cell').textContent === '12/31 23:59:00', 'First page starts with oldest record');
  first.querySelector('.pagination button:last-child').click();
  assert(first.querySelectorAll('tbody tr').length === 3 && first.querySelector('tbody tr .time-cell').textContent === '01/01 00:00:40', 'Next page has final three records in continuous order across year boundary');
  assert(first.querySelectorAll('tbody .key-value').length === 1 && first.querySelector('tbody .key-value').rowSpan === 3, 'Merged key adjusts to final page size');
  first.querySelector('.pagination button:first-of-type').click();
  assert(first.querySelectorAll('tbody tr').length === 100, 'Detail previous page works');
  assert(alice.rows[0].time.raw === '2026-12-31 23:59:00' && alice.rows[102].time.raw === '2027-01-01 00:00:42', 'Sort uses full year rather than displayed month and day');
  assert(document.querySelectorAll('.group-card').length === 12, 'Outer page is limited to 12 keys');
  document.querySelector('#group-pagination button:last-child').click();
  assert(document.querySelectorAll('.group-card').length === 2, 'Remaining keys accessible on outer next page');
  document.querySelector('#group-pagination button:first-of-type').click();
  assert(document.querySelectorAll('.group-card').length === 12, 'Outer previous page restores 12 keys');
  const search = document.getElementById('search');
  search.value = 'Message 1'; search.dispatchEvent(new Event('input', {bubbles:true}));
  await new Promise(r => setTimeout(r, 160));
  assert(document.querySelectorAll('.group-card').length === 1 && document.querySelector('.count-badge').textContent.includes('103'), 'Search keeps every trigger within matched group');
  search.value = 'not-a-real-value'; search.dispatchEvent(new Event('input', {bubbles:true}));
  await new Promise(r => setTimeout(r, 160));
  assert(!document.getElementById('empty-state').hidden, 'No-result feedback');
  document.querySelectorAll('[role=tab]')[1].click();
  assert(search.value === '' && viewer.data.reports[viewer.state.sheet].rows.length === 3, 'Sheet switch resets search and updates totals');
  assert(document.querySelector('tbody tr .time-cell').textContent === '10/20 13:53:42', 'Resource Accessed Time is converted');
  assert(document.querySelectorAll('thead th').length === 4 && document.querySelectorAll('tbody tr').length === 2, 'Identical times and shared non-key contents retain one row per source record');
  assert(!document.getElementById('time-warning').hidden && viewer.getGroups(1).invalidTimes === 1, 'Invalid timestamp retained and flagged');
  const formats = [
    ['Active Alerts (Trouble)','Source','Time Generated','Alert Message'],
    ['All First Time Resource Access','Reference Resource','Resource Accessed Time','Message'],
    ['All Threshold Breach Activity','Name','Time Generated','Mean Count','Threshold Count','Message'],
    ['All Unusual Time Activity','Name','Time Generated','General Start Time','General End Time','Message'],
    ['First Time Remote Access on Host','Unusual Remote Access From Computer','Resource Accessed Time','Message'],
    ['Unusual Logon Activity Time','User Name','Time Generated','General Start Time','General End Time','Message'],
    ['Unusual Logon Activity Time on Host','Computer Name','Time Generated','General Start Time','General End Time','Message'],
    ['Unusual User Activity Time','User Name','Time Generated','General Start Time','General End Time','Message'],
    ['Unusual Volume of Lockout','User Name','Time Generated','Mean Count','Threshold Count','Message'],
    ['Unusual Volume of Logon Failure','User Name','Time Generated','Mean Count','Threshold Count','Message'],
    ['Unusual Volume of Logon Failure on Host','Computer Name','Time Generated','Mean Count','Threshold Count','Message'],
    ['Unusual Volume of User Management Activity','User Name','Time Generated','Mean Count','Threshold Count','Message']
  ];
  for (const [i, format] of formats.entries()) {
    viewer.selectSheet(i + 2);
    const report = viewer.data.reports[i + 2];
    const headers = [...document.querySelectorAll('thead th')].map(x => x.textContent.replace(/\s+/g, ' ').trim());
    assert(report.title === format[0] && report.columns[report.keyIndex] === format[1] && report.keyIndex !== 0, `Key mapped by name: ${format[0]}`);
    assert(headers.join('|') === format.slice(1).join('|'), `Exact detail order and hidden columns: ${format[0]}`);
    assert(document.querySelectorAll('tbody tr').length === 2 && document.querySelector('.count-badge').textContent === '2 筆', `Duplicate key counted, shared fields retained: ${format[0]}`);
    assert(document.querySelectorAll('tbody .key-value').length === 1 && document.querySelector('tbody .key-value').rowSpan === 2 && document.querySelector('tbody .key-value').textContent === 'same-key', "Merged key appears once: " + format[0]);
    assert(!document.querySelector('tbody').textContent.includes('shared-Domain Name'), `Excluded content hidden: ${format[0]}`);
  }
  const mixed = viewer.buildGroups({keyIndex: 0, timeIndex: 1, rows: [
    {values:['key','invalid','bad-first']},
    {values:['key','2027-01-01 00:00:00','new-year']},
    {values:['key','2026-12-31 23:59:59','tie-first']},
    {values:['key','2026-12-31 23:59:59','tie-second']},
    {values:['key','','bad-second']}
  ]});
  assert(mixed.groups[0].rows.map(r => r.values[2]).join('|') === 'tie-first|tie-second|new-year|bad-first|bad-second', 'Invalid times last; equal timestamps and invalid times preserve input order');
  viewer.selectSheet(0);
  const sort = document.getElementById('sort');
  sort.value = 'name'; sort.dispatchEvent(new Event('change'));
  const byName = [...groups].sort((a,b) => a.identifier.localeCompare(b.identifier));
  assert(document.querySelector('.group-card').dataset.identifier === byName[0].identifier, 'Name group sort works');
  sort.value = 'latest'; sort.dispatchEvent(new Event('change'));
  assert(document.querySelector('.group-card').dataset.identifier === 'alice', 'Latest group sort independent from ascending detail order');
  const closed = document.querySelectorAll('.group-card')[1]; closed.open = true;
  await new Promise(r => setTimeout(r, 30));
  assert(closed.querySelector('tbody'), 'Collapsed groups lazily build detail tables on expansion');
  sort.value = 'count'; sort.dispatchEvent(new Event('change'));
  assert(document.getElementById('report-title').textContent === 'Logon fixture', 'Return to original sheet');
  return {passed: checks.length, checks};
})()
