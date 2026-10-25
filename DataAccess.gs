function requireSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: "' + name + '". Check it matches the actual tab name exactly.');
  return sheet;
}

function num_(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

/**
 * Confirms the configured column indices don't exceed the sheet's actual
 * width, so a reordered/trimmed sheet fails loudly instead of silently
 * reading the wrong column.
 */
function validateHeader_(headerRow, colsConfig, sheetLabel) {
  const width = headerRow ? headerRow.length : 0;
  const maxIndex = Math.max.apply(null, Object.values(colsConfig));
  if (maxIndex >= width) {
    throw new Error(
      'Column mismatch in "' + sheetLabel + '": expected at least ' + (maxIndex + 1) +
      ' columns but the sheet only has ' + width + '. Check it hasn\'t been reordered or trimmed.'
    );
  }
}

function getAgentDirectory_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.AGENT_SHEET);
  if (!sheet) return {};
  const values = sheet.getDataRange().getValues();
  validateHeader_(values[0], CONFIG.AGENT_COLS, CONFIG.AGENT_SHEET);
  const rows = values.slice(1);
  const A = CONFIG.AGENT_COLS;
  const dir = {};
  rows.forEach(function (r) {
    const ldap = r[A.LDAP];
    if (!ldap) return;
    dir[ldap] = {
      name: r[A.NAME] || ldap,
      googleId: r[A.GOOGLE_ID] || '',
      role: r[A.ROLE] || '',
      lob: r[A.LOB] || '',
      workflow: r[A.WORKFLOW] || '',
      channel: r[A.CHANNEL] || '',
      team: r[A.TEAM] || ''
    };
  });
  return dir;
}

function computeSharedData_() {
  const sheet = requireSheet_(CONFIG.RAW_SHEET);
  const data = sheet.getDataRange().getValues();
  validateHeader_(data[CONFIG.HEADER_ROW - 1], CONFIG.COLS, CONFIG.RAW_SHEET);
  const rows = data.slice(CONFIG.HEADER_ROW);
  const C = CONFIG.COLS;
  const tz = Session.getScriptTimeZone();

  const index = {};
  const agents = new Set();
  let minDate = null, maxDate = null;

  let skipped = 0;
  const skipReasons = {};
  rows.forEach(function (r) {
    try {
      const day = r[C.DAY];
      const hour = num_(r[C.HOUR]);
      const agent = r[C.AGENT];
      if (!day || hour === null || hour < 0 || hour > 23 || !agent) {
        skipped++;
        const reason = !day ? 'missing date' : !agent ? 'missing agent' : 'bad hour';
        skipReasons[reason] = (skipReasons[reason] || 0) + 1;
        return;
      }

      const dateKey = Utilities.formatDate(new Date(day), tz, 'yyyy-MM-dd');
      agents.add(agent);
      if (!minDate || dateKey < minDate) minDate = dateKey;
      if (!maxDate || dateKey > maxDate) maxDate = dateKey;

      if (!index[agent]) index[agent] = {};
      if (!index[agent][dateKey]) index[agent][dateKey] = {};
      index[agent][dateKey][hour] = {
        chatAHT: num_(r[C.CHAT_AHT]),
        phoneAHT: num_(r[C.PHONE_AHT]),
        calls: num_(r[C.CALLS]) || 0,
        emails: num_(r[C.EMAILS]) || 0,
        chats: num_(r[C.CHATS]) || 0,
        reopens: num_(r[C.REOPENED]) || 0,
        casesClosed: num_(r[C.CASES_CLOSED]) || 0
      };
    } catch (rowErr) {
      skipped++;
      skipReasons['error'] = (skipReasons['error'] || 0) + 1;
    }
  });
  if (skipped) Logger.log('computeSharedData_: skipped %s row(s): %s', skipped, JSON.stringify(skipReasons));

  const agentInfo = getAgentDirectory_();
  const warnings = [];
  if (agents.size && Object.keys(agentInfo).length === 0) {
    warnings.push('Agent Data directory is empty or the sheet name doesn\'t match — names, roles, and teams will show as LDAPs only.');
  }

  return {
    index: index,
    agents: Array.from(agents).sort(),
    minDate: minDate,
    maxDate: maxDate,
    agentInfo: agentInfo,
    warnings: warnings,
    computedAt: new Date().toISOString()
  };
}
