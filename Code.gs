const CONFIG = {
  RAW_SHEET: 'Heatmap  Raw', // note: double space in the actual tab name
  HEADER_ROW: 3, // last header row (1-indexed); data starts the row after
  COLS: { MONTH: 0, DAY: 1, HOUR: 2, AGENT: 3, CHAT_AHT: 4, PHONE_AHT: 5, CALLS: 6, EMAILS: 7, CHATS: 8, REOPENED: 9, CASES_CLOSED: 10 },
  AGENT_SHEET: 'Agent Data',
  AGENT_COLS: { LDAP: 0, NAME: 1, GOOGLE_ID: 2, ROLE: 3, LOB: 4, WORKFLOW: 5, CHANNEL: 6, TEAM: 7 }
};

function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('MyPlayCases')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px;">Failed to load MyPlayCases: ' +
      String(err.message || err) + '</p>'
    );
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Client entry point. Viewer identity is resolved fresh every call
 * (never cached); the heavy shared dataset is cached/locked separately.
 */
function getAllData(bypassCache) {
  if (bypassCache && !checkRateLimit_('bypass_refresh', 4, 30)) {
    throw new Error('Refreshing too often — please wait a few seconds and try again.');
  }

  const shared = getSharedData_(!!bypassCache);

  const viewerEmail = Session.getActiveUser().getEmail();
  const viewerLdap = viewerEmail ? viewerEmail.split('@')[0] : '';
  const viewerInfo = shared.agentInfo[viewerLdap] || null;
  const isLead = !!(viewerInfo && /supervisor|manager/i.test(viewerInfo.role || ''));

  return {
    index: shared.index,
    agents: shared.agents,
    minDate: shared.minDate,
    maxDate: shared.maxDate,
    agentInfo: shared.agentInfo,
    warnings: shared.warnings || [],
    computedAt: shared.computedAt,
    viewer: { ldap: viewerLdap, info: viewerInfo, isLead: isLead },
    triggerSetup: hasBackgroundTrigger_()
  };
}

function refreshCacheBackground() {
  getSharedData_(true); // bypass cache, force recompute and save
}

function setupBackgroundTrigger() {
  const viewerEmail = Session.getActiveUser().getEmail();
  const viewerLdap = viewerEmail ? viewerEmail.split('@')[0] : '';
  const agentInfo = getAgentDirectory_();
  const viewerInfo = agentInfo[viewerLdap] || null;
  const isLead = !!(viewerInfo && /supervisor|manager/i.test(viewerInfo.role || ''));

  if (!isLead) {
    throw new Error('Only leads can setup background triggers.');
  }

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('auto_refresh_setup') === 'true') {
      return true;
  }

  removeBackgroundTrigger(); // Clear any existing first
  ScriptApp.newTrigger('refreshCacheBackground')
    .timeBased()
    .everyMinutes(10)
    .create();

  props.setProperty('auto_refresh_setup', 'true');
  return true;
}

function removeBackgroundTrigger() {
  const props = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshCacheBackground') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  props.deleteProperty('auto_refresh_setup');
  return false;
}

function hasBackgroundTrigger_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('auto_refresh_setup') === 'true';
}

// Temporary debug helper — run this if getAllData() can't find the sheet.
function listSheetNames() {
  const names = SpreadsheetApp.getActive().getSheets().map(function (s) {
    return '"' + s.getName() + '"';
  });
  Logger.log(names.join('\n'));
}

// Run manually from the Apps Script editor to sanity-check before wiring the UI.
function testGetAllData() {
  const result = getAllData();
  Logger.log('Agents: %s', result.agents.length);
  Logger.log('Date range: %s to %s', result.minDate, result.maxDate);
  Logger.log('Sample agents: %s', JSON.stringify(result.agents.slice(0, 5)));
}
