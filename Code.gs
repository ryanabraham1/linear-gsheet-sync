/**
 * Machining Tracker -> Linear one-way issue sync.
 *
 * Bound Google Apps Script for:
 * https://docs.google.com/spreadsheets/d/1WlHAodErTZT3_4QGyWwXvK8GIcZfhQbtOCWkGZ7PArk/
 *
 * Target: WarriorBorgs 3256 (WB) / bot-selected robot project /
 * Fabrication label / matching fabrication milestone.
 * Code.local.gs is the ignored local source of truth. Code.gs is its
 * publishable mirror; the API-key value must be their only difference.
 */
const CONFIG = Object.freeze({
  LINEAR_API_KEY: 'PASTE_LINEAR_PERSONAL_API_KEY_HERE',
  LINEAR_TEAM_ID: '16cf7cf2-dfa0-4d50-9639-48ec86de52e5',
  LINEAR_TEAM_KEY: 'WB',
  BOT_PROJECTS: {
    aimbot: {
      botName: 'Aimbot',
      projectId: 'c86bfc64-472f-41a3-a6ac-a10d02d4f07d',
      projectName: 'Aimbot Changes',
      milestoneId: '262fbb62-adca-4122-954e-b6550609cad2',
      milestoneName: 'Fully Fabbed',
    },
    dumper: {
      botName: 'Dumper',
      projectId: 'c626362d-0b26-4d3c-afe7-d125581a0383',
      projectName: 'Dumper',
      milestoneId: '95d255a1-0227-43a5-962d-0fbf07427a21',
      milestoneName: 'Dumper Fabbed',
    },
    everybot: {
      botName: 'EveryBot',
      projectId: 'f70a2ef5-4e36-41d0-bfe9-c8cc4352cfda',
      projectName: 'EveryBot',
      milestoneId: 'fabd7578-8b04-40bf-896e-b999e8666362',
      milestoneName: 'Fully Fabbed',
    },
  },
  LINEAR_LABEL_IDS: ['8b5ea928-8841-43a0-9ec5-bc6c02e0f836'],
  LINEAR_LABEL_NAMES: ['Fabrication'],
  LINEAR_TITLE_PREFIX: 'Fab: ',
  DYNAMIC_LABEL_GROUPS: {
    subsystem: { name: 'Subsystem', color: '#5E6AD2' },
    machine: { name: 'Fab Machine', color: '#4CB782' },
  },
  SPREADSHEET_URL:
    'https://docs.google.com/spreadsheets/d/1WlHAodErTZT3_4QGyWwXvK8GIcZfhQbtOCWkGZ7PArk/edit?gid=0#gid=0',
  SHEET_NAME: 'Machining Tracker',
  HEADER_ROW: 3,
  FIRST_DATA_ROW: 4,
  OUTPUT_START_COLUMN: 16, // P
  RECONCILE_EVERY_MINUTES: 5,
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,
  ERROR_RETRY_AFTER_MS: 15 * 60 * 1000,
});

const LINEAR_API_URL = 'https://api.linear.app/graphql';
const REGISTRY_PREFIX = 'LINEAR_SYNC_ISSUE_';
const ARCHIVED_MESSAGE =
  '[Archived by Sheets sync] Part #_Name was cleared. Restore the name to unarchive it.';

const INPUT_HEADERS = Object.freeze([
  'Status',
  'Bot',
  'Subsystem',
  'Part #_Name',
  'Priority',
  'Qty',
  'Spare Qty',
  'Stock Material/Type',
  'Stock Dimensions',
  'Length',
  'Tapped?',
  'Machine',
  'Drawing/CAM File (STEP FILES)',
  'Notes',
]);

const OUTPUT_HEADERS = Object.freeze([
  'Linear ID',
  'Linear Key',
  'Linear URL',
  'Last Synced',
  'Sync Error',
  '_Sync Hash',
]);

const REQUIRED_HEADERS = Object.freeze(INPUT_HEADERS.concat(OUTPUT_HEADERS));

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Linear Sync')
    .addItem('Install / reconnect', 'setupLinearSync')
    .addSeparator()
    .addItem('Sync selected rows now', 'syncSelectedRows')
    .addItem('Sync all changed rows now', 'syncAllRowsFromMenu')
    .addItem('Test connection', 'testLinearConnection')
    .addSeparator()
    .addItem('Remove automatic triggers', 'removeLinearTriggers')
    .addToUi();
}

/**
 * Run once from Linear Sync > Install / reconnect.
 * Adds only the sync metadata columns P:U, installs triggers, and performs the
 * initial upload of every populated machining row.
 */
function setupLinearSync() {
  const spreadsheet = SpreadsheetApp.getActive();
  ensureLinearApiKey_();
  PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());

  initializeSyncColumns_();
  testConfiguredLinearObjects_();
  ensureConfiguredLabelGroups_(getLinearApiKey_(), CONFIG.LINEAR_TEAM_ID);
  installLinearTriggers_();
  syncAllRows();

  spreadsheet.toast(
    'Connected to Linear team ' + CONFIG.LINEAR_TEAM_KEY + '. Bot selections now choose ' +
      'the robot project; blank Bot cells stay projectless. Label: ' +
      CONFIG.LINEAR_LABEL_NAMES.join(', ') + '.',
    'Linear Sync installed',
    10
  );
}

/** Installable edit-trigger handler. Do not run this manually. */
function handleLinearEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME || e.range.getLastRow() < CONFIG.FIRST_DATA_ROW) {
    return;
  }

  const headerMap = getHeaderMap_(sheet);
  if (!rangeTouchesInputs_(e.range, headerMap)) return;

  const rows = [];
  const firstRow = Math.max(CONFIG.FIRST_DATA_ROW, e.range.getRow());
  for (let row = firstRow; row <= e.range.getLastRow(); row += 1) rows.push(row);

  runWithLock_(function () {
    syncRows_(sheet, rows, false);
  });
}

/** Installable change-trigger handler for physically removed spreadsheet rows. */
function handleLinearChange(e) {
  if (!e || e.changeType !== 'REMOVE_ROW') return;
  runWithLock_(function () {
    archiveMissingRegisteredIssues_();
  });
}

/** Five-minute reconciliation trigger and safe manual entry point. */
function syncAllRows() {
  const sheet = getSyncSheet_();
  const rows = [];
  for (let row = CONFIG.FIRST_DATA_ROW; row <= sheet.getLastRow(); row += 1) rows.push(row);
  runWithLock_(function () {
    syncRows_(sheet, rows, false);
    archiveMissingRegisteredIssues_(sheet);
  });
}

function syncAllRowsFromMenu() {
  syncAllRows();
  SpreadsheetApp.getActive().toast(
    'Finished checking changed machining rows. See Last Synced or Sync Error.',
    'Linear Sync',
    7
  );
}

function syncSelectedRows() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const selection = sheet.getActiveRange();
  if (sheet.getName() !== CONFIG.SHEET_NAME || !selection) {
    throw new Error('Select rows on the "' + CONFIG.SHEET_NAME + '" tab first.');
  }

  const rows = [];
  const firstRow = Math.max(CONFIG.FIRST_DATA_ROW, selection.getRow());
  for (let row = firstRow; row <= selection.getLastRow(); row += 1) rows.push(row);
  if (!rows.length) throw new Error('Select at least one machining data row.');

  runWithLock_(function () {
    syncRows_(sheet, rows, true);
  });
  spreadsheet.toast('Selected rows were sent to Linear.', 'Linear Sync', 6);
}

function testLinearConnection() {
  ensureLinearApiKey_();
  const data = testConfiguredLinearObjects_();
  SpreadsheetApp.getUi().alert(
    'Connected as ' + data.viewer.name + '.\n\n' +
      'Team: ' + data.team.name + ' (' + CONFIG.LINEAR_TEAM_KEY + ')\n' +
      'Bot projects: ' + data.botProjects.map(function (project) {
        return project.botName + ' -> ' + project.project.name;
      }).join(', ') + '\n' +
      'Label: ' + data.issueLabel.name
  );
}

function removeLinearTriggers() {
  removeLinearTriggers_();
  SpreadsheetApp.getActive().toast('Automatic Linear triggers removed.', 'Linear Sync', 5);
}

function syncRows_(sheet, rows, force) {
  const startedAt = Date.now();
  const headerMap = getHeaderMap_(sheet);
  const apiKey = getLinearApiKey_();
  const labelGroups = ensureConfiguredLabelGroups_(apiKey, CONFIG.LINEAR_TEAM_ID);
  const context = {
    apiKey: apiKey,
    teamId: CONFIG.LINEAR_TEAM_ID,
    labelIds: CONFIG.LINEAR_LABEL_IDS.slice(),
    labelGroups: labelGroups,
    dynamicLabelCache: {},
    stateByName: null,
    stateNames: null,
  };

  for (let index = 0; index < rows.length; index += 1) {
    if (Date.now() - startedAt > CONFIG.MAX_RUNTIME_MS) return;
    syncOneRow_(sheet, rows[index], headerMap, context, force);
  }
}

function syncOneRow_(sheet, rowNumber, headerMap, context, force) {
  const lastNeededColumn = CONFIG.OUTPUT_START_COLUMN + OUTPUT_HEADERS.length - 1;
  const values = sheet.getRange(rowNumber, 1, 1, lastNeededColumn).getValues()[0];
  const raw = rowObject_(values, headerMap);
  const currentHash = machiningRowHash_(raw, context);
  const storedHash = stringValue_(raw['_Sync Hash']);
  const linearId = stringValue_(raw['Linear ID']);

  if (!force && shouldSkipHash_(storedHash, currentHash, linearId)) return;

  try {
    const title = stringValue_(raw['Part #_Name']);
    if (!title && !linearId) {
      clearRowError_(sheet, rowNumber, headerMap);
      return;
    }
    if (!title) {
      if (isArchivedBySync_(raw['Sync Error'])) {
        const existingArchivedIssue = {
          id: linearId,
          identifier: stringValue_(raw['Linear Key']),
          url: stringValue_(raw['Linear URL']),
        };
        writeArchiveSuccess_(
          sheet,
          rowNumber,
          headerMap,
          existingArchivedIssue,
          raw,
          currentHash
        );
        registerIssue_(existingArchivedIssue, true);
        return;
      }
      const archivedIssue = archiveIssue_(context.apiKey, linearId);
      writeArchiveSuccess_(
        sheet,
        rowNumber,
        headerMap,
        archivedIssue,
        raw,
        currentHash
      );
      registerIssue_(archivedIssue, true);
      return;
    }

    if (linearId && isArchivedBySync_(raw['Sync Error'])) {
      unarchiveIssue_(context.apiKey, linearId);
    }

    const status = linearStatusName_(raw['Status']);
    const botProject = resolveBotProject_(raw['Bot']);
    const dynamicLabelIds = resolveDynamicLabelIds_(context, raw);
    const commonInput = {
      title: linearIssueTitle_(title),
      description: machiningDescription_(raw, rowNumber),
      priority: priorityValue_(raw['Priority']),
    };
    if (botProject) {
      commonInput.projectId = botProject.projectId;
      commonInput.projectMilestoneId = botProject.milestoneId;
    } else if (linearId) {
      commonInput.projectId = null;
      commonInput.projectMilestoneId = null;
    }
    if (status) commonInput.stateId = resolveStateId_(context, status);

    let issue;
    if (linearId) {
      commonInput.labelIds = getFinalIssueLabelIds_(context, linearId, dynamicLabelIds);
      issue = updateIssue_(context.apiKey, linearId, commonInput);
    } else {
      commonInput.teamId = context.teamId;
      commonInput.labelIds = uniqueStrings_(context.labelIds.concat(dynamicLabelIds));
      issue = createIssue_(context.apiKey, commonInput);
    }

    writeSyncSuccess_(sheet, rowNumber, headerMap, issue, currentHash);
    registerIssue_(issue, false);
  } catch (error) {
    writeSyncError_(sheet, rowNumber, headerMap, error, currentHash);
  }
}

function resolveBotProject_(value) {
  const bot = stringValue_(value);
  if (!bot) return null;
  const project = CONFIG.BOT_PROJECTS[bot.toLowerCase()];
  if (!project) {
    throw new Error(
      'Unknown Bot "' + bot + '". Available: ' +
        Object.keys(CONFIG.BOT_PROJECTS).map(function (key) {
          return CONFIG.BOT_PROJECTS[key].botName;
        }).join(', ')
    );
  }
  return project;
}

function machiningDescription_(raw, rowNumber) {
  const lines = [];
  addDescriptionLine_(lines, 'Subsystem', raw['Subsystem']);
  addDescriptionLine_(lines, 'Quantity', raw['Qty']);
  addDescriptionLine_(lines, 'Spare quantity', raw['Spare Qty']);
  addDescriptionLine_(lines, 'Stock material/type', raw['Stock Material/Type']);
  addDescriptionLine_(lines, 'Stock dimensions', raw['Stock Dimensions']);
  addDescriptionLine_(lines, 'Length', raw['Length']);
  addDescriptionLine_(lines, 'Tapped?', raw['Tapped?']);
  addDescriptionLine_(lines, 'Machine', raw['Machine']);

  const drawing = stringValue_(raw['Drawing/CAM File (STEP FILES)']);
  if (drawing) {
    lines.push(
      '**Drawing/CAM file:** ' +
        (/^https?:\/\//i.test(drawing) ? '[Open file](' + drawing + ')' : drawing)
    );
  }
  addDescriptionLine_(lines, 'Notes', raw['Notes']);

  const rowUrl = CONFIG.SPREADSHEET_URL.replace(/#.*$/, '') + '#gid=0&range=E' + rowNumber;
  lines.push('');
  lines.push('---');
  lines.push('**Source:** [Machining Tracker row ' + rowNumber + '](' + rowUrl + ')');
  lines.push('_Synced automatically from Google Sheets._');
  return lines.join('\n');
}

function addDescriptionLine_(lines, label, value) {
  const text = stringValue_(value);
  if (text) lines.push('**' + label + ':** ' + text);
}

function machiningRowHash_(raw, context) {
  const fields = {};
  INPUT_HEADERS.forEach(function (header) {
    fields[header] = serializableCell_(raw[header]);
  });
  return hash_(
    JSON.stringify({
      teamId: context.teamId,
      botProjects: CONFIG.BOT_PROJECTS,
      labelIds: context.labelIds,
      titlePrefix: CONFIG.LINEAR_TITLE_PREFIX,
      dynamicLabelGroups: CONFIG.DYNAMIC_LABEL_GROUPS,
      syncBehaviorVersion: 5,
      fields: fields,
    })
  );
}

function createIssue_(apiKey, input) {
  const data = linearRequest_(
    'mutation CreateIssue($input: IssueCreateInput!) {\n' +
      '  issueCreate(input: $input) { success issue { id identifier url } }\n' +
      '}',
    { input: input },
    apiKey
  );
  if (!data.issueCreate || !data.issueCreate.success || !data.issueCreate.issue) {
    throw new Error('Linear did not confirm that the issue was created.');
  }
  return data.issueCreate.issue;
}

function updateIssue_(apiKey, issueId, input) {
  const data = linearRequest_(
    'mutation UpdateIssue($id: String!, $input: IssueUpdateInput!) {\n' +
      '  issueUpdate(id: $id, input: $input) { success issue { id identifier url } }\n' +
      '}',
    { id: issueId, input: input },
    apiKey
  );
  if (!data.issueUpdate || !data.issueUpdate.success || !data.issueUpdate.issue) {
    throw new Error('Linear did not confirm that the issue was updated.');
  }
  return data.issueUpdate.issue;
}

function archiveIssue_(apiKey, issueId) {
  const data = linearRequest_(
    'mutation ArchiveIssue($id: String!) {\n' +
      '  issueArchive(id: $id, trash: true) {\n' +
      '    success\n' +
      '    entity { id identifier url }\n' +
      '  }\n' +
      '}',
    { id: issueId },
    apiKey
  );
  if (!data.issueArchive || !data.issueArchive.success) {
    throw new Error('Linear did not confirm that the issue was archived.');
  }
  return data.issueArchive.entity || { id: issueId, identifier: '', url: '' };
}

function unarchiveIssue_(apiKey, issueId) {
  const data = linearRequest_(
    'mutation UnarchiveIssue($id: String!) {\n' +
      '  issueUnarchive(id: $id) { success entity { id identifier url } }\n' +
      '}',
    { id: issueId },
    apiKey
  );
  if (!data.issueUnarchive || !data.issueUnarchive.success) {
    throw new Error('Linear did not confirm that the issue was restored.');
  }
  return data.issueUnarchive.entity;
}

function ensureConfiguredLabelGroups_(apiKey, teamId) {
  const result = {};
  Object.keys(CONFIG.DYNAMIC_LABEL_GROUPS).forEach(function (key) {
    const definition = CONFIG.DYNAMIC_LABEL_GROUPS[key];
    result[key] = {
      id: getOrCreateLabelGroup_(apiKey, teamId, key, definition),
      name: definition.name,
      color: definition.color,
    };
  });
  return result;
}

function getOrCreateLabelGroup_(apiKey, teamId, groupKey, definition) {
  const properties = PropertiesService.getDocumentProperties();
  const propertyKey = 'LINEAR_LABEL_GROUP_' + groupKey.toUpperCase();
  const cachedId = properties.getProperty(propertyKey);
  if (cachedId) return cachedId;

  const data = linearRequest_(
    'query FindLabelGroup($teamId: ID!, $name: String!) {\n' +
      '  issueLabels(first: 10, filter: {\n' +
      '    team: { id: { eq: $teamId } }\n' +
      '    name: { eqIgnoreCase: $name }\n' +
      '    isGroup: { eq: true }\n' +
      '  }) { nodes { id name isGroup } }\n' +
      '}',
    { teamId: teamId, name: definition.name },
    apiKey
  );
  let group = data.issueLabels.nodes[0];
  if (!group) {
    group = createIssueLabel_(apiKey, {
      name: definition.name,
      color: definition.color,
      teamId: teamId,
      isGroup: true,
    });
  }
  properties.setProperty(propertyKey, group.id);
  return group.id;
}

function resolveDynamicLabelIds_(context, raw) {
  const selected = [];
  const subsystem = labelNameFromCell_(raw['Subsystem']);
  const machine = labelNameFromCell_(raw['Machine']);
  if (subsystem) {
    selected.push(
      getOrCreateChildLabel_(context, context.labelGroups.subsystem, subsystem)
    );
  }
  if (machine) {
    selected.push(getOrCreateChildLabel_(context, context.labelGroups.machine, machine));
  }
  return uniqueStrings_(selected);
}

function getOrCreateChildLabel_(context, group, labelName) {
  const memoryKey = group.id + '|' + labelName.toLowerCase();
  if (context.dynamicLabelCache[memoryKey]) {
    return context.dynamicLabelCache[memoryKey];
  }

  const properties = PropertiesService.getDocumentProperties();
  const propertyKey = 'LINEAR_CHILD_LABEL_' + hash_(memoryKey).slice(0, 24);
  const cachedId = properties.getProperty(propertyKey);
  if (cachedId) {
    context.dynamicLabelCache[memoryKey] = cachedId;
    return cachedId;
  }

  const data = linearRequest_(
    'query FindChildLabel($teamId: ID!, $parentId: ID!, $name: String!) {\n' +
      '  issueLabels(first: 10, filter: {\n' +
      '    team: { id: { eq: $teamId } }\n' +
      '    parent: { id: { eq: $parentId } }\n' +
      '    name: { eqIgnoreCase: $name }\n' +
      '    isGroup: { eq: false }\n' +
      '  }) { nodes { id name parent { id } } }\n' +
      '}',
    { teamId: context.teamId, parentId: group.id, name: labelName },
    context.apiKey
  );
  let label = data.issueLabels.nodes[0];
  if (!label) {
    label = createIssueLabel_(context.apiKey, {
      name: labelName,
      color: group.color,
      teamId: context.teamId,
      parentId: group.id,
      isGroup: false,
    });
  }
  properties.setProperty(propertyKey, label.id);
  context.dynamicLabelCache[memoryKey] = label.id;
  return label.id;
}

function createIssueLabel_(apiKey, input) {
  const data = linearRequest_(
    'mutation CreateIssueLabel($input: IssueLabelCreateInput!) {\n' +
      '  issueLabelCreate(input: $input) {\n' +
      '    success\n' +
      '    issueLabel { id name isGroup parent { id } }\n' +
      '  }\n' +
      '}',
    { input: input },
    apiKey
  );
  if (!data.issueLabelCreate || !data.issueLabelCreate.success) {
    throw new Error('Linear did not confirm that the label was created.');
  }
  return data.issueLabelCreate.issueLabel;
}

function getFinalIssueLabelIds_(context, issueId, selectedIds) {
  const data = linearRequest_(
    'query CurrentIssueLabels($id: String!) {\n' +
      '  issue(id: $id) {\n' +
      '    labels { nodes { id name parent { id } } }\n' +
      '  }\n' +
      '}',
    { id: issueId },
    context.apiKey
  );
  if (!data.issue) throw new Error('Could not read the current Linear issue labels.');

  return finalIssueLabelIds_(context, data.issue.labels.nodes, selectedIds);
}

function finalIssueLabelIds_(context, currentLabels, selectedIds) {
  const dynamicGroupIds = {};
  Object.keys(context.labelGroups).forEach(function (key) {
    dynamicGroupIds[context.labelGroups[key].id] = true;
  });

  const preservedLabelIds = [];
  currentLabels.forEach(function (label) {
    const parentId = label.parent && label.parent.id;
    if (!parentId || !dynamicGroupIds[parentId]) preservedLabelIds.push(label.id);
  });
  return uniqueStrings_(preservedLabelIds.concat(context.labelIds, selectedIds));
}

function linearRequest_(query, variables, apiKey) {
  const response = UrlFetchApp.fetch(LINEAR_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: apiKey },
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  let body;
  try {
    body = JSON.parse(responseText);
  } catch (error) {
    throw new Error('Linear returned HTTP ' + statusCode + ' with an unreadable response.');
  }

  if (body.errors && body.errors.length) {
    throw new Error(
      'Linear API: ' +
        body.errors
          .map(function (item) {
            return item.message;
          })
          .join(' | ')
    );
  }
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Linear returned HTTP ' + statusCode + '.');
  }
  if (!body.data) throw new Error('Linear returned no data.');
  return body.data;
}

function testConfiguredLinearObjects_() {
  const apiKey = getLinearApiKey_();
  const data = linearRequest_(
    'query SyncContext($teamId: String!, $labelId: String!) {\n' +
      '  viewer { id name email }\n' +
      '  team(id: $teamId) { id name }\n' +
      '  issueLabel(id: $labelId) { id name }\n' +
      '}',
    {
      teamId: CONFIG.LINEAR_TEAM_ID,
      labelId: CONFIG.LINEAR_LABEL_IDS[0],
    },
    apiKey
  );
  if (!data.team || !data.issueLabel) {
    throw new Error('One or more configured Linear objects could not be found.');
  }
  data.botProjects = Object.keys(CONFIG.BOT_PROJECTS).map(function (key) {
    const botProject = CONFIG.BOT_PROJECTS[key];
    const projectData = linearRequest_(
      'query BotProject($projectId: String!, $milestoneId: String!) {\n' +
        '  project(id: $projectId) { id name }\n' +
        '  projectMilestone(id: $milestoneId) { id name project { id } }\n' +
        '}',
      {
        projectId: botProject.projectId,
        milestoneId: botProject.milestoneId,
      },
      apiKey
    );
    if (
      !projectData.project ||
      !projectData.projectMilestone ||
      !projectData.projectMilestone.project ||
      projectData.projectMilestone.project.id !== projectData.project.id
    ) {
      throw new Error('Invalid Linear project or milestone for Bot "' + botProject.botName + '".');
    }
    return {
      botName: botProject.botName,
      project: projectData.project,
      milestone: projectData.projectMilestone,
    };
  });
  return data;
}

function ensureLinearApiKey_() {
  const configured = configuredLinearApiKey_();
  if (configured) return configured;

  const properties = PropertiesService.getUserProperties();
  const stored = properties.getProperty('LINEAR_API_KEY');
  if (stored) return stored;

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Linear personal API key',
    'Create one in Linear Settings > Security & access > Personal API keys, then paste it here.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error('Linear setup was canceled before an API key was saved.');
  }
  const apiKey = response.getResponseText().trim();
  if (!apiKey) throw new Error('No Linear API key was entered.');
  properties.setProperty('LINEAR_API_KEY', apiKey);
  return apiKey;
}

function getLinearApiKey_() {
  const apiKey =
    configuredLinearApiKey_() ||
    PropertiesService.getUserProperties().getProperty('LINEAR_API_KEY');
  if (!apiKey) {
    throw new Error('Linear API key is missing. Run Linear Sync > Install / reconnect.');
  }
  return apiKey;
}

function configuredLinearApiKey_() {
  const value = String(CONFIG.LINEAR_API_KEY || '').trim();
  return value && value.indexOf('PASTE_') !== 0 ? value : '';
}

function getTeamStates_(apiKey, teamId) {
  const data = linearRequest_(
    'query TeamStates($teamId: ID!) {\n' +
      '  workflowStates(first: 100, filter: { team: { id: { eq: $teamId } } }) {\n' +
      '    nodes { id name type }\n' +
      '  }\n' +
      '}',
    { teamId: teamId },
    apiKey
  );
  return data.workflowStates.nodes;
}

function resolveStateId_(context, requestedName) {
  if (!context.stateByName) {
    const states = getTeamStates_(context.apiKey, context.teamId);
    context.stateByName = {};
    context.stateNames = [];
    states.forEach(function (state) {
      context.stateByName[state.name.toLowerCase()] = state.id;
      context.stateNames.push(state.name);
    });
  }
  const stateId = context.stateByName[requestedName.toLowerCase()];
  if (!stateId) {
    throw new Error(
      'Unknown Linear status "' + requestedName + '". Available: ' + context.stateNames.join(', ')
    );
  }
  return stateId;
}

function initializeSyncColumns_() {
  const sheet = getSyncSheet_();
  const requiredLastColumn = CONFIG.OUTPUT_START_COLUMN + OUTPUT_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredLastColumn) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredLastColumn - sheet.getMaxColumns()
    );
  }

  const outputRange = sheet.getRange(
    CONFIG.HEADER_ROW,
    CONFIG.OUTPUT_START_COLUMN,
    1,
    OUTPUT_HEADERS.length
  );
  const existing = outputRange.getDisplayValues()[0].map(function (value) {
    return value.trim();
  });
  const hasExisting = existing.some(function (value) {
    return value !== '';
  });
  if (hasExisting && existing.join('|') !== OUTPUT_HEADERS.join('|')) {
    throw new Error(
      'Columns P:U are not empty and do not contain the expected Linear sync headers. ' +
        'Move that data elsewhere before installing.'
    );
  }

  const styleSource = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.OUTPUT_START_COLUMN - 1);
  outputRange
    .setValues([OUTPUT_HEADERS])
    .setBackground(styleSource.getBackground())
    .setFontColor(styleSource.getFontColor())
    .setFontFamily(styleSource.getFontFamily())
    .setFontSize(styleSource.getFontSize())
    .setFontWeight(styleSource.getFontWeight())
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(CONFIG.OUTPUT_START_COLUMN + 1, 100);
  sheet.setColumnWidth(CONFIG.OUTPUT_START_COLUMN + 2, 280);
  sheet.setColumnWidth(CONFIG.OUTPUT_START_COLUMN + 3, 165);
  sheet.setColumnWidth(CONFIG.OUTPUT_START_COLUMN + 4, 320);
  sheet
    .getRange(
      CONFIG.FIRST_DATA_ROW,
      CONFIG.OUTPUT_START_COLUMN + 3,
      Math.max(1, sheet.getMaxRows() - 3),
      1
    )
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.hideColumns(CONFIG.OUTPUT_START_COLUMN);
  sheet.hideColumns(CONFIG.OUTPUT_START_COLUMN + OUTPUT_HEADERS.length - 1);
}

function installLinearTriggers_() {
  removeLinearTriggers_();
  const spreadsheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('handleLinearEdit').forSpreadsheet(spreadsheet).onEdit().create();
  ScriptApp.newTrigger('handleLinearChange').forSpreadsheet(spreadsheet).onChange().create();
  ScriptApp.newTrigger('syncAllRows')
    .timeBased()
    .everyMinutes(CONFIG.RECONCILE_EVERY_MINUTES)
    .create();
}

function removeLinearTriggers_() {
  const handlers = { handleLinearEdit: true, handleLinearChange: true, syncAllRows: true };
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (handlers[trigger.getHandlerFunction()]) ScriptApp.deleteTrigger(trigger);
  });
}

function getSyncSheet_() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    const spreadsheetId = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('Spreadsheet connection is missing. Run Linear Sync > Install / reconnect.');
    }
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  }
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Missing tab "' + CONFIG.SHEET_NAME + '".');
  return sheet;
}

function getHeaderMap_(sheet) {
  const lastNeededColumn = CONFIG.OUTPUT_START_COLUMN + OUTPUT_HEADERS.length - 1;
  const values = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, lastNeededColumn)
    .getDisplayValues()[0];
  const map = {};
  values.forEach(function (value, index) {
    const normalized = value.trim();
    if (normalized) map[normalized] = index;
  });
  const missing = REQUIRED_HEADERS.filter(function (header) {
    return map[header] === undefined;
  });
  if (missing.length) throw new Error('Missing Machining Tracker headers: ' + missing.join(', '));
  return map;
}

function rowObject_(values, headerMap) {
  const object = {};
  REQUIRED_HEADERS.forEach(function (header) {
    object[header] = values[headerMap[header]];
  });
  return object;
}

function rangeTouchesInputs_(range, headerMap) {
  const firstIndex = range.getColumn() - 1;
  const lastIndex = range.getLastColumn() - 1;
  return INPUT_HEADERS.some(function (header) {
    const index = headerMap[header];
    return index >= firstIndex && index <= lastIndex;
  });
}

function writeSyncSuccess_(sheet, rowNumber, headerMap, issue, currentHash) {
  sheet
    .getRange(rowNumber, headerMap['Linear ID'] + 1, 1, OUTPUT_HEADERS.length)
    .setValues([[
      issue.id,
      issue.identifier,
      issue.url,
      new Date(),
      '',
      currentHash,
    ]]);
}

function writeArchiveSuccess_(sheet, rowNumber, headerMap, issue, raw, currentHash) {
  sheet
    .getRange(rowNumber, headerMap['Linear ID'] + 1, 1, OUTPUT_HEADERS.length)
    .setValues([[
      issue.id || stringValue_(raw['Linear ID']),
      issue.identifier || stringValue_(raw['Linear Key']),
      issue.url || stringValue_(raw['Linear URL']),
      new Date(),
      ARCHIVED_MESSAGE,
      currentHash,
    ]]);
}

function writeSyncError_(sheet, rowNumber, headerMap, error, currentHash) {
  const message = String(error && error.message ? error.message : error).slice(0, 500);
  sheet.getRange(rowNumber, headerMap['Sync Error'] + 1).setValue(message);
  sheet
    .getRange(rowNumber, headerMap['_Sync Hash'] + 1)
    .setValue('ERROR|' + Date.now() + '|' + currentHash);
}

function clearRowError_(sheet, rowNumber, headerMap) {
  sheet.getRange(rowNumber, headerMap['Sync Error'] + 1).clearContent();
}

function shouldSkipHash_(storedHash, currentHash, linearId) {
  if (linearId && storedHash === currentHash) return true;
  const parts = storedHash.split('|');
  if (parts.length === 3 && parts[0] === 'ERROR' && parts[2] === currentHash) {
    const lastAttempt = Number(parts[1]);
    if (lastAttempt && Date.now() - lastAttempt < CONFIG.ERROR_RETRY_AFTER_MS) return true;
  }
  return false;
}

function linearIssueTitle_(partName) {
  const title = stringValue_(partName);
  if (!title) return '';
  return /^fab:\s*/i.test(title) ? title : CONFIG.LINEAR_TITLE_PREFIX + title;
}

function labelNameFromCell_(value) {
  return stringValue_(value)
    .replace(/^\d+\s*[.\-_:]\s*/, '')
    .trim();
}

function uniqueStrings_(values) {
  const seen = {};
  return values.filter(function (value) {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function isArchivedBySync_(message) {
  return stringValue_(message).indexOf('[Archived by Sheets sync]') === 0;
}

function registerIssue_(issue, archived) {
  if (!issue || !issue.id) return;
  PropertiesService.getDocumentProperties().setProperty(
    REGISTRY_PREFIX + issue.id,
    JSON.stringify({
      id: issue.id,
      identifier: issue.identifier || '',
      url: issue.url || '',
      archived: Boolean(archived),
      updatedAt: new Date().toISOString(),
    })
  );
}

function archiveMissingRegisteredIssues_(providedSheet) {
  const sheet = providedSheet || getSyncSheet_();
  const headerMap = getHeaderMap_(sheet);
  const currentIds = {};
  const rowCount = Math.max(0, sheet.getLastRow() - CONFIG.FIRST_DATA_ROW + 1);
  if (rowCount) {
    sheet
      .getRange(CONFIG.FIRST_DATA_ROW, headerMap['Linear ID'] + 1, rowCount, 1)
      .getDisplayValues()
      .forEach(function (row) {
        const id = stringValue_(row[0]);
        if (id) currentIds[id] = true;
      });
  }

  const properties = PropertiesService.getDocumentProperties();
  const allProperties = properties.getProperties();
  Object.keys(allProperties).forEach(function (key) {
    if (key.indexOf(REGISTRY_PREFIX) !== 0) return;
    let entry;
    try {
      entry = JSON.parse(allProperties[key]);
    } catch (error) {
      return;
    }
    if (!entry.id || entry.archived || currentIds[entry.id]) return;

    try {
      const archivedIssue = archiveIssue_(getLinearApiKey_(), entry.id);
      registerIssue_(archivedIssue, true);
    } catch (error) {
      console.error('Could not archive missing Linear issue ' + entry.id + ': ' + error.message);
    }
  });
}

function linearStatusName_(value) {
  const normalized = stringValue_(value).toLowerCase();
  if (!normalized) return '';
  const mapping = {
    'not started': 'Todo',
    'todo': 'Todo',
    'in progress': 'In Progress',
    'finished': 'Done',
    'done': 'Done',
    'backlog': 'Backlog',
    'in review': 'In Review',
    'canceled': 'Canceled',
    'cancelled': 'Canceled',
  };
  return mapping[normalized] || stringValue_(value);
}

function priorityValue_(value) {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 0 && value <= 4) {
    return value;
  }
  const normalized = stringValue_(value).toLowerCase();
  const numberMatch = normalized.match(/^#?([0-4])$/);
  if (numberMatch) return Number(numberMatch[1]);
  const priorities = {
    '': 0,
    'none': 0,
    'no priority': 0,
    'urgent': 1,
    'high': 2,
    'normal': 3,
    'medium': 3,
    'low': 4,
  };
  if (!Object.prototype.hasOwnProperty.call(priorities, normalized)) {
    throw new Error('Priority must be blank, #1, #2, #3, #4, Urgent, High, Medium, or Low.');
  }
  return priorities[normalized];
}

function stringValue_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function serializableCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value.toISOString();
  return value === null || value === undefined ? '' : String(value);
}

function hash_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );
  return bytes
    .map(function (byte) {
      const normalized = byte < 0 ? byte + 256 : byte;
      return ('0' + normalized.toString(16)).slice(-2);
    })
    .join('');
}

function runWithLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try {
    callback();
  } finally {
    lock.releaseLock();
  }
}
