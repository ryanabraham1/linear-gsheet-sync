const assert = require('node:assert');
const crypto = require('node:crypto');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync('Code.gs', 'utf8');
const context = {
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' },
    Charset: { UTF_8: 'utf8' },
    computeDigest(algorithm, value) {
      return Array.from(crypto.createHash(algorithm).update(value).digest()).map((byte) =>
        byte > 127 ? byte - 256 : byte
      );
    },
    formatDate(date) {
      return date.toISOString().slice(0, 10);
    },
  },
};
vm.createContext(context);
vm.runInContext(source, context);

assert.equal(context.priorityValue_('Urgent'), 1);
assert.equal(context.priorityValue_('medium'), 3);
assert.equal(context.priorityValue_('No priority'), 0);
assert.equal(context.priorityValue_('#1'), 1);
assert.equal(context.priorityValue_('#4'), 4);
assert.equal(context.linearStatusName_('Finished'), 'Done');
assert.equal(context.linearStatusName_('Not Started'), 'Todo');
assert.equal(context.linearStatusName_('In Progress'), 'In Progress');
assert.equal(context.linearIssueTitle_('0100_Bellypan'), 'Fab: 0100_Bellypan');
assert.equal(context.linearIssueTitle_('Fab: 0100_Bellypan'), 'Fab: 0100_Bellypan');
assert.equal(context.resolveBotProject_(''), null);
assert.equal(context.resolveBotProject_('Dumper').projectName, 'Dumper');
assert.equal(context.resolveBotProject_('everybot').projectName, 'EveryBot');
assert.equal(context.resolveBotProject_('AIMBOT').projectName, 'Aimbot Changes');
assert.throws(() => context.resolveBotProject_('Mystery Bot'), /Unknown Bot/);
assert.equal(context.labelNameFromCell_('01. Drivebase '), 'Drivebase');
assert.equal(context.labelNameFromCell_('05 - Intake'), 'Intake');
assert.equal(context.labelNameFromCell_('Bridgeport Mill'), 'Bridgeport Mill');
assert.deepEqual(
  Array.from(context.uniqueStrings_(['a', 'b', 'a'])),
  ['a', 'b']
);
assert.deepEqual(
  Array.from(context.finalIssueLabelIds_(
    {
      labelIds: ['fabrication'],
      labelGroups: {
        subsystem: { id: 'subsystem-group' },
        machine: { id: 'machine-group' },
      },
    },
    [
      { id: 'fabrication', parent: null },
      { id: 'manual-label', parent: null },
      { id: 'old-subsystem', parent: { id: 'subsystem-group' } },
      { id: 'old-machine', parent: { id: 'machine-group' } },
    ],
    ['new-subsystem', 'new-machine']
  )),
  ['fabrication', 'manual-label', 'new-subsystem', 'new-machine']
);
assert.doesNotMatch(source, /addedLabelIds|removedLabelIds/);
assert.equal(
  context.isArchivedBySync_('[Archived by Sheets sync] Part #_Name was cleared.'),
  true
);
assert.equal(context.shouldSkipHash_('abc', 'abc', 'issue-id'), true);
assert.equal(context.shouldSkipHash_('abc', 'abc', ''), false);
assert.match(context.hash_('test'), /^[a-f0-9]{64}$/);

const raw = {
  Subsystem: '05. Intake',
  Qty: 4,
  'Spare Qty': '',
  'Stock Material/Type': 'Aluminum Sheet',
  'Stock Dimensions': '1/8" thick',
  Length: '',
  'Tapped?': '',
  Machine: 'Mill',
  'Drawing/CAM File (STEP FILES)': 'https://example.com/file',
  Notes: 'Test note',
};
const description = context.machiningDescription_(raw, 8);
assert.match(description, /\*\*Subsystem:\*\* 05\. Intake/);
assert.match(description, /\[Open file\]\(https:\/\/example\.com\/file\)/);
assert.match(description, /range=E8/);

console.log('All pure-function checks passed.');
