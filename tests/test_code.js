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
assert.match(description, /range=D8/);

console.log('All pure-function checks passed.');

