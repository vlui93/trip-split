// Runs the test files one after another and fails if any of them does.
//   node tools/run-tests.mjs          unit tests, then browser tests
//   node tools/run-tests.mjs unit     tests/*.test.mjs
//   node tools/run-tests.mjs e2e      tests/e2e/*.e2e.mjs (needs Playwright)
import { readdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));
const sets = {
  unit: ['tests', /\.test\.mjs$/],
  e2e: ['tests/e2e', /\.e2e\.mjs$/],
};
const want = process.argv[2] ? [process.argv[2]] : ['unit', 'e2e'];
const failed = [];
let ran = 0;
for (const name of want) {
  if (!sets[name]) { console.error(`Unknown set "${name}" (unit or e2e)`); process.exit(2); }
  const [dir, rx] = sets[name];
  for (const f of readdirSync(path.join(root, dir)).filter(f => rx.test(f)).sort()) {
    const rel = `${dir}/${f}`;
    console.log(`\n--- ${rel}`);
    const r = spawnSync(process.execPath, [rel], { cwd: root, stdio: 'inherit' });
    ran++;
    if (r.status !== 0) failed.push(rel);
  }
}
console.log(`\n${ran - failed.length} of ${ran} test files passed`);
if (failed.length) { console.error('Failed: ' + failed.join(', ')); process.exit(1); }
