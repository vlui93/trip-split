// Makes sure the Apps Script manifest still describes a web app that runs as the
// owner and is open to anyone with the link, which is what the phones rely on.
// Everything else in the manifest (time zone, scopes) is left as the project has it.
import { readFileSync, writeFileSync } from 'node:fs';

const file = process.argv[2];
const m = JSON.parse(readFileSync(file, 'utf8'));
m.runtimeVersion = m.runtimeVersion || 'V8';
m.webapp = Object.assign({ executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' }, m.webapp);
writeFileSync(file, JSON.stringify(m, null, 2) + '\n');
console.log('Manifest: web app runs as', m.webapp.executeAs + ', access', m.webapp.access);
