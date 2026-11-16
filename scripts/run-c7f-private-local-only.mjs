#!/usr/bin/env node
// TRADE public-runner wrapper: execute the private C7F workflow's 24 local-only shell blocks.
// All shell output remains ephemeral; public logs only show safe PASS/failed step metadata.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';

const root = path.resolve(process.argv[2] || '');
const sourceFile = path.join(root, '.github/workflows/73x39j-c18-p0c7f-writer-boundary-inventory-local.yml');
function hold(reason) { console.error('TRADE_C7F_HOLD_' + reason); process.exit(2); }
if (!process.argv[2] || !fs.existsSync(sourceFile)) hold('SOURCE_WORKFLOW_NOT_FOUND');
if (fs.existsSync(path.join(root, '.clasp.json')) || fs.existsSync(path.join(root, '.clasprc.json'))) hold('CLASP_CONFIG_FOUND');
let config;
try { config = JSON.parse(fs.readFileSync(path.join(root, 'config.json'), 'utf8')); }
catch { hold('CONFIG_UNREADABLE'); }
if (config.liveDeploymentEnabled !== false || config.automaticStagingDeployment !== false ||
    config.destination !== 'LOCAL_ONLY_NO_DEPLOY') hold('LOCAL_ONLY_CONFIG_GUARD');
let core;
try { core = fs.readFileSync(path.join(root, 'candidate/Codigo.gs')); }
catch { hold('CORE_MISSING'); }
const coreSha = crypto.createHash('sha256').update(core).digest('hex');
if (coreSha !== 'b4415d13b6efcc0036a5220edb9a3f388170a41924031dbd5516b02278aac25b') hold('CORE_SHA_CHANGED');
const lines = fs.readFileSync(sourceFile, 'utf8').replace(/\r\n/g, '\n').split('\n');
const start = lines.indexOf('    steps:');
if (start < 0) hold('SOURCE_STEPS_MISSING');
const steps = [];
let name = null;
for (let i = start + 1; i < lines.length; i++) {
  const title = lines[i].match(/^      - name: (.+)$/);
  if (title) {
    name = title[1].replace(/[^a-z0-9 _.,:-]/gi, '').slice(0, 90);
    continue;
  }
  if (lines[i] === '        run: |') {
    if (!name) hold('UNNAMED_STEP');
    const scriptLines = [];
    for (i += 1; i < lines.length; i++) {
      const line = lines[i];
      if (line !== '' && !line.startsWith('          ')) { i--; break; }
      scriptLines.push(line === '' ? '' : line.slice(10));
    }
    steps.push({ name, script: scriptLines.join('\n') });
    name = null;
  }
}
if (steps.length !== 24 || steps.some(step => !step.script.trim())) hold('SOURCE_STRUCTURE_CHANGED');
const logsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'trade-c7f-'));
try {
  for (let i = 0; i < steps.length; i++) {
    const log = path.join(logsDir, 'step-' + (i + 1) + '.log');
    const fd = fs.openSync(log, 'w', 0o600);
    const result = spawnSync('bash', ['--noprofile', '--norc', '-e', '-o', 'pipefail', '-c', steps[i].script], {
      cwd: root, env: { ...process.env, CI: 'true' }, stdio: ['ignore', fd, fd], timeout: 12 * 60 * 1000
    });
    fs.closeSync(fd);
    if (result.error || result.status !== 0) {
      console.error('TRADE_C7F_FAIL step=' + (i + 1) + '/24 name=' + JSON.stringify(steps[i].name));
      process.exitCode = 1;
      break;
    }
  }
  if (!process.exitCode) console.log('TRADE_C7F_PASS_24_SOURCE_STEPS_1048_HISTORICAL_LOCAL_CASES_NO_GAS_NO_LIVE');
} finally {
  fs.rmSync(logsDir, { recursive: true, force: true });
}
