#!/usr/bin/env node
'use strict';
// TRADE 73X39J / C7M: PAPER 39J HEAD observer. No private checkout, clasp, Sheets, writes or scripts.run.
// A Google OAuth access token MUST be minted with ONLY script.projects.readonly and kept as an environment secret.
const crypto = require('node:crypto');
const assert = require('node:assert/strict');

const SCRIPT = '1RuAU2ofIUzbBj15BnpVheu9iC37uklCPujopAvI_kt7nRtRMOrVj8Dmq';
const BOOK = '1qz2YHslOULJ6fRi7VdZ41XiUQ4Geq1Tq654kL8UMSww';
const NAME = 'AUDITOR_73X39J_C18_PRECHECK_READ_ONLY';
const EXPECTED_SHA = '3966afd6a2971965b0f4a1ea3b478ae7869462f45362462fd0a870d8b375f4a6';
const MANIFEST = {timeZone:'America/New_York',dependencies:{},exceptionLogging:'STACKDRIVER',runtimeVersion:'V8'};
const SHA = s => crypto.createHash('sha256').update(s,'utf8').digest('hex');
function hold(code) { throw Error('C7M_HOLD_' + code); }
function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object')
    return Object.fromEntries(Object.keys(value).sort().map(k => [k,canonical(value[k])]));
  return value;
}
function check(project, content, expected = EXPECTED_SHA) {
  if (!project || project.scriptId !== SCRIPT) hold('PROJECT_ID');
  if (project.parentId !== BOOK) hold('PAPER_PARENT_ID');
  if (!content || content.scriptId !== SCRIPT || !Array.isArray(content.files)) hold('CONTENT_ID');
  if (content.files.length !== 2) hold('UNEXPECTED_FILE_COUNT');
  const code = content.files.find(f => f && f.name === NAME && f.type === 'SERVER_JS');
  const manifest = content.files.find(f => f && f.name === 'appsscript' && f.type === 'JSON');
  if (!code || !manifest || typeof code.source !== 'string' || typeof manifest.source !== 'string')
    hold('UNEXPECTED_FILE_SET');
  const checksum = SHA(code.source);
  if (checksum !== expected) hold('REMOTE_CODE_SHA_DRIFT');
  let parsed;
  try { parsed=JSON.parse(manifest.source); } catch { hold('MANIFEST_NOT_JSON'); }
  if (JSON.stringify(canonical(parsed)) !== JSON.stringify(canonical(MANIFEST))) hold('MANIFEST_DRIFT');
  return {sha256:checksum, manifestSha256:SHA(manifest.source), filesCount:2};
}
async function googleGet(endpoint, token) {
  if (!endpoint.startsWith('/v1/projects/' + SCRIPT)) hold('URL_NOT_ALLOWLISTED');
  let response;
  try {
    response = await fetch('https://script.googleapis.com' + endpoint, {
      method:'GET', redirect:'error', headers:{Authorization:'Bearer ' + token, Accept:'application/json'},
      signal:AbortSignal.timeout(12000)
    });
  } catch { hold('GOOGLE_CONNECTION_OR_TIMEOUT'); }
  if (!response.ok) hold('GOOGLE_HTTP_' + response.status);
  let body;
  try {
    const t=await response.text();
    if (t.length > 250000) hold('OVERSIZE_REMOTE_RESPONSE');
    body=JSON.parse(t);
  } catch(e) {
    if (String(e.message).startsWith('C7M_HOLD_')) throw e;
    hold('INVALID_REMOTE_JSON');
  }
  return body;
}
async function observe(token) {
  const metadata=await googleGet('/v1/projects/' + SCRIPT,token);
  const head=await googleGet('/v1/projects/' + SCRIPT + '/content',token);
  return check(metadata,head);
}
function selfTest() {
  const source='function testOnly() { return 1; }\n';
  const expected=SHA(source);
  const original=()=>({
    project:{scriptId:SCRIPT,parentId:BOOK},
    content:{scriptId:SCRIPT,files:[
      {name:NAME,type:'SERVER_JS',source},
      {name:'appsscript',type:'JSON',source:JSON.stringify(MANIFEST)}
    ]}
  });
  let n=0;
  const good=original();
  assert.equal(check(good.project,good.content,expected).sha256,expected); n++;
  const negative=(code,fn)=>{assert.throws(fn,new RegExp('^Error: C7M_HOLD_'+code));n++;};
  let a=original();a.project.parentId='OTHER';negative('PAPER_PARENT_ID',()=>check(a.project,a.content,expected));
  a=original();a.project.scriptId='OTHER';negative('PROJECT_ID',()=>check(a.project,a.content,expected));
  a=original();a.content.scriptId='OTHER';negative('CONTENT_ID',()=>check(a.project,a.content,expected));
  a=original();a.content.files.push({name:'extra',type:'SERVER_JS',source:'x'});negative('UNEXPECTED_FILE_COUNT',()=>check(a.project,a.content,expected));
  a=original();a.content.files[0].name='Codigo';negative('UNEXPECTED_FILE_SET',()=>check(a.project,a.content,expected));
  a=original();a.content.files[0].source+='// drift';negative('REMOTE_CODE_SHA_DRIFT',()=>check(a.project,a.content,expected));
  a=original();a.content.files[1].source='{invalid';negative('MANIFEST_NOT_JSON',()=>check(a.project,a.content,expected));
  a=original();a.content.files[1].source=JSON.stringify({...MANIFEST,timeZone:'Europe/Madrid'});negative('MANIFEST_DRIFT',()=>check(a.project,a.content,expected));
  return n;
}
async function main() {
  if (process.argv.length === 3 && process.argv[2] === '--self-test') {
    console.log('TRADE_C7M_LOCAL_ONLY_SELF_TEST_PASS_' + selfTest() + '_NO_GAS'); return;
  }
  if (process.argv.length !== 2) hold('INVALID_ARGUMENTS');
  if (process.env.GITHUB_EVENT_NAME !== 'workflow_dispatch' ||
      process.env.GITHUB_REPOSITORY !== 'supersodan-netizen/trade' ||
      process.env.GITHUB_REF !== 'refs/heads/main' ||
      process.env.GITHUB_ACTOR !== 'supersodan-netizen' ||
      process.env.TRADE_PAPER_READBACK_CONFIRM !== 'READ_PAPER_HEAD_ONLY') hold('WORKFLOW_CONTEXT');
  const token=process.env.TRADE_GAS_PAPER_READONLY_ACCESS_TOKEN;
  if (!token || token.length < 20 || token.length > 8192 || /[\r\n]/.test(token)) hold('READONLY_TOKEN_NOT_CONFIGURED');
  const one=await observe(token);
  const two=await observe(token);
  if (JSON.stringify(one) !== JSON.stringify(two)) hold('DOUBLE_READBACK_DRIFT');
  // This is an observed HEAD match, never a proof that all writers are excluded or that GAS ran.
  console.log('TRADE_C7M_REMOTE_HEAD_MATCH_OBSERVED_TWICE_READONLY_NO_GAS_EXECUTION_NO_DEPLOY_W_HOLD');
}
main().catch(e=>{
  const m=String(e && e.message || '');
  console.error(m.startsWith('C7M_HOLD_')?m:'C7M_HOLD_UNCLASSIFIED_FAILURE');
  process.exitCode=2;
});