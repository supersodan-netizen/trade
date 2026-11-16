#!/usr/bin/env node
'use strict';
// 73X39J C7N — READONLY OAuth refresh + PAPER HEAD verification.
// NO clasp, scripts.run, Sheets, updateContent, or source/credential in public logs.
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const SCOPE = 'https://www.googleapis.com/auth/script.projects.readonly';
const SCRIPT = '1RuAU2ofIUzbBj15BnpVheu9iC37uklCPujopAvI_kt7nRtRMOrVj8Dmq';
const BOOK = '1qz2YHslOULJ6fRi7VdZ41XiUQ4Geq1Tq654kL8UMSww';
const NAME = 'AUDITOR_73X39J_C18_PRECHECK_READ_ONLY';
const EXPECTED_SHA = '3966afd6a2971965b0f4a1ea3b478ae7869462f45362462fd0a870d8b375f4a6';
const MANIFEST = {timeZone:'America/New_York',dependencies:{},exceptionLogging:'STACKDRIVER',runtimeVersion:'V8'};
const SHA = s => crypto.createHash('sha256').update(s,'utf8').digest('hex');
function hold(code) { throw Error('C7N_HOLD_' + code); }
function canon(value) {
  if (Array.isArray(value)) return value.map(canon);
  if (value && typeof value === 'object')
    return Object.fromEntries(Object.keys(value).sort().map(k => [k,canon(value[k])]));
  return value;
}
function check(project, content, expected=EXPECTED_SHA) {
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
  if (JSON.stringify(canon(parsed)) !== JSON.stringify(canon(MANIFEST))) hold('MANIFEST_DRIFT');
  return {sha256:checksum,manifestSha256:SHA(manifest.source),filesCount:2};
}
function credentials(env) {
  const id=env.TRADE_GAS_RO_CLIENT_ID;
  const secret=env.TRADE_GAS_RO_CLIENT_SECRET;
  const refresh=env.TRADE_GAS_RO_REFRESH_TOKEN;
  if (typeof id !== 'string' || !/^[a-zA-Z0-9._-]{8,200}\.apps\.googleusercontent\.com$/.test(id))
    hold('CLIENT_ID_NOT_CONFIGURED');
  if (typeof secret !== 'string' || secret.length < 5 || secret.length > 1024 || /[\r\n]/.test(secret))
    hold('CLIENT_SECRET_NOT_CONFIGURED');
  if (typeof refresh !== 'string' || refresh.length < 20 || refresh.length > 4096 || /[\r\n]/.test(refresh))
    hold('REFRESH_TOKEN_NOT_CONFIGURED');
  return {id,secret,refresh};
}
function validateTokenResponse(t) {
  if (!t || typeof t.access_token !== 'string' || t.access_token.length < 20 ||
      t.access_token.length > 8192 || /[\r\n]/.test(t.access_token) || t.token_type !== 'Bearer')
    hold('ACCESS_TOKEN_INVALID');
  if (t.scope !== undefined && t.scope !== SCOPE) hold('TOKEN_RESPONSE_SCOPE');
  if (!Number.isFinite(Number(t.expires_in)) || Number(t.expires_in) <= 60)
    hold('ACCESS_TOKEN_SHORT_LIVED');
  return t.access_token;
}
function validateTokenInfo(t,id) {
  if (!t || t.scope !== SCOPE) hold('TOKENINFO_SCOPE_NOT_MINIMAL');
  if ((t.aud || t.audience) !== id || (t.azp && t.azp !== id))
    hold('TOKENINFO_AUDIENCE');
  if (!Number.isFinite(Number(t.expires_in)) || Number(t.expires_in) <= 60)
    hold('TOKENINFO_EXPIRY');
}
async function safeJson(url,opts,maxChars=40000) {
  let res;
  try {
    res = await fetch(url,{...opts,redirect:'error',signal:AbortSignal.timeout(12000)});
  } catch { hold('HTTPS_FAILED_OR_TIMEOUT'); }
  // Never echo the HTTP response body, as it could contain sensitive data.
  if (!res.ok) hold('GOOGLE_HTTP_' + res.status);
  let payload;
  try {
    const txt=await res.text();
    if (txt.length > maxChars) hold('REMOTE_RESPONSE_OVERSIZE');
    payload=JSON.parse(txt);
  } catch(e) {
    if (String(e.message).startsWith('C7N_HOLD_')) throw e;
    hold('REMOTE_RESPONSE_INVALID_JSON');
  }
  return payload;
}
async function freshToken(creds) {
  const form=new URLSearchParams({
    grant_type:'refresh_token',refresh_token:creds.refresh,client_id:creds.id,
    client_secret:creds.secret,scope:SCOPE
  });
  const token=validateTokenResponse(await safeJson('https://oauth2.googleapis.com/token',{
    method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:form.toString()
  }));
  // Independently check ACTUAL access token scopes, audience and expiry.
  const info=await safeJson('https://oauth2.googleapis.com/tokeninfo?access_token='+encodeURIComponent(token),{
    method:'GET',headers:{Accept:'application/json'}
  });
  validateTokenInfo(info,creds.id);
  return token;
}
async function googleGet(endpoint,token) {
  if (endpoint !== '/v1/projects/' + SCRIPT && endpoint !== '/v1/projects/' + SCRIPT + '/content')
    hold('GOOGLE_ENDPOINT_NOT_ALLOWED');
  return safeJson('https://script.googleapis.com'+endpoint,{
    method:'GET',headers:{Authorization:'Bearer '+token,Accept:'application/json'}
  },250000);
}
async function observe(token) {
  const metadata=await googleGet('/v1/projects/'+SCRIPT,token);
  const head=await googleGet('/v1/projects/'+SCRIPT+'/content',token);
  return check(metadata,head);
}
function selfTest() {
  const source='function testOnly() { return 1; }\n';
  const expected=SHA(source);
  const orig=()=>({project:{scriptId:SCRIPT,parentId:BOOK},
    content:{scriptId:SCRIPT,files:[{name:NAME,type:'SERVER_JS',source},
      {name:'appsscript',type:'JSON',source:JSON.stringify(MANIFEST)}]}});
  let n=0,a=orig();
  assert.equal(check(a.project,a.content,expected).sha256,expected);n++;
  const bad=(code,fn)=>{assert.throws(fn,new RegExp('C7N_HOLD_'+code));n++;};
  a=orig();a.project.parentId='OTHER';bad('PAPER_PARENT_ID',()=>check(a.project,a.content,expected));
  a=orig();a.project.scriptId='OTHER';bad('PROJECT_ID',()=>check(a.project,a.content,expected));
  a=orig();a.content.scriptId='OTHER';bad('CONTENT_ID',()=>check(a.project,a.content,expected));
  a=orig();a.content.files.push({name:'extra',type:'SERVER_JS',source:'x'});bad('UNEXPECTED_FILE_COUNT',()=>check(a.project,a.content,expected));
  a=orig();a.content.files[0].name='Codigo';bad('UNEXPECTED_FILE_SET',()=>check(a.project,a.content,expected));
  a=orig();a.content.files[0].source+='// drift';bad('REMOTE_CODE_SHA_DRIFT',()=>check(a.project,a.content,expected));
  a=orig();a.content.files[1].source='{invalid';bad('MANIFEST_NOT_JSON',()=>check(a.project,a.content,expected));
  a=orig();a.content.files[1].source=JSON.stringify({...MANIFEST,timeZone:'Europe/Madrid'});bad('MANIFEST_DRIFT',()=>check(a.project,a.content,expected));
  bad('CLIENT_ID_NOT_CONFIGURED',()=>credentials({}));
  const cid='123456789012-abc.apps.googleusercontent.com';
  const goodEnv={TRADE_GAS_RO_CLIENT_ID:cid,TRADE_GAS_RO_CLIENT_SECRET:'valid_secret',TRADE_GAS_RO_REFRESH_TOKEN:'long_refresh_token_value_123456789'};
  const creds=credentials(goodEnv);assert.equal(creds.id,cid);n++;
  bad('CLIENT_SECRET_NOT_CONFIGURED',()=>credentials({...goodEnv,TRADE_GAS_RO_CLIENT_SECRET:''}));
  bad('REFRESH_TOKEN_NOT_CONFIGURED',()=>credentials({...goodEnv,TRADE_GAS_RO_REFRESH_TOKEN:'x'}));
  const t={access_token:'valid_access_token_length_123456789',token_type:'Bearer',expires_in:3600,scope:SCOPE};
  assert.equal(validateTokenResponse(t),t.access_token);n++;
  bad('TOKEN_RESPONSE_SCOPE',()=>validateTokenResponse({...t,scope:SCOPE+' https://www.googleapis.com/auth/script.projects'}));
  bad('ACCESS_TOKEN_SHORT_LIVED',()=>validateTokenResponse({...t,expires_in:15}));
  const info={aud:cid,azp:cid,scope:SCOPE,expires_in:3500};
  validateTokenInfo(info,cid);n++;
  bad('TOKENINFO_SCOPE_NOT_MINIMAL',()=>validateTokenInfo({...info,scope:SCOPE+' https://www.googleapis.com/auth/script.projects'},cid));
  bad('TOKENINFO_AUDIENCE',()=>validateTokenInfo({...info,aud:'other'},cid));
  bad('TOKENINFO_EXPIRY',()=>validateTokenInfo({...info,expires_in:0},cid));
  return n;
}
async function main() {
  if (process.argv.length===3 && process.argv[2]==='--self-test') {
    console.log('TRADE_C7N_OFFLINE_OAUTH_GUARDS_PASS_'+selfTest()+'_NO_GOOGLE');return;
  }
  if (process.argv.length!==2) hold('INVALID_ARGUMENTS');
  if (process.env.GITHUB_EVENT_NAME!=='workflow_dispatch' ||
      process.env.GITHUB_REPOSITORY!=='supersodan-netizen/trade' ||
      process.env.GITHUB_REF!=='refs/heads/main' ||
      process.env.GITHUB_ACTOR!=='supersodan-netizen' ||
      process.env.TRADE_PAPER_READBACK_CONFIRM!=='READ_PAPER_HEAD_ONLY')
    hold('WORKFLOW_CONTEXT');
  const creds=credentials(process.env);
  const token=await freshToken(creds);
  const first=await observe(token),second=await observe(token);
  if (JSON.stringify(first)!==JSON.stringify(second)) hold('DOUBLE_READBACK_DRIFT');
  console.log('TRADE_C7N_PAPER_HEAD_MATCH_TWICE_READONLY_NO_GAS_NO_DEPLOY_W_HOLD');
}
main().catch(e=>{
  const message=String(e && e.message || '');
  console.error(message.startsWith('C7N_HOLD_') ? message : 'C7N_HOLD_UNCLASSIFIED_FAILURE');
  process.exitCode=2;
});
