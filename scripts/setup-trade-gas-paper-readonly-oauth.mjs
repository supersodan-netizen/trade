#!/usr/bin/env node
// TRADE 73X39J C7N: one-time LOCAL desktop OAuth PKCE setup, never run in GitHub Actions.
// Reads a Google Desktop OAuth JSON file from a LOCAL path. Never writes tokens to disk.
import fs from 'node:fs';
import http from 'node:http';
import crypto from 'node:crypto';

const SCOPE='https://www.googleapis.com/auth/script.projects.readonly';
const fail=code=>{throw Error('C7N_SETUP_'+code);};
const credFile=process.argv[2];
function parseDesktop(raw) {
  const config=JSON.parse(raw);
  if (!config || !config.installed || typeof config.installed!=='object') fail('NOT_A_DESKTOP_OAUTH_CLIENT_JSON');
  const {client_id:id,client_secret:secret}=config.installed;
  if (!/^[a-zA-Z0-9._-]{8,200}\.apps\.googleusercontent\.com$/.test(id||'')) fail('CLIENT_ID_INVALID');
  if (typeof secret!=='string' || secret.length<5 || secret.length>1024 || /[\r\n]/.test(secret))
    fail('CLIENT_SECRET_INVALID');
  return {id,secret};
}
function selfTest(){
  const sample={installed:{client_id:'1234567890-test.apps.googleusercontent.com',client_secret:'example_test_secret'}};
  if(parseDesktop(JSON.stringify(sample)).id!==sample.installed.client_id)fail('SELF_TEST');
  let negative=false;
  try{parseDesktop(JSON.stringify({web:sample.installed}));}catch(e){negative=String(e.message).startsWith('C7N_SETUP_NOT_A_DESKTOP');}
  if(!negative)fail('SELF_TEST_NEGATIVE');
  console.log('TRADE_C7N_LOCAL_OAUTH_SETUP_2_TESTS_PASS_NO_GOOGLE');
}
async function main() {
  if(credFile==='--self-test' && process.argv.length===3){selfTest();return;}
  if(process.env.GITHUB_ACTIONS || process.env.CI==='true')fail('INTERACTIVE_LOCAL_ONLY');
  if(process.argv.length!==3 || !credFile)fail('USAGE_NODE_SCRIPT_CLIENT_JSON_FILE');
  let config;
  try{config=parseDesktop(fs.readFileSync(credFile,'utf8'));}catch(e){
    if(String(e.message).startsWith('C7N_SETUP_'))throw e;
    fail('CLIENT_JSON_UNREADABLE');
  }
  const verifier=crypto.randomBytes(48).toString('base64url');
  const challenge=crypto.createHash('sha256').update(verifier).digest('base64url');
  const state=crypto.randomBytes(24).toString('base64url');
  const server=http.createServer();
  await new Promise((resolve,reject)=>{
    server.once('error',reject);
    server.listen(0,'127.0.0.1',resolve);
  });
  const port=server.address().port;
  const callback='http://127.0.0.1:'+port+'/oauth-callback';
  const url=new URL('https://accounts.google.com/o/oauth2/v2/auth');
  for(const [k,v] of Object.entries({
    client_id:config.id,redirect_uri:callback,response_type:'code',
    scope:SCOPE,access_type:'offline',prompt:'consent',state,
    code_challenge:challenge,code_challenge_method:'S256'
  }))url.searchParams.set(k,v);
  console.log('\nTRADE · autorización SOLO LECTURA de Google Apps Script');
  console.log('Abre este enlace en el navegador de ESTE ordenador:\n');
  console.log(url.toString());
  console.log('\nInicia sesión con la cuenta propietaria de la copia PAPER 39J.');
  console.log('Acepta SOLO el permiso de lectura del código de Apps Script.');
  console.log('La autorización terminará aquí automáticamente, sin pegar códigos ni tokens en chats.');
  const result=await new Promise((resolve,reject)=>{
    const timer=setTimeout(()=>reject(Error('C7N_SETUP_TIMEOUT_5_MIN')),300000);
    server.on('request',(req,res)=>{
      res.setHeader('Content-Type','text/html; charset=utf-8');
      res.setHeader('Cache-Control','no-store');
      const wrong=()=>{res.writeHead(400);res.end('Solicitud no válida.');};
      try{
        if(req.method!=='GET' || !req.url || req.url.length>4096)return wrong();
        const target=new URL(req.url,callback);
        if(target.pathname!=='/oauth-callback')return wrong();
        if(target.searchParams.get('state')!==state)return wrong();
        const code=target.searchParams.get('code');
        if(!code || code.length>4096 || target.searchParams.has('error')){
          res.writeHead(400);res.end('Autorización cancelada.');clearTimeout(timer);
          reject(Error('C7N_SETUP_AUTH_NOT_GRANTED'));return;
        }
        res.writeHead(200);
        res.end('<!doctype html><meta charset="utf-8"><title>TRADE</title><p>Google ha devuelto la autorización a este ordenador. Puedes cerrar esta pestaña.</p>');
        clearTimeout(timer);resolve(code);
      }catch{wrong();}
    });
  }).finally(()=>server.close());
  // This code and its secrets exist ONLY in local process memory.
  const body=new URLSearchParams({
    grant_type:'authorization_code',code:result,client_id:config.id,
    client_secret:config.secret,redirect_uri:callback,code_verifier:verifier
  });
  let resp;
  try{resp=await fetch('https://oauth2.googleapis.com/token',{
    method:'POST',redirect:'error',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:body.toString(),signal:AbortSignal.timeout(12000)
  });}catch{fail('TOKEN_EXCHANGE_CONNECTION');}
  if(!resp.ok)fail('TOKEN_EXCHANGE_HTTP_'+resp.status);
  let token;try{token=await resp.json();}catch{fail('TOKEN_EXCHANGE_JSON');}
  if(token.scope!==SCOPE || token.token_type!=='Bearer')fail('GRANT_SCOPE_MISMATCH');
  if(typeof token.refresh_token!=='string'||token.refresh_token.length<20)fail('NO_OFFLINE_REFRESH_TOKEN');
  if(token.refresh_token_expires_in)console.log('\nAVISO: Google ha emitido refresh token con caducidad indicada.');
  console.log('\nCONEXIÓN AUTORIZADA SOLO LECTURA. Guarda estos tres valores como secrets del environment trade-gas-paper-readonly en GitHub:');
  console.log('\nTRADE_GAS_RO_CLIENT_ID='+config.id);
  console.log('TRADE_GAS_RO_CLIENT_SECRET='+config.secret);
  console.log('TRADE_GAS_RO_REFRESH_TOKEN='+token.refresh_token);
  console.log('\nNO compartas esta salida, NO hagas capturas, NO subas el JSON o los tokens a GitHub.');
  console.log('Una vez guardados los 3 secrets, limpia la terminal y borra el JSON descargado si no lo necesitas.');
}
main().catch(e=>{
  const msg=String(e&&e.message||'');
  console.error(msg.startsWith('C7N_SETUP_')?msg:'C7N_SETUP_UNCLASSIFIED');
  process.exitCode=2;
});
