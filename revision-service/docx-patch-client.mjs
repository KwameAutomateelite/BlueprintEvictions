import {spawn} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {normalizeWordLayout} from './normalize-word.mjs';
export function createDocxPatchClient({python,normalizeWord=normalizeWordLayout}) {
 if(typeof python!=='string'||!python.startsWith('/'))throw Error('Absolute Python runtime required');
 return ({mode,bytes,...input})=>new Promise((resolve,reject)=>{
  if(!Buffer.isBuffer(bytes)||bytes.length>10*1024*1024)return reject(Error('Document size requires review'));
  const child=spawn(python,[fileURLToPath(new URL('./docx_patch_cli.py',import.meta.url))],{stdio:['pipe','pipe','pipe'],shell:false});let output=Buffer.alloc(0),done=false;
  const finish=(err,value)=>{if(done)return;done=true;clearTimeout(timer);err?reject(err):resolve(value)};
  const timer=setTimeout(()=>{child.kill();finish(Error('Revision timed out'))},30000);
  child.on('error',()=>finish(Error('Revision worker unavailable')));child.stderr.resume();child.stdin.on('error',()=>{});
  child.stdout.on('data',b=>{output=Buffer.concat([output,b]);if(output.length>15*1024*1024){child.kill();finish(Error('Revision output too large'))}});
  child.on('close',async code=>{try{if(code!==0)throw Error('Revision requires clarification');let r=JSON.parse(output);if(r.docx_base64 && r.printed_signer)r=await normalizeWord(r);if(r.docx_base64){r.bytes=Buffer.from(r.docx_base64,'base64');delete r.docx_base64}finish(null,r)}catch(e){finish(e)}});
  child.stdin.end(JSON.stringify({...input,mode,docx_base64:bytes.toString('base64')}));
 });
}
