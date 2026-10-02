import {fileError} from './file-rules.mjs';
const $ = id => document.getElementById(id);
let chosen = null, worker = null, active = null, busy = false, timer = null, labels = {};
function status(text, running=false) { $('status').textContent=text; $('status').classList.toggle('busy',running); }
function error(text) { $('error').textContent=text; $('error').hidden=!text; }
function clearResult() { $('scored').hidden=true; $('empty').hidden=false; }
function choose(file) {
  if(busy) return;
  chosen=null;clearResult();const problem=fileError(file);error(problem);
  $('file-name').textContent=problem?'Drop your spreadsheet here':file.name;
  $('file-detail').textContent=problem?'Excel .xlsx · up to 10 MB':`${(file.size/1024).toFixed(0)} KB · Ready to check`;
  if(!problem) {chosen=file;status('Ready. Confirm the lab rubric, then check your file.');}
}
function finish() { clearTimeout(timer);busy=false;active=null; $('check').disabled=false;$('lab').disabled=false;$('file').disabled=false;$('cancel').hidden=true;$('status').classList.remove('busy'); }
function stop() {worker?.terminate();worker=null;finish();}
function getWorker() {
  if(worker) return worker;
  worker=new Worker(new URL('./worker.mjs',import.meta.url),{type:'module'});
  worker.onmessage=({data})=>{
    if(data.id!==active) return;
    if(data.stage) {status(data.stage,true);return;}
    if(data.error) {finish();error(data.error);status('No score was assigned. Use the message above to fix the file.');return;}
    if(data.result) {
      const result=data.result;finish();error('');$('empty').hidden=true;$('scored').hidden=false;
      $('score').textContent=String(result.score);$('meter-fill').style.width=`${result.score*10}%`;
      $('result-lab').textContent=labels[result.lab];$('feedback').replaceChildren();
      for(const note of result.feedback) {const item=document.createElement('li');item.textContent=note;if(note.startsWith('Table/header placement'))item.classList.add('placement');$('feedback').append(item);}
      status('Check complete. Review the feedback, revise if needed, and submit your work in Canvas.');
      $('results').scrollIntoView({behavior:'smooth',block:'nearest'});
    }
  };
  worker.onerror=()=>{stop();error('The checker couldn’t start. Check your internet connection, refresh, and try again.');status('Checker unavailable. Your file has not been submitted anywhere.');};
  return worker;
}
$('file').addEventListener('change',e=>choose(e.target.files[0]));
$('lab').addEventListener('change',()=>{clearResult();error('');});
const drop=$('drop-zone');
for(const name of ['dragenter','dragover'])drop.addEventListener(name,e=>{e.preventDefault();if(!busy)drop.classList.add('dragging');});
for(const name of ['dragleave','drop'])drop.addEventListener(name,e=>{e.preventDefault();drop.classList.remove('dragging');});
drop.addEventListener('drop',e=>{if(e.dataTransfer.files.length!==1)error('Choose one spreadsheet at a time.');else choose(e.dataTransfer.files[0]);});
$('cancel').addEventListener('click',()=>{stop();status('Check canceled. You can choose another file or try again.');});
$('check-form').addEventListener('submit',async e=>{
  e.preventDefault();if(busy)return;error('');
  const problem=fileError(chosen);if(problem){error(problem);return;}
  if(!$('lab').value){error('Choose the rubric for your lab.');return;}
  const lab=$('lab').value;clearResult();busy=true;$('check').disabled=true;$('lab').disabled=true;$('file').disabled=true;$('cancel').hidden=false;
  active=crypto.randomUUID();const id=active;
  status('Preparing the checker. Your first visit may take a little longer…',true);
  timer=setTimeout(()=>{stop();error('The check took too long. Try again, or ask your instructor to review this file.');status('No score was assigned.');},120000);
  try {const buffer=await chosen.arrayBuffer();if(active!==id)return;getWorker().postMessage({id,lab,buffer},[buffer]);}
  catch {stop();error('Unable to read this file. Choose the saved spreadsheet again.');}
});
try {
  const response=await fetch(new URL('./labs.json',import.meta.url));if(!response.ok)throw Error();labels=await response.json();
  for(const [value,label]of Object.entries(labels)){const option=document.createElement('option');option.value=value;option.textContent=label;$('lab').append(option);}
} catch {$('check').disabled=true;error('Lab rubrics couldn’t load. Refresh the page and try again.');}
