const DAYS=[['ISNIN','Isnin'],['SELASA','Selasa'],['RABU','Rabu'],['KHAMIS','Khamis'],['JUMAAT','Jumaat']];
const $=id=>document.getElementById(id);const state={};
function iso(date){return date.toISOString().slice(0,10)}
function toMonday(date){const d=new Date(date);const day=d.getDay()||7;d.setDate(d.getDate()-day+1);return d}
function msDate(date){return new Intl.DateTimeFormat('ms-MY',{weekday:'long',day:'numeric',month:'long'}).format(date)}
function outputMode(){return 'jawi'}
function outputLabel(){return 'Jawi'}
function initialise(){
  for(let n=1;n<=45;n++)$('week').add(new Option(`Minggu ${n}`,n));
  const monday=toMonday(new Date());$('monday').value=iso(monday);$('week').value=String(getIsoWeek(monday));
  renderDays();
  $('monday').addEventListener('change',()=>{const selected=new Date($('monday').value+'T12:00:00');$('monday').value=iso(toMonday(selected));renderDays();hideClassroomCard()});
  $('week').addEventListener('change',hideClassroomCard);
  $('selectAll').addEventListener('click',()=>{document.querySelectorAll('[data-day]').forEach(el=>el.checked=true)});
  $('generateButton').addEventListener('click',generate);
  $('prepareClassroomButton').addEventListener('click',prepareClassroom);
  $('copySheetLink').addEventListener('click',copySheetLink)
}
function updateGenerateButtonUi(){
  $('generateButton').innerHTML='<span>✦</span> Jana eRPH Mingguan — Jawi'
}
function getIsoWeek(d){const x=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));const day=x.getUTCDay()||7;x.setUTCDate(x.getUTCDate()+4-day);const y=new Date(Date.UTC(x.getUTCFullYear(),0,1));return Math.ceil((((x-y)/86400000)+1)/7)}
function renderDays(){const monday=new Date($('monday').value+'T12:00:00');$('mondayHint').textContent=`Minggu ini bermula ${msDate(monday)}.`;$('daysList').innerHTML=DAYS.map(([key,label],i)=>{const date=new Date(monday);date.setDate(monday.getDate()+i);return `<div class="day-row"><div class="day-date"><strong>${label}</strong><span>${msDate(date)}</span></div><label class="switch"><input data-day="${key}" type="checkbox" checked aria-label="Jana ${label}"><span class="slider"></span></label></div>`}).join('')}
async function postToErph(body){let response;try{response=await fetch('/api/generate-erph',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)})}catch{throw new Error('Sambungan terputus sebelum pelayan membalas. Penjanaan mungkin masih berjalan — tunggu 2–3 minit dan semak Google Sheet sebelum cuba lagi. Pastikan skrin tidak dikunci semasa menjana.')}const raw=await response.text();let result;try{result=raw?JSON.parse(raw):null}catch{throw new Error('PWA menerima respons tidak sah. Sila muat semula aplikasi dan cuba lagi.')}if(!response.ok||!result||!result.ok)throw new Error(result?.error||'Permintaan eRPH gagal.');return result}
async function generate(){
  const days=Object.fromEntries([...document.querySelectorAll('[data-day]')].map(input=>[input.dataset.day,input.checked]));
  if(!Object.values(days).some(Boolean)){showStatus('Pilih sekurang-kurangnya satu hari','Tandakan hari yang hendak dijana.',false);return}
  hideClassroomCard();$('generateButton').disabled=true;
  showStatus(`AI sedang menjana eRPH ${outputLabel()}`,outputMode()==='jawi'?'Kandungan sedang dijana terus dalam tulisan Jawi.':'Semua slot bagi hari yang dipilih sedang diproses.',true);
  try{
    const result=await postToErph({action:'generate_week',outputScript:outputMode(),week:Number($('week').value),monday:$('monday').value,days});
    state.lastWorkbookUrl=result.workbookUrl||'';
    const summary=result.summary&&typeof result.summary==='object'?result.summary:{};
    const detail=Object.entries(summary).map(([day,count])=>`${day}: ${count} slot`).join(' · ');
    showStatus(`eRPH ${result.outputLabel||outputLabel()} berjaya dijana`,`${result.generatedSlots||0} slot siap.${detail?' '+detail:''}`,false);
    $('classroomCard').classList.remove('hidden')
  }catch(error){showStatus('Penjanaan tidak berjaya',error.message||'Sila semak sambungan dan cuba lagi.',false)}
  finally{$('generateButton').disabled=false;updateGenerateButtonUi()}
}
function hideClassroomCard(){$('classroomCard').classList.add('hidden');$('classroomResult').classList.add('hidden');$('openSheetLink').removeAttribute('href');$('classroomFileName').textContent='';$('prepareClassroomButton').disabled=false;$('prepareClassroomButton').innerHTML='<span>▣</span> Sediakan Google Sheet'}
async function prepareClassroom(){
  const button=$('prepareClassroomButton');let completed=false;button.disabled=true;button.innerHTML='<span>◌</span> Menyediakan fail…';
  try{
    const result=await postToErph({action:'prepare_classroom',outputScript:outputMode(),week:Number($('week').value),monday:$('monday').value});
    completed=true;$('classroomFileName').textContent=result.fileName||'Google Sheet eRPH';$('openSheetLink').href=result.fileUrl;$('classroomResult').classList.remove('hidden');
    showStatus('Fail Classroom sudah disediakan',`Salinan eRPH ${outputLabel()} telah dibuat. Pentadbir boleh menandatangani tab PENGESAHAN dalam fail yang sama.`,false)
  }catch(error){showStatus('Penyediaan fail tidak berjaya',error.message||'Sila cuba lagi.',false)}
  finally{button.disabled=!completed;button.innerHTML=completed?'<span>✓</span> Salinan sudah disediakan':'<span>▣</span> Sediakan Google Sheet'}
}
async function copySheetLink(){const link=$('openSheetLink').href;if(!link)return;try{await navigator.clipboard.writeText(link);$('copySheetLink').textContent='Pautan disalin';setTimeout(()=>{$('copySheetLink').textContent='Salin pautan'},1800)}catch(error){showStatus('Tidak dapat menyalin pautan','Tekan “Buka Google Sheet”, kemudian salin pautan dari pelayar.',false)}}
function showStatus(title,text,loading){$('statusCard').classList.remove('hidden');$('statusTitle').textContent=title;$('statusText').textContent=text;$('spinner').style.display=loading?'block':'none'}
initialise();updateGenerateButtonUi();if('serviceWorker'in navigator)navigator.serviceWorker.register('sw.js');
