const filesInput = document.getElementById('files');
const dropzone = document.getElementById('dropzone');
const fileList = document.getElementById('fileList');
const runBtn = document.getElementById('runBtn');
const statusEl = document.getElementById('status');
let files = [];

function render(){ fileList.innerHTML = files.map(f=>`<li>${f.name}</li>`).join(''); }

dropzone.onclick = ()=>filesInput.click();
filesInput.onchange = e=>{ files = [...files, ...e.target.files].filter(f=>f.name.endsWith('.xlsx')); render(); };
['dragenter','dragover'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.add('drag')}));
['dragleave','drop'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.remove('drag')}));
dropzone.addEventListener('drop', e=>{ files = [...files, ...e.dataTransfer.files].filter(f=>f.name.endsWith('.xlsx')); render(); });

runBtn.onclick = async ()=>{
  if(!files.length){statusEl.textContent='יש לבחור קבצים'; return;}
  statusEl.textContent='מעבד...';
  const fd = new FormData(); files.forEach(f=>fd.append('files',f));
  const r = await fetch('/process',{method:'POST', body:fd});
  if(!r.ok){ const j = await r.json(); statusEl.textContent = j.error || 'שגיאה'; return; }
  const blob = await r.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const cd = r.headers.get('content-disposition') || '';
  const m = cd.match(/filename="?([^";]+)"?/i);
  a.download = m ? m[1] : 'SOX_Status_Summary.xlsx';
  a.click();
  statusEl.textContent='הושלם';
}
