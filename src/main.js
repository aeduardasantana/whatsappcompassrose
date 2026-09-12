import './style.css';
import * as XLSX from 'xlsx';

const KEY='compass-wa-v1';
const defaultState={contacts:[],campaign:{name:'',sender:'WhatsApp 1',intervalMin:60,intervalMax:60,mode:'balanced',variations:['','','','','']},queue:[],running:false,current:0,results:[],daily:{date:'',count:0},waitUntil:0,resumeFrom:1};
let state=load();
let timer=null;

function load(){try{return {...structuredClone(defaultState),...(JSON.parse(localStorage.getItem(KEY))||{})}}catch{return structuredClone(defaultState)}}
function save(){localStorage.setItem(KEY,JSON.stringify(state))}
function migrateQueue(){
  state.queue=(state.queue||[]).map((q,i)=>({
    queueItemId:q.queueItemId||('q-'+Date.now()+'-'+i),
    position:q.position||i+1,
    attempts:q.attempts||0,
    ...q
  }));
  if(!state.resumeFrom||state.resumeFrom<1)state.resumeFrom=1;
}
function recoverInterruptedQueue(){
  let changed=false;
  (state.queue||[]).forEach(q=>{
    if(q.status==='aberto'){
      q.status='incerto';
      q.uncertainAt=new Date().toISOString();
      changed=true;
    }
  });
  if(changed)save();
}
function firstActionablePosition(){
  const q=(state.queue||[]).find(x=>['pendente','erro','incerto'].includes(x.status));
  return q?.position||1;
}
function todayKey(){return new Date().toLocaleDateString('sv-SE')}
function syncDaily(){const d=todayKey();if(!state.daily||state.daily.date!==d)state.daily={date:d,count:0}}
function remainingToday(){syncDaily();return Math.max(0,40-state.daily.count)}
function openWhatsAppHome(){window.open('https://web.whatsapp.com/','compassWhatsApp')}
function normPhone(v=''){return String(v).replace(/\D/g,'')}
function esc(v=''){return String(v).replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[s]))}
function pick(obj,names){for(const n of names){const k=Object.keys(obj).find(x=>x.trim().toLowerCase()===n);if(k!==undefined)return obj[k]}return ''}
function normalizeContact(row,i){return {id:i+1,name:pick(row,['nome','name']),phone:normPhone(pick(row,['telefone','celular','whatsapp','phone','numero','número'])),company:pick(row,['empresa','company']),tags:pick(row,['tags','tag','segmento']),selected:true,...row}}

const app=document.querySelector('#app');
app.innerHTML=`<div class="shell"><aside class="sidebar"><div class="brand">Compass WhatsApp</div><nav class="nav">
<button data-view="dashboard" class="active">Dashboard</button><button data-view="whatsapp">WhatsApp</button><button data-view="contacts">Contatos</button><button data-view="campaign">Campanha</button><button data-view="dispatch">Disparo</button><button data-view="report">Relatório</button></nav></aside><main class="main"><div id="view"></div></main></div>`;

document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>show(b.dataset.view));
function navActive(v){document.querySelectorAll('[data-view]').forEach(b=>b.classList.toggle('active',b.dataset.view===v))}

function show(v='dashboard'){navActive(v);({dashboard,whatsapp,contacts,campaign,dispatch,report}[v]||dashboard)()}
function header(title,sub=''){return `<div class="top"><div><h1>${title}</h1>${sub?`<div class="muted">${sub}</div>`:''}</div><span class="badge">Dados locais · sem banco</span></div>`}

function dashboard(){syncDaily();const valid=state.contacts.filter(c=>c.phone).length;const done=state.results.length;view.innerHTML=header('Dashboard','Operação local com Excel e WhatsApp Web')+`<div class="grid">
<div class="card metric"><span>Contatos</span><strong>${state.contacts.length}</strong><small class="muted">${valid} com telefone válido</small></div>
<div class="card metric"><span>Selecionados</span><strong>${state.contacts.filter(c=>c.selected).length}</strong><small class="muted">para a próxima campanha</small></div>
<div class="card metric"><span>Enviados hoje</span><strong>${state.daily.count}/40</strong><small class="muted">${remainingToday()} disponíveis hoje</small></div>
<div class="card metric"><span>Variações</span><strong>${state.campaign.variations.filter(Boolean).length}</strong><small class="muted">mensagens configuradas</small></div></div>
<div class="card section"><h3>Fluxo</h3><p>Importe o Excel → configure a campanha → abra o WhatsApp Web → monte a fila → abra a conversa → envie no WhatsApp → confirme no Compass → aguarde 60 segundos → próximo contato.</p></div>`}

function whatsapp(){syncDaily();view.innerHTML=header('WhatsApp','A conexão é feita diretamente no WhatsApp Web')+`<div class="card"><h3>Conta ativa</h3><p>Informe na campanha qual conta está usando. Para conectar ou trocar de usuário, faça isso diretamente no WhatsApp Web.</p><div class="row"><button class="btn primary" id="openWa">Abrir WhatsApp Web</button></div><p class="muted" style="margin-top:14px">Se ainda não houver sessão, o próprio WhatsApp exibirá o QR Code. Para trocar de conta, use os recursos de sessão/aparelhos conectados do WhatsApp Web.</p></div><div class="card section"><h3>Regras desta versão</h3><p><strong>1 conta por vez · máximo 40 envios/dia · intervalo mínimo 60 segundos.</strong></p><p class="muted">O Compass não lê sua senha, QR Code ou sessão. Ele apenas abre as conversas com a mensagem preparada.</p></div>`;openWa.onclick=openWhatsAppHome}
function contacts(){view.innerHTML=header('Contatos','Importe .xlsx, .xls ou .csv; os dados permanecem neste navegador')+`<div class="card"><div class="row">
<div class="field"><label>Planilha</label><input id="file" type="file" accept=".xlsx,.xls,.csv"/></div><div class="field"><label>Pesquisar</label><input id="search" placeholder="Nome, telefone, empresa ou tag"/></div></div>
<div class="row" style="margin-top:12px"><button class="btn secondary" id="all">Selecionar todos</button><button class="btn secondary" id="none">Desmarcar todos</button><button class="btn danger" id="clear">Limpar contatos</button></div></div><div class="card section table-wrap" id="ctable"></div>`;
file.onchange=importExcel;search.oninput=renderContacts;all.onclick=()=>{state.contacts.forEach(c=>c.selected=true);save();renderContacts()};none.onclick=()=>{state.contacts.forEach(c=>c.selected=false);save();renderContacts()};clear.onclick=()=>{if(confirm('Remover todos os contatos deste navegador?')){state.contacts=[];save();renderContacts()}};renderContacts()}
function renderContacts(){const q=(document.querySelector('#search')?.value||'').toLowerCase();const rows=state.contacts.filter(c=>JSON.stringify(c).toLowerCase().includes(q));ctable.innerHTML=rows.length?`<table class="table"><thead><tr><th></th><th>Nome</th><th>Telefone</th><th>Empresa</th><th>Tags</th><th>Validação</th></tr></thead><tbody>${rows.map(c=>`<tr><td><input type="checkbox" data-id="${c.id}" ${c.selected?'checked':''}></td><td>${esc(c.name)}</td><td>${esc(c.phone)}</td><td>${esc(c.company)}</td><td>${esc(c.tags)}</td><td class="${c.phone.length>=10?'ok':'err'}">${c.phone.length>=10?'OK':'Revisar'}</td></tr>`).join('')}</tbody></table>`:`<div class="empty">Nenhum contato importado.</div>`;ctable.querySelectorAll('input[type=checkbox]').forEach(x=>x.onchange=()=>{const c=state.contacts.find(c=>c.id==x.dataset.id);c.selected=x.checked;save()})}
async function importExcel(e){const f=e.target.files[0];if(!f)return;const data=await f.arrayBuffer();const wb=XLSX.read(data,{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{defval:''});state.contacts=rows.map(normalizeContact).filter(c=>Object.values(c).some(Boolean));save();renderContacts()}

function campaign(){const c=state.campaign;c.intervalMin=60;c.intervalMax=60;view.innerHTML=header('Campanha','Configure o conteúdo antes de montar a fila')+`<div class="card"><div class="row"><div class="field"><label>Nome da campanha</label><input id="cname" value="${esc(c.name)}" placeholder="Ex.: Prospecção NR-1 Setembro"></div><div class="field"><label>Conta WhatsApp em uso</label><input id="sender" value="${esc(c.sender)}" placeholder="Ex.: GEB Empresarial"></div></div><div class="row" style="margin-top:12px"><div class="field"><label>Intervalo</label><input value="60 segundos" disabled></div><div class="field"><label>Limite diário</label><input value="40 envios" disabled></div><div class="field"><label>Rodízio das mensagens</label><select id="mode"><option value="balanced" ${c.mode==='balanced'?'selected':''}>Equilibrado</option><option value="sequential" ${c.mode==='sequential'?'selected':''}>Sequencial</option><option value="random" ${c.mode==='random'?'selected':''}>Aleatório</option></select></div></div><div class="row" style="margin-top:12px"><button class="btn secondary" id="openWa">Abrir / trocar WhatsApp Web</button></div></div><div class="card section"><h3>Variações de mensagem</h3><p class="muted">Use {{nome}}, {{empresa}} e {{telefone}} como variáveis.</p><div class="variations">${c.variations.map((m,i)=>`<div class="field"><label>Variação ${i+1}</label><textarea data-var="${i}" placeholder="Mensagem ${i+1}">${esc(m)}</textarea></div>`).join('')}</div><div style="margin-top:14px"><button id="saveCampaign" class="btn primary">Salvar campanha</button></div></div>`;
openWa.onclick=openWhatsAppHome;saveCampaign.onclick=()=>{state.campaign.name=cname.value.trim();state.campaign.sender=sender.value.trim();state.campaign.intervalMin=60;state.campaign.intervalMax=60;state.campaign.mode=mode.value;document.querySelectorAll('[data-var]').forEach(t=>state.campaign.variations[+t.dataset.var]=t.value);save();alert('Campanha salva neste navegador.')}}
function buildQueue(){syncDaily();const vars=state.campaign.variations.filter(x=>x.trim());if(!vars.length)return alert('Cadastre ao menos uma variação de mensagem.');const selected=state.contacts.filter(c=>c.selected&&c.phone.length>=10);if(!selected.length)return alert('Nenhum contato válido selecionado.');const allowed=remainingToday();if(!allowed)return alert('O limite de 40 envios de hoje já foi atingido.');const picked=selected.slice(0,allowed);state.queue=picked.map((c,i)=>{let idx=0;if(state.campaign.mode==='random')idx=Math.floor(Math.random()*vars.length);else idx=i%vars.length;return {queueItemId:'q-'+Date.now()+'-'+i,position:i+1,contactId:c.id,variation:idx,status:'pendente',attempts:0}});state.results=[];state.running=false;state.waitUntil=0;state.resumeFrom=1;save();if(selected.length>allowed)alert('Foram selecionados '+selected.length+' contatos, mas apenas '+allowed+' entraram na fila por causa do limite diário.');dispatch()}

function dispatch(){
  syncDaily();
  migrateQueue();
  const total=state.queue.length;
  const sent=state.queue.filter(q=>q.status==='enviado').length;
  const pending=state.queue.filter(q=>q.status==='pendente').length;
  const skipped=state.queue.filter(q=>q.status==='pulado').length;
  const errors=state.queue.filter(q=>q.status==='erro').length;
  const uncertain=state.queue.filter(q=>q.status==='incerto').length;
  const pct=total?Math.round(sent/total*100):0;
  const waiting=Math.max(0,Math.ceil(((state.waitUntil||0)-Date.now())/1000));
  const next=state.queue.find(q=>q.position>=state.resumeFrom&&q.status==='pendente');
  const canOpen=next&&waiting===0&&remainingToday()>0;
  const nextPos=next?.position||'-';

  view.innerHTML=header('Disparo assistido','Checkpoint automático · retomada por posição')+`<div class="card">
    <div class="statusbar"><strong>${sent}/${total}</strong><div class="progress"><span style="width:${pct}%"></span></div><span>${pct}%</span></div>
    <div class="row" style="margin-top:14px">
      <span><strong>Enviados:</strong> ${sent}</span>
      <span><strong>Pendentes:</strong> ${pending}</span>
      <span><strong>Erros:</strong> ${errors}</span>
      <span><strong>Incertos:</strong> ${uncertain}</span>
      <span><strong>Pulados:</strong> ${skipped}</span>
    </div>
    <p style="margin-top:14px"><strong>Hoje:</strong> ${state.daily.count}/40 · <strong>Próxima posição:</strong> ${nextPos}</p>
    ${waiting>0?`<p class="warn"><strong>Aguarde ${waiting}s</strong> para liberar o próximo contato.</p>`:''}
    <div class="row" style="margin-top:16px">
      <button class="btn secondary" id="openWa">Abrir WhatsApp Web</button>
      <button class="btn secondary" id="build">Montar/reiniciar fila</button>
      <button class="btn secondary" id="resumeFirst">Continuar do primeiro pendente</button>
      <div class="field" style="max-width:190px;min-width:150px">
        <label>Ir para posição</label>
        <input id="goPosition" type="number" min="1" max="${Math.max(total,1)}" value="${state.resumeFrom||1}">
      </div>
      <button class="btn secondary" id="applyPosition">Aplicar posição</button>
      <button class="btn primary" id="openNext" ${canOpen?'':'disabled'}>Abrir próximo contato</button>
      <button class="btn danger" id="cancel">Cancelar fila</button>
    </div>
  </div>
  <div class="card section"><h3>Campanha</h3><p><strong>${esc(state.campaign.name||'Sem nome')}</strong> · conta em uso: <strong>${esc(state.campaign.sender)}</strong></p><p class="muted">Cada alteração de status é salva imediatamente neste navegador. Itens enviados não são reenviados automaticamente.</p></div>
  <div class="card section table-wrap" id="qtable"></div>`;

  openWa.onclick=openWhatsAppHome;
  build.onclick=buildQueue;
  resumeFirst.onclick=()=>{state.resumeFrom=firstActionablePosition();save();dispatch()};
  applyPosition.onclick=()=>{
    const pos=Math.max(1,Math.min(total||1,+goPosition.value||1));
    state.resumeFrom=pos;
    save();
    dispatch();
  };
  openNext.onclick=openNextContact;
  cancel.onclick=()=>{if(confirm('Cancelar e apagar a fila atual?')){state.queue=[];state.results=[];state.waitUntil=0;state.resumeFrom=1;save();dispatch()}};
  renderQueue();
  if(waiting>0)setTimeout(()=>{if(document.querySelector('#qtable'))dispatch()},1000);
}
function renderQueue(){
  if(!qtable)return;
  const rows=state.queue.slice(0,200).map((q,i)=>{
    const ct=state.contacts.find(x=>x.id===q.contactId)||{};
    let action='';
    if(q.status==='aberto') action=`<button class="btn primary" data-confirm="${i}">Confirmar enviado</button> <button class="btn secondary" data-skip="${i}">Pular</button>`;
    if(q.status==='incerto') action=`<button class="btn primary" data-mark-sent="${i}">Marcar como enviado</button> <button class="btn secondary" data-retry="${i}">Enviar novamente</button> <button class="btn secondary" data-skip="${i}">Pular</button>`;
    if(q.status==='erro') action=`<button class="btn secondary" data-retry="${i}">Tentar novamente</button> <button class="btn secondary" data-skip="${i}">Pular</button>`;
    return `<tr><td>${q.position}</td><td>${esc(ct.name)}</td><td>${esc(ct.phone)}</td><td>V${q.variation+1}</td><td>${esc(q.status)}</td><td>${q.attempts||0}</td><td>${action}</td></tr>`;
  });
  qtable.innerHTML=rows.length?`<table class="table"><thead><tr><th>Posição</th><th>Nome</th><th>Telefone</th><th>Variação</th><th>Status</th><th>Tentativas</th><th>Ação</th></tr></thead><tbody>${rows.join('')}</tbody></table>`:`<div class="empty">Monte a fila para iniciar.</div>`;
  qtable.querySelectorAll('[data-confirm]').forEach(b=>b.onclick=()=>confirmSent(+b.dataset.confirm));
  qtable.querySelectorAll('[data-mark-sent]').forEach(b=>b.onclick=()=>confirmUncertainAsSent(+b.dataset.markSent));
  qtable.querySelectorAll('[data-retry]').forEach(b=>b.onclick=()=>retryItem(+b.dataset.retry));
  qtable.querySelectorAll('[data-skip]').forEach(b=>b.onclick=()=>skipContact(+b.dataset.skip));
}
function renderTemplate(t,c){return t.replace(/{{\s*nome\s*}}/gi,c.name||'').replace(/{{\s*empresa\s*}}/gi,c.company||'').replace(/{{\s*telefone\s*}}/gi,c.phone||'')}

function openNextContact(){
  syncDaily();
  migrateQueue();
  if(remainingToday()<=0)return alert('Limite diário de 40 envios atingido.');
  if(Date.now()<(state.waitUntil||0))return dispatch();
  const idx=state.queue.findIndex(q=>q.position>=state.resumeFrom&&q.status==='pendente');
  if(idx<0)return alert('Não há contatos pendentes a partir da posição selecionada.');
  const q=state.queue[idx];
  const ct=state.contacts.find(x=>x.id===q.contactId);
  const vars=state.campaign.variations.filter(Boolean);
  const msg=renderTemplate(vars[q.variation%vars.length],ct);
  const url='https://web.whatsapp.com/send?phone='+encodeURIComponent(ct.phone)+'&text='+encodeURIComponent(msg);
  q.status='aberto';
  q.openedAt=new Date().toISOString();
  q.attempts=(q.attempts||0)+1;
  state.resumeFrom=q.position;
  save();
  window.open(url,'compassWhatsApp');
  dispatch();
}
function confirmSent(i){
  syncDaily();
  const q=state.queue[i];
  if(!q||q.status!=='aberto')return;
  if(remainingToday()<=0)return alert('Limite diário atingido.');
  finalizeSent(q);
  dispatch();
}

function confirmUncertainAsSent(i){
  syncDaily();
  const q=state.queue[i];
  if(!q||q.status!=='incerto')return;
  if(remainingToday()<=0)return alert('Limite diário atingido.');
  finalizeSent(q);
  dispatch();
}

function finalizeSent(q){
  const ct=state.contacts.find(x=>x.id===q.contactId);
  const vars=state.campaign.variations.filter(Boolean);
  const msg=renderTemplate(vars[q.variation%vars.length],ct);
  q.status='enviado';
  q.sentAt=new Date().toISOString();
  delete q.uncertainAt;
  state.daily.count++;
  state.daily.date=todayKey();
  state.waitUntil=Date.now()+60000;
  state.resumeFrom=q.position+1;
  const existing=state.results.find(r=>r.queueItemId===q.queueItemId);
  const row={queueItemId:q.queueItemId,position:q.position,campaign:state.campaign.name,sender:state.campaign.sender,name:ct.name,phone:ct.phone,company:ct.company,tags:ct.tags,variation:q.variation+1,message:msg,sentAt:q.sentAt,status:'enviado_confirmado',attempts:q.attempts||0};
  if(existing)Object.assign(existing,row);else state.results.push(row);
  save();
}

function retryItem(i){
  const q=state.queue[i];
  if(!q)return;
  q.status='pendente';
  delete q.uncertainAt;
  delete q.error;
  state.resumeFrom=q.position;
  save();
  dispatch();
}
function skipContact(i){const q=state.queue[i];if(!q)return;q.status='pulado';if(state.resumeFrom<=q.position)state.resumeFrom=q.position+1;save();dispatch()}
function report(){view.innerHTML=header('Relatório','Exporte o resultado e use o Excel como histórico')+`<div class="card"><div class="row"><button class="btn primary" id="exportResults">Exportar relatório .xlsx</button><button class="btn secondary" id="exportContacts">Exportar contatos atuais .xlsx</button><button class="btn danger" id="clearResults">Limpar resultado local</button></div></div><div class="card section table-wrap" id="rtable"></div>`;
exportResults.onclick=()=>exportXlsx(state.results,'relatorio-compass-whatsapp.xlsx','Relatório');exportContacts.onclick=()=>exportXlsx(state.contacts,'contatos-compass-whatsapp.xlsx','Contatos');clearResults.onclick=()=>{state.results=[];state.queue=[];state.current=0;save();report()};rtable.innerHTML=state.results.length?`<table class="table"><thead><tr><th>Posição</th><th>Nome</th><th>Telefone</th><th>Empresa</th><th>Variação</th><th>Status</th><th>Tentativas</th><th>Enviado em</th></tr></thead><tbody>${state.results.map(r=>`<tr><td>${r.position||''}</td><td>${esc(r.name)}</td><td>${esc(r.phone)}</td><td>${esc(r.company)}</td><td>${r.variation}</td><td>${esc(r.status)}</td><td>${r.attempts||0}</td><td>${esc(r.sentAt||'')}</td></tr>`).join('')}</tbody></table>`:`<div class="empty">Ainda não há resultados nesta campanha.</div>`}
function exportXlsx(rows,name,sheet){if(!rows.length)return alert('Não há dados para exportar.');const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);XLSX.writeFile(wb,name)}

if('serviceWorker' in navigator){navigator.serviceWorker.register('./sw.js').catch(()=>{})}
migrateQueue();recoverInterruptedQueue();syncDaily();save();show('dashboard');
