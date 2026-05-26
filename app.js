// ===== Utilidades de data =====
function toYMD(d){const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`}
function fromYMD(s){return new Date(`${s}T00:00:00`)}
function addDays(d,n){const nd=new Date(d);nd.setDate(nd.getDate()+n);return nd}
function diffDays(a,b){const A=new Date(a.getFullYear(),a.getMonth(),a.getDate());const B=new Date(b.getFullYear(),b.getMonth(),b.getDate());return Math.round((A-B)/(1000*60*60*24))}
function clampDate(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate())}

// ===== Segurança de UI: escape de strings vindas de dados (evita XSS via importação) =====
function escHTML(v){
  return String(v ?? '').replace(/[&<>"']/g, (ch)=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[ch]));
}
function escAttr(v){ return escHTML(v); }

function uuid(){if (crypto && crypto.randomUUID) return crypto.randomUUID(); const s=()=>Math.floor((1+Math.random())*0x10000).toString(16).substring(1); return `${s()}${s()}-${s()}-${s()}-${s()}-${s()}${s()}${s()}`}

function encodeInlineText(v){
  return String(v ?? '').replace(/\r?\n/g, ' \\n ').trim();
}

function decodeInlineText(v){
  return String(v ?? '').replace(/\s*\\n\s*/g, '\n').trim();
}

function normalizeActivityCommentEntry(entry){
  if(!entry || typeof entry !== 'object') return null;
  const text = String(entry.text || entry.comentario || '').trim();
  if(!text) return null;
  return {
    id: String(entry.id || uuid()),
    ts: String(entry.ts || entry.timestamp || new Date().toISOString()),
    user: String(entry.user || entry.autor || '').trim(),
    text
  };
}

function parseActivityComments(rawJson, legacyText, fallbackTs){
  let list = [];
  const raw = String(rawJson || '').trim();
  if(raw){
    try{
      const parsed = JSON.parse(decodeInlineText(raw));
      if(Array.isArray(parsed)) list = parsed.map(normalizeActivityCommentEntry).filter(Boolean);
    }catch(_){ }
  }
  if(!list.length){
    const legacy = String(legacyText || '').trim();
    if(legacy){
      list = [{
        id: uuid(),
        ts: fallbackTs ? new Date(Number(fallbackTs) || Date.now()).toISOString() : new Date().toISOString(),
        user: '',
        text: legacy
      }];
    }
  }
  return list;
}

function serializeActivityComments(list){
  try{
    return encodeInlineText(JSON.stringify((list || []).map(normalizeActivityCommentEntry).filter(Boolean)));
  }catch(_){
    return '';
  }
}

function formatActivityCommentsForDisplay(list){
  return (list || []).map(c => {
    const when = c?.ts ? formatDateTimeBR(c.ts) : '';
    const who = String(c?.user || '').trim();
    const header = [when, who].filter(Boolean).join(' • ');
    return `${header ? header + '\n' : ''}${String(c?.text || '').trim()}`.trim();
  }).filter(Boolean).join('\n\n');
}

function renderActivityCommentsHtml(list){
  if(!list || !list.length) return '<div class="muted small">Sem comentários publicados.</div>';
  return list.map(c => {
    const when = c?.ts ? formatDateTimeBR(c.ts) : '';
    const who = String(c?.user || '').trim();
    const meta = [when, who].filter(Boolean).join(' • ');
    return `<div class="comment-item" style="padding:8px 10px; border:1px solid var(--border-color, #d9dee8); border-radius:10px; margin-bottom:8px; background:rgba(255,255,255,.55);">\n      ${meta ? `<div class="muted small" style="margin-bottom:4px;"><strong>${escHTML(meta)}</strong></div>` : ''}\n      <div class="small" style="white-space:pre-wrap;">${escHTML(c.text || '')}</div>\n    </div>`;
  }).join('');
}
function validateCommentLength(text){
  const normalized = String(text || '');
  if(normalized.length > MAX_COMMENT_LENGTH){
    alert(`O comentário pode ter no máximo ${MAX_COMMENT_LENGTH} caracteres.`);
    return false;
  }
  return true;
}
function formatDateTimeBR(v){
  try{
    const d = v instanceof Date ? v : new Date(v);
    if(!d || isNaN(d.getTime())) return '';
    return d.toLocaleString('pt-BR', {
      day:'2-digit', month:'2-digit', year:'numeric',
      hour:'2-digit', minute:'2-digit'
    });
  }catch(_){
    return '';
  }
}


function normalizeCommentRow(row){
  if(!row || typeof row !== 'object') return null;
  const activityId = String(row.activityId || row.atividadeId || '').trim();
  const text = String(row.texto ?? row.text ?? row.comentario ?? '').trim();
  if(!activityId || !text) return null;
  const tsSource = row.ts || row.timestamp || row.dataHora || row.createdAt || row.updatedAt || Date.now();
  const d = tsSource instanceof Date ? tsSource : new Date(Number(tsSource) || tsSource);
  const safeDate = (!d || isNaN(d.getTime())) ? new Date() : d;
  const createdAt = Number(row.createdAt || 0) || safeDate.getTime();
  const updatedAt = Number(row.updatedAt || row.atualizadoAt || 0) || createdAt;
  return {
    commentId: String(row.commentId || row.id || uuid()),
    activityId,
    texto: text.slice(0, MAX_COMMENT_LENGTH),
    usuario: String(row.usuario || row.user || row.autor || '').trim(),
    ts: safeDate.toISOString(),
    createdAt,
    updatedAt,
    deletedAt: row.deletedAt ? (Number(row.deletedAt) || null) : null,
    version: Math.max(1, Number(row.version || row.versao || 1) || 1)
  };
}
function hydrateLoadedComments(loadedActivities, loadedComments){
  const out = [];
  const seen = new Set();
  const activityIdsWithStructuredComments = new Set();

  (Array.isArray(loadedComments) ? loadedComments : []).forEach(row=>{
    const n = normalizeCommentRow(row);
    if(!n || seen.has(n.commentId)) return;
    seen.add(n.commentId);
    activityIdsWithStructuredComments.add(String(n.activityId || ''));
    out.push(n);
  });

  (Array.isArray(loadedActivities) ? loadedActivities : []).forEach(a=>{
    if(!a || !a.id) return;
    const activityId = String(a.id || '');
    if(!activityId || activityIdsWithStructuredComments.has(activityId)) return;

    const legacy = parseActivityComments(a.comentariosJson, a.comentarios, a.updatedAt);
    legacy.forEach(c=>{
      const n = normalizeCommentRow({
        commentId:c.id,
        activityId,
        texto:c.text,
        usuario:c.user,
        ts:c.ts,
        createdAt:Date.parse(c.ts)||a.updatedAt||Date.now(),
        updatedAt:Date.parse(c.ts)||a.updatedAt||Date.now()
      });
      if(!n || seen.has(n.commentId)) return;
      seen.add(n.commentId);
      out.push(n);
    });
  });

  out.sort((a,b)=>(Number(a.createdAt||0)-Number(b.createdAt||0)) || String(a.commentId).localeCompare(String(b.commentId)));
  return out;
}
function indexCommentsByActivity(list){
  const map = new Map();
  (list||[]).forEach(c=>{
    if(!c || c.deletedAt || !c.activityId) return;
    const key = String(c.activityId);
    if(!map.has(key)) map.set(key, []);
    map.get(key).push(c);
  });
  for(const arr of map.values()) arr.sort((a,b)=>(Number(b.updatedAt||b.createdAt||0)-Number(a.updatedAt||a.createdAt||0)) || String(b.commentId).localeCompare(String(a.commentId)));
  return map;
}
let commentsByActivityId = new Map();
function rebuildCommentsIndex(){ commentsByActivityId = indexCommentsByActivity(comments || []); return commentsByActivityId; }
function getAllCommentsForActivity(activityId){ return (commentsByActivityId.get(String(activityId || '')) || []).slice(); }
function getComments(activityId, opts={}){ const offset=Math.max(0, Number(opts.offset||0)||0); const limit=Math.max(1, Number(opts.limit||10)||10); const list=getAllCommentsForActivity(activityId); return { total:list.length, items:list.slice(offset, offset+limit) }; }
function getLastCommentPreview(activityId, maxLen=140){ const list=getAllCommentsForActivity(activityId); if(!list.length) return ''; const txt=String(list[0].texto||'').replace(/\s+/g,' ').trim(); return txt.length>maxLen ? txt.slice(0,maxLen-1)+'…' : txt; }
function mergeComentarios(baseList, incomingList){ const merged = new Map(); [...(baseList||[]), ...(incomingList||[])].forEach(item=>{ const n = normalizeCommentRow(item); if(!n) return; const existing=merged.get(n.commentId); if(!existing){ merged.set(n.commentId,n); return; } const er=Number(existing.updatedAt||existing.createdAt||0); const nr=Number(n.updatedAt||n.createdAt||0); if(nr>=er) merged.set(n.commentId,n); }); return Array.from(merged.values()).sort((a,b)=>(Number(a.createdAt||0)-Number(b.createdAt||0)) || String(a.commentId).localeCompare(String(b.commentId))); }
function syncActivityCommentFields(activity){ if(!activity) return activity; const list=getAllCommentsForActivity(activity.id).map(c=>({ id:c.commentId, ts:c.ts, user:c.usuario, text:c.texto })); activity.comentariosLista=list; activity.comentariosJson=serializeActivityComments(list); activity.comentarios=formatActivityCommentsForDisplay(list); return activity; }
function syncAllActivityCommentFields(){ (activities||[]).forEach(syncActivityCommentFields); }
const COMMENT_DRAFT_PREFIX='comentarioDraft:';
function getCommentDraftKey(activityId, usuario){ return `${COMMENT_DRAFT_PREFIX}${String(activityId||'')}:${String(usuario||'anonimo')}`; }
function salvarComentarioDraft(activityId, usuario, payload){ return saveLS(getCommentDraftKey(activityId, usuario), payload || {}); }
function carregarComentarioDraft(activityId, usuario){ return loadLS(getCommentDraftKey(activityId, usuario), null); }
function clearCommentDraft(activityId, usuario){ try{ localStorage.removeItem(getCommentDraftKey(activityId, usuario)); }catch(_){ } }
function appendComment(activityId, texto, usuario){ const text=String(texto||'').trim(); if(!text) return {ok:true, skipped:true}; if(!validateCommentLength(text)) return {ok:false, reason:'length'}; const nowTs=Date.now(); const row=normalizeCommentRow({ commentId:uuid(), activityId, texto:text, usuario:usuario||'', ts:new Date(nowTs).toISOString(), createdAt:nowTs, updatedAt:nowTs, version:1 }); comments=mergeComentarios(comments,[row]); rebuildCommentsIndex(); syncAllActivityCommentFields(); saveLS(LS.comments, comments); clearCommentDraft(activityId, usuario || ''); return {ok:true, comment:row}; }

function getActivityComparablePayload(activity){
  if(!activity) return null;
  return {
    id:String(activity.id||''),
    codigoAtividade:String(activity.codigoAtividade||''),
    titulo:String(activity.titulo||''),
    resourceId:String(activity.resourceId||''),
    linkedOriginId:String(activity.linkedOriginId||''),
    linkedOriginCode:String(activity.linkedOriginCode||''),
    inicio:String(activity.inicio||''),
    fim:String(activity.fim||''),
    status:String(activity.status||''),
    alocacao:Number(activity.alocacao||0),
    tags:Array.isArray(activity.tags) ? activity.tags.map(t=>String(t||'')).sort() : [],
    deletedAt:activity.deletedAt ? Number(activity.deletedAt||0) : null
  };
}
function isSameActivityComparable(a,b){
  try { return JSON.stringify(getActivityComparablePayload(a)) === JSON.stringify(getActivityComparablePayload(b)); } catch(_) { return false; }
}
async function loadPersistedBDState(){
  if(!bdHandle) return null;
  const file = await bdHandle.getFile();
  const text = await file.text();
  const parsed = (bdFileExt === 'csv') ? parseCSVBDUnico(text) : parseHTMLBDTables(text);
  parsed.recursos = (parsed.recursos || []).map(coerceResource);
  parsed.atividades = hydrateLoadedActivities(parsed.atividades || []);
  parsed.comments = hydrateLoadedComments(parsed.atividades || [], parsed.comments || []);
  return parsed;
}
async function reconcileCommentWrite(activityBase, pendingActivity, draftText, usuario){
  if(!bdHandle || !activityBase || !draftText) return { ok:true, resolvedBase:activityBase, resolvedPending:pendingActivity, message:'' };
  try{
    const persisted = await loadPersistedBDState();
    const persistedActivity = (persisted?.atividades || []).find(a => a && a.id === activityBase.id);
    if(!persistedActivity) return { ok:true, resolvedBase:activityBase, resolvedPending:pendingActivity, message:'' };
    const sameVersion = Number(persistedActivity.version || 0) === Number(activityBase.version || 0)
      && Number(persistedActivity.updatedAt || 0) === Number(activityBase.updatedAt || 0);
    if(sameVersion) return { ok:true, resolvedBase:activityBase, resolvedPending:pendingActivity, message:'' };
    const onlyCommentAppend = isSameActivityComparable(activityBase, pendingActivity);
    if(onlyCommentAppend){
      comments = mergeComentarios(persisted.comments || [], comments || []);
      rebuildCommentsIndex();
      const mergedActivity = syncActivityCommentFields({ ...persistedActivity });
      const idx = activities.findIndex(a => a && a.id === mergedActivity.id);
      if(idx >= 0) activities[idx] = { ...activities[idx], ...mergedActivity };
      else activities.push(mergedActivity);
      syncAllActivityCommentFields();
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      return { ok:true, resolvedBase:mergedActivity, resolvedPending:{ ...mergedActivity }, message:'Comentário salvo após reconciliação.' };
    }
    salvarComentarioDraft(activityBase.id, usuario || '', {
      texto:String(draftText||''),
      ts:Date.now(),
      baseActivityVersion:Number(activityBase.version || 0),
      baseActivityUpdatedAt:Number(activityBase.updatedAt || 0)
    });
    const reopen = confirm('A atividade foi alterada por outra sessão e existe conflito com os dados em edição. O rascunho do comentário foi preservado.\n\nDeseja reabrir a atividade com os dados mais atuais e manter o rascunho?');
    if(reopen){
      const hydratedPersisted = syncActivityCommentFields({ ...persistedActivity });
      const idx = activities.findIndex(a => a && a.id === hydratedPersisted.id);
      if(idx >= 0) activities[idx] = { ...activities[idx], ...hydratedPersisted };
      else activities.push(hydratedPersisted);
      syncAllActivityCommentFields();
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      renderAll();
      setTimeout(()=>openActivityModalById(hydratedPersisted.id), 0);
    } else {
      const discard = confirm('Deseja descartar o rascunho preservado para este comentário?');
      if(discard) clearCommentDraft(activityBase.id, usuario || '');
    }
    return { ok:false, unresolved:true };
  }catch(err){
    console.warn('Falha ao reconciliar comentário com o BD persistido', err);
    salvarComentarioDraft(activityBase.id, usuario || '', {
      texto:String(draftText||''),
      ts:Date.now(),
      baseActivityVersion:Number(activityBase.version || 0),
      baseActivityUpdatedAt:Number(activityBase.updatedAt || 0)
    });
    alert('Não foi possível reconciliar o comentário com o BD apontado. O rascunho foi preservado para nova tentativa.');
    return { ok:false, unresolved:true, error:err };
  }
}
function loadCommentDraftIntoForm(activityId, usuario){
  if(!atividadeComentariosInput) return;
  const draft = activityId ? carregarComentarioDraft(activityId, usuario || '') : null;
  atividadeComentariosInput.value = (draft && draft.texto) ? String(draft.texto) : '';
  updateActivityCommentCounter();
}
function openActivityModalById(activityId){
  const activity = (activities || []).find(a => a && a.id === activityId);
  if(!activity || !formAtividade || !dlgAtividade) return false;
  dlgAtividadeTitulo.textContent='Editar Atividade';
  fillRecursoOptions();
  formAtividade.elements['id'].value=activity.id;
  if(formAtividade.elements['codigoAtividade']) formAtividade.elements['codigoAtividade'].value=activity.codigoAtividade||'';
  formAtividade.elements['titulo'].value=activity.titulo||'';
  const recObj = resources.find(x=>x.id===activity.resourceId);
  formAtividade.elements['resourceName'].value = recObj ? recObj.nome : '';
  formAtividade.elements['inicio'].value=activity.inicio||'';
  formAtividade.elements['fim'].value=activity.fim||'';
  formAtividade.elements['status'].value=activity.status||'Planejada';
  formAtividade.elements['alocacao'].value=activity.alocacao||100;
  loadCommentDraftIntoForm(activity.id, currentUser || '');
  refreshActivityCommentsPanel(activity);
  formAtividade.elements['tags'].value = (activity.tags || []).join(', ');
  if(formAtividade.elements['linkedOriginId']) formAtividade.elements['linkedOriginId'].value = activity.linkedOriginId || '';
  if(formAtividade.elements['linkedOriginSearch']) formAtividade.elements['linkedOriginSearch'].value = '';
  fillActivityLinkOptions(activity.linkedOriginId || '', '', activity.id);
  dlgAtividade.showModal();
  return true;
}

function getActivityCodeSeed(list){
  return (Array.isArray(list) ? list : []).reduce((max, a) => {
    const raw = String(a?.codigoAtividade || a?.codigo || a?.atividadeCodigo || '').trim();
    const m = raw.match(/ATV-(\d+)/i);
    if(!m) return max;
    const n = Number(m[1]);
    return Number.isFinite(n) ? Math.max(max, n) : max;
  }, 0);
}
function generateNextActivityCode(list){
  const next = getActivityCodeSeed(list) + 1;
  return `ATV-${String(next).padStart(4,'0')}`;
}
function ensureActivityCodes(list){
  const arr = Array.isArray(list) ? list.map(a => ({...a})) : [];
  let changed = false;
  let seed = getActivityCodeSeed(arr);
  const used = new Set();
  arr.forEach(a => {
    let code = String(a?.codigoAtividade || a?.codigo || a?.atividadeCodigo || '').trim();
    if(!code || used.has(code)){
      seed += 1;
      code = `ATV-${String(seed).padStart(4,'0')}`;
      a.codigoAtividade = code;
      changed = true;
    }
    used.add(code);
  });
  return { list: arr, changed };
}

function formatActivityLinkOption(activity){
  const code = String(activity?.codigoAtividade || activity?.id || '').trim();
  const title = String(activity?.titulo || '').trim();
  if(code && title) return `${code} — ${title}`;
  return code || title;
}
const activityLinkOptionMap = new Map();
function fillActivityLinkOptions(selectedId = '', searchValue = '', excludeId = ''){
  const list = document.getElementById('activityLinkList');
  const searchInput = formAtividade?.elements?.['linkedOriginSearch'];
  const hiddenInput = formAtividade?.elements?.['linkedOriginId'];
  if(!list || !searchInput || !hiddenInput) return;
  const currentId = excludeId || formAtividade?.elements?.['id']?.value || '';
  const typedValue = String(searchValue ?? searchInput.value ?? '');
  const query = typedValue.trim().toLowerCase();
  activityLinkOptionMap.clear();
  list.innerHTML = '';
  (activities || []).filter(a=>!a.deletedAt && a.id !== currentId).slice().sort((a,b)=>{
    return String(a?.codigoAtividade || '').localeCompare(String(b?.codigoAtividade || '')) || String(a?.titulo || '').localeCompare(String(b?.titulo || ''), undefined, { sensitivity:'base' });
  }).filter(a=>{
    if(!query) return true;
    return formatActivityLinkOption(a).toLowerCase().includes(query);
  }).slice(0,50).forEach(a=>{
    const opt = document.createElement('option');
    opt.value = formatActivityLinkOption(a);
    list.appendChild(opt);
    activityLinkOptionMap.set(opt.value, a.id);
  });
  hiddenInput.value = selectedId || '';
  if(selectedId && !typedValue.trim()){
    const selected = (activities || []).find(a=>a.id === selectedId);
    if(selected) searchInput.value = formatActivityLinkOption(selected);
  }
}
function resolveLinkedActivityReference(inputValue, excludeId = ''){
  const raw = String(inputValue || '').trim();
  if(!raw) return null;
  if(activityLinkOptionMap.has(raw)){
    const mappedId = activityLinkOptionMap.get(raw);
    const act = (activities || []).find(a=>!a.deletedAt && a.id === mappedId && a.id !== excludeId);
    if(act) return act;
  }
  const key = raw.toLowerCase();
  let matches = (activities || []).filter(a=>!a.deletedAt && a.id !== excludeId).filter(a=>{
    const code = String(a.codigoAtividade || '').toLowerCase();
    const title = String(a.titulo || '').toLowerCase();
    const combo = formatActivityLinkOption(a).toLowerCase();
    return code === key || title === key || combo === key;
  });
  if(matches.length === 1) return matches[0];
  matches = (activities || []).filter(a=>!a.deletedAt && a.id !== excludeId).filter(a=>{
    const code = String(a.codigoAtividade || '').toLowerCase();
    const title = String(a.titulo || '').toLowerCase();
    const combo = formatActivityLinkOption(a).toLowerCase();
    return code.includes(key) || title.includes(key) || combo.includes(key);
  });
  return matches.length === 1 ? matches[0] : null;
}

// ===== Importação em massa =====
// Normaliza nomes para comparação (trim e lower case). Utilizado no merge de recursos e atividades importados.
/**
 * Normaliza um nome para comparação, removendo acentuação, espaços nas extremidades
 * e convertendo para minúsculas. Isso permite que nomes com
 * diacríticos sejam comparados de forma insensível a acentos (ex.:
 * "João" e "Joao" serão tratados como iguais). Se o valor for
 * nulo/undefined, retorna uma string vazia.
 *
 * @param {string} str Texto a normalizar
 * @returns {string} Nome normalizado
 */
function normalizeName(str) {
  if (!str) return '';
  // Converte para string, remove espaços extras, normaliza unicode
  const s = String(str).trim().toLowerCase().normalize('NFD');
  // Substitui alguns caracteres especiais antes de remover diacríticos.
  // Por exemplo: 'ç' -> 'c', 'Ç' -> 'c', 'ñ' -> 'n'.
  let normalized = s
    .replace(/ç/g, 'c')
    .replace(/Ç/g, 'c')
    .replace(/ñ/g, 'n')
    .replace(/Ñ/g, 'n');
  // Remove diacríticos (acentos)
  normalized = normalized.replace(/[\u0300-\u036f]/g, '');
  return normalized;
}

// ===== Utilidade de Tags =====
function normalizeTag(tag) {
  if (!tag) return '';
  const trimmed = tag.trim();
  // Primeira letra maiúscula, resto minúsculo
  return trimmed.charAt(0).toUpperCase() + trimmed.slice(1).toLowerCase();
}

// ===== Domínio =====
const TIPOS=["Interno","Externo"];
const SENIORIDADES=["Jr","Pl","Sr","NA"];
const STATUS=["Planejada","Em Execução","Bloqueada","Concluída","Cancelada"];

// ===== Persistência =====
// Principais chaves para arrays de domínio
const LS={res:"rp_resources_v2",act:"rp_activities_v2",comments:"rp_comments_v1",trail:"rp_trail_v1",user:"rp_user_v1",base:"rp_baselines_v1",baseItems:"rp_baseline_items_v1",showInactiveRes:"rp_show_inactive_resources_v1"};
const LS_ACCESS_SESSION = "rp_access_session_v1";
const ACCESS_VERSION = "1.1";
const ACCESS_PROFILES = ["Administrador","Planejador","Executor","Consulta/Gestor"];
const ACCESS_PERMISSIONS = {
  Administrador: { db:true, manageUsers:true, plan:true, execute:true, export:true, audit:true, delete:true, admin:true },
  Planejador: { db:true, manageUsers:false, plan:true, execute:true, export:true, audit:true, delete:false, admin:false },
  Executor: { db:true, manageUsers:false, plan:false, execute:true, export:true, audit:false, delete:false, admin:false },
  "Consulta/Gestor": { db:true, manageUsers:false, plan:false, execute:false, export:true, audit:true, delete:false, admin:false }
};
let systemUsers = [];
let auditEvents = loadLS("rp_audit_events_v1", []);
let auditPage = 1;
const AUDIT_PAGE_SIZE = 50;
let accessSession = null;
let accessGateReady = false;


// Chaves adicionais para log de eventos e snapshots.  O log de eventos registra cada
// alteração de recurso/atividade (criação, atualização, exclusão) de forma
// append-only.  Snapshots são capturas periódicas do estado completo e
// permitem reconstruir rapidamente o estado sem precisar reprocessar todos
// os eventos.
const LS_LOG="rp_event_log_v1";
const LS_SNAP="rp_snapshot_v1";
const MAX_COMMENT_LENGTH = 2000;
const MAX_BASELINES = 10;
let lastStorageWarning = '';
function isQuotaExceededError(err){
  if(!err) return false;
  return err.name === 'QuotaExceededError' || err.code === 22 || err.code === 1014 || /quota/i.test(String(err.message || ''));
}
function notifyStorageIssue(message){
  if(lastStorageWarning === message) return;
  lastStorageWarning = message;
  try{ alert(message); }catch(_){ }
  setTimeout(()=>{ if(lastStorageWarning === message) lastStorageWarning = ''; }, 1200);
}
function safeSetLocalStorageItem(key, value, contextLabel){
  try{
    localStorage.setItem(key, value);
    return true;
  }catch(err){
    if(isQuotaExceededError(err)){
      console.error(`localStorage cheio ao salvar ${contextLabel || key}`, err);
      notifyStorageIssue('Armazenamento local cheio. Exclua baselines antigos ou reduza o volume de dados antes de salvar novamente.');
      return false;
    }
    throw err;
  }
}
function loadLS(k,fallback){try{const raw=localStorage.getItem(k);return raw?JSON.parse(raw):fallback}catch{return fallback}}
function saveLS(k,v){return safeSetLocalStorageItem(k,JSON.stringify(v),k)}


// ===== Calendário operacional / Feriados =====
const LS_CALENDAR_PREFS = "rp_calendar_prefs_v1";
const LS_OVERTIME = "rp_overtime_v1";
const LS_ADMIN_HASH = "rp_admin_hash_v1";
const DEFAULT_CALENDAR_PREFS = { destacarFeriados:true, destacarFds:true };
let calendarPrefs = Object.assign({}, DEFAULT_CALENDAR_PREFS, loadLS(LS_CALENDAR_PREFS, {}));
let overtimeEntries = normalizeOvertimeList(loadLS(LS_OVERTIME, []));
let adminUnlocked = true;




// ===== Execução operacional por atividade =====
function parseExecArray(value){
  if(Array.isArray(value)) return value.filter(Boolean);
  const raw = String(value || '').trim();
  if(!raw) return [];
  try{
    const decoded = decodeInlineText(raw);
    const parsed = JSON.parse(decoded);
    return Array.isArray(parsed) ? parsed.filter(Boolean) : [];
  }catch(_){ return []; }
}
function normalizeExecSubtask(row){
  if(!row || typeof row !== 'object') return null;
  const title = String(row.titulo || row.title || '').trim();
  if(!title) return null;
  return {
    id: String(row.id || row.subtaskId || uuid()),
    titulo: title,
    tipoSubatividade: getExecSubtaskType(row),
    responsavelId: String(row.responsavelId || row.resourceId || '').trim(),
    horasPlanejadas: Math.max(0, Number(String(row.horasPlanejadas ?? row.plannedHours ?? 0).replace(',', '.')) || 0),
    inicioPrevisto: normalizeDateField(row.inicioPrevisto || row.start || ''),
    fimPrevisto: normalizeDateField(row.fimPrevisto || row.due || ''),
    status: String(row.status || 'Planejada').trim() || 'Planejada',
    createdAt: Number(row.createdAt || 0) || Date.now(),
    createdBy: String(row.createdBy || row.user || '').trim()
  };
}
function limitExecComment(value){
  return String(value || '').trim().slice(0, 360);
}
const EXEC_SUBTASK_TYPES = ['Elaboração','Revisão','Execução','Aprovação','Teste','Documentação','Investigação','Homologação','Implantação','Outros'];
function normalizeExecSubtaskType(value){
  const v = String(value || '').trim().slice(0, 80);
  return v || 'Outros';
}
function getExecSubtaskType(row){
  if(!row || typeof row !== 'object') return 'Outros';
  const raw = String(row.tipoSubatividade || row.subtaskType || row.tipo || '').trim();
  const other = String(row.tipoSubatividadeOutro || row.tipoOutro || row.otherType || '').trim();
  if(raw === 'Outros') return normalizeExecSubtaskType(other || raw);
  return normalizeExecSubtaskType(raw || other || 'Outros');
}
function execStats(values){
  const nums = (values || []).map(v => Number(v) || 0).filter(v => Number.isFinite(v));
  const n = nums.length;
  if(!n) return { quantidade:0, total:0, media:0, mediana:0, desvio:0 };
  const total = nums.reduce((a,b)=>a+b,0);
  const media = total / n;
  const sorted = nums.slice().sort((a,b)=>a-b);
  const mediana = n % 2 ? sorted[(n-1)/2] : (sorted[n/2-1] + sorted[n/2]) / 2;
  const variancia = nums.reduce((acc,x)=>acc + Math.pow(x-media,2),0) / n;
  return { quantidade:n, total, media, mediana, desvio:Math.sqrt(variancia) };
}
function normalizeExecEntry(row){
  if(!row || typeof row !== 'object') return null;
  const hours = Number(String(row.horas ?? row.hours ?? 0).replace(',', '.')) || 0;
  if(hours <= 0) return null;
  return {
    id: String(row.id || row.entryId || uuid()),
    subtaskId: String(row.subtaskId || row.subatividadeId || '').trim(),
    data: normalizeDateField(row.data || row.date || toYMD(new Date())),
    horas: Math.max(0.25, hours),
    comentario: limitExecComment(row.comentario || row.comment || ''),
    createdAt: Number(row.createdAt || 0) || Date.now(),
    createdBy: String(row.createdBy || row.user || '').trim()
  };
}
function normalizeExecIssue(row){
  if(!row || typeof row !== 'object') return null;
  const desc = limitExecComment(row.descricao || row.description || '');
  const impactHours = Math.max(0, Number(String(row.impactoHoras ?? row.impactHours ?? 0).replace(',', '.')) || 0);
  const impactDays = Math.max(0, Number(row.impactoPrazoDias ?? row.impactDays ?? 0) || 0);
  if(!desc && !impactHours && !impactDays) return null;
  return {
    id: String(row.id || row.issueId || uuid()),
    tipo: String(row.tipo || row.type || 'Ocorrência').trim() || 'Ocorrência',
    descricao: desc,
    impactoHoras: impactHours,
    impactoPrazoDias: impactDays,
    data: normalizeDateField(row.data || row.date || toYMD(new Date())),
    createdAt: Number(row.createdAt || 0) || Date.now(),
    createdBy: String(row.createdBy || row.user || '').trim()
  };
}
function normalizeExecData(activity){
  // Prefer in-memory arrays over JSON snapshots.
  // During execution edits, mutators update execSubtasks/execEntries/execIssues first;
  // if we prioritize exec*Json here, an old encoded "[]" snapshot can discard the new item.
  const subtasksSource = Array.isArray(activity?.execSubtasks) ? activity.execSubtasks : (activity?.execSubtasksJson || activity?.execSubtasks);
  const entriesSource = Array.isArray(activity?.execEntries) ? activity.execEntries : (activity?.execEntriesJson || activity?.execEntries);
  const issuesSource = Array.isArray(activity?.execIssues) ? activity.execIssues : (activity?.execIssuesJson || activity?.execIssues);
  const subtasks = parseExecArray(subtasksSource).map(normalizeExecSubtask).filter(Boolean);
  const entries = parseExecArray(entriesSource).map(normalizeExecEntry).filter(Boolean);
  const issues = parseExecArray(issuesSource).map(normalizeExecIssue).filter(Boolean);
  return { subtasks, entries, issues };
}
function syncExecFields(activity){
  if(!activity) return activity;
  const data = normalizeExecData(activity);
  activity.execSubtasks = data.subtasks;
  activity.execEntries = data.entries;
  activity.execIssues = data.issues;
  activity.execSubtasksJson = encodeInlineText(JSON.stringify(data.subtasks));
  activity.execEntriesJson = encodeInlineText(JSON.stringify(data.entries));
  activity.execIssuesJson = encodeInlineText(JSON.stringify(data.issues));
  activity.controleExecucao = calcExecutionMetrics(activity);
  return activity;
}
function estimatePlannedHoursFromActivity(activity){
  try{
    const r = resources.find(x=>x.id === activity.resourceId);
    const daily = getDailyHours(r) * ((Number(activity?.alocacao ?? 100) || 100) / 100);
    const start = fromYMD(activity.inicio), end = fromYMD(activity.fim);
    let days = 0;
    for(let d=new Date(start); d<=end; d=addDays(d,1)){
      const ctx = typeof getDayContext === 'function' ? getDayContext(d, r?.id) : null;
      const dow = d.getDay();
      const nonWork = (ctx && ctx.isHoliday) || dow === 0 || dow === 6;
      if(!nonWork) days++;
    }
    return Math.max(0, days * (daily || 0));
  }catch(_){ return 0; }
}
function addBusinessDaysFromToday(daysToAdd, resourceId){
  let d = new Date();
  let remaining = Math.max(0, Math.ceil(Number(daysToAdd)||0));
  let guard = 0;
  while(remaining > 0 && guard < 730){
    d = addDays(d, 1); guard++;
    const dow = d.getDay();
    const ctx = typeof getDayContext === 'function' ? getDayContext(d, resourceId) : null;
    if(dow !== 0 && dow !== 6 && !(ctx && ctx.isHoliday)) remaining--;
  }
  return toYMD(d);
}
function calcExecutionMetrics(activity){
  const data = normalizeExecData(activity);

  // A atividade principal permanece como o pacote total planejado.
  // As subatividades representam parcelas/abatimentos desse pacote, não o novo total da atividade.
  const horasPlanejadasAtividade = Number(String(activity?.horasPlanejadas ?? activity?.execHorasPlanejadas ?? 0).replace(',', '.')) || estimatePlannedHoursFromActivity(activity);
  const horasPlanejadasSubatividades = data.subtasks.reduce((acc,x)=>acc + (Number(x.horasPlanejadas)||0), 0);
  const horasRealizadas = data.entries.reduce((acc,x)=>acc + (Number(x.horas)||0), 0);
  const impactoHoras = data.issues.reduce((acc,x)=>acc + (Number(x.impactoHoras)||0), 0);
  const impactoPrazoDias = data.issues.reduce((acc,x)=>acc + (Number(x.impactoPrazoDias)||0), 0);

  const entriesBySubtask = {};
  data.entries.forEach(en => {
    const key = en.subtaskId || '';
    entriesBySubtask[key] = (entriesBySubtask[key] || 0) + (Number(en.horas)||0);
  });
  const consumedBySubtasks = data.subtasks.reduce((acc, st) => {
    const planned = Number(st.horasPlanejadas)||0;
    const done = entriesBySubtask[st.id] || 0;
    return acc + Math.max(planned, done);
  }, 0);
  const consumedWithoutSubtask = entriesBySubtask[''] || 0;
  const horasAbatidas = consumedBySubtasks + consumedWithoutSubtask;

  const totalAjustado = horasPlanejadasAtividade + impactoHoras;
  const horasRestantes = Math.max(0, totalAjustado - horasAbatidas);
  const percentualReal = totalAjustado > 0 ? Math.min(999, (horasRealizadas / totalAjustado) * 100) : 0;
  const r = resources.find(x=>x.id === activity.resourceId);
  const capacidadeDia = Math.max(0.25, getDailyHours(r) * ((Number(activity?.alocacao ?? 100) || 100) / 100));
  const diasNecessarios = horasRestantes > 0 ? Math.ceil(horasRestantes / capacidadeDia) : 0;
  const terminoPlanejado = activity.fim || '';
  const previsaoFim = horasRestantes > 0 ? addBusinessDaysFromToday(diasNecessarios + impactoPrazoDias, activity.resourceId) : (terminoPlanejado || '');
  const atrasoDias = previsaoFim && terminoPlanejado ? Math.max(0, diffDays(fromYMD(previsaoFim), fromYMD(terminoPlanejado))) : 0;
  return {
    horasPlanejadas: horasPlanejadasAtividade,
    horasPlanejadasAtividade,
    horasPlanejadasSubatividades,
    horasRealizadas,
    horasAbatidas,
    impactoHoras,
    impactoPrazoDias,
    horasRestantes,
    percentualReal,
    capacidadeDia,
    diasNecessarios,
    terminoPlanejado,
    previsaoFim,
    atrasoDias
  };
}

function getExecStatusBadge(m){
  if(!m) return {label:'Sem dados', cls:''};
  if(m.atrasoDias > 0) return {label:'Atraso previsto', cls:'danger'};
  if((m.percentualReal || 0) < 70 && m.horasRestantes > 0) return {label:'Em acompanhamento', cls:'warn'};
  return {label:'Dentro do previsto', cls:'ok'};
}
function renderExecutionStrip(activity){
  const m = calcExecutionMetrics(activity);
  const badge = getExecStatusBadge(m);
  const pct = Math.max(0, Math.min(100, m.percentualReal || 0));
  return `<div class="execution-strip">
    <div style="display:flex;justify-content:space-between;gap:8px;align-items:center;flex-wrap:wrap;">
      <strong>Execução</strong><span class="exec-badge ${badge.cls}">${escHTML(badge.label)}</span>
    </div>
    <div class="execution-progress" title="${pct.toFixed(1)}% executado"><div style="width:${pct}%;"></div></div>
    <div class="exec-kpis">
      <div class="exec-kpi"><strong>${m.horasRealizadas.toFixed(1)}h</strong><span>realizadas</span></div>
      <div class="exec-kpi"><strong>${m.horasPlanejadas.toFixed(1)}h</strong><span>planejadas</span></div>
      <div class="exec-kpi"><strong>${m.horasAbatidas.toFixed(1)}h</strong><span>abatidas</span></div>
      <div class="exec-kpi"><strong>${m.horasRestantes.toFixed(1)}h</strong><span>saldo</span></div>
      <div class="exec-kpi"><strong>${escHTML(m.previsaoFim || '—')}</strong><span>previsão fim</span></div>
    </div>
  </div>`;
}
function persistExecutionChange(activityId, mutator){
  const idx = activities.findIndex(a=>a.id === activityId);
  if(idx < 0) return null;
  const current = syncExecFields({ ...activities[idx] });
  mutator(current);
  current.version = (current.version || 0) + 1;
  current.updatedAt = Date.now();
  activities[idx] = syncExecFields(current);
  saveLS(LS.act, activities);
  try{ recordEvent('activity','update', current.id, current); }catch(_){ }
  renderAll();
  saveBDDebounced();
  return activities[idx];
}
function fillExecResourceOptions(selectedId){
  const sel = document.getElementById('execSubResponsavel');
  if(!sel) return;
  sel.innerHTML = '';
  (resources || []).filter(r=>!r.deletedAt && r.ativo !== false).slice().sort((a,b)=>(a.nome||'').localeCompare(b.nome||'', undefined, {sensitivity:'base'})).forEach(r=>{
    const opt = document.createElement('option'); opt.value = r.id; opt.textContent = r.nome || r.id; if(r.id === selectedId) opt.selected = true; sel.appendChild(opt);
  });
}
function fillExecSubtaskTypeOptions(selectedType){
  const sel = document.getElementById('execSubTipo');
  if(!sel) return;
  const current = selectedType || sel.value || 'Execução';
  sel.innerHTML = '';
  EXEC_SUBTASK_TYPES.forEach(t=>{
    const opt = document.createElement('option'); opt.value = t; opt.textContent = t; if(t === current) opt.selected = true; sel.appendChild(opt);
  });
  const manual = document.getElementById('execSubTipoOutro');
  if(manual){
    const isOther = sel.value === 'Outros';
    manual.disabled = !isOther;
    manual.style.display = isOther ? '' : 'none';
    if(!isOther) manual.value = '';
  }
}
function updateExecSubtaskOtherVisibility(){
  const sel = document.getElementById('execSubTipo');
  const manual = document.getElementById('execSubTipoOutro');
  if(!sel || !manual) return;
  const isOther = sel.value === 'Outros';
  manual.disabled = !isOther;
  manual.style.display = isOther ? '' : 'none';
  if(!isOther) manual.value = '';
}
function renderExecSelects(activity){
  const data = normalizeExecData(activity);
  const sel = document.getElementById('execEntrySub');
  if(sel){
    sel.innerHTML = '';
    if(!data.subtasks.length){ const opt=document.createElement('option'); opt.value=''; opt.textContent='Sem subatividade cadastrada'; sel.appendChild(opt); }
    data.subtasks.forEach(st=>{ const opt=document.createElement('option'); opt.value=st.id; opt.textContent=st.titulo; sel.appendChild(opt); });
  }
  fillExecResourceOptions(activity.resourceId);
  fillExecSubtaskTypeOptions();
}
function renderExecDialog(activity){
  activity = syncExecFields(activity);
  const data = normalizeExecData(activity);
  const m = calcExecutionMetrics(activity);
  const badge = getExecStatusBadge(m);
  const title = document.getElementById('dlgExecucaoTitulo');
  const hidden = document.getElementById('execActivityId');
  if(title) title.textContent = `Execução — ${activity.codigoAtividade || ''} ${activity.titulo || ''}`.trim();
  if(hidden) hidden.value = activity.id;
  const resumo = document.getElementById('execResumo');
  if(resumo) resumo.innerHTML = `
    <div class="exec-summary-card"><strong>${m.percentualReal.toFixed(1)}%</strong><span>execução real</span></div>
    <div class="exec-summary-card"><strong>${m.horasPlanejadas.toFixed(1)}h</strong><span>total planejado</span></div>
    <div class="exec-summary-card"><strong>${m.horasAbatidas.toFixed(1)}h</strong><span>horas abatidas</span></div>
    <div class="exec-summary-card"><strong>${m.horasRestantes.toFixed(1)}h</strong><span>saldo restante</span></div>
    <div class="exec-summary-card"><strong>${escHTML(m.terminoPlanejado || '—')}</strong><span>término planejado</span></div>
    <div class="exec-summary-card"><strong>${escHTML(m.previsaoFim || '—')}</strong><span>previsão atual de término</span></div>
    <div class="exec-summary-card"><strong><span class="exec-badge ${badge.cls}">${escHTML(badge.label)}</span></strong><span>${m.atrasoDias ? m.atrasoDias + ' dia(s) de atraso previsto' : 'sem atraso previsto'}</span></div>`;
  const subById = Object.fromEntries(data.subtasks.map(s=>[s.id,s]));
  const subList = document.getElementById('execSubList');
  if(subList) subList.innerHTML = data.subtasks.length ? `<table class="exec-table"><thead><tr><th>Tipo</th><th>Subatividade</th><th>Responsável</th><th>Planejado</th><th>Realizado</th><th>Fim previsto</th><th></th></tr></thead><tbody>${data.subtasks.map(st=>{
    const r = resources.find(x=>x.id===st.responsavelId);
    const done = data.entries.filter(e=>e.subtaskId===st.id).reduce((a,e)=>a+(Number(e.horas)||0),0);
    return `<tr><td>${escHTML(st.tipoSubatividade || 'Outros')}</td><td>${escHTML(st.titulo)}</td><td>${escHTML(r?.nome || '—')}</td><td>${Number(st.horasPlanejadas||0).toFixed(1)}h</td><td>${done.toFixed(1)}h</td><td>${escHTML(st.fimPrevisto || '—')}</td><td><button type="button" class="btn ghost exec-del-sub" data-id="${escAttr(st.id)}">Excluir</button></td></tr>`;
  }).join('')}</tbody></table>` : '<div class="muted small">Nenhuma subatividade cadastrada.</div>';
  const entryList = document.getElementById('execEntryList');
  if(entryList) entryList.innerHTML = data.entries.length ? `<table class="exec-table"><thead><tr><th>Data</th><th>Subatividade</th><th>Horas</th><th>Comentário</th><th>Usuário</th><th></th></tr></thead><tbody>${data.entries.slice().sort((a,b)=>(b.data||'').localeCompare(a.data||'')).map(en=>`<tr><td>${escHTML(en.data)}</td><td>${escHTML(subById[en.subtaskId]?.titulo || '—')}</td><td>${Number(en.horas||0).toFixed(1)}h</td><td>${escHTML(en.comentario || '')}</td><td>${escHTML(en.createdBy || '')}</td><td><button type="button" class="btn ghost exec-del-entry" data-id="${escAttr(en.id)}">Excluir</button></td></tr>`).join('')}</tbody></table>` : '<div class="muted small">Nenhum apontamento registrado.</div>';
  const issueList = document.getElementById('execIssueList');
  if(issueList) issueList.innerHTML = data.issues.length ? `<table class="exec-table"><thead><tr><th>Data</th><th>Tipo</th><th>Impacto h</th><th>Impacto dias</th><th>Descrição</th><th></th></tr></thead><tbody>${data.issues.slice().sort((a,b)=>(b.data||'').localeCompare(a.data||'')).map(is=>`<tr><td>${escHTML(is.data)}</td><td>${escHTML(is.tipo)}</td><td>${Number(is.impactoHoras||0).toFixed(1)}h</td><td>${Number(is.impactoPrazoDias||0)}</td><td>${escHTML(is.descricao || '')}</td><td><button type="button" class="btn ghost exec-del-issue" data-id="${escAttr(is.id)}">Excluir</button></td></tr>`).join('')}</tbody></table>` : '<div class="muted small">Nenhuma ocorrência registrada.</div>';
  renderExecSelects(activity);
  ['execSubList','execEntryList','execIssueList'].forEach(id=>{
    const box = document.getElementById(id); if(!box) return;
    box.querySelectorAll('.exec-del-sub').forEach(btn=>btn.onclick=()=>{ const actId=document.getElementById('execActivityId')?.value; persistExecutionChange(actId, a=>{ const d=normalizeExecData(a); a.execSubtasks=d.subtasks.filter(x=>x.id!==btn.dataset.id); a.execEntries=d.entries.filter(x=>x.subtaskId!==btn.dataset.id); a.execIssues=d.issues; }); const updated=activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated); });
    box.querySelectorAll('.exec-del-entry').forEach(btn=>btn.onclick=()=>{ const actId=document.getElementById('execActivityId')?.value; persistExecutionChange(actId, a=>{ const d=normalizeExecData(a); a.execSubtasks=d.subtasks; a.execEntries=d.entries.filter(x=>x.id!==btn.dataset.id); a.execIssues=d.issues; }); const updated=activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated); });
    box.querySelectorAll('.exec-del-issue').forEach(btn=>btn.onclick=()=>{ const actId=document.getElementById('execActivityId')?.value; persistExecutionChange(actId, a=>{ const d=normalizeExecData(a); a.execSubtasks=d.subtasks; a.execEntries=d.entries; a.execIssues=d.issues.filter(x=>x.id!==btn.dataset.id); }); const updated=activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated); });
  });
}
function openExecutionDialog(activity){
  const dlg = document.getElementById('dlgExecucao');
  if(!dlg || !activity) return;
  renderExecDialog(activity);
  const todayStr = toYMD(new Date());
  const entryDate = document.getElementById('execEntryData'); if(entryDate && !entryDate.value) entryDate.value = todayStr;
  const subFim = document.getElementById('execSubFim'); if(subFim && !subFim.value) subFim.value = activity.fim || todayStr;
  dlg.showModal();
}
function bindExecutionDialog(){
  const form = document.getElementById('formExecucao');
  if(!form || form.dataset.execBound === '1') return;
  form.dataset.execBound = '1';
  const typeSel = document.getElementById('execSubTipo');
  if(typeSel) typeSel.addEventListener('change', updateExecSubtaskOtherVisibility);
  updateExecSubtaskOtherVisibility();

  form.addEventListener('click', (ev)=>{
    const btn = ev.target && ev.target.closest ? ev.target.closest('button') : null;
    if(!btn) return;
    if(!['btnExecAddSub','btnExecAddEntry','btnExecAddIssue'].includes(btn.id)) return;
    ev.preventDefault();
    ev.stopPropagation();

    const actId = document.getElementById('execActivityId')?.value;
    if(!actId || !activities.some(x=>x.id===actId)){
      alert('Não foi possível identificar a atividade em execução. Feche e abra novamente o painel de execução.');
      return;
    }

    if(btn.id === 'btnExecAddSub'){
      const tituloEl = document.getElementById('execSubTitulo');
      const titulo = String(tituloEl?.value || '').trim();
      if(!titulo) return alert('Informe o título da subatividade.');
      const horasPlanejadas = Number(String(document.getElementById('execSubHoras')?.value || '0').replace(',', '.')) || 0;
      if(horasPlanejadas <= 0) return alert('Informe as horas planejadas da subatividade.');

      persistExecutionChange(actId, a=>{
        const d = normalizeExecData(a);
        const selectedTipo = document.getElementById('execSubTipo')?.value || 'Outros';
        const tipoOutro = document.getElementById('execSubTipoOutro')?.value || '';
        if(selectedTipo === 'Outros' && !String(tipoOutro).trim()) return alert('Informe o tipo manual da subatividade.');
        const sub = normalizeExecSubtask({
          titulo,
          tipoSubatividade: selectedTipo,
          tipoSubatividadeOutro: tipoOutro,
          responsavelId: document.getElementById('execSubResponsavel')?.value || a.resourceId,
          horasPlanejadas,
          fimPrevisto: document.getElementById('execSubFim')?.value || '',
          createdBy: currentUser || ''
        });
        if(sub) d.subtasks.push(sub);
        a.execSubtasks = d.subtasks;
        a.execEntries = d.entries;
        a.execIssues = d.issues;
      });

      if(tituloEl) tituloEl.value = '';
      const tipoOutroEl = document.getElementById('execSubTipoOutro'); if(tipoOutroEl) tipoOutroEl.value = '';
      const tipoSel = document.getElementById('execSubTipo'); if(tipoSel) tipoSel.value = 'Execução';
      updateExecSubtaskOtherVisibility();
      const horasEl = document.getElementById('execSubHoras'); if(horasEl) horasEl.value = '1';
      const updated = activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated);
      return;
    }

    if(btn.id === 'btnExecAddEntry'){
      const horas = Number(String(document.getElementById('execEntryHoras')?.value || '').replace(',', '.')) || 0;
      if(horas <= 0) return alert('Informe as horas realizadas.');
      const subtaskId = document.getElementById('execEntrySub')?.value || '';
      const activity = activities.find(x=>x.id===actId);
      const hasSubtasks = normalizeExecData(activity || {}).subtasks.length > 0;
      if(hasSubtasks && !subtaskId) return alert('Selecione uma subatividade para registrar as horas.');

      persistExecutionChange(actId, a=>{
        const d = normalizeExecData(a);
        const entry = normalizeExecEntry({
          subtaskId,
          data: document.getElementById('execEntryData')?.value || toYMD(new Date()),
          horas,
          comentario: limitExecComment(document.getElementById('execEntryComentario')?.value || ''),
          createdBy: currentUser || ''
        });
        if(entry) d.entries.push(entry);
        a.execSubtasks = d.subtasks;
        a.execEntries = d.entries;
        a.execIssues = d.issues;
      });

      const comentarioEl = document.getElementById('execEntryComentario'); if(comentarioEl) comentarioEl.value = '';
      const horasEl = document.getElementById('execEntryHoras'); if(horasEl) horasEl.value = '1';
      const updated = activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated);
      return;
    }

    if(btn.id === 'btnExecAddIssue'){
      const descricaoEl = document.getElementById('execIssueDescricao');
      const descricao = limitExecComment(descricaoEl?.value || '');
      const impactoHoras = Number(String(document.getElementById('execIssueHoras')?.value || '0').replace(',', '.')) || 0;
      const impactoPrazoDias = Number(document.getElementById('execIssueDias')?.value || 0) || 0;
      if(!descricao && impactoHoras <= 0 && impactoPrazoDias <= 0) return alert('Informe uma descrição ou impacto da ocorrência.');

      persistExecutionChange(actId, a=>{
        const d = normalizeExecData(a);
        const issue = normalizeExecIssue({
          tipo: document.getElementById('execIssueTipo')?.value || 'Ocorrência',
          descricao,
          impactoHoras,
          impactoPrazoDias,
          data: toYMD(new Date()),
          createdBy: currentUser || ''
        });
        if(issue) d.issues.push(issue);
        a.execSubtasks = d.subtasks;
        a.execEntries = d.entries;
        a.execIssues = d.issues;
      });

      if(descricaoEl) descricaoEl.value = '';
      const hEl = document.getElementById('execIssueHoras'); if(hEl) hEl.value = '0';
      const dEl = document.getElementById('execIssueDias'); if(dEl) dEl.value = '0';
      const updated = activities.find(x=>x.id===actId); if(updated) renderExecDialog(updated);
      return;
    }
  });
}

if(document.readyState === 'loading'){
  document.addEventListener('DOMContentLoaded', bindExecutionDialog);
}else{
  bindExecutionDialog();
}

function normalizeOvertimeEntry(row){
  if(!row || typeof row !== 'object') return null;
  const resourceId = String(row.resourceId || row.recursoId || row.idRecurso || row.id || '').trim();
  const date = normalizeDateField(row.date || row.data || row.dia || '');
  const rawHours = row.horas ?? row.hours ?? row.overtimeHours ?? row.horasExtras ?? '';
  const rawMin = row.minutos ?? row.minutes ?? '';
  let hours = Number(String(rawHours).replace(',', '.'));
  if((!hours || isNaN(hours)) && rawMin !== '') hours = Number(rawMin) / 60;
  if(!resourceId || !date || !isFinite(hours) || hours <= 0) return null;
  const status = String(row.status || row.situacao || 'aprovada').trim().toLowerCase();
  return {
    id: String(row.overtimeId || row.horaExtraId || row.entryId || row.uuid || row.registroId || row.rowId || row.idRegistro || row.idHoraExtra || (row.tabela === 'hora_extra' ? row.id : '') || uuid()),
    resourceId,
    date,
    hours: Math.min(24, Math.max(0.25, hours)),
    status: ['aprovada','pendente','reprovada'].includes(status) ? status : 'aprovada',
    reason: String(row.reason || row.justificativa || row.motivo || '').trim(),
    createdAt: Number(row.createdAt || 0) || Date.now(),
    createdBy: String(row.createdBy || row.user || row.usuario || '').trim()
  };
}

function normalizeOvertimeList(list){
  const seen = new Set();
  return (Array.isArray(list) ? list : []).map(normalizeOvertimeEntry).filter(Boolean).filter(x=>{
    const key = [x.resourceId, x.date, x.hours, x.status, x.reason].join('|');
    if(seen.has(key)) return false;
    seen.add(key);
    return true;
  }).sort((a,b)=>(a.date||'').localeCompare(b.date||'') || String(a.resourceId||'').localeCompare(String(b.resourceId||'')));
}

function getOvertimeEntries(){ return normalizeOvertimeList(overtimeEntries || []); }
function setOvertimeEntries(list){
  overtimeEntries = normalizeOvertimeList(list);
  saveLS(LS_OVERTIME, overtimeEntries);
  renderOvertimeUI();
  try{ renderAll(); }catch(_){ }
  try{ saveBDDebounced(); }catch(_){ }
}
function getApprovedOvertimeHours(resource, dateOrYmd){
  const resourceId = typeof resource === 'string' ? resource : String(resource?.id || '');
  const ymd = typeof dateOrYmd === 'string' ? normalizeDateField(dateOrYmd) : toYMD(dateOrYmd);
  return getOvertimeEntries().filter(x=>x.resourceId === resourceId && x.date === ymd && x.status === 'aprovada').reduce((acc,x)=>acc + Number(x.hours || 0), 0);
}

async function sha256Text(text){
  const value = String(text || '');
  if(window.crypto && crypto.subtle){
    const data = new TextEncoder().encode(value);
    const hash = await crypto.subtle.digest('SHA-256', data);
    return Array.from(new Uint8Array(hash)).map(b=>b.toString(16).padStart(2,'0')).join('');
  }
  let h = 0;
  for(let i=0;i<value.length;i++){ h = ((h<<5)-h) + value.charCodeAt(i); h |= 0; }
  return 'fallback-' + Math.abs(h);
}

function normalizeMatricula(value){
  return String(value || '').replace(/\D+/g, '').trim();
}
function normalizeSystemUser(row){
  if(!row || typeof row !== 'object') return null;
  const matricula = normalizeMatricula(row.matricula || row.Matricula || row.id || row.ID || '');
  if(!matricula) return null;
  const perfilRaw = String(row.perfil || row.Perfil || 'Executor').trim();
  const perfil = ACCESS_PROFILES.includes(perfilRaw) ? perfilRaw : 'Executor';
  const ativoRaw = row.ativo ?? row.Ativo ?? true;
  const ativo = typeof ativoRaw === 'boolean' ? ativoRaw : !String(ativoRaw || 'S').trim().toUpperCase().startsWith('N');
  return {
    matricula,
    nome: String(row.nome || row.Nome || '').trim(),
    perfil,
    ativo,
    pinHash: String(row.pinHash || row.PinHash || row.hashPin || '').trim(),
    trocarPinNoPrimeiroAcesso: String(row.trocarPinNoPrimeiroAcesso ?? row.trocarPin ?? 'N').trim().toUpperCase().startsWith('S') || row.trocarPinNoPrimeiroAcesso === true,
    administradorInicial: String(row.administradorInicial ?? row.adminInicial ?? 'N').trim().toUpperCase().startsWith('S') || row.administradorInicial === true,
    createdAt: row.createdAt || row.CriadoEm || Date.now(),
    updatedAt: row.updatedAt || row.AtualizadoEm || Date.now(),
    createdBy: String(row.createdBy || row.CriadoPor || '').trim()
  };
}
function normalizeSystemUsers(list){
  const seen = new Set();
  return (list || []).map(normalizeSystemUser).filter(u=>{
    if(!u || seen.has(u.matricula)) return false;
    seen.add(u.matricula);
    return true;
  });
}
async function hashUserPin(matricula, pin){
  return sha256Text('rp-user-pin|' + normalizeMatricula(matricula) + '|' + String(pin || ''));
}

function normalizeAuditEvent(row){
  if(!row || typeof row !== 'object') return null;
  const tsRaw = row.ts || row.timestamp || row.createdAt || row.dataEvento || Date.now();
  const d = tsRaw instanceof Date ? tsRaw : new Date(Number(tsRaw) || tsRaw);
  const safe = (!d || isNaN(d.getTime())) ? new Date() : d;
  return {
    eventId: String(row.eventId || row.id || uuid()),
    ts: safe.toISOString(),
    timestamp: safe.getTime(),
    eventType: String(row.eventType || row.tipoEvento || row.tipo || row.action || 'EVENTO').trim(),
    action: String(row.action || row.acao || row.eventType || '').trim(),
    entityType: String(row.entityType || row.tipoEntidade || row.entidadeTipo || '').trim(),
    entityId: String(row.entityId || row.entidadeId || row.activityId || '').trim(),
    entityLabel: String(row.entityLabel || row.entidade || row.titulo || '').trim(),
    reason: String(row.reason || row.justificativa || row.motivo || row.detail || '').trim(),
    matricula: normalizeMatricula(row.matricula || row.userMatricula || row.usuario || ''),
    nome: String(row.nome || row.userName || '').trim(),
    perfil: String(row.perfil || row.profile || '').trim(),
    beforeJson: String(row.beforeJson || row.antes || '').trim(),
    afterJson: String(row.afterJson || row.depois || '').trim(),
    source: String(row.source || row.origem || 'sistema').trim()
  };
}
function normalizeAuditEvents(list){
  const seen = new Set();
  return (Array.isArray(list) ? list : []).map(normalizeAuditEvent).filter(Boolean).filter(ev=>{
    const key = ev.eventId || [ev.ts, ev.eventType, ev.entityType, ev.entityId, ev.reason, ev.matricula].join('|');
    if(seen.has(key)) return false;
    seen.add(key);
    return true;
  }).sort((a,b)=>Number(b.timestamp||0)-Number(a.timestamp||0));
}
function currentAuditUserInfo(){
  const u = getCurrentAccessUser ? getCurrentAccessUser() : null;
  return {
    matricula: u?.matricula || normalizeMatricula(accessSession?.matricula || ''),
    nome: u?.nome || accessSession?.nome || '',
    perfil: u?.perfil || accessSession?.perfil || ''
  };
}
function recordAuditEvent(eventType, payload){
  try{
    const who = currentAuditUserInfo();
    const ev = normalizeAuditEvent({
      ...(payload || {}),
      eventId: uuid(),
      ts: new Date().toISOString(),
      timestamp: Date.now(),
      eventType,
      action: payload?.action || eventType,
      matricula: payload?.matricula || who.matricula,
      nome: payload?.nome || who.nome,
      perfil: payload?.perfil || who.perfil,
      source: payload?.source || 'app'
    });
    auditEvents = normalizeAuditEvents([ev, ...(auditEvents || [])]).slice(0, 10000);
    saveLS('rp_audit_events_v1', auditEvents);
    try{ renderAuditTrail(); }catch(_){ }
    return ev;
  }catch(e){ console.warn('[audit] falha ao registrar evento', e); return null; }
}
function auditEventTypeLabel(type){
  const map = {
    LOGIN:'Login', LOGOUT:'Logout', PRIMEIRO_ACESSO_TROCA_PIN:'Troca de PIN inicial', ALTERACAO_PIN:'Alteração de PIN',
    USUARIO_CRIADO:'Usuário criado', USUARIO_ATUALIZADO:'Usuário atualizado', USUARIO_STATUS:'Status de usuário', PIN_RESETADO:'PIN resetado',
    ADMIN_INICIAL:'Administrador inicial', ALTERACAO_DATAS:'Alteração de datas', COMENTARIO:'Comentário', EXECUCAO_APONTAMENTO:'Apontamento de execução', OCORRENCIA:'Ocorrência',
    HISTORICO_ATIVIDADE:'Histórico da atividade'
  };
  return map[type] || String(type || 'Evento');
}
function buildConsolidatedAuditRows(){
  const rows = normalizeAuditEvents(auditEvents || []);
  try{
    Object.keys(trails || {}).forEach(activityId=>{
      const act = (activities || []).find(a=>a.id===activityId) || {};
      (trails[activityId] || []).forEach(h=>{
        rows.push(normalizeAuditEvent({
          eventId:'hist-'+activityId+'-'+String(h.ts||'')+'-'+String(h.type||''),
          ts:h.ts || h.timestamp || Date.now(),
          eventType:h.type || 'HISTORICO_ATIVIDADE',
          action:h.type || 'Histórico',
          entityType:h.entityType || 'atividade',
          entityId:activityId,
          entityLabel:act.codigoAtividade ? `${act.codigoAtividade} - ${act.titulo||''}` : (act.titulo || activityId),
          reason:h.justificativa || h.motivo || '',
          matricula:h.user || h.usuario || '',
          nome:h.user || '',
          perfil:'',
          source:'historico'
        }));
      });
    });
    (comments || []).filter(c=>c && !c.deletedAt).forEach(c=>{
      const act=(activities||[]).find(a=>a.id===c.activityId)||{};
      rows.push(normalizeAuditEvent({
        eventId:'coment-'+(c.commentId||uuid()), ts:c.ts||c.createdAt||Date.now(), eventType:'COMENTARIO', action:'Comentário publicado',
        entityType:'atividade', entityId:c.activityId, entityLabel:act.codigoAtividade ? `${act.codigoAtividade} - ${act.titulo||''}` : (act.titulo || c.activityId),
        reason:String(c.texto||'').slice(0,360), matricula:c.usuario||'', nome:c.usuario||'', source:'comentarios'
      }));
    });
    (activities || []).forEach(a=>{
      getExecIssues(a).forEach(i=>rows.push(normalizeAuditEvent({
        eventId:'occ-'+(i.id||uuid()), ts:i.createdAt||Date.now(), eventType:'OCORRENCIA', action:i.tipo||'Ocorrência', entityType:'atividade', entityId:a.id,
        entityLabel:a.codigoAtividade ? `${a.codigoAtividade} - ${a.titulo||''}` : (a.titulo||a.id), reason:i.descricao||i.reason||'', matricula:i.createdBy||'', nome:i.createdBy||'', source:'execucao'
      })));
      getExecEntries(a).forEach(e=>rows.push(normalizeAuditEvent({
        eventId:'apont-'+(e.id||uuid()), ts:e.createdAt||Date.now(), eventType:'EXECUCAO_APONTAMENTO', action:'Apontamento de horas', entityType:'atividade', entityId:a.id,
        entityLabel:a.codigoAtividade ? `${a.codigoAtividade} - ${a.titulo||''}` : (a.titulo||a.id), reason:`${Number(e.horas||0)}h — ${e.comentario||''}`.trim(), matricula:e.createdBy||e.recursoId||'', nome:e.createdBy||'', source:'execucao'
      })));
    });
  }catch(e){ console.warn('[audit] consolidação parcial', e); }
  return normalizeAuditEvents(rows);
}
function getAuditFilters(){
  return {
    from: document.getElementById('auditFrom')?.value || '',
    to: document.getElementById('auditTo')?.value || '',
    user: String(document.getElementById('auditUser')?.value || '').toLowerCase(),
    type: document.getElementById('auditType')?.value || '',
    search: String(document.getElementById('auditSearch')?.value || '').toLowerCase()
  };
}
function filterAuditRows(rows){
  const f = getAuditFilters();
  const fromTs = f.from ? fromYMD(f.from).getTime() : null;
  const toTs = f.to ? addDays(fromYMD(f.to),1).getTime()-1 : null;
  return (rows || []).filter(r=>{
    const t = Number(r.timestamp || 0);
    if(fromTs && t < fromTs) return false;
    if(toTs && t > toTs) return false;
    if(f.type && r.eventType !== f.type) return false;
    const userText = `${r.matricula||''} ${r.nome||''} ${r.perfil||''}`.toLowerCase();
    if(f.user && !userText.includes(f.user)) return false;
    const all = `${r.eventType||''} ${r.action||''} ${r.entityType||''} ${r.entityId||''} ${r.entityLabel||''} ${r.reason||''} ${r.matricula||''} ${r.nome||''}`.toLowerCase();
    if(f.search && !all.includes(f.search)) return false;
    return true;
  });
}
function refreshAuditTypeOptions(rows){
  const sel = document.getElementById('auditType'); if(!sel) return;
  const current = sel.value;
  const types = Array.from(new Set((rows || []).map(r=>r.eventType).filter(Boolean))).sort();
  sel.innerHTML = '<option value="">Todos</option>' + types.map(t=>`<option value="${escAttr(t)}">${escHTML(auditEventTypeLabel(t))}</option>`).join('');
  if(types.includes(current)) sel.value = current;
}
function renderAuditTrail(){
  const tbody = document.querySelector('#auditTable tbody');
  if(!tbody) return;
  const perms = getCurrentPermissions ? getCurrentPermissions() : null;
  const panel = document.getElementById('tab-audit');
  if(panel && (!perms || !perms.audit)){
    tbody.innerHTML = '<tr><td colspan="6" class="muted">Seu perfil não possui permissão para visualizar a trilha de auditoria.</td></tr>';
    return;
  }
  const allRows = buildConsolidatedAuditRows();
  refreshAuditTypeOptions(allRows);
  const rows = filterAuditRows(allRows);
  const kpis = document.getElementById('auditKpis');
  if(kpis){
    const users = new Set(rows.map(r=>r.matricula||r.nome).filter(Boolean)).size;
    const todayStr = toYMD(new Date());
    const todayCount = rows.filter(r=>toYMD(new Date(r.timestamp))===todayStr).length;
    const critical = rows.filter(r=>['PIN_RESETADO','USUARIO_STATUS','USUARIO_CRIADO','USUARIO_ATUALIZADO','ADMIN_INICIAL'].includes(r.eventType)).length;
    kpis.innerHTML = `<div class="card"><strong>${rows.length}</strong><span class="muted small">eventos filtrados</span></div><div class="card"><strong>${users}</strong><span class="muted small">usuários envolvidos</span></div><div class="card"><strong>${todayCount}</strong><span class="muted small">eventos hoje</span></div><div class="card"><strong>${critical}</strong><span class="muted small">eventos administrativos</span></div>`;
  }
  const totalPages = Math.max(1, Math.ceil(rows.length / AUDIT_PAGE_SIZE));
  auditPage = Math.min(Math.max(1, auditPage || 1), totalPages);
  const pageRows = rows.slice((auditPage-1)*AUDIT_PAGE_SIZE, auditPage*AUDIT_PAGE_SIZE);
  tbody.innerHTML = pageRows.map(r=>`<tr><td>${escHTML(formatDateTimeBR(r.ts))}</td><td>${escHTML([r.matricula,r.nome].filter(Boolean).join(' - '))}</td><td>${escHTML(r.perfil||'')}</td><td>${escHTML(auditEventTypeLabel(r.eventType))}</td><td>${escHTML([r.entityType,r.entityLabel||r.entityId].filter(Boolean).join(' · '))}</td><td style="white-space:pre-wrap;max-width:420px;">${escHTML(r.reason||r.action||'')}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">Nenhum evento encontrado.</td></tr>';
  const pag = document.getElementById('auditPagination');
  if(pag){
    pag.innerHTML = `<button class="btn" type="button" id="auditPrev" ${auditPage<=1?'disabled':''}>Anterior</button><span class="muted small">Página ${auditPage} de ${totalPages} · ${rows.length} registros</span><button class="btn" type="button" id="auditNext" ${auditPage>=totalPages?'disabled':''}>Próxima</button>`;
    const prev=document.getElementById('auditPrev'); if(prev) prev.onclick=()=>{ auditPage--; renderAuditTrail(); };
    const next=document.getElementById('auditNext'); if(next) next.onclick=()=>{ auditPage++; renderAuditTrail(); };
  }
}
function exportAuditCsv(){
  const rows = filterAuditRows(buildConsolidatedAuditRows());
  const headers = ['quando','matricula','nome','perfil','tipoEvento','acao','tipoEntidade','entidade','justificativa','origem'];
  const csv = [headers.join(';')].concat(rows.map(r=>[r.ts,r.matricula,r.nome,r.perfil,auditEventTypeLabel(r.eventType),r.action,r.entityType,r.entityLabel||r.entityId,r.reason,r.source].map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(';'))).join('\n');
  download('trilha_auditoria.csv', '\ufeff'+csv, 'text/csv;charset=utf-8');
}
function bindAuditUI(){
  ['auditFrom','auditTo','auditUser','auditType','auditSearch'].forEach(id=>{ const el=document.getElementById(id); if(el) el.addEventListener(id==='auditType'?'change':'input',()=>{ auditPage=1; renderAuditTrail(); }); });
  const apply=document.getElementById('btnAuditApply'); if(apply) apply.onclick=()=>{ auditPage=1; renderAuditTrail(); };
  const clear=document.getElementById('btnAuditClear'); if(clear) clear.onclick=()=>{ ['auditFrom','auditTo','auditUser','auditSearch'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; }); const t=document.getElementById('auditType'); if(t) t.value=''; auditPage=1; renderAuditTrail(); };
  const exp=document.getElementById('btnAuditExportCsv'); if(exp) exp.onclick=exportAuditCsv;
  renderAuditTrail();
}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', bindAuditUI); else bindAuditUI();
function getCurrentAccessUser(){
  if(!accessSession || !accessSession.matricula) return null;
  return (systemUsers || []).find(u=>u.matricula === accessSession.matricula && u.ativo) || null;
}
function getCurrentPermissions(){
  const u = getCurrentAccessUser();
  return u ? (ACCESS_PERMISSIONS[u.perfil] || ACCESS_PERMISSIONS['Executor']) : null;
}
function isAccessAdmin(){
  const p = getCurrentPermissions();
  return !!(p && p.admin);
}
function setAccessSession(user){
  accessSession = user ? { matricula:user.matricula, nome:user.nome, perfil:user.perfil, ts:Date.now() } : null;
  try{ if(accessSession) localStorage.setItem(LS_ACCESS_SESSION, JSON.stringify(accessSession)); else localStorage.removeItem(LS_ACCESS_SESSION); }catch(_){ }
  currentUser = user ? `${user.matricula} - ${user.nome || user.perfil}` : '';
  saveLS(LS.user, currentUser);
  if(currentUserInput) currentUserInput.value = currentUser;
  applyAccessPermissions();
  renderAccessUI();
  try{ renderAuditTrail(); }catch(_){ }
}
function restoreAccessSession(){
  try{
    const raw = localStorage.getItem(LS_ACCESS_SESSION);
    if(!raw) return null;
    const sess = JSON.parse(raw);
    if(!sess || !sess.matricula) return null;
    const u = (systemUsers || []).find(x=>x.matricula === sess.matricula && x.ativo);
    if(!u) return null;
    accessSession = { matricula:u.matricula, nome:u.nome, perfil:u.perfil, ts:Date.now() };
    currentUser = `${u.matricula} - ${u.nome || u.perfil}`;
    saveLS(LS.user, currentUser);
    if(currentUserInput) currentUserInput.value = currentUser;
    return accessSession;
  }catch(_){ return null; }
}
function hasAccessConfigured(){ return Array.isArray(systemUsers) && systemUsers.length > 0; }
function bdHasLegacyData(){ return (resources||[]).length > 0 || (activities||[]).length > 0 || (comments||[]).length > 0; }
function ensureAccessStyles(){
  if(document.getElementById('rp-access-styles')) return;
  const st = document.createElement('style');
  st.id = 'rp-access-styles';
  st.textContent = `
    .rp-access-overlay{position:fixed;inset:0;z-index:20000;background:rgba(15,23,42,.82);display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(3px)}
    .rp-access-card{width:min(560px,96vw);background:#fff;color:#0f172a;border-radius:18px;box-shadow:0 30px 70px rgba(0,0,0,.35);padding:22px;border:1px solid #e2e8f0}
    .rp-access-card h2{margin:0 0 8px 0}.rp-access-card p{margin:4px 0 12px 0;color:#475569;line-height:1.45}.rp-access-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}.rp-access-card label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:#334155}.rp-access-card input,.rp-access-card select{padding:10px;border:1px solid #cbd5e1;border-radius:10px}.rp-access-actions{display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;margin-top:14px}.rp-access-error{color:#b91c1c;font-size:13px;margin-top:8px;min-height:18px}.rp-user-pill{display:inline-flex;align-items:center;gap:6px;border:1px solid #cbd5e1;background:#f8fafc;color:#334155;border-radius:999px;padding:6px 10px;font-size:12px}.rp-users-table{width:100%;border-collapse:collapse}.rp-users-table th,.rp-users-table td{border-bottom:1px solid #e2e8f0;padding:7px;text-align:left;font-size:13px}.rp-users-config{margin-top:16px}.rp-users-form{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:8px;align-items:end}.rp-users-form label{font-size:12px;color:#475569;display:flex;flex-direction:column;gap:4px}.rp-users-form input,.rp-users-form select{padding:8px;border:1px solid #cbd5e1;border-radius:8px}@media(max-width:900px){.rp-access-grid,.rp-users-form{grid-template-columns:1fr}.rp-access-actions{justify-content:stretch}.rp-access-actions .btn{width:100%}}
  `;
  document.head.appendChild(st);
}
function getAccessOverlay(){
  ensureAccessStyles();
  let el = document.getElementById('rpAccessOverlay');
  if(el) return el;
  el = document.createElement('div');
  el.id = 'rpAccessOverlay';
  el.className = 'rp-access-overlay';
  document.body.appendChild(el);
  return el;
}
function renderAccessUI(){
  const el = getAccessOverlay();
  const user = getCurrentAccessUser();
  const needsSetup = bdHandle && !hasAccessConfigured();
  const needsBD = !bdHandle && !hasAccessConfigured();
  if(user){ el.style.display='none'; renderAccessStatusPill(); return; }
  el.style.display='flex';
  if(needsBD){
    el.innerHTML = `<div class="rp-access-card"><h2>Conectar banco compartilhado</h2><p>Para ativar o controle de acesso, conecte o BD oficial. Se o banco for legado, o sistema preservará os dados e solicitará a criação do administrador inicial.</p><div class="rp-access-actions"><button class="btn primary" id="rpAccessPickBD">Conectar banco compartilhado</button></div><div class="rp-access-error" id="rpAccessErr"></div></div>`;
    const b = document.getElementById('rpAccessPickBD');
    if(b) b.onclick = async()=>{ try{ if(typeof selectAndLoadBDFile === 'function') await selectAndLoadBDFile(); else document.getElementById('btnPickBDFile')?.click(); }catch(e){ const er=document.getElementById('rpAccessErr'); if(er) er.textContent=e?.message||'Erro ao conectar BD.'; } };
    return;
  }
  if(needsSetup){
    el.innerHTML = `<div class="rp-access-card"><h2>Configuração inicial de acesso</h2><p>${bdHasLegacyData() ? 'Banco legado detectado. Os dados existentes serão mantidos. Crie o administrador inicial para liberar o uso controlado.' : 'Nenhum usuário de acesso foi encontrado. Crie o administrador inicial.'}</p><div class="rp-access-grid"><label>Matrícula<input id="rpSetupMatricula" inputmode="numeric" autocomplete="off"></label><label>Nome<input id="rpSetupNome" autocomplete="off"></label><label>PIN<input id="rpSetupPin" type="password" inputmode="numeric" autocomplete="new-password"></label><label>Confirmar PIN<input id="rpSetupPin2" type="password" inputmode="numeric" autocomplete="new-password"></label></div><div class="rp-access-actions"><button class="btn primary" id="rpCreateAdmin">Criar administrador inicial</button></div><div class="rp-access-error" id="rpAccessErr"></div></div>`;
    document.getElementById('rpCreateAdmin').onclick = createInitialAdminFromOverlay;
    return;
  }
  el.innerHTML = `<div class="rp-access-card"><h2>Acesso obrigatório</h2><p>Informe sua matrícula e PIN para acessar o Planejador de Recursos.</p><div class="rp-access-grid"><label>Matrícula<input id="rpLoginMatricula" inputmode="numeric" autocomplete="username"></label><label>PIN<input id="rpLoginPin" type="password" inputmode="numeric" autocomplete="current-password"></label></div><div class="rp-access-actions"><button class="btn" id="rpLoginPickBD">Trocar/conectar BD</button><button class="btn primary" id="rpLoginBtn">Entrar</button></div><div class="rp-access-error" id="rpAccessErr"></div></div>`;
  document.getElementById('rpLoginBtn').onclick = loginFromOverlay;
  const p=document.getElementById('rpLoginPin'); if(p) p.addEventListener('keydown',e=>{ if(e.key==='Enter') loginFromOverlay(); });
  const b=document.getElementById('rpLoginPickBD'); if(b) b.onclick = async()=>{ if(typeof selectAndLoadBDFile === 'function') await selectAndLoadBDFile(); };
}

function renderInitialPinChange(user){
  const el = getAccessOverlay();
  el.style.display='flex';
  el.innerHTML = `<div class="rp-access-card"><h2>Alterar PIN inicial</h2><p>Este é seu primeiro acesso ou seu PIN foi redefinido. Defina um novo PIN para continuar.</p><div class="rp-access-grid"><label>Novo PIN<input id="rpNewPin1" type="password" inputmode="numeric" autocomplete="new-password"></label><label>Confirmar novo PIN<input id="rpNewPin2" type="password" inputmode="numeric" autocomplete="new-password"></label></div><div class="rp-access-actions"><button class="btn primary" id="rpConfirmPinChange" type="button">Salvar novo PIN</button></div><div class="rp-access-error" id="rpAccessErr"></div></div>`;
  document.getElementById('rpConfirmPinChange').onclick = async()=>{
    const err=document.getElementById('rpAccessErr'); if(err) err.textContent='';
    const p1=String(document.getElementById('rpNewPin1')?.value||'');
    const p2=String(document.getElementById('rpNewPin2')?.value||'');
    if(p1.length<4 || p1!==p2){ if(err) err.textContent='Informe PIN com pelo menos 4 dígitos e confirmação igual.'; return; }
    user.pinHash = await hashUserPin(user.matricula, p1);
    user.trocarPinNoPrimeiroAcesso = false;
    user.updatedAt = Date.now();
    recordAuditEvent('PRIMEIRO_ACESSO_TROCA_PIN', { entityType:'usuario_sistema', entityId:user.matricula, entityLabel:user.nome||user.matricula, reason:'PIN inicial alterado pelo usuário.' });
    setAccessSession(user);
    renderUsersConfigUI();
    saveBDDebounced();
    showToast('PIN alterado', 'Acesso liberado.', 'success', 3500);
  };
}

async function createInitialAdminFromOverlay(){
  const err = document.getElementById('rpAccessErr'); if(err) err.textContent='';
  const matricula = normalizeMatricula(document.getElementById('rpSetupMatricula')?.value || '');
  const nome = String(document.getElementById('rpSetupNome')?.value || '').trim();
  const pin = String(document.getElementById('rpSetupPin')?.value || '');
  const pin2 = String(document.getElementById('rpSetupPin2')?.value || '');
  if(!matricula || !nome || pin.length < 4 || pin !== pin2){ if(err) err.textContent='Informe matrícula, nome e PIN com pelo menos 4 dígitos iguais.'; return; }
  const now = Date.now();
  const user = { matricula, nome, perfil:'Administrador', ativo:true, pinHash:await hashUserPin(matricula,pin), trocarPinNoPrimeiroAcesso:false, administradorInicial:true, createdAt:now, updatedAt:now, createdBy:'setup' };
  systemUsers = normalizeSystemUsers([user]);
  setAccessSession(user);
  recordAuditEvent('ADMIN_INICIAL', { entityType:'usuario_sistema', entityId:user.matricula, entityLabel:user.nome, reason:'Administrador inicial criado em banco legado/novo.' });
  renderUsersConfigUI();
  await saveBD();
  showToast('Administrador inicial criado', 'Controle de acesso ativado para este BD.', 'success', 5000);
}
async function loginFromOverlay(){
  const err = document.getElementById('rpAccessErr'); if(err) err.textContent='';
  const matricula = normalizeMatricula(document.getElementById('rpLoginMatricula')?.value || '');
  const pin = String(document.getElementById('rpLoginPin')?.value || '');
  const user = (systemUsers || []).find(u=>u.matricula === matricula && u.ativo);
  if(!user){ if(err) err.textContent='Usuário não localizado ou inativo.'; return; }
  const hash = await hashUserPin(matricula, pin);
  if(hash !== user.pinHash){ if(err) err.textContent='PIN inválido.'; return; }
  user.ultimoLogin = Date.now();
  user.updatedAt = Date.now();
  recordAuditEvent('LOGIN', { entityType:'usuario_sistema', entityId:user.matricula, entityLabel:user.nome||user.matricula, reason:'Login realizado.' });
  if(user.trocarPinNoPrimeiroAcesso){
    accessSession = { matricula:user.matricula, nome:user.nome, perfil:user.perfil, ts:Date.now() };
    renderInitialPinChange(user);
    saveBDDebounced();
    return;
  }
  setAccessSession(user);
  renderUsersConfigUI();
  saveBDDebounced();
}
function renderAccessStatusPill(){
  let host = document.querySelector('.topbar-meta');
  if(!host) return;
  let pill = document.getElementById('rpAccessUserPill');
  if(!pill){ pill = document.createElement('div'); pill.id='rpAccessUserPill'; pill.className='rp-user-pill'; host.prepend(pill); }
  const u = getCurrentAccessUser();
  if(!u){ pill.textContent='Sem usuário'; return; }
  pill.innerHTML = `<span>👤 ${escHTML(u.nome || u.matricula)} · ${escHTML(u.perfil)}</span> <button type="button" class="btn" id="rpLogoutBtn" style="padding:2px 6px;font-size:11px;">Sair</button>`;
  const b = document.getElementById('rpLogoutBtn');
  if(b) b.onclick = ()=>{ recordAuditEvent('LOGOUT', { entityType:'usuario_sistema', entityId:u.matricula, entityLabel:u.nome||u.matricula, reason:'Logout realizado.' }); setAccessSession(null); renderAccessUI(); }; 
}
function applyAccessPermissions(){
  const p = getCurrentPermissions();
  const noLogin = !p;
  const set = (sel, disabled)=>document.querySelectorAll(sel).forEach(el=>{ el.disabled = !!disabled; el.classList.toggle('is-muted', !!disabled); el.setAttribute('aria-disabled', disabled?'true':'false'); });
  set('#btnNovoRecurso,#btnNovaAtividade,#btnSalvarRecurso,#btnSalvarAtividade', noLogin || !p.plan);
  set('#btnBaselineExcluir,.btn.danger', noLogin || !(p.delete || p.admin));
  set('#btnExportCSV,#btnExportXLS,#btnExportPBI,#btnExportAtrasadasMes,#btnHistAll,#btnBackup,#btnExportPDF,#btnExportExecIndicators,#btnExportExecPdf,#btnExportComparacao', noLogin || !p.export);
  set('#btnPickBDFile,#btnReauthBD', noLogin || !p.db);
  document.querySelectorAll('[data-tab="audit"]').forEach(el=>{ el.style.display = (!noLogin && p.audit) ? '' : 'none'; });
  const cfg = document.getElementById('rpUsersConfig'); if(cfg) cfg.style.display = p && p.manageUsers ? '' : 'none';
}
function getRoleOptions(selected){ return ACCESS_PROFILES.map(x=>`<option ${x===selected?'selected':''}>${escHTML(x)}</option>`).join(''); }
function renderUsersConfigUI(){
  let host = document.getElementById('rpUsersConfig');
  if(!host){
    const dbPanel = document.querySelector('#tab-db section.panel');
    if(!dbPanel) return;
    host = document.createElement('div'); host.id='rpUsersConfig'; host.className='card rp-users-config';
    const after = document.getElementById('calendarOvertimeCard');
    dbPanel.insertBefore(host, after || null);
  }
  const p = getCurrentPermissions();
  host.style.display = p && p.manageUsers ? '' : 'none';
  const rows = (systemUsers || []).map(u=>`<tr><td>${escHTML(u.matricula)}</td><td>${escHTML(u.nome)}</td><td>${escHTML(u.perfil)}</td><td>${u.ativo?'Ativo':'Inativo'}</td><td>${u.trocarPinNoPrimeiroAcesso?'Pendente':'Não'}</td><td><button class="btn" data-rp-reset-pin="${escHTML(u.matricula)}">Resetar PIN</button> <button class="btn" data-rp-toggle-user="${escHTML(u.matricula)}">${u.ativo?'Inativar':'Ativar'}</button></td></tr>`).join('') || '<tr><td colspan="5" class="muted">Nenhum usuário cadastrado.</td></tr>';
  host.innerHTML = `<h3>Configuração de Usuários e Perfis</h3><p class="muted small">Disponível somente para administradores. O usuário comum continua podendo conectar e reautorizar o banco quando possuir perfil ativo.</p><div class="rp-users-form"><label>Matrícula<input id="rpUserMatricula" inputmode="numeric"></label><label>Nome<input id="rpUserNome"></label><label>Perfil<select id="rpUserPerfil">${getRoleOptions('Executor')}</select></label><label>PIN inicial<input id="rpUserPin" type="password" inputmode="numeric" placeholder="Mín. 4 dígitos"></label><button class="btn primary" id="rpAddUserBtn" type="button">Salvar usuário</button></div><div class="table-wrap" style="margin-top:10px"><table class="rp-users-table"><thead><tr><th>Matrícula</th><th>Nome</th><th>Perfil</th><th>Status</th><th>Troca PIN</th><th>Ações</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  document.getElementById('rpAddUserBtn').onclick = saveUserFromConfig;
  host.querySelectorAll('[data-rp-toggle-user]').forEach(b=>b.onclick=()=>toggleSystemUser(b.getAttribute('data-rp-toggle-user')));
  host.querySelectorAll('[data-rp-reset-pin]').forEach(b=>b.onclick=()=>resetSystemUserPin(b.getAttribute('data-rp-reset-pin')));
}
async function saveUserFromConfig(){
  if(!isAccessAdmin()) return alert('Apenas administradores podem configurar usuários.');
  const matricula = normalizeMatricula(document.getElementById('rpUserMatricula')?.value || '');
  const nome = String(document.getElementById('rpUserNome')?.value || '').trim();
  const perfil = String(document.getElementById('rpUserPerfil')?.value || 'Executor');
  const pin = String(document.getElementById('rpUserPin')?.value || '');
  if(!matricula || !nome || !ACCESS_PROFILES.includes(perfil)) return alert('Informe matrícula, nome e perfil válido.');
  let user = (systemUsers || []).find(u=>u.matricula===matricula);
  if(!user && pin.length < 4) return alert('Para novo usuário, informe PIN inicial com pelo menos 4 dígitos.');
  const now=Date.now();
  const isNew = !user;
  if(!user){ user = { matricula, createdAt:now, createdBy:accessSession?.matricula || '' }; systemUsers.push(user); }
  user.nome=nome; user.perfil=perfil; user.ativo=true; user.updatedAt=now;
  if(pin){ user.pinHash = await hashUserPin(matricula,pin); user.trocarPinNoPrimeiroAcesso = true; }
  if(isNew && !pin) user.trocarPinNoPrimeiroAcesso = true;
  systemUsers = normalizeSystemUsers(systemUsers);
  recordAuditEvent(isNew ? 'USUARIO_CRIADO' : 'USUARIO_ATUALIZADO', { entityType:'usuario_sistema', entityId:user.matricula, entityLabel:user.nome, reason:isNew ? 'Usuário criado pelo administrador.' : 'Usuário atualizado pelo administrador.' });
  renderUsersConfigUI(); saveBDDebounced(); showToast('Usuário salvo', `${nome} foi atualizado.`, 'success', 3500);
}
function toggleSystemUser(matricula){
  if(!isAccessAdmin()) return;
  const u=(systemUsers||[]).find(x=>x.matricula===matricula); if(!u) return;
  if(accessSession && accessSession.matricula === u.matricula && u.ativo) return alert('Não é permitido inativar o usuário logado.');
  u.ativo=!u.ativo; u.updatedAt=Date.now(); recordAuditEvent('USUARIO_STATUS', { entityType:'usuario_sistema', entityId:u.matricula, entityLabel:u.nome||u.matricula, reason:u.ativo?'Usuário ativado.':'Usuário inativado.' }); renderUsersConfigUI(); saveBDDebounced();
}
async function resetSystemUserPin(matricula){
  if(!isAccessAdmin()) return;
  const u=(systemUsers||[]).find(x=>x.matricula===matricula); if(!u) return;
  const pin = prompt(`Informe o novo PIN para ${u.nome || u.matricula} (mín. 4 dígitos):`);
  if(!pin) return;
  if(String(pin).length < 4) return alert('PIN deve ter pelo menos 4 dígitos.');
  u.pinHash = await hashUserPin(u.matricula, pin); u.trocarPinNoPrimeiroAcesso = true; u.updatedAt=Date.now(); recordAuditEvent('PIN_RESETADO', { entityType:'usuario_sistema', entityId:u.matricula, entityLabel:u.nome||u.matricula, reason:'PIN redefinido pelo administrador. Troca obrigatória no próximo acesso.' }); saveBDDebounced(); showToast('PIN redefinido', 'Informe o PIN temporário ao usuário; ele deverá trocar no próximo acesso.', 'success', 4500);
}
function ingestParsedAccess(parsed){
  systemUsers = normalizeSystemUsers(parsed && parsed.usuariosSistema ? parsed.usuariosSistema : []);
  auditEvents = normalizeAuditEvents(parsed && parsed.auditEvents ? parsed.auditEvents : auditEvents || []);
  saveLS('rp_audit_events_v1', auditEvents);
  restoreAccessSession();
  renderUsersConfigUI();
  applyAccessPermissions();
  renderAccessUI();
}
function bootAccessGate(){
  if(accessGateReady) return;
  accessGateReady = true;
  ensureAccessStyles();
  restoreAccessSession();
  renderUsersConfigUI();
  applyAccessPermissions();
  renderAccessUI();
}
function hasAdminPassword(){ return !!localStorage.getItem(LS_ADMIN_HASH); }
async function verifyAdminPassword(password){
  const stored = localStorage.getItem(LS_ADMIN_HASH) || '';
  if(!stored) return false;
  return (await sha256Text('rp-recursos-admin|' + String(password || ''))) === stored;
}
async function setAdminPassword(password){
  if(String(password || '').length < 4) throw new Error('A senha deve ter pelo menos 4 caracteres.');
  localStorage.setItem(LS_ADMIN_HASH, await sha256Text('rp-recursos-admin|' + String(password || '')));
}
function updateAdminAuthUI(){
  const box = document.getElementById('adminAuthBox');
  const panel = document.getElementById('adminRestrictedPanel');
  const msg = document.getElementById('adminAuthMsg');
  const confirmWrap = document.getElementById('adminPasswordConfirmLabel');
  const btn = document.getElementById('btnAdminLogin');
  if(!box || !panel) return;
  const firstSetup = !hasAdminPassword();
  if(confirmWrap) confirmWrap.style.display = firstSetup ? '' : 'none';
  if(btn) btn.textContent = firstSetup ? 'Criar senha e acessar' : 'Acessar';
  box.style.display = adminUnlocked ? 'none' : '';
  panel.style.display = adminUnlocked ? '' : 'none';
  if(msg) msg.textContent = firstSetup ? 'Primeiro acesso: defina uma senha local para restringir este cadastro.' : 'Informe a senha administrativa para liberar o cadastro.';
  if(adminUnlocked){ renderCalendarUI(); renderOvertimeUI(); }
}

function renderOvertimeUI(){
  const sel = document.getElementById('overtimeResource');
  const listEl = document.getElementById('overtimeList');
  if(sel){
    const cur = sel.value;
    sel.innerHTML = '<option value="">Selecione...</option>' + (resources || []).filter(r=>r && r.ativo && !r.deletedAt).slice().sort((a,b)=>String(a.nome||'').localeCompare(String(b.nome||''), undefined, {sensitivity:'base'})).map(r=>`<option value="${escAttr(r.id)}">${escHTML(r.nome)}</option>`).join('');
    if(cur) sel.value = cur;
  }
  if(listEl){
    const rows = getOvertimeEntries();
    if(!rows.length){ listEl.innerHTML = '<div class="muted small">Nenhuma hora extra cadastrada.</div>'; return; }
    const resourceName = (id)=>((resources||[]).find(r=>r.id===id)?.nome || id || '');
    const makeOvertimeKey = (x)=>[x.resourceId, x.date, Number(x.hours||0).toFixed(2), x.status || '', x.reason || ''].join('|');
    listEl.innerHTML = `<table class="tbl"><thead><tr><th>Data</th><th>Recurso</th><th>Horas</th><th>Status</th><th>Justificativa</th><th>Ação</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${escHTML(x.date)}</td><td>${escHTML(resourceName(x.resourceId))}</td><td>${Number(x.hours||0).toFixed(2)}</td><td>${escHTML(x.status)}</td><td>${escHTML(x.reason||'')}</td><td><button class="btn danger btn-del-overtime" data-id="${escAttr(x.id)}" data-key="${escAttr(makeOvertimeKey(x))}">Remover</button></td></tr>`).join('')}</tbody></table>`;
    listEl.querySelectorAll('.btn-del-overtime').forEach(btn=>{
      btn.onclick = async ()=>{
        const id = btn.getAttribute('data-id') || '';
        const key = btn.getAttribute('data-key') || '';
        const next = getOvertimeEntries().filter(x=> x.id !== id && makeOvertimeKey(x) !== key);
        setOvertimeEntries(next);
        try{ if(bdHandle) await saveBD(); }catch(_){ }
      };
    });
  }
}

function bindAdminRestrictedUI(){
  const btnLogin = document.getElementById('btnAdminLogin');
  const btnLogout = document.getElementById('btnAdminLogout');
  const btnChange = document.getElementById('btnAdminChangePassword');
  const btnAddOvertime = document.getElementById('btnAddOvertime');
  if(btnLogin && !btnLogin.__bound){
    btnLogin.__bound = true;
    btnLogin.onclick = async ()=>{
      const pass = document.getElementById('adminPassword')?.value || '';
      const conf = document.getElementById('adminPasswordConfirm')?.value || '';
      const msg = document.getElementById('adminAuthMsg');
      try{
        if(!hasAdminPassword()){
          if(pass !== conf){ if(msg) msg.textContent = 'A confirmação da senha não confere.'; return; }
          await setAdminPassword(pass);
          adminUnlocked = true;
        }else{
          if(!(await verifyAdminPassword(pass))){ if(msg) msg.textContent = 'Senha inválida.'; return; }
          adminUnlocked = true;
        }
        const p = document.getElementById('adminPassword'); if(p) p.value = '';
        const c = document.getElementById('adminPasswordConfirm'); if(c) c.value = '';
        updateAdminAuthUI();
      }catch(e){ if(msg) msg.textContent = e.message || 'Falha ao autenticar.'; }
    };
  }
  if(btnLogout && !btnLogout.__bound){ btnLogout.__bound = true; btnLogout.onclick = ()=>{ adminUnlocked = false; updateAdminAuthUI(); }; }
  if(btnChange && !btnChange.__bound){
    btnChange.__bound = true;
    btnChange.onclick = async ()=>{
      const np = prompt('Informe a nova senha administrativa local:');
      if(np == null) return;
      try{ await setAdminPassword(np); alert('Senha alterada.'); }catch(e){ alert(e.message || 'Não foi possível alterar a senha.'); }
    };
  }
  if(btnAddOvertime && !btnAddOvertime.__bound){
    btnAddOvertime.__bound = true;
    btnAddOvertime.onclick = ()=>{
      const resourceId = document.getElementById('overtimeResource')?.value || '';
      const date = normalizeDateField(document.getElementById('overtimeDate')?.value || '');
      const hours = Number(String(document.getElementById('overtimeHours')?.value || '').replace(',', '.'));
      const status = String(document.getElementById('overtimeStatus')?.value || 'aprovada').toLowerCase();
      const reason = String(document.getElementById('overtimeReason')?.value || '').trim();
      if(!resourceId){ alert('Selecione o recurso.'); return; }
      if(!date){ alert('Informe a data.'); return; }
      if(!isFinite(hours) || hours <= 0 || hours > 24){ alert('Informe horas extras entre 0,25 e 24 horas.'); return; }
      if(!reason){ alert('Informe uma justificativa.'); return; }
      setOvertimeEntries([...getOvertimeEntries(), { id: uuid(), resourceId, date, hours, status, reason, createdAt: Date.now(), createdBy: currentUser || '' }]);
      const h=document.getElementById('overtimeHours'); if(h) h.value='';
      const r=document.getElementById('overtimeReason'); if(r) r.value='';
    };
  }
  renderCalendarUI();
  renderOvertimeUI();
}

function normalizeHolidayEntry(f){
  if(!f) return null;
  if(typeof f === 'string'){
    const line = f.trim();
    if(!line) return null;
    const [date, ...rest] = line.split(/\s+/);
    const ymd = normalizeDateField ? normalizeDateField(date) : date;
    return ymd ? { date: ymd, legend: rest.join(' ').trim() } : null;
  }
  const date = normalizeDateField(f.date || f.data || f.dia || '');
  if(!date) return null;
  return {
    date,
    legend: String(f.legend || f.nome || f.descricao || '').trim(),
    tipo: String(f.tipo || 'geral').trim(),
    abrangencia: String(f.abrangencia || 'todos').trim(),
    departamentos: Array.isArray(f.departamentos) ? f.departamentos : [],
    usuarios: Array.isArray(f.usuarios) ? f.usuarios : [],
    ativo: f.ativo === false ? false : true
  };
}

function normalizeHolidayList(list){
  const seen = new Set();
  return (Array.isArray(list) ? list : []).map(normalizeHolidayEntry).filter(Boolean).filter(f=>{
    if(f.ativo === false) return false;
    const key = `${f.date}|${String(f.legend||'').toLowerCase()}`;
    if(seen.has(key)) return false;
    seen.add(key);
    return true;
  }).sort((a,b)=>String(a.date).localeCompare(String(b.date)) || String(a.legend||'').localeCompare(String(b.legend||'')));
}

function getCalendarHolidays(){
  if(typeof window !== 'undefined' && typeof window.getFeriados === 'function'){
    return normalizeHolidayList(window.getFeriados() || []);
  }
  return [];
}

function setCalendarHolidays(list){
  const clean = normalizeHolidayList(list);
  if(typeof window !== 'undefined' && typeof window.setFeriados === 'function') window.setFeriados(clean);
  renderCalendarUI();
  try{ renderAll(); }catch(_){ }
  try{ saveBDDebounced(); }catch(_){ }
}

function parseHolidayLines(text){
  return String(text || '').split(/\r?\n/).map(line=>normalizeHolidayEntry(line)).filter(Boolean);
}

function getHolidayForDate(dateOrYmd, resource){
  const ymd = typeof dateOrYmd === 'string' ? normalizeDateField(dateOrYmd) : toYMD(dateOrYmd);
  const list = getCalendarHolidays();
  return list.find(f => f.date === ymd && isHolidayApplicableToResource(f, resource)) || null;
}

function isHolidayApplicableToResource(holiday, resource){
  if(!holiday || holiday.ativo === false) return false;
  const scope = String(holiday.abrangencia || 'todos').toLowerCase();
  if(scope === 'todos' || scope === 'geral' || !resource) return true;
  if(scope.includes('usuario') && Array.isArray(holiday.usuarios)) return holiday.usuarios.map(String).includes(String(resource.id));
  if(scope.includes('depart') && Array.isArray(holiday.departamentos)) return holiday.departamentos.map(x=>String(x).toLowerCase()).includes(String(resource.departamento || resource.tipo || '').toLowerCase());
  return true;
}

function getDayContext(dateOrYmd, resource){
  const d = typeof dateOrYmd === 'string' ? fromYMD(normalizeDateField(dateOrYmd)) : dateOrYmd;
  const ymd = toYMD(d);
  const isWeekend = d.getDay() === 0 || d.getDay() === 6;
  const holiday = getHolidayForDate(ymd, resource);
  const isHoliday = !!holiday;
  const nonWorking = isWeekend || isHoliday;
  const overtimeApproved = getApprovedOvertimeHours(resource, ymd);
  return { ymd, isWeekend, isHoliday, holiday, nonWorking, overtimeApproved };
}

function renderCalendarUI(){
  const txt = document.getElementById('feriadosTexto');
  const listEl = document.getElementById('feriadosLista');
  const toggleHol = document.getElementById('toggleDestacarFeriados');
  const toggleFds = document.getElementById('toggleDestacarFds');
  if(toggleHol) toggleHol.checked = calendarPrefs.destacarFeriados !== false;
  if(toggleFds) toggleFds.checked = calendarPrefs.destacarFds !== false;
  const holidays = getCalendarHolidays();
  if(txt && document.activeElement !== txt){
    txt.value = holidays.map(f => `${f.date} ${f.legend || ''}`.trim()).join('\n');
  }
  if(listEl){
    if(!holidays.length){
      listEl.innerHTML = '<div class="muted small">Nenhum feriado cadastrado.</div>';
    }else{
      listEl.innerHTML = `<table class="tbl"><thead><tr><th>Data</th><th>Descrição</th><th>Ação</th></tr></thead><tbody>${holidays.map(f=>`<tr><td>${escHTML(f.date)}</td><td><span class="holiday-dot"></span>${escHTML(f.legend || 'Feriado')}</td><td><button class="btn danger btn-del-feriado" data-date="${escAttr(f.date)}" data-legend="${escAttr(f.legend || '')}">Remover</button></td></tr>`).join('')}</tbody></table>`;
      listEl.querySelectorAll('.btn-del-feriado').forEach(btn=>{
        btn.onclick = ()=>{
          const date = btn.getAttribute('data-date');
          const legend = btn.getAttribute('data-legend') || '';
          setCalendarHolidays(holidays.filter(f=>!(f.date===date && String(f.legend||'')===legend)));
        };
      });
    }
  }
}

function bindCalendarUI(){
  const btnAdd = document.getElementById('btnAddFeriado');
  const btnImport = document.getElementById('btnImportFeriadosTexto');
  const toggleHol = document.getElementById('toggleDestacarFeriados');
  const toggleFds = document.getElementById('toggleDestacarFds');
  if(btnAdd && !btnAdd.__bound){
    btnAdd.__bound = true;
    btnAdd.onclick = ()=>{
      const dataEl = document.getElementById('feriadoData');
      const legEl = document.getElementById('feriadoLegenda');
      const date = normalizeDateField(dataEl?.value || '');
      if(!date){ alert('Informe a data do feriado.'); return; }
      const legend = String(legEl?.value || '').trim();
      setCalendarHolidays([...getCalendarHolidays(), {date, legend}]);
      if(dataEl) dataEl.value = '';
      if(legEl) legEl.value = '';
    };
  }
  if(btnImport && !btnImport.__bound){
    btnImport.__bound = true;
    btnImport.onclick = ()=>{
      const txt = document.getElementById('feriadosTexto');
      setCalendarHolidays(parseHolidayLines(txt?.value || ''));
    };
  }
  const bindToggle = (el, key)=>{
    if(el && !el.__bound){
      el.__bound = true;
      el.onchange = ()=>{
        calendarPrefs[key] = !!el.checked;
        saveLS(LS_CALENDAR_PREFS, calendarPrefs);
        renderAll();
      };
    }
  };
  bindToggle(toggleHol, 'destacarFeriados');
  bindToggle(toggleFds, 'destacarFds');
  renderCalendarUI();
  try{ if(window.bindTeamAnnualCapacityUI) window.bindTeamAnnualCapacityUI(); if(window.renderTeamAnnualCapacity) window.renderTeamAnnualCapacity(); }catch(_){ }
}

if(typeof window !== 'undefined'){
  window.renderCalendarUI = renderCalendarUI;
  window.getHorasExtrasData = getOvertimeEntries;
  window.setHorasExtrasData = setOvertimeEntries;
}

// ===== Event Log e Snapshot =====
// Carrega log de eventos e snapshot, se existirem.  O log é uma lista de
// objetos {ts, type, action, id, data}.  O snapshot é um objeto
// {ts, resources, activities, trails}.  Essas estruturas são usadas para
// reconstruir o estado no início da aplicação e reduzir conflitos de
// concorrência.
let eventLog = loadLS(LS_LOG, []);
let snapshot = loadLS(LS_SNAP, null);

/**
 * Reinicia o log de eventos e o snapshot. Deve ser chamado quando o usuário
 * aponta um novo arquivo de Banco de Dados como origem, para evitar que
 * eventos antigos (antes de selecionar o BD) sejam reaplicados e causem
 * "dados fantasmas" ao reconstruir o estado. Zera o eventLog e o
 * snapshot persistidos em localStorage.
 */
function resetEventLogAndSnapshot(){
  try {
    eventLog = [];
    saveLS(LS_LOG, eventLog);
  } catch(e){ console.warn('Erro ao zerar o log de eventos', e); }
  try {
    snapshot = null;
    saveLS(LS_SNAP, snapshot);
  } catch(e){ console.warn('Erro ao zerar snapshot', e); }
}

/**
 * Marca o estado atual como baseline persistido.
 *
 * Use esta função quando o estado em memória já reflete fielmente o conteúdo
 * persistido no BD/FSA. Nesses casos, manter o eventLog anterior faria com que
 * updates/deletes locais já salvos fossem reaplicados sobre um snapshot mais
 * novo, podendo ressuscitar registros excluídos em outra sessão.
 */
function adoptCurrentStateAsPersistedBaseline(){
  try {
    snapshot = {
      ts: Date.now(),
      resources: (resources || []).map(r => ({ ...r })),
      activities: (activities || []).map(a => ({ ...a })),
      trails: { ...(trails || {}) }
    };
    saveLS(LS_SNAP, snapshot);
    eventLog = [];
    saveLS(LS_LOG, eventLog);
  } catch(e) {
    console.error('Erro ao atualizar baseline persistido', e);
  }
}

/**
 * Aplica eventos da lista eventLog sobre os estados base (resources/activities)
 * e substitui os arrays globais.  Caso exista um snapshot válido, ele é
 * utilizado como base; caso contrário, os arrays atuais de resources e
 * activities são usados como base.  Ao final, os estados globais e o
 * localStorage são atualizados para refletir o resultado.
 */
function loadSnapshotAndEvents() {
  try {
    let baseRes, baseActs, baseTrails;
    if (snapshot && Array.isArray(snapshot.resources) && Array.isArray(snapshot.activities)) {
      baseRes = snapshot.resources.map(r => ({ ...r }));
      baseActs = snapshot.activities.map(a => ({ ...a }));
      baseTrails = { ...(snapshot.trails || {}) };
    } else {
      // Caso não haja snapshot, utiliza os arrays atuais como base
      baseRes = (resources || []).map(r => ({ ...r }));
      baseActs = (activities || []).map(a => ({ ...a }));
      baseTrails = { ...(trails || {}) };
    }
    // Mapas para acesso rápido por id
    const resMap = {};
    baseRes.forEach(r => { resMap[r.id] = { ...r }; });
    const actMap = {};
    baseActs.forEach(a => { actMap[a.id] = { ...a }; });

    // Ordena o log por timestamp crescente para garantir reprodutibilidade
    const sortedEvents = (eventLog || []).slice().sort((a, b) => (a.ts || 0) - (b.ts || 0));
    for (const ev of sortedEvents) {
      if (!ev || !ev.type || !ev.action) continue;
      if (ev.type === 'resource') {
        if (ev.action === 'create' || ev.action === 'update') {
          resMap[ev.id] = { ...ev.data };
        } else if (ev.action === 'delete') {
          if (resMap[ev.id]) {
            resMap[ev.id] = { ...resMap[ev.id], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (resMap[ev.id].version || 0) + 1 };
          }
          // marcas todas as atividades relacionadas como deletadas
          for (const aid in actMap) {
            if (actMap[aid] && actMap[aid].resourceId === ev.id) {
              actMap[aid] = { ...actMap[aid], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (actMap[aid].version || 0) + 1 };
            }
          }
        }
      } else if (ev.type === 'activity') {
        if (ev.action === 'create' || ev.action === 'update') {
          actMap[ev.id] = { ...ev.data };
        } else if (ev.action === 'delete') {
          if (actMap[ev.id]) {
            actMap[ev.id] = { ...actMap[ev.id], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (actMap[ev.id].version || 0) + 1 };
          }
        }
      }
    }
    resources = Object.values(resMap);
    activities = hydrateLoadedActivities(Object.values(actMap));
    trails = baseTrails;
    // Persistir o estado reconstruído
    saveLS(LS.res, resources);
    saveLS(LS.act, activities);
    saveLS(LS.trail, trails);
  } catch (e) {
    console.error('Erro ao aplicar snapshot/eventos', e);
  }
}

/**
 * Registra um evento no log.  O evento inclui: timestamp atual,
 * tipo (resource|activity), ação (create|update|delete), id do item
 * e dados (objeto com os campos relevantes ou no caso de delete
 * apenas deletedAt).  Após registrar, salva o log no localStorage e
 * invoca maybeSaveSnapshot() para decidir se um snapshot deve ser gerado.
 * @param {string} type Tipo da entidade ('resource' ou 'activity')
 * @param {string} action Ação realizada ('create','update','delete')
 * @param {string} id ID do item
 * @param {object} data Dados atuais do item (ou {deletedAt:ts} para delete)
 */
function recordEvent(type, action, id, data) {
  try {
    const ev = { ts: Date.now(), type, action, id, data };
    eventLog.push(ev);
    saveLS(LS_LOG, eventLog);
    maybeSaveSnapshot();
  } catch (e) {
    console.error('Erro ao registrar evento', e);
  }
}

/**
 * Gera um snapshot do estado atual se o log atingir um limiar definido.
 * O snapshot captura resources, activities e trails e armazena em LS_SNAP.
 * Após salvar o snapshot, o log é esvaziado para evitar crescimento
 * descontrolado.  Ajuste o threshold conforme necessário.
 */
function maybeSaveSnapshot() {
  const threshold = 30;
  if ((eventLog || []).length >= threshold) {
    try {
      snapshot = {
        ts: Date.now(),
        resources: (resources || []).map(r => ({ ...r })),
        activities: (activities || []).map(a => ({ ...a })),
        trails: { ...(trails || {}) }
      };
      saveLS(LS_SNAP, snapshot);
      // Limpa o log
      eventLog = [];
      saveLS(LS_LOG, eventLog);
    } catch (e) {
      console.error('Erro ao salvar snapshot', e);
    }
  }
}

// ===== IndexedDB helpers (tiny) =====
function idbProm(req){return new Promise((res,rej)=>{req.onsuccess=()=>res(req.result); req.onerror=()=>rej(req.error);})}
function idbOpen(name, store){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(name,1);
    req.onupgradeneeded=()=>{const db=req.result; if(!db.objectStoreNames.contains(store)) db.createObjectStore(store);};
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
  });
}
async function idbSet(dbname, store, key, value){
  const db=await idbOpen(dbname, store);
  const tx=db.transaction(store,"readwrite"); const st=tx.objectStore(store);
  const req=st.put(value, key); await idbProm(req); await idbProm(tx.done||tx); db.close();
}
async function idbGet(dbname, store, key){
  const db=await idbOpen(dbname, store);
  const tx=db.transaction(store,"readonly"); const st=tx.objectStore(store);
  const req=st.get(key); const val=await idbProm(req); await idbProm(tx.done||tx); db.close(); return val;
}

// ===== File System Access (Salvar na pasta) =====
const DATAFILES = {
  [LS.res]: "resources.json",
  [LS.act]: "activities.json",
  [LS.trail]: "trails.json"
};
const FSA_DB="planner_fs";
const FSA_STORE="handles";
let dirHandle=null;

// === Observador (watcher) para sincronização multiusuário ===
let fsaWatchTimer = null;
let fsaLastSeen = {};

function startFsaWatcher() {
  if (fsaWatchTimer) clearInterval(fsaWatchTimer);
  fsaLastSeen = {};
  if (!dirHandle) return;
  fsaWatchTimer = setInterval(async () => {
    try {
      for (const key of [LS.res, LS.act, LS.trail]) {
        const fileName = DATAFILES[key];
        let fh;
        try {
          fh = await dirHandle.getFileHandle(fileName, { create: false });
        } catch (e) {
          continue; 
        }
        const file = await fh.getFile();
        const lm = file.lastModified;
        if (!fsaLastSeen[fileName] || lm > fsaLastSeen[fileName]) {
          fsaLastSeen[fileName] = lm;
          const text = await file.text();
          let changed = false;
          try {
            let data;
            if (key === LS.trail) {
              data = JSON.parse(text || '{}');
            } else {
              data = JSON.parse(text || '[]');
            }
            if (key === LS.res) {
              // Comparação leve: usa updatedAt/version em vez de JSON.stringify completo
              const newSig = data.map(r => `${r.id}:${r.updatedAt || 0}:${r.deletedAt || ''}`).join('|');
              const curSig = resources.map(r => `${r.id}:${r.updatedAt || 0}:${r.deletedAt || ''}`).join('|');
              if (newSig !== curSig) {
                resources = data;
                saveLS(LS.res, resources);
                changed = true;
              }
            } else if (key === LS.act) {
              const newSig = data.map(a => `${a.id}:${a.updatedAt || 0}:${a.status}:${a.deletedAt || ''}`).join('|');
              const curSig = activities.map(a => `${a.id}:${a.updatedAt || 0}:${a.status}:${a.deletedAt || ''}`).join('|');
              if (newSig !== curSig) {
                activities = data;
                saveLS(LS.act, activities);
                changed = true;
              }
            } else if (key === LS.trail) {
              // Trail não tem updatedAt — compara contagem por entidade (rápido e suficiente)
              const newSig = Object.entries(data).map(([k,v]) => `${k}:${Array.isArray(v)?v.length:0}`).sort().join('|');
              const curSig = Object.entries(trails).map(([k,v]) => `${k}:${Array.isArray(v)?v.length:0}`).sort().join('|');
              if (newSig !== curSig) {
                trails = data;
                saveLS(LS.trail, trails);
                changed = true;
              }
            }
          } catch (e) {
            // JSON inválido: ignorar alteração
          }
          if (changed) {
            try {
              adoptCurrentStateAsPersistedBaseline();
            } catch (e) {
              console.error('Erro ao atualizar baseline após mudança no FSA:', e);
            }
            renderAll();
            const fmtTime = new Date(lm).toLocaleTimeString();
            updateFolderStatus('Atualizado por outra sessão às ' + fmtTime);
            showToast('BD atualizado', `Outra sessão salvou alterações às ${fmtTime}. Dados recarregados.`, 'info', 6000);
          }
        }
      }
    } catch (e) {
      // falhas silenciosas
    }
  }, 3000);
}

// === Banco de Dados (BD) ===
let bdHandle = null;
let bdFileExt = '';
let bdFileName = '';
let _saveBDTimer = null;

// === Rastreio simples da última gravação do BD ===
// Este valor é atualizado sempre que um salvamento ocorre ou quando o watcher detecta
// uma mudança externa. É utilizado para detectar se o arquivo foi modificado por
// outra sessão para evitar sobrescrever alterações recentes.
if (typeof window !== 'undefined') {
  window.__bdLastWrite = 0;
}

// === Observador para o arquivo de Banco de Dados (Excel/CSV) ===
let bdWatchTimer = null;
let bdLastSeen = null;

function startBDWatcher() {
  if (bdWatchTimer) clearInterval(bdWatchTimer);
  bdLastSeen = null;
  if (!bdHandle) return;
  bdWatchTimer = setInterval(async () => {
    try {
      const file = await bdHandle.getFile();
      const lm = file.lastModified;
      if (!bdLastSeen || lm > bdLastSeen) {
        bdLastSeen = lm;
        let parsed;
        if (bdFileExt === 'csv') {
      try {
        const existingFile = await bdHandle.getFile();
        const existingText = await existingFile.text();
        const persisted = parseCSVBDUnico(existingText);
        comments = mergeComentarios(persisted.comments || [], comments || []);
        auditEvents = normalizeAuditEvents([...(persisted.auditEvents || []), ...(auditEvents || [])]);
        saveLS('rp_audit_events_v1', auditEvents);
        rebuildCommentsIndex();
        syncAllActivityCommentFields();
      } catch(_e){}
          const txt = await file.text();
          parsed = parseCSVBDUnico(txt);
        } else {
          const txt = await file.text();
          parsed = parseHTMLBDTables(txt);
        }
        const newResources = (parsed.recursos || []).map(coerceResource);
        const newActivities = hydrateLoadedActivities(parsed.atividades || []);
        const newHoras = parsed.horas || [];
        const newCfg = parsed.cfg || [];
        const newFeriados = parsed.feriados || [];
        const newHorasExtras = parsed.horasExtras || [];
        const newTrails = {};
        (parsed.historico || []).forEach(h => {
          const id = h.activityId;
          if (!id) return;
          if (!newTrails[id]) newTrails[id] = [];
          // Suporta eventos de histórico além de alteração de datas.
          // O campo 'legend' pode carregar um JSON com metadados do evento.
          let meta = null;
          try {
            if (h.legend && String(h.legend).trim().startsWith('{')) meta = JSON.parse(h.legend);
          } catch(e) { meta = null; }
          newTrails[id].push({
            ts: h.timestamp,
            type: (meta && meta.type) ? meta.type : (h.type || h.tipoEvento || 'ALTERACAO_DATAS'),
            entityType: (meta && meta.entityType) ? meta.entityType : (h.entityType || h.entidadeTipo || 'atividade'),
            oldInicio: h.oldInicio,
            oldFim: h.oldFim,
            newInicio: h.newInicio,
            newFim: h.newFim,
            oldResourceId: (meta && meta.oldResourceId) ? meta.oldResourceId : (h.oldResourceId || ''),
            oldResourceName: (meta && meta.oldResourceName) ? meta.oldResourceName : (h.oldResourceName || ''),
            newResourceId: (meta && meta.newResourceId) ? meta.newResourceId : (h.newResourceId || ''),
            newResourceName: (meta && meta.newResourceName) ? meta.newResourceName : (h.newResourceName || ''),
            justificativa: h.justificativa,
            user: h.user,
            legend: h.legend || ''
          });
        });

        let changed = false;
        // Comparação leve por updatedAt/version — evita JSON.stringify de arrays grandes
        const _resSig = (arr) => (arr||[]).map(r=>`${r.id}:${r.updatedAt||0}:${r.deletedAt||''}`).join('|');
        const _actSig = (arr) => (arr||[]).map(a=>`${a.id}:${a.updatedAt||0}:${a.status}:${a.deletedAt||''}`).join('|');
        const _trlSig = (obj) => Object.entries(obj||{}).map(([k,v])=>`${k}:${Array.isArray(v)?v.length:0}`).sort().join('|');
        if (_resSig(resources) !== _resSig(newResources)) {
          resources = newResources;
          saveLS(LS.res, resources);
          changed = true;
        }
        if (_actSig(activities) !== _actSig(newActivities)) {
          activities = newActivities;
          saveLS(LS.act, activities);
          changed = true;
        }
        if (_trlSig(trails) !== _trlSig(newTrails)) {
          trails = newTrails;
          saveLS(LS.trail, trails);
          changed = true;
        }

        let horasChanged = false;
        try {
          if (typeof window.getHorasExternosData === 'function' && typeof window.setHorasExternosData === 'function') {
            const curHoras = window.getHorasExternosData() || [];
            // Horas: compara por id+minutos+data (campos mutáveis relevantes)
            const _hSig = (arr) => (arr||[]).map(h=>`${h.id}:${h.date||''}:${h.minutos||0}`).sort().join('|');
            if (_hSig(curHoras) !== _hSig(newHoras)) {
              window.setHorasExternosData(newHoras);
              horasChanged = true;
            }
          }
        } catch(e) {}
        try {
          if (typeof window.getHorasExternosConfig === 'function' && typeof window.setHorasExternosConfig === 'function') {
            const curCfg = window.getHorasExternosConfig() || [];
            const _cfgSig = (arr) => (arr||[]).map(c=>`${c.id}:${c.horasDia||''}:${c.dias||''}`).sort().join('|');
            if (_cfgSig(curCfg) !== _cfgSig(newCfg)) {
              window.setHorasExternosConfig(newCfg);
              horasChanged = true;
            }
          }
        } catch(e) {}
        try {
          if (typeof window.getFeriados === 'function' && typeof window.setFeriados === 'function') {
            const curFeriados = window.getFeriados() || [];
            const _fSig = (arr) => (arr||[]).map(f=>f.date||'').sort().join('|');
            if (_fSig(curFeriados) !== _fSig(newFeriados)) {
              window.setFeriados(newFeriados);
              horasChanged = true;
            }
          }
        } catch(e) {}
        if (changed || horasChanged) {
          try {
            adoptCurrentStateAsPersistedBaseline();
          } catch (e) {
            console.error('Erro ao atualizar baseline após mudança no BD:', e);
          }
          renderAll();
          const _fmtLm = new Date(lm).toLocaleTimeString();
          updateBDStatus('BD recarregado após alteração externa às ' + _fmtLm);
          updateDBStatusBanner('synced');
          showToast('BD atualizado', `Outra sessão salvou alterações às ${_fmtLm}. Dados recarregados automaticamente.`, 'info', 7000);
          // Atualiza o rastreio de última modificação quando o arquivo é alterado por outra sessão
          if (typeof window !== 'undefined') {
            window.__bdLastWrite = lm;
          }
        }
      }
    } catch (e) {
      // Erros silenciosos
    }
  }, 3000);
}

// ===== Overlay de espera para gravação concorrente =====
let _bdWaitInterval = null;
// ===== Sistema de Toast (notificações visíveis) =====
(function _initToastSystem() {
  if (document.getElementById('rp-toast-container')) return;
  const style = document.createElement('style');
  style.textContent = `
    #rp-toast-container {
      position: fixed; bottom: 24px; right: 24px; z-index: 9999;
      display: flex; flex-direction: column; gap: 10px; pointer-events: none;
    }
    .rp-toast {
      display: flex; align-items: flex-start; gap: 10px;
      min-width: 280px; max-width: 420px;
      padding: 14px 16px; border-radius: 10px;
      font-family: system-ui, sans-serif; font-size: 14px; line-height: 1.45;
      box-shadow: 0 4px 20px rgba(0,0,0,0.18);
      pointer-events: all; cursor: default;
      animation: rp-toast-in 0.25s ease both;
    }
    .rp-toast.out { animation: rp-toast-out 0.3s ease both; }
    .rp-toast-icon { font-size: 18px; flex-shrink: 0; margin-top: 1px; }
    .rp-toast-body { flex: 1; }
    .rp-toast-title { font-weight: 600; margin-bottom: 2px; }
    .rp-toast-msg { opacity: .85; font-size: 13px; }
    .rp-toast-close { background: none; border: none; cursor: pointer; opacity: .5; font-size: 16px; padding: 0 0 0 6px; flex-shrink: 0; line-height: 1; }
    .rp-toast-close:hover { opacity: 1; }
    .rp-toast.info  { background: #1e40af; color: #fff; }
    .rp-toast.success { background: #15803d; color: #fff; }
    .rp-toast.warning { background: #92400e; color: #fff; }
    .rp-toast.error { background: #991b1b; color: #fff; }
    @keyframes rp-toast-in  { from { opacity:0; transform: translateX(40px); } to { opacity:1; transform: translateX(0); } }
    @keyframes rp-toast-out { from { opacity:1; transform: translateX(0);    } to { opacity:0; transform: translateX(40px); } }
  `;
  document.head.appendChild(style);
  const container = document.createElement('div');
  container.id = 'rp-toast-container';
  document.body.appendChild(container);
})();

/**
 * Exibe um toast visível no canto inferior direito.
 * @param {string} title  - Título em negrito
 * @param {string} msg    - Mensagem secundária
 * @param {'info'|'success'|'warning'|'error'} type
 * @param {number} duration - ms antes de fechar automaticamente (0 = manual)
 */
function showToast(title, msg, type = 'info', duration = 5000) {
  try {
    const icons = { info: 'ℹ️', success: '✅', warning: '⚠️', error: '❌' };
    const container = document.getElementById('rp-toast-container');
    if (!container) return;
    const el = document.createElement('div');
    el.className = `rp-toast ${type}`;

    const icon = document.createElement('span');
    icon.className = 'rp-toast-icon';
    icon.textContent = icons[type] || 'ℹ️';

    const body = document.createElement('div');
    body.className = 'rp-toast-body';

    const titleEl = document.createElement('div');
    titleEl.className = 'rp-toast-title';
    titleEl.textContent = String(title ?? '');
    body.appendChild(titleEl);

    if (msg) {
      const msgEl = document.createElement('div');
      msgEl.className = 'rp-toast-msg';
      msgEl.textContent = String(msg);
      body.appendChild(msgEl);
    }

    const closeBtn = document.createElement('button');
    closeBtn.className = 'rp-toast-close';
    closeBtn.type = 'button';
    closeBtn.title = 'Fechar';
    closeBtn.setAttribute('aria-label', 'Fechar notificação');
    closeBtn.textContent = '✕';

    el.appendChild(icon);
    el.appendChild(body);
    el.appendChild(closeBtn);

    const dismiss = () => {
      el.classList.add('out');
      el.addEventListener('animationend', () => el.remove(), { once: true });
    };
    closeBtn.onclick = dismiss;
    container.appendChild(el);
    if (duration > 0) setTimeout(dismiss, duration);
    return el;
  } catch (e) { /* não impede o fluxo principal */ }
}

// ===== Overlay de espera para gravação concorrente (redesenhado) =====
function showBDWaitOverlay(seconds, reason) {
  const overlay = document.getElementById('bdWaitOverlay');
  if (overlay) {
    overlay.style.display = 'flex';
    const countdownEl = document.getElementById('bdWaitCountdown');
    const reasonEl    = document.getElementById('bdWaitReason');
    const barEl       = document.getElementById('bdWaitBar');
    let remaining = seconds;
    if (countdownEl) countdownEl.textContent = remaining;
    if (reasonEl && reason) reasonEl.textContent = reason;
    // Barra de progresso: começa em 100% e esvazia linearmente
    if (barEl) { barEl.style.transition = 'none'; barEl.style.width = '100%'; }
    requestAnimationFrame(() => {
      if (barEl) { barEl.style.transition = `width ${seconds}s linear`; barEl.style.width = '0%'; }
    });
    if (_bdWaitInterval) clearInterval(_bdWaitInterval);
    _bdWaitInterval = setInterval(() => {
      remaining -= 1;
      if (countdownEl) countdownEl.textContent = remaining >= 0 ? remaining : 0;
      if (remaining <= 0) { clearInterval(_bdWaitInterval); _bdWaitInterval = null; }
    }, 1000);
    return;
  }
  // Fallback: toast se o overlay não existir no DOM
  showToast('Aguardando BD...', reason || `Outra sessão salvou recentemente. Aguardando ${seconds}s.`, 'warning', seconds * 1000 + 500);
}

function hideBDWaitOverlay() {
  const overlay = document.getElementById('bdWaitOverlay');
  if (!overlay) return;
  overlay.style.display = 'none';
  if (_bdWaitInterval) { clearInterval(_bdWaitInterval); _bdWaitInterval = null; }
}

/**
 * Lock robusto com retry e detecção de conflito real.
 *
 * Estratégia:
 * 1. Lê o lastModified atual do arquivo.
 * 2. Se foi modificado por OUTRA sessão nos últimos WAIT_WINDOW_MS, aguarda
 *    com exibição de contagem regressiva clara.
 * 3. Repete até a janela expirar — máximo MAX_RETRIES tentativas para
 *    evitar espera infinita em ambientes de rede com timestamps instáveis.
 * 4. Se esgotar tentativas, aborta a gravação para não escrever sem uma
 *    janela mínima de estabilidade.
 */
async function acquireBDLockIfBusy() {
  if (!bdHandle) return true;
  const WAIT_WINDOW_MS = 4000;
  const MAX_RETRIES = 8;
  let attempt = 0;
  let waitedAtLeastOnce = false;
  while (attempt < MAX_RETRIES) {
    try {
      const file = await bdHandle.getFile();
      const lm = file.lastModified;
      const now = Date.now();
      const modifiedByOther = lm > (window.__bdLastWrite || 0);
      const withinWindow = (now - lm) < WAIT_WINDOW_MS;
      if (modifiedByOther && withinWindow) {
        const remainingMs = WAIT_WINDOW_MS - (now - lm);
        const secondsLeft = Math.max(1, Math.ceil(remainingMs / 1000));
        showBDWaitOverlay(secondsLeft, 'O BD foi modificado por outra sessão recentemente. Aguardando uma janela segura para evitar sobrescrita.');
        waitedAtLeastOnce = true;
        window.__bdShowSaveConfirm = true;
        await new Promise(resolve => setTimeout(resolve, remainingMs + 200));
        hideBDWaitOverlay();
        attempt++;
        continue;
      }
      hideBDWaitOverlay();
      if (waitedAtLeastOnce) {
        showToast('Janela segura liberada', 'A gravação pode prosseguir.', 'info', 2200);
      }
      return true;
    } catch (e) {
      hideBDWaitOverlay();
      return false;
    }
  }
  hideBDWaitOverlay();
  console.warn('[rp] acquireBDLockIfBusy: máximo de tentativas atingido; gravação abortada para evitar escrita sem estabilidade mínima.');
  showToast('Fila de gravação ocupada', 'O banco continuou recebendo alterações durante toda a espera. Tente salvar novamente em instantes.', 'warning', 5000);
  return false;
}

async function hasFSA(){ return 'showDirectoryPicker' in window; }

async function fsaLoadHandle(){
  try{
    const h=await idbGet(FSA_DB,FSA_STORE,'dir');
    if(!h) return null;
    const perm = await h.queryPermission({mode:'readwrite'});
    if(perm==='granted') return h;
    const perm2 = await h.requestPermission({mode:'readwrite'});
    return perm2==='granted'?h:null;
  }catch(e){ console.warn('fsaLoadHandle error',e); return null; }
}


async function ensureBDWritePermission(options = {}){
  const request = options.request === true;
  const silent = options.silent === true;
  if (!bdHandle) {
    if (!silent) showToast('BD não selecionado', 'Selecione novamente o BD em "Selecionar BD (ler/gravar)".', 'warning', 5000);
    return false;
  }
  if (typeof bdHandle.queryPermission !== 'function') {
    return true;
  }
  try {
    let permission = await bdHandle.queryPermission({ mode: 'readwrite' });
    if (permission === 'granted') return true;
    if (!request) {
      if (!silent) {
        showToast('Reautorize a gravação', 'Clique em "Reautorizar gravação" na aba Banco de Dados para liberar escrita no BD apontado.', 'warning', 7000);
      }
      return false;
    }
    permission = await bdHandle.requestPermission({ mode: 'readwrite' });
    if (permission === 'granted') {
      try { await idbSet(FSA_DB, FSA_STORE, 'bd', bdHandle); } catch(_e) {}
      updateBDStatus('Gravação autorizada para ' + (bdFileName || 'BD'));
      updateDBStatusBanner('synced');
      showToast('Gravação autorizada', 'Permissão de escrita do BD foi confirmada.', 'success', 4500);
      return true;
    }
    updateBDStatus('Sem permissão de gravação no BD');
    updateDBStatusBanner('error');
    showToast('Permissão não concedida', 'O navegador não liberou a escrita. Selecione novamente o BD em "Selecionar BD (ler/gravar)".', 'error', 7000);
    return false;
  } catch (e) {
    console.warn('[rp] Falha ao verificar/solicitar permissão do BD:', e);
    updateBDStatus('Falha ao reautorizar gravação');
    updateDBStatusBanner('error');
    if (!silent) {
      showToast('Falha ao reautorizar gravação', e?.message || 'Selecione novamente o BD para renovar o acesso.', 'error', 8000);
    }
    return false;
  }
}

async function reauthorizeBDWritePermission(){
  if (!('showOpenFilePicker' in window)) {
    alert('Seu navegador não suporta permissão de gravação por arquivo. Use Chrome/Edge via http(s)://.');
    return false;
  }
  if (!bdHandle) {
    const pick = document.getElementById('btnPickBDFile');
    if (pick) pick.click();
    return false;
  }
  const ok = await ensureBDWritePermission({ request: true });
  if (!ok) {
    showToast('Selecione o BD novamente', 'Caso o navegador não mostre a caixa de permissão, use "Selecionar BD (ler/gravar)" para renovar o handle do arquivo.', 'warning', 8000);
  }
  return ok;
}

async function fsaPickFolder(){
  try{
    const h=await window.showDirectoryPicker();
    await idbSet(FSA_DB,FSA_STORE,'dir',h);
    dirHandle=h;
    updateFolderStatus();
    try {
      await loadAllFromFolder();
    } catch(e) { /* ignore */ }
    startFsaWatcher();
    alert('Pasta definida: '+(h.name||'(sem nome)'));
  }catch(e){ if(e&&e.name!=='AbortError') alert('Não foi possível selecionar a pasta: '+e.message); }
}

async function writeFile(handle, name, content){
  const fhandle = await handle.getFileHandle(name, {create:true});
  const ws = await fhandle.createWritable();
  await ws.write(content);
  await ws.close();
}

async function readFile(handle, name){
  const fhandle = await handle.getFileHandle(name, {create:false});
  const file = await fhandle.getFile();
  const text = await file.text();
  return text;
}

async function saveAllToFolder(){
  if(!dirHandle) return false;
  try{
    await writeFile(dirHandle, DATAFILES[LS.res], JSON.stringify(resources,null,2));
    await writeFile(dirHandle, DATAFILES[LS.act], JSON.stringify(activities,null,2));
    await writeFile(dirHandle, DATAFILES[LS.trail], JSON.stringify(trails,null,2));
    updateFolderStatus('Salvo em '+new Date().toLocaleTimeString());
    return true;
  }catch(e){ console.error(e); alert('Falha ao salvar na pasta: '+e.message); return false; }
}

async function loadAllFromFolder(){
  if(!dirHandle) return false;
  try{
    const rtxt = await readFile(dirHandle, DATAFILES[LS.res]).catch(e=>{ if(e && e.name==='NotFoundError') return '[]'; else throw e; });
    const atxt = await readFile(dirHandle, DATAFILES[LS.act]).catch(e=>{ if(e && e.name==='NotFoundError') return '[]'; else throw e; });
    const ttxt = await readFile(dirHandle, DATAFILES[LS.trail]).catch(e=>{ if(e && e.name==='NotFoundError') return '{}'; else throw e; });
    const r = JSON.parse(rtxt); const a = JSON.parse(atxt); const t = JSON.parse(ttxt);
    if(Array.isArray(r)&&Array.isArray(a)&&t&&typeof t==='object'){
      resources=(r||[]).map(coerceResource); activities=hydrateLoadedActivities(a||[]); trails=t;
      saveLS(LS.res,resources); saveLS(LS.act,activities); saveLS(LS.trail,trails);
      renderAll();
      updateFolderStatus('Carregado da pasta');
      return true;
    } else { alert('Arquivos inválidos na pasta.'); return false; }
  }catch(e){ console.error(e); alert('Falha ao carregar da pasta: '+e.message); return false; }
}

function updateFolderStatus(extra){
  const el=document.getElementById('folderStatus');
  if(!el) return;
  if(!dirHandle){ el.textContent='(nenhuma pasta definida)'; return; }
  el.textContent='Pasta: '+(dirHandle.name||'(sem nome)') + (extra? ' — '+extra : '');
}

// ===== Indicação de status de sincronização do Banco de Dados =====
/**
 * Atualiza o banner de status do banco de dados no topo da aplicação.
 * Quando `synced` for verdadeiro, exibe mensagem positiva; caso contrário, exibe mensagem de não sincronizado.
 */
function updateDBStatusBanner(state, detail=''){
  const el = document.getElementById('dbStatusBanner');
  if(!el) return;
  let mode = state;
  if (typeof state === 'boolean') {
    mode = state ? 'synced' : 'not_synced';
  }
  let text = 'Banco de Dados: Não sincronizado';
  let ok = false;
  if (mode === 'synced') {
    text = detail || 'Banco de Dados: Sincronizado';
    ok = true;
  } else if (mode === 'syncing') {
    text = detail || 'Banco de Dados: Atualizando dados...';
  } else if (mode === 'stale') {
    text = detail || 'Banco de Dados: Atualizado por outra sessão';
  } else if (mode === 'error') {
    text = detail || 'Banco de Dados: Falha ao atualizar';
  } else {
    text = detail || 'Banco de Dados: Não sincronizado';
  }
  el.textContent = text;
  el.classList.toggle('ok', ok);
  el.classList.toggle('warning', !ok);
}

// ===== Gerenciar caminho de BD em rede (campo de texto) =====
// Permite ao usuário salvar um caminho (string) para o banco de dados em rede.
// Este caminho é persistido em localStorage e pode ser copiado para a área de transferência.
const bdPathInput = document.getElementById('bdPathInput');
const btnSavePath = document.getElementById('btnSavePath');
const btnCopyPath = document.getElementById('btnCopyPath');

// Carrega o caminho salvo no carregamento inicial do aplicativo.
if (bdPathInput) {
  try {
    const saved = window.localStorage.getItem('rp_bd_path') || '';
    bdPathInput.value = saved;
  } catch (e) {
    console.warn('Não foi possível ler rp_bd_path', e);
  }
}

// Salva o caminho informado pelo usuário em localStorage
if (btnSavePath && bdPathInput) {
  btnSavePath.onclick = () => {
    const path = (bdPathInput.value || '').trim();
    try {
      safeSetLocalStorageItem('rp_bd_path', path, 'caminho do BD');
      alert('Caminho salvo!');
    } catch (e) {
      alert('Não foi possível salvar o caminho no armazenamento local.');
      console.warn('Erro ao salvar rp_bd_path', e);
    }
  };
}

// Copia o caminho salvo/atual para a área de transferência
if (btnCopyPath && bdPathInput) {
  btnCopyPath.onclick = async () => {
    const path = (bdPathInput.value || '').trim();
    if (!path) {
      alert('Nenhum caminho definido.');
      return;
    }
    try {
      await navigator.clipboard.writeText(path);
      alert('Caminho copiado!');
    } catch (e) {
      alert('Não foi possível copiar para a área de transferência. Por favor, copie manualmente.');
      console.warn('Erro ao copiar rp_bd_path', e);
    }
  };
}

const btnPickFolder=document.getElementById('btnPickFolder');
const btnSaveNow=document.getElementById('btnSaveNow');
const btnReloadFromFolder=document.getElementById('btnReloadFromFolder');
if(btnPickFolder) btnPickFolder.onclick=()=>fsaPickFolder();
if(btnSaveNow) btnSaveNow.onclick=()=>saveAllToFolder();
if(btnReloadFromFolder) btnReloadFromFolder.onclick=()=>loadAllFromFolder();

(async()=>{
  if(await hasFSA()){
    dirHandle = await fsaLoadHandle();
    updateFolderStatus();
    // Sempre atualizar status do BD para não sincronizado por padrão
    updateDBStatusBanner(false);
    if(dirHandle){
      try{ await loadAllFromFolder(); }catch{ /* ignore */ }
      startFsaWatcher();
    }
    // Carregar BD padrão previamente configurado (se houver)
    try {
      const savedBD = await idbGet(FSA_DB, FSA_STORE, 'bd');
      if(savedBD){
        bdHandle = savedBD;
        try {
          // Tentar ler o arquivo salvo e carregar dados
          const file = await bdHandle.getFile();
          bdFileName = file.name || '';
          bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
          const text = await file.text();
          let parsed;
          if(bdFileExt === 'csv') parsed = parseCSVBDUnico(text);
          else parsed = parseHTMLBDTables(text);
          resources = (parsed.recursos || []).map(coerceResource);
          activities = hydrateLoadedActivities(parsed.atividades || []);
          // horas/cfg/feriados se disponíveis
          if (parsed.horas && typeof window.setHorasExternosData === 'function') {
            window.setHorasExternosData(parsed.horas);
          }
          if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
            window.setHorasExternosConfig(parsed.cfg);
          }
          if (parsed.feriados && typeof window.setFeriados === 'function') {
            window.setFeriados(parsed.feriados);
          }
          if (parsed.horasExtras && typeof window.setHorasExtrasData === 'function') {
            window.setHorasExtrasData(parsed.horasExtras);
          }
          const newTrails = {};
          (parsed.historico || []).forEach(h => {
            const id = h.activityId;
            if (!id) return;
            if (!newTrails[id]) newTrails[id] = [];
            newTrails[id].push(buildTrailEntryFromBDRow(h));
          });
          trails = newTrails;
          // Ao carregar BD padrão no início, limpar eventLog e snapshot para evitar
          // reaplicar eventos antigos e trazer dados fora do BD.
          adoptCurrentStateAsPersistedBaseline();
          saveLS(LS.res, resources);
          saveLS(LS.act, activities);
          saveLS(LS.trail, trails);
          renderAll();
          updateBDStatus('BD carregado do padrão: ' + bdFileName);
          // armazena o lastModified atual como referência de última versão carregada
          try {
            const fileLm = file.lastModified;
            if (typeof window !== 'undefined') {
              window.__bdLastWrite = fileLm;
            }
          } catch(e){}
          startBDWatcher();
          updateDBStatusBanner(true);
        } catch (e) {
          console.warn('Falha ao carregar BD padrão', e);
          updateBDStatus('Falha ao carregar BD padrão');
          updateDBStatusBanner(false);
        }
      }
    } catch(e){
      console.warn('Erro ao recuperar BD padrão', e);
    }
  } else {
    const el=document.getElementById('folderStatus'); if(el) el.textContent='(navegador sem suporte de salvar em pasta — usando armazenamento do navegador)';
    // Também informar status de BD como não sincronizado
    updateDBStatusBanner(false);
  }
})();

const _saveLS_orig = saveLS;
saveLS = function(k,v){
  _saveLS_orig(k,v);
  if(dirHandle) { try{ saveAllToFolder(); }catch{} }
};

// ===== Dados iniciais =====
const today=clampDate(new Date());
const sampleResources = [];
const sampleActivities = [];

try {
  const VERSION_KEY = 'rv-version';
  const CUR_VERSION = '4';
  const storedV = localStorage.getItem(VERSION_KEY);
  if (!storedV || storedV < CUR_VERSION) {
    localStorage.removeItem(LS.res);
    localStorage.removeItem(LS.act);
    localStorage.removeItem(LS.trail);
    localStorage.removeItem('rv-enhancer-v1');
    safeSetLocalStorageItem(VERSION_KEY, CUR_VERSION, 'versão local');
  }
} catch (e) {
  // Ignorar erros de armazenamento
}

let resources = loadLS(LS.res, sampleResources);
// Garante que todos os recursos tenham campos de versionamento e exclusão
resources = (resources || []).map(r => {
  const nowTs = Date.now();
  return {
    ...r,
    cargaHorariaDiaria: Math.max(0.5, Number(r.cargaHorariaDiaria ?? r.cargaDiaria ?? 9) || 9),
    version: typeof r.version === 'number' ? r.version : 1,
    updatedAt: typeof r.updatedAt === 'number' ? r.updatedAt : nowTs,
    deletedAt: r.deletedAt || null
  };
});
let activities = loadLS(LS.act, sampleActivities);
// Garante que todas as atividades tenham campos derivados (incluindo comentários hidratados),
// versionamento, exclusão e código amigável logo na carga inicial.
activities = hydrateLoadedActivities((activities || []).map(a => {
  const nowTs = Date.now();
  return {
    ...a,
    codigoAtividade: String(a.codigoAtividade || a.codigo || a.atividadeCodigo || '').trim(),
    version: typeof a.version === 'number' ? a.version : 1,
    updatedAt: typeof a.updatedAt === 'number' ? a.updatedAt : nowTs,
    deletedAt: a.deletedAt || null
  };
}));
let comments = hydrateLoadedComments(activities, loadLS(LS.comments, []));
rebuildCommentsIndex();
syncAllActivityCommentFields();
saveLS(LS.act, activities);
saveLS(LS.comments, comments);
let trails=loadLS(LS.trail,{});
function saveTrails(){ saveLS(LS.trail, trails); }
function addTrail(atividadeId, entry){
  if(!trails[atividadeId]) trails[atividadeId]=[];
  trails[atividadeId].push(entry);
  saveTrails();
}

// Converte uma linha da tabela 'historico' do BD em uma entrada de trilha.
function buildTrailEntryFromBDRow(h){
  let meta = null;
  const legend = (h && typeof h.legend === 'string') ? h.legend.trim() : '';
  if(legend){
    try{
      // Alguns BDs antigos podem usar legend apenas como texto; só parseia se parecer JSON.
      if(legend.startsWith('{') && legend.endsWith('}')) meta = JSON.parse(legend);
    }catch(e){ meta = null; }
  }
  const entry = {
    ts: h.timestamp,
    oldInicio: h.oldInicio,
    oldFim: h.oldFim,
    newInicio: h.newInicio,
    newFim: h.newFim,
    justificativa: h.justificativa,
    user: h.user
  };
  if(meta && typeof meta === 'object'){
    if(meta.type) entry.type = meta.type;
    if(meta.entityType) entry.entityType = meta.entityType;
    if(meta.oldResourceId) entry.oldResourceId = meta.oldResourceId;
    if(meta.oldResourceName) entry.oldResourceName = meta.oldResourceName;
    if(meta.newResourceId) entry.newResourceId = meta.newResourceId;
    if(meta.newResourceName) entry.newResourceName = meta.newResourceName;
    if(meta.activityTitle) entry.activityTitle = meta.activityTitle;
    if(meta.activityCode) entry.activityCode = meta.activityCode;
    if(meta.commentSnapshot) entry.commentSnapshot = meta.commentSnapshot;
    if(meta.deletedByCascade) entry.deletedByCascade = meta.deletedByCascade;
  }
  return entry;
}

// Após carregar arrays de recursos/atividades e garantir campos de versionamento,
// aplicamos snapshot e eventos do log para reconstruir o estado.  Isso
// preserva compatibilidade: se não houver snapshot ou log, nada é alterado.
loadSnapshotAndEvents();
rebuildCommentsIndex();
syncAllActivityCommentFields();

// ===== Estado dos filtros =====
const selectedStatus=new Set(STATUS);
let filtroTipo="";
let filtroSenioridade="";
let buscaTitulo="";
let buscaTag = "";
let buscaRecurso="";

// Ao iniciar a aplicação, definir o início da visão para a data atual (hoje) ao invés de duas semanas antes.
let rangeStart=toYMD(today);
let rangeEnd=toYMD(addDays(today,60));

// Persistência de filtros (UX): mantém contexto entre recarregamentos.
const LS_FILTERS = "rp_filters_state_v1";
function loadFiltersState(){
  try{
    const raw = localStorage.getItem(LS_FILTERS);
    if(!raw) return null;
    const st = JSON.parse(raw);
    return st && typeof st === 'object' ? st : null;
  }catch{ return null }
}
function saveFiltersState(){
  try{
    const st = {
      filtroTipo,
      filtroSenioridade,
      buscaTitulo,
      buscaTag,
      buscaRecurso,
      rangeStart,
      rangeEnd,
      selectedStatus: Array.from(selectedStatus)
    };
    safeSetLocalStorageItem(LS_FILTERS, JSON.stringify(st), 'filtros');
  }catch{ /* ignore */ }
}

// Aplica o estado persistido (se existir) aos valores em memória.
// Observação: os elementos de UI são preenchidos mais adiante.
const __loadedFilters = loadFiltersState();
if(__loadedFilters){
  if(typeof __loadedFilters.filtroTipo === 'string') filtroTipo = __loadedFilters.filtroTipo;
  if(typeof __loadedFilters.filtroSenioridade === 'string') filtroSenioridade = __loadedFilters.filtroSenioridade;
  if(typeof __loadedFilters.buscaTitulo === 'string') buscaTitulo = __loadedFilters.buscaTitulo;
  if(typeof __loadedFilters.buscaTag === 'string') buscaTag = __loadedFilters.buscaTag;
  if(typeof __loadedFilters.buscaRecurso === 'string') buscaRecurso = __loadedFilters.buscaRecurso;
  if(typeof __loadedFilters.rangeStart === 'string') rangeStart = __loadedFilters.rangeStart;
  if(typeof __loadedFilters.rangeEnd === 'string') rangeEnd = __loadedFilters.rangeEnd;
  if(Array.isArray(__loadedFilters.selectedStatus)){
    selectedStatus.clear();
    __loadedFilters.selectedStatus.forEach(s=>{ if(STATUS.includes(s)) selectedStatus.add(s); });
    // Se por algum motivo ficou vazio, volta ao padrão
    if(selectedStatus.size===0) STATUS.forEach(s=>selectedStatus.add(s));
  }
}

// ===== UI refs =====
const statusChips=document.getElementById("statusChips");
const tipoSel=document.getElementById("tipoSel");
const senioridadeSel=document.getElementById("senioridadeSel");
const buscaTituloInput=document.getElementById("buscaTitulo");
const buscaTagInput = document.getElementById("buscaTag");
const buscaRecursoInput=document.getElementById("buscaRecurso");
const inicioVisao=document.getElementById("inicioVisao");
const fimVisao=document.getElementById("fimVisao");
const btnClearFilters = document.getElementById("btnClearFilters");
const viewKpi = document.getElementById("viewKpi");

// As referências às tabelas são mantidas para compatibilidade com outros scripts (enhancer.js)
// mas a função renderTables agora usa os contêineres de cards.
const tblRecursos=document.querySelector("#recursos-container"); // Anteriormente '#tblRecursos tbody'
const tblAtividades=document.querySelector("#atividades-container"); // Anteriormente '#tblAtividades tbody'
const gantt=document.getElementById("gantt");

const dlgRecurso=document.getElementById("dlgRecurso");
const formRecurso=document.getElementById("formRecurso");
const dlgRecursoTitulo=document.getElementById("dlgRecursoTitulo");

const dlgAtividade=document.getElementById("dlgAtividade");
const formAtividade=document.getElementById("formAtividade");
const dlgAtividadeTitulo=document.getElementById("dlgAtividadeTitulo");
const activityLinkSearchInput=document.getElementById("activityLinkSearchInput");

const dlgJust=document.getElementById("dlgJust");
const formJust=document.getElementById("formJust");
const justResumo=document.getElementById("justResumo");
const btnJustConfirm=document.getElementById("btnJustConfirm");

const dlgHist=document.getElementById("dlgHist");
const histList=document.getElementById("histList");
const atividadeComentariosBox=document.getElementById("atividadeComentariosBox");
const atividadeComentariosLista=document.getElementById("atividadeComentariosLista");
const btnHistExport=document.getElementById("btnHistExport");
let histCurrentId=null;
let currentCommentsActivityId=null;
let currentCommentsOffset=0;
const COMMENT_PAGE_SIZE=10;

const currentUserInput=document.getElementById("currentUser");
const btnHistAll=document.getElementById("btnHistAll");
const btnBackup=document.getElementById("btnBackup");
const fileRestore=document.getElementById("fileRestore");
const tooltip=document.getElementById("tooltip");
const aggGran=document.getElementById("aggGran");
const aggMode=document.getElementById("aggMode");
const aggCharts=document.getElementById("aggCharts");

// ===== Baseline mensal & KPI =====
const baselineSel = document.getElementById('baselineSel');
const baselineNome = document.getElementById('baselineNome');
const baselineUsarFiltros = document.getElementById('baselineUsarFiltros');
const baselineCompararFiltros = document.getElementById('baselineCompararFiltros');
const btnBaselineCriar = document.getElementById('btnBaselineCriar');
const btnBaselineComparar = document.getElementById('btnBaselineComparar');
const btnBaselineAplicarFiltros = document.getElementById('btnBaselineAplicarFiltros');
const btnBaselineBaixar = document.getElementById('btnBaselineBaixar');
const btnBaselineExcluir = document.getElementById('btnBaselineExcluir');
const kpiCards = document.getElementById('kpiCards');
const baselineTabela = document.getElementById('baselineTabela');
const btnExportComparacao = document.getElementById('btnExportComparacao');

let baselines = loadLS(LS.base, []);
let baselineItems = loadLS(LS.baseItems, []);
const atividadeComentariosInput = formAtividade?.elements?.["comentarios"] || null;
const atividadeComentariosCounter = document.getElementById('atividadeComentariosCounter');
let atividadeComentariosVerMaisBtn = document.getElementById('atividadeComentariosVerMais');
if(!atividadeComentariosVerMaisBtn && atividadeComentariosBox){
  atividadeComentariosVerMaisBtn = document.createElement('button');
  atividadeComentariosVerMaisBtn.type='button';
  atividadeComentariosVerMaisBtn.id='atividadeComentariosVerMais';
  atividadeComentariosVerMaisBtn.className='btn ghost';
  atividadeComentariosVerMaisBtn.textContent='Ver mais';
  atividadeComentariosVerMaisBtn.style.marginTop='8px';
  atividadeComentariosVerMaisBtn.style.display='none';
  atividadeComentariosBox.appendChild(atividadeComentariosVerMaisBtn);
}
let lastComparisonRows = [];

let currentUser = loadLS(LS.user, "");
let showInactiveResources = !!loadLS(LS.showInactiveRes, false);
const showInactiveResourcesInput = document.getElementById("showInactiveResources");
if(showInactiveResourcesInput){
  showInactiveResourcesInput.checked = showInactiveResources;
  showInactiveResourcesInput.onchange = ()=>{
    showInactiveResources = !!showInactiveResourcesInput.checked;
    saveLS(LS.showInactiveRes, showInactiveResources);
    renderAll();
  };
}
if(currentUserInput){ currentUserInput.value=currentUser; currentUserInput.readOnly=true; currentUserInput.placeholder='Usuário logado (matrícula + nome)'; currentUserInput.title='Identificação preenchida automaticamente após login'; }


// ===== Abas (tabs) =====
function activateTab(name){
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active', b.dataset.tab===name));
  document.querySelectorAll('.tabpanel').forEach(p=>p.classList.toggle('active', p.id==='tab-'+name));
}
document.addEventListener('click', (ev)=>{
  const b=ev.target.closest('.tab'); if(!b) return;
  activateTab(b.dataset.tab);
});

bindCalendarUI();
bindAdminRestrictedUI();
renderCalendarUI();
renderOvertimeUI();

// ===== Chips de status =====
function renderStatusChips(){
  statusChips.innerHTML="";
  STATUS.forEach(s=>{
    const span=document.createElement("span");
    span.className="chip"+(selectedStatus.has(s)?" active":"");
    span.textContent=s;
    span.onclick=()=>{ if(selectedStatus.has(s)) selectedStatus.delete(s); else selectedStatus.add(s); renderStatusChips(); renderAll(); };
    statusChips.appendChild(span);
  });
}

// Preenche UI com os filtros persistidos (quando existir)
if(tipoSel) tipoSel.value = filtroTipo || "";
if(senioridadeSel) senioridadeSel.value = filtroSenioridade || "";
if(buscaTituloInput) buscaTituloInput.value = buscaTitulo || "";
if(buscaTagInput) buscaTagInput.value = buscaTag || "";
if(buscaRecursoInput) buscaRecursoInput.value = buscaRecurso || "";

// ===== Filtros =====
tipoSel.onchange=()=>{filtroTipo=tipoSel.value; renderAll();};
senioridadeSel.onchange=()=>{filtroSenioridade=senioridadeSel.value; renderAll();};
buscaTituloInput.oninput=()=>{buscaTitulo=buscaTituloInput.value.toLowerCase(); renderAll();};
buscaTagInput.oninput = () => { buscaTag = buscaTagInput.value.toLowerCase(); renderAll(); };
if(buscaRecursoInput){
 buscaRecursoInput.oninput=()=>{
    buscaRecurso=buscaRecursoInput.value.toLowerCase().trim();
    renderAll();
  };
}

// Botão: limpar filtros e voltar para o padrão
if(btnClearFilters){
  btnClearFilters.onclick=()=>{
    try{ localStorage.removeItem(LS_FILTERS); }catch{}
    filtroTipo = "";
    filtroSenioridade = "";
    buscaTitulo = "";
    buscaTag = "";
    buscaRecurso = "";
    selectedStatus.clear();
    STATUS.forEach(s=>selectedStatus.add(s));
    rangeStart = toYMD(today);
    rangeEnd = endOfMonthYMD(rangeStart);
    // Sincroniza UI
    if(tipoSel) tipoSel.value = "";
    if(senioridadeSel) senioridadeSel.value = "";
    if(buscaTituloInput) buscaTituloInput.value = "";
    if(buscaTagInput) buscaTagInput.value = "";
    if(buscaRecursoInput) buscaRecursoInput.value = "";
    if(inicioVisao) inicioVisao.value = rangeStart;
    if(fimVisao) fimVisao.value = rangeEnd;
    renderStatusChips();
    renderAll();
  };
}

// ===== Range =====
/**
 * Retorna o último dia (Y-M-D) do mês referente à data informada (Y-M-D).
 * Ex.: 2026-01-10 -> 2026-01-31
 */
function endOfMonthYMD(ymd){
  const d = fromYMD(ymd);
  const last = new Date(d.getFullYear(), d.getMonth()+1, 0); // dia 0 do próximo mês
  return toYMD(last);
}

inicioVisao.value=rangeStart; fimVisao.value=rangeEnd;
inicioVisao.onchange=()=>{
  rangeStart=inicioVisao.value;
  // Ao alterar o início da visão, automaticamente considerar o fim do mês correspondente
  rangeEnd=endOfMonthYMD(rangeStart);
  if(fimVisao) fimVisao.value=rangeEnd;
  renderAll();
};
fimVisao.onchange=()=>{rangeEnd=fimVisao.value; renderAll();};

// ===== CRUD Recurso =====
document.getElementById("btnNovoRecurso").onclick=()=>{
  dlgRecursoTitulo.textContent="Novo Recurso";
  formRecurso.reset();
  formRecurso.elements["id"].value="";
  formRecurso.elements["capacidade"].value=100;
  if(formRecurso.elements["cargaHorariaDiaria"]) formRecurso.elements["cargaHorariaDiaria"].value=9;
  dlgRecurso.showModal();
};
document.getElementById("btnSalvarRecurso").onclick=()=>{
  const f=formRecurso.elements;
  // Build resource object with versioning and timestamps
  const idValue = f["id"].value || uuid();
  const existingIndex = resources.findIndex(r => r.id === idValue);
  const nowTs = Date.now();
  let version = 1;
  let deletedAt = null;
  if (existingIndex >= 0) {
    const existing = resources[existingIndex];
    version = (existing.version || 0) + 1;
    deletedAt = existing.deletedAt || null;
  }
  const rec={
    id: idValue,
    nome:f["nome"].value.trim(),
    tipo:f["tipo"].value,
    senioridade:f["senioridade"].value,
    ativo:f["ativo"].checked,
    capacidade:Math.max(1,Number(f["capacidade"].value||100)),
    cargaHorariaDiaria: Math.max(0.5, Number(f["cargaHorariaDiaria"].value || 9)),
    inicioAtivo:f["inicioAtivo"].value||"",
    fimAtivo:f["fimAtivo"].value||"",
    version: version,
    updatedAt: nowTs,
    deletedAt: deletedAt
  };
  if(!rec.nome){alert("Informe o nome.");return}
  if(existingIndex>=0) resources[existingIndex]=rec; else resources.push(rec);
  saveLS(LS.res,resources);
  // Registra evento no log: criação ou atualização de recurso
  recordEvent('resource', existingIndex>=0 ? 'update' : 'create', rec.id, rec);
  dlgRecurso.close();
  renderAll();
  saveBDDebounced();
};

function applyStatusToLinkedChain(baseActivity, newStatus){
  if(!baseActivity || !newStatus) return false;
  if(newStatus !== 'Concluída' && newStatus !== 'Cancelada') return false;
  const chain = getLinkedChainActivities(baseActivity.id).filter(a=>a.id !== baseActivity.id);
  if(!chain.length) return false;
  const targets = chain.filter(a => (a.status||'') !== newStatus);
  if(!targets.length) return false;
  const actionLabel = newStatus === 'Concluída' ? 'concluir' : 'cancelar';
  const preview = targets.slice(0,5).map(a => '- ' + (a.codigoAtividade || a.id) + ' | ' + a.titulo).join('\n');
  const extra = targets.length > 5 ? '\n... e mais ' + (targets.length - 5) + ' atividade(s).' : '';
  const msg = 'Esta atividade possui vínculo com ' + chain.length + ' outra(s) atividade(s).\n\nDeseja ' + actionLabel + ' também as demais atividades vinculadas?\n\n' + preview + extra;
  if(!confirm(msg)) return false;
  const nowTs = Date.now();
  targets.forEach(item=>{
    const idx = activities.findIndex(a=>a.id===item.id);
    if(idx<0) return;
    const updated = {
      ...activities[idx],
      status: newStatus,
      version: (activities[idx].version || 0) + 1,
      updatedAt: nowTs
    };
    activities[idx] = updated;
    try{ recordEvent('activity','update', updated.id, updated); }catch(_){ }
  });
  return true;
}

// ===== CRUD Atividade =====
function updateActivityCommentCounter(){
  if(!atividadeComentariosInput || !atividadeComentariosCounter) return;
  atividadeComentariosCounter.textContent = `${atividadeComentariosInput.value.length}/${MAX_COMMENT_LENGTH}`;
}
if(atividadeComentariosInput){
  atividadeComentariosInput.maxLength = MAX_COMMENT_LENGTH;
  atividadeComentariosInput.addEventListener('input', ()=>{
    if(atividadeComentariosInput.value.length > MAX_COMMENT_LENGTH){
      atividadeComentariosInput.value = atividadeComentariosInput.value.slice(0, MAX_COMMENT_LENGTH);
    }
    updateActivityCommentCounter();
  });
  updateActivityCommentCounter();
}
document.getElementById("btnNovaAtividade").onclick=()=>{
  dlgAtividadeTitulo.textContent="Nova Atividade";
  formAtividade.reset();
  fillRecursoOptions();
  formAtividade.elements["id"].value="";
  if(formAtividade.elements["codigoAtividade"]) formAtividade.elements["codigoAtividade"].value=generateNextActivityCode(activities);
  if(formAtividade.elements["linkedOriginSearch"]) formAtividade.elements["linkedOriginSearch"].value="";
  if(formAtividade.elements["linkedOriginId"]) formAtividade.elements["linkedOriginId"].value="";
  loadCommentDraftIntoForm('', currentUser || '');
  fillActivityLinkOptions();
  currentCommentsActivityId = null;
  currentCommentsOffset = 0;
  refreshActivityCommentsPanel(null);
  formAtividade.elements["inicio"].value=toYMD(today);
  formAtividade.elements["fim"].value=toYMD(addDays(today,5));
  formAtividade.elements["alocacao"].value=100;
  dlgAtividade.showModal();
};
/**
 * Preenche a lista de opções de recursos (datalist) para o campo de seleção de recurso na criação/edição de atividade.
 * Em vez de um <select>, utilizamos um <input> com datalist para permitir digitação e sugestão automática.
 */
function fillRecursoOptions(){
  const list = document.getElementById('resourceList');
  if(!list) return;
  list.innerHTML = '';
  // Ordena os recursos alfabeticamente (apenas ordem de exibição).
  // Mantém exatamente os mesmos filtros e regras já aplicados abaixo.
  const ganttResources = (resources || []).slice().sort((a,b)=>{
    const an = (a?.nome || '').toString();
    const bn = (b?.nome || '').toString();
    return an.localeCompare(bn, undefined, { sensitivity: 'base' });
  });

  ganttResources.forEach(r=>{
    if(!r.ativo) return;
    const opt = document.createElement('option');
    opt.value = r.nome;
    list.appendChild(opt);
  });
}
if(activityLinkSearchInput){
  activityLinkSearchInput.addEventListener('input', ()=>{
    const hidden = formAtividade?.elements?.['linkedOriginId'];
    const currentId = formAtividade?.elements?.['id']?.value || '';
    const currentLinkedId = hidden?.value || '';
    const currentLinked = currentLinkedId ? (activities || []).find(a => !a.deletedAt && a.id === currentLinkedId && a.id !== currentId) : null;
    const currentLabel = currentLinked ? formatActivityLinkOption(currentLinked) : '';
    if(hidden && activityLinkSearchInput.value.trim() !== currentLabel){
      hidden.value = '';
    }
    fillActivityLinkOptions(hidden?.value || '', activityLinkSearchInput.value, currentId);
  });
  activityLinkSearchInput.addEventListener('change', ()=>{
    const resolved = resolveLinkedActivityReference(activityLinkSearchInput.value, formAtividade?.elements?.['id']?.value || '');
    if(formAtividade?.elements?.['linkedOriginId']) formAtividade.elements['linkedOriginId'].value = resolved ? resolved.id : '';
    if(resolved) activityLinkSearchInput.value = formatActivityLinkOption(resolved);
  });
}
document.getElementById("btnSalvarAtividade").onclick=async ()=>{
  const f=formAtividade.elements;

  const tagsInput = f["tags"].value || '';
  const tags = tagsInput.split(',')
                       .map(normalizeTag)
                       .filter(t => t);

  // Obtém o nome do recurso digitado e valida sua existência
  const resourceName = (f["resourceName"].value || '').trim();
  const recByName = resources.find(r => (r.nome || '').toLowerCase() === resourceName.toLowerCase());

  // Prepare id, código amigável e versioning
  const atId = f["id"].value || uuid();
  const existingAtIndex = activities.findIndex(a => a.id === atId);
  const existingActivity = existingAtIndex >= 0 ? activities[existingAtIndex] : null;
  const activityCode = String((f["codigoAtividade"] && f["codigoAtividade"].value) || existingActivity?.codigoAtividade || generateNextActivityCode(activities)).trim();
  const nowAtTs = Date.now();
  let atVersion = 1;
  let atDeletedAt = null;
  if (existingAtIndex >= 0) {
    const existingA = activities[existingAtIndex];
    atVersion = (existingA.version || 0) + 1;
    atDeletedAt = existingA.deletedAt || null;
  }
  const linkedOriginResolved = resolveLinkedActivityReference((f["linkedOriginSearch"] && f["linkedOriginSearch"].value) || '', atId) || ((f["linkedOriginId"] && f["linkedOriginId"].value) ? activities.find(a=>!a.deletedAt && a.id === f["linkedOriginId"].value && a.id !== atId) : null);
  const newCommentText = String(f["comentarios"].value || '').trim();
  if(!validateCommentLength(newCommentText)) return;
  const existingDerivedComments = getAllCommentsForActivity(atId).map(c => ({ id:c.commentId, ts:c.ts, user:c.usuario, text:c.texto }));
  const at={
    id: atId,
    codigoAtividade: activityCode,
    titulo:f["titulo"].value.trim(),
    // ResourceId será atribuído a partir do recurso encontrado pelo nome
    resourceId: recByName ? recByName.id : '',
    linkedOriginId: linkedOriginResolved ? linkedOriginResolved.id : '',
    linkedOriginCode: linkedOriginResolved ? String(linkedOriginResolved.codigoAtividade || '').trim() : '',
    inicio:f["inicio"].value,
    fim:f["fim"].value,
    status:f["status"].value,
    alocacao:Math.max(1,Number(f["alocacao"].value||100)),
    comentariosLista: existingDerivedComments,
    comentarios: formatActivityCommentsForDisplay(existingDerivedComments),
    comentariosJson: serializeActivityComments(existingDerivedComments),
    tags: [...new Set(tags)],
    execSubtasks: normalizeExecData(existingActivity || {}).subtasks,
    execEntries: normalizeExecData(existingActivity || {}).entries,
    execIssues: normalizeExecData(existingActivity || {}).issues,
    version: atVersion,
    updatedAt: nowAtTs,
    deletedAt: atDeletedAt
  };
  if(!at.titulo) return alert("Informe o título.");
  if(!resourceName) return alert("Selecione o recurso.");
  if(!recByName) return alert("O nome do recurso não corresponde a nenhum recurso cadastrado.");
  if(fromYMD(at.fim)<fromYMD(at.inicio)) return alert("Fim não pode ser menor que início.");

  const rec = recByName;
  if(rec){
    if(rec.inicioAtivo && fromYMD(at.inicio) < fromYMD(rec.inicioAtivo)){
      return alert(`Início da atividade (${at.inicio}) menor que início ativo do recurso (${rec.inicioAtivo}).`);
    }
    if(rec.fimAtivo && fromYMD(at.fim) > fromYMD(rec.fimAtivo)){
      return alert(`Fim da atividade (${at.fim}) maior que fim ativo do recurso (${rec.fimAtivo}).`);
    }
  }
  let over=false;
  const start=fromYMD(at.inicio), end=fromYMD(at.fim);
  for(let d=new Date(start); d<=end; d=addDays(d,1)){
    const sum = activities.filter(x=>x.id!==at.id && x.resourceId===at.resourceId && fromYMD(x.inicio)<=d && d<=fromYMD(x.fim))
                          .reduce((acc,x)=>acc+(x.alocacao||100),0) + (at.alocacao||100);
    const cap = rec? (rec.capacidade||100) : 100;
    if(sum>cap){ over=true; break; }
  }
  if(over && !confirm("Aviso: esta alteração resultará em sobrealocação (>100%) em pelo menos um dia. Deseja continuar?")) return;

  const idx=activities.findIndex(a=>a.id===at.id);
  if(idx>=0){
    const prev=activities[idx];
    const mudouDatas = prev.inicio!==at.inicio || prev.fim!==at.fim;
    if(mudouDatas){
      window.__pendingAt=at;
      window.__pendingIdx=idx;
      justResumo.textContent=`${prev.titulo} — Início: ${prev.inicio} → ${at.inicio} | Fim: ${prev.fim} → ${at.fim}`;
      formJust.elements["just"].value="";
      dlgJust.showModal();
      return; 
    } else {
      // Registra delegação de atividade (mudança de responsável) no histórico
      const mudouRecurso = (prev.resourceId || '') !== (at.resourceId || '');
      if(mudouRecurso){
        const byIdRes = Object.fromEntries(resources.map(r=>[r.id,r]));
        const oldName = (prev.resourceId && byIdRes[prev.resourceId]) ? byIdRes[prev.resourceId].nome : '';
        const newName = (at.resourceId && byIdRes[at.resourceId]) ? byIdRes[at.resourceId].nome : '';
        addTrail(at.id, {
          ts: new Date().toISOString(),
          type: 'DELEGACAO_ATIVIDADE',
          entityType: 'atividade',
          oldResourceId: prev.resourceId || '',
          oldResourceName: oldName || '',
          newResourceId: at.resourceId || '',
          newResourceName: newName || '',
          justificativa: '',
          user: currentUser||''
        });
      }
      if(newCommentText){
        const reconcileResult = await reconcileCommentWrite(prev, at, newCommentText, currentUser || '');
        if(!reconcileResult.ok) return;
        if(reconcileResult.message) showToast('Comentários', reconcileResult.message, 'success', 3200);
      }
      activities[idx]=syncExecFields(syncActivityCommentFields(at));
      if(newCommentText){
        const appendResult = appendComment(at.id, newCommentText, currentUser || '');
        if(!appendResult.ok){
          salvarComentarioDraft(at.id, currentUser || '', { texto:newCommentText, ts:Date.now(), baseActivityVersion: prev.version || 0, baseActivityUpdatedAt: prev.updatedAt || 0 });
          return;
        }
        activities[idx]=syncActivityCommentFields(activities[idx]);
      }
      const statusChanged = (prev.status||'') !== (at.status||'');
      if(statusChanged){
        try{ applyStatusToLinkedChain(at, at.status); }catch(_){ }
      }
      saveLS(LS.act,activities);
      // Registra evento de atualização de atividade
      recordEvent('activity','update', at.id, at);
      dlgAtividade.close();
      renderAll();
      saveBDDebounced();
    }
  } else {
    activities.push(syncExecFields(syncActivityCommentFields(at)));
    if(newCommentText){
      const appendResult = appendComment(at.id, newCommentText, currentUser || '');
      if(!appendResult.ok){
        salvarComentarioDraft(at.id, currentUser || '', { texto:newCommentText, ts:Date.now(), baseActivityVersion: 0, baseActivityUpdatedAt: 0 });
        return;
      }
      activities[activities.length-1]=syncActivityCommentFields(activities[activities.length-1]);
    }
    saveLS(LS.act,activities);
    // Registra evento de criação de atividade
    recordEvent('activity','create', at.id, at);
    dlgAtividade.close();
    renderAll();
    saveBDDebounced();
  }
};

btnJustConfirm.onclick=(e)=>{
  e.preventDefault();
  const txt=formJust.elements["just"].value.trim();
  if(!txt){ alert("Informe a justificativa."); return; }
  // Modo 1: justificativa para exclusão (atividade/recurso)
  if(window.__pendingDelete){
    const pd = window.__pendingDelete;
    const nowIso = new Date().toISOString();
    const nowTs = Date.now();
    const byIdRes = Object.fromEntries(resources.map(r=>[r.id,r]));
    try{
      if(pd.entityType === 'atividade'){
        const aIdx = activities.findIndex(a=>a.id===pd.id);
        if(aIdx>=0){
          const prevA = activities[aIdx];
          const v = (prevA.version || 0) + 1;
          const updated = { ...prevA, deletedAt: nowTs, updatedAt: nowTs, version: v };
          activities[aIdx]=updated;
          saveLS(LS.act,activities);
          recordEvent('activity','delete', updated.id, { deletedAt: nowTs });
          // Evento de exclusão (auditável)
          addTrail(updated.id, {
            ts: nowIso,
            type: 'EXCLUSAO_ATIVIDADE',
            entityType: 'atividade',
            justificativa: txt,
            user: currentUser||'',
            oldResourceId: prevA.resourceId || '',
            oldResourceName: (prevA.resourceId && byIdRes[prevA.resourceId]) ? byIdRes[prevA.resourceId].nome : '',
            newResourceId: '',
            newResourceName: '',
            activityTitle: prevA.titulo || '',
            activityCode: prevA.codigoAtividade || '',
            commentSnapshot: prevA.comentarios || ''
          });
        }
      } else if(pd.entityType === 'recurso'){
        const rIdx = resources.findIndex(r=>r.id===pd.id);
        if(rIdx>=0){
          const prevR = resources[rIdx];
          const v = (prevR.version || 0) + 1;
          resources[rIdx] = { ...prevR, deletedAt: nowTs, updatedAt: nowTs, version: v };
          // marca todas as atividades desse recurso como deletadas
          activities = activities.map(item => {
            if(item.resourceId === prevR.id && !item.deletedAt){
              const vv = (item.version || 0) + 1;
              const updated = { ...item, deletedAt: nowTs, updatedAt: nowTs, version: vv };
              // registra evento de exclusão por atividade (para auditoria)
              addTrail(updated.id, {
                ts: nowIso,
                type: 'EXCLUSAO_ATIVIDADE',
                entityType: 'atividade',
                justificativa: txt,
                user: currentUser||'',
                oldResourceId: updated.resourceId || '',
                oldResourceName: prevR.nome || '',
                newResourceId: '',
                newResourceName: '',
                activityTitle: updated.titulo || '',
                activityCode: updated.codigoAtividade || '',
                commentSnapshot: updated.comentarios || '',
                deletedByCascade: true
              });
              recordEvent('activity','delete', updated.id, { deletedAt: nowTs });
              return updated;
            }
            return item;
          });
          saveLS(LS.res, resources);
          saveLS(LS.act, activities);
          recordEvent('resource','delete', prevR.id, { deletedAt: nowTs });
          // Evento de exclusão do recurso (auditável)
          addTrail(prevR.id, {
            ts: nowIso,
            type: 'EXCLUSAO_RECURSO',
            entityType: 'recurso',
            justificativa: txt,
            user: currentUser||'',
            oldResourceId: prevR.id,
            oldResourceName: prevR.nome || '',
            newResourceId: '',
            newResourceName: ''
          });
        }
      }
    }catch(e){
      console.error('Erro ao excluir com justificativa', e);
      alert('Não foi possível concluir a exclusão: '+(e.message||e));
    }
    window.__pendingDelete = null;
    dlgJust.close();
    renderAll();
    saveBDDebounced();
    return;
  }

  const at=window.__pendingAt;
  const idx=window.__pendingIdx;
  if(at==null || idx==null){ dlgJust.close(); return; }
  const prev=activities[idx];
  // Se também houve mudança de recurso, registra delegação separadamente (além do evento de datas)
  const mudouRecurso = (prev.resourceId || '') !== (at.resourceId || '');
  if(mudouRecurso){
    const byIdRes = Object.fromEntries(resources.map(r=>[r.id,r]));
    const oldName = (prev.resourceId && byIdRes[prev.resourceId]) ? byIdRes[prev.resourceId].nome : '';
    const newName = (at.resourceId && byIdRes[at.resourceId]) ? byIdRes[at.resourceId].nome : '';
    addTrail(at.id, {
      ts: new Date().toISOString(),
      type: 'DELEGACAO_ATIVIDADE',
      entityType: 'atividade',
      oldResourceId: prev.resourceId || '',
      oldResourceName: oldName || '',
      newResourceId: at.resourceId || '',
      newResourceName: newName || '',
      justificativa: '',
      user: currentUser||""
    });
  }
  addTrail(at.id, {
    ts: new Date().toISOString(),
    oldInicio: prev.inicio, oldFim: prev.fim,
    newInicio: at.inicio, newFim: at.fim,
    justificativa: txt,
    user: currentUser||""
  });
  activities[idx]=syncExecFields(at);
  saveLS(LS.act,activities);
  // Registra evento de atualização após justificativa de alteração de datas
  recordEvent('activity','update', at.id, at);
  dlgJust.close();
  dlgAtividade.close();
  renderAll();
  saveBDDebounced();
};

// ===== NOVA FUNÇÃO renderTables PARA LAYOUT DE CARDS =====
function renderTables(filteredActs){
  const recursosContainer = document.getElementById('recursos-container');
  const atividadesContainer = document.getElementById('atividades-container');
  if (!recursosContainer || !atividadesContainer) return;

  // ===== Renderização dos Cards de Recursos =====
  recursosContainer.innerHTML = "";
  // Filtra recursos de acordo com os mesmos critérios aplicados às atividades.
  // Além dos filtros de tipo/senioridade/nomes, exibe apenas recursos que tenham
  // pelo menos uma atividade no conjunto filtrado (filteredActs). Isso garante
  // que a listagem de recursos reflita fielmente a seleção feita nas buscas.
  const recursosComAtividades = new Set(filteredActs.map(a => a.resourceId));
  const visibleResources = resources.filter(r => {
    // Ignorar recursos marcados como excluídos
    if(r.deletedAt) return false;
    // Mostrar apenas recursos ativos, a menos que o usuário peça para gerenciar inativos
    if(!showInactiveResources && !r.ativo) return false;
    // Filtrar por tipo (Interno/Externo) se definido
    if(filtroTipo && (r.tipo||'').toLowerCase() !== filtroTipo.toLowerCase()) return false;
    // Filtrar por senioridade se definido
    if(filtroSenioridade && r.senioridade !== filtroSenioridade) return false;
    // Filtrar por busca de recurso se definido
    if(buscaRecurso && !(r.nome || '').toLowerCase().includes(buscaRecurso)) return false;
    // Se houver atividades filtradas, normalmente o componente escondia recursos
    // que não possuíssem nenhuma atividade visível. Entretanto, isso impedia
    // editar/inativar/excluir recursos sem atividades. Agora, sempre exibe
    // recursos que passam pelos demais filtros, mesmo que não tenham
    // nenhuma atividade no conjunto filtrado. Isso permite manipular
    // recursos "órfãos" de forma explícita.
    return true;
  });
  const groupedResources = visibleResources.reduce((acc, r) => {
    const tipo = r.tipo || 'Outros';
    if (!acc[tipo]) acc[tipo] = [];
    acc[tipo].push(r);
    return acc;
  }, {});

  // Ordem desejada: Interno, Externo, e depois outros
  const groupOrder = ['Interno', 'Externo'];
  const sortedGroups = [...Object.keys(groupedResources)].sort((a, b) => {
    const aIndex = groupOrder.indexOf(a);
    const bIndex = groupOrder.indexOf(b);
    if (aIndex > -1 && bIndex > -1) return aIndex - bIndex;
    if (aIndex > -1) return -1;
    if (bIndex > -1) return 1;
    return a.localeCompare(b);
  });

  for (const tipo of sortedGroups) {
    const group = groupedResources[tipo];
    const details = document.createElement('details');
    details.className = 'resource-group';
    details.open = true; // Inicia aberto

    const summary = document.createElement('summary');
    summary.textContent = tipo + ` (${group.length})`;
    details.appendChild(summary);

    group.forEach(r => {
      const card = document.createElement('div');
      card.className = `resource-card type-${tipo.toLowerCase()}`;
      card.innerHTML = `
        <div class="card-header">
          <strong class="resource-name">👤 ${escHTML(r.nome)}</strong>
          <div class="card-actions">
            <button class="btn-icon edit" title="Editar"></button>
            <button class="btn-icon duplicate" title="Duplicar"></button>
            ${r.ativo ? `<button class="btn-icon delete" title="Excluir"></button>` : `<button class="btn-icon reactivate" title="Reativar"></button>`}
          </div>
        </div>
        <div class="card-body">
          <span class="chip">${escHTML(r.senioridade)}</span>
          <span class="chip">${r.ativo ? "Ativo" : "Inativo"}</span>
          <span class="chip">📊 Cap: ${r.capacidade || 100}%</span>
          <span class="chip">⏱ ${getDailyHours(r)}h/dia</span>
        </div>
        ${(r.inicioAtivo || r.fimAtivo) ? `<div class="card-footer muted small">📅 ${escHTML(r.inicioAtivo || '...')} → ${escHTML(r.fimAtivo || '...')}</div>` : ''}
      `;
      // Adiciona eventos aos botões de ação
      card.querySelector('.edit').onclick = () => {
        dlgRecursoTitulo.textContent="Editar Recurso";
        formRecurso.elements["id"].value=r.id;
        formRecurso.elements["nome"].value=r.nome;
        formRecurso.elements["tipo"].value=r.tipo;
        formRecurso.elements["senioridade"].value=r.senioridade;
        formRecurso.elements["ativo"].checked=!!r.ativo;
        formRecurso.elements["capacidade"].value=r.capacidade||100;
        if(formRecurso.elements["cargaHorariaDiaria"]) formRecurso.elements["cargaHorariaDiaria"].value=getDailyHours(r);
        formRecurso.elements["inicioAtivo"].value=r.inicioAtivo||"";
        formRecurso.elements["fimAtivo"].value=r.fimAtivo||"";
        dlgRecurso.showModal();
      };
      card.querySelector('.duplicate').onclick = () => {
        const nowTs = Date.now();
        const copy={...r,id:uuid(),nome:"Cópia de "+r.nome, version:1, updatedAt: nowTs, deletedAt:null};
        resources.push(copy);
        saveLS(LS.res,resources);
        // Registra evento de criação (duplicação) de recurso
        recordEvent('resource','create',copy.id, copy);
        renderAll();
        // Persiste no BD assíncrono para evitar reaparecimento após sincronização
        saveBDDebounced();
      };
      const deleteBtn = card.querySelector('.delete');
      if(deleteBtn){
        deleteBtn.onclick = () => {
          // Exclusão com justificativa obrigatória (equivalente ao padrão de alteração de datas)
          window.__pendingDelete = {
            entityType: 'recurso',
            id: r.id
          };
          justResumo.textContent = `Excluir recurso: ${r.nome} (isso também excluirá as atividades vinculadas).`;
          formJust.elements["just"].value = "";
          dlgJust.showModal();
        };
      }
      const reactivateBtn = card.querySelector('.reactivate');
      if(reactivateBtn){
        reactivateBtn.onclick = () => {
          const idx = resources.findIndex(x => x.id === r.id);
          if(idx < 0) return;
          const nowTs = Date.now();
          resources[idx] = { ...resources[idx], ativo: true, version: (resources[idx].version || 0) + 1, updatedAt: nowTs };
          saveLS(LS.res, resources);
          recordEvent('resource', 'update', resources[idx].id, resources[idx]);
          renderAll();
          saveBDDebounced();
        };
      }
      details.appendChild(card);
    });
    recursosContainer.appendChild(details);
  }

  // ===== Renderização dos Cards de Atividades =====
  atividadesContainer.innerHTML = "";
  filteredActs.forEach(a => {
    const r = resources.find(x => x.id === a.resourceId);
    const card = document.createElement('div');
    card.className = 'activity-card';
    const statusClass = (a.status || '').toLowerCase().replace(/\s+/g, '-');

    const tagsHtml = (a.tags && a.tags.length)
      ? `<div class="tags-container">${a.tags.map(t => `<span class="chip tag">${escHTML(t)}</span>`).join(' ')}</div>`
      : '';
    const linkedOrigin = a.linkedOriginId ? activities.find(x=>x.id === a.linkedOriginId) : null;
    const linkedLabel = linkedOrigin ? formatActivityLinkOption(linkedOrigin) : String(a.linkedOriginCode || '').trim();
    const linkedHtml = linkedLabel ? `<div class="muted small">🔗 ${escHTML(linkedLabel)}</div>` : '';
    const extrapolated = isActivityExtrapolated(a);
    const extrapolatedHtml = extrapolated ? `<span class="status-badge status-bloqueada" title="Data fim vencida e atividade ainda não concluída/cancelada">Extrapolada</span>` : '';

    card.innerHTML = `
      <div class="card-header">
        <div><div class="muted small">${escHTML(a.codigoAtividade || "")}</div><strong class="activity-title">${escHTML(a.titulo)}</strong></div>
        <div class="card-actions">
          <button class="btn-icon edit" title="Editar"></button>
          <button class="btn-icon history" title="Histórico"></button>
          <button class="btn-icon execution" title="Execução" aria-label="Abrir execução da atividade"></button>
          <button class="btn-icon duplicate" title="Duplicar"></button>
          <button class="btn-icon delete" title="Excluir"></button>
        </div>
      </div>
      <div class="card-body">
        <div class="activity-meta">
          <span class="status-badge status-${statusClass}">${escHTML(a.status)}</span>
          ${extrapolatedHtml}
          <span class="muted small">👤 ${escHTML(r ? r.nome : 'N/A')}</span>
        </div>
        <div class="allocation-bar-container" title="Alocação: ${a.alocacao || 100}%">
          <div class="allocation-bar" style="width: ${Math.min(a.alocacao || 100, 100)}%;"></div>
          ${(a.alocacao || 100) > 100 ? '<div class="allocation-overload"></div>' : ''}
        </div>
        ${linkedHtml}
        ${tagsHtml}
        ${getAllCommentsForActivity(a.id).length ? `<div style="margin-top:8px;"><div class="muted small" style="margin-bottom:6px;"><strong>Comentários (${getAllCommentsForActivity(a.id).length})</strong></div><div class="small" style="white-space:pre-wrap;">${escHTML(getLastCommentPreview(a.id) || '')}</div></div>` : ''}
      </div>
      <div class="card-footer muted small">
        📅 ${escHTML(a.inicio)} → ${escHTML(a.fim)}
      </div>
    `;
    // Adiciona eventos aos botões de ação
    card.querySelector('.edit').onclick = () => {
      dlgAtividadeTitulo.textContent="Editar Atividade";
      fillRecursoOptions();
      formAtividade.elements["id"].value=a.id;
      if(formAtividade.elements["codigoAtividade"]) formAtividade.elements["codigoAtividade"].value=a.codigoAtividade||"";
      formAtividade.elements["titulo"].value=a.titulo;
      // Preenche campo de recurso com o nome correspondente
      const recObj = resources.find(x=>x.id===a.resourceId);
      formAtividade.elements["resourceName"].value = recObj ? recObj.nome : '';
      formAtividade.elements["inicio"].value=a.inicio;
      formAtividade.elements["fim"].value=a.fim;
      formAtividade.elements["status"].value=a.status;
      formAtividade.elements["alocacao"].value=a.alocacao||100;
      if(formAtividade.elements["comentarios"]) formAtividade.elements["comentarios"].value = '';
      updateActivityCommentCounter();
      refreshActivityCommentsPanel(a);
      formAtividade.elements["tags"].value = (a.tags || []).join(', ');
      if(formAtividade.elements["linkedOriginId"]) formAtividade.elements["linkedOriginId"].value = a.linkedOriginId || '';
      if(formAtividade.elements["linkedOriginSearch"]) formAtividade.elements["linkedOriginSearch"].value = '';
      fillActivityLinkOptions(a.linkedOriginId || '', '', a.id);
      dlgAtividade.showModal();
    };
    card.querySelector('.history').onclick = () => {
      histCurrentId=a.id;
      const list = trails[a.id]||[];
      if(list.length===0){
        histList.innerHTML='<div class="muted">Sem alterações de datas registradas para esta atividade.</div>';
      }else{
        const s=document.getElementById('histStart').value;
        const e=document.getElementById('histEnd').value;
        const rows=list.slice().reverse().filter(it=>{
          const t=new Date(it.ts);
          return (!s || t>=fromYMD(s)) && (!e || t<=addDays(fromYMD(e),0));
        }).map(it=>{
          const safeJust = it.justificativa ? escHTML(it.justificativa) : '';
          const safeComment = it.commentSnapshot ? escHTML(it.commentSnapshot).replace(/\n/g,'<br>') : '';
          const safeCode = escHTML(it.activityCode || '');
          const safeTitle = escHTML(it.activityTitle || a.titulo || '');
          if((it.type || '') === 'EXCLUSAO_ATIVIDADE'){
            return `<div style="padding:6px 8px; background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin:6px 0">
              <div style="font-size:12px;color:#475569">${new Date(it.ts).toLocaleString()}${it.user? ' • ' + it.user : ''}</div>
              <div style="font-weight:600">Exclusão de atividade${it.deletedByCascade ? ' (por exclusão do recurso)' : ''}</div>
              <div style="margin-top:4px">${safeCode ? safeCode + ' — ' : ''}${safeTitle}</div>
              ${safeComment ? `<div style="margin-top:4px"><strong>Comentário salvo:</strong><br>${safeComment}</div>` : ''}
              <div style="margin-top:4px">${safeJust}</div>
            </div>`;
          }
          if((it.type || '') === 'DELEGACAO_ATIVIDADE'){
            return `<div style="padding:6px 8px; background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin:6px 0">
              <div style="font-size:12px;color:#475569">${new Date(it.ts).toLocaleString()}${it.user? ' • ' + it.user : ''}</div>
              <div style="font-weight:600">Delegação de atividade</div>
              <div style="margin-top:4px">${escHTML(it.oldResourceName || '')} → ${escHTML(it.newResourceName || '')}</div>
              <div style="margin-top:4px">${safeJust}</div>
            </div>`;
          }
          return `<div style="padding:6px 8px; background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin:6px 0">
              <div style="font-size:12px;color:#475569">${new Date(it.ts).toLocaleString()}${it.user? ' • ' + it.user : ''}</div>
              <div style="font-weight:600">Início: ${it.oldInicio} → ${it.newInicio} | Fim: ${it.oldFim} → ${it.newFim}</div>
              <div style="margin-top:4px">${safeJust}</div>
            </div>`;
        }).join("");
        histList.innerHTML=rows || '<div class="muted">Sem registros no período.</div>';
      }
      dlgHist.showModal();
      const btn=document.getElementById('histApply'); if(btn){ btn.onclick=()=>{ card.querySelector('.history').onclick(); }; }
    };
    const execBtn = card.querySelector('.execution');
    if(execBtn) execBtn.onclick = () => openExecutionDialog(a);
    card.querySelector('.duplicate').onclick = () => {
      const nowTs = Date.now();
      const copy={...a,id:uuid(),codigoAtividade:generateNextActivityCode(activities),titulo:"Cópia de "+a.titulo, linkedOriginId:"", version:1, updatedAt: nowTs, deletedAt:null, execSubtasks:[], execEntries:[], execIssues:[]};
      activities.push(syncExecFields(copy));
      saveLS(LS.act,activities);
      // Registra evento de criação de atividade (cópia)
      recordEvent('activity','create', copy.id, copy);
      renderAll();
      // Persiste no BD assíncrono para evitar reaparecimento após sincronização
      saveBDDebounced();
    };
    card.querySelector('.delete').onclick = () => {
      // Exclusão com justificativa obrigatória (equivalente ao padrão de alteração de datas)
      window.__pendingDelete = {
        entityType: 'atividade',
        id: a.id
      };
      justResumo.textContent = `Excluir atividade: ${a.titulo}`;
      formJust.elements["just"].value = "";
      dlgJust.showModal();
    };
    atividadesContainer.appendChild(card);
  });
}

// ===== Interseção/intervalos =====
function mergeIntervals(intervals){
  if(!intervals.length) return [];
  const sorted=[...intervals].sort((a,b)=>fromYMD(a.inicio)-fromYMD(b.inicio));
  const res=[sorted[0]];
  for(let i=1;i<sorted.length;i++){
    const prev=res[res.length-1];
    const cur=sorted[i];
    const prevEnd=fromYMD(prev.fim);
    const curStart=fromYMD(cur.inicio);
    if(curStart<=addDays(prevEnd,1)){
      const newEnd=fromYMD(cur.fim)>prevEnd?cur.fim:prev.fim;
      res[res.length-1]={inicio:prev.inicio,fim:newEnd};
    }else res.push(cur);
  }
  return res;
}
function invertIntervals(intervals,startYMD,endYMD){
  const free=[]; let cursor=startYMD;
  const sDate=fromYMD(startYMD); const eDate=fromYMD(endYMD);
  const merged=mergeIntervals(intervals);
  for(const intv of merged){
    const iStart=fromYMD(intv.inicio);
    if(iStart>sDate && fromYMD(cursor)<iStart){
      free.push({inicio:cursor,fim:toYMD(addDays(iStart,-1))});
    }
    const iEnd=fromYMD(intv.fim);
    cursor=toYMD(addDays(iEnd,1));
  }
  if(fromYMD(cursor)<=eDate) free.push({inicio:cursor,fim:endYMD});
  return free;
}

// ===== Gantt =====
function statusClass(s){
  switch(s){
    case "Planejada": return "planejada";
    case "Em Execução": return "execucao";
    case "Bloqueada": return "bloqueada";
    case "Concluída": return "concluida";
    case "Cancelada": return "cancelada";
    default: return "";
  }
}
function buildDays(){
  const start=fromYMD(rangeStart), end=fromYMD(rangeEnd);
  const out=[]; for(let d=new Date(start); d<=end; d=addDays(d,1)) out.push(new Date(d));
  return out;
}

function renderGantt(filteredActs){
  gantt.innerHTML="";
  const days=buildDays();
  const todayYMD = toYMD(today);
  const todayIdx = days.findIndex(d => toYMD(d) === todayYMD);
  const header=document.createElement("div");
  header.className="header";
  const left=document.createElement("div");
  left.className="col-fixed";
  left.innerHTML=`<div class="muted" style="font-size:12px;font-weight:600">RECURSO</div>`;
  const right=document.createElement("div");
  right.className="col-grid";

  const gridDays=document.createElement("div");
  gridDays.className="grid-days";
  gridDays.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  // -------------------------------------------------------------------------
  // Header row showing month names.  We display a short month label
  // ("jan", "fev", "mar", etc.) on the first cell of each month and leave
  // other cells blank.  This helps users identify the month boundaries
  // without introducing confusing year labels (e.g. "fev. 25").  If your
  // locale is set to Portuguese, this will show month abbreviations in
  // Portuguese; otherwise it will use the browser's default locale.
  const rowMonths=document.createElement("div");
  rowMonths.className="row-months";
  rowMonths.style.display="grid";
  rowMonths.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach((d,i)=>{
    const isFirstOfMonth=d.getDate()===1 || i===0;
    const cell=document.createElement("div");
    cell.className="cell-day";
    if(isFirstOfMonth) cell.classList.add("month-start");
    if(toYMD(d) === todayYMD) cell.classList.add("today");
    cell.style.fontWeight=isFirstOfMonth?"600":"400";
    // Use only the short month name (no year) when this cell marks
    // the beginning of a month.  For example, on 1st February it will
    // show "fev".  Leaving other cells blank preserves alignment.
    cell.textContent=isFirstOfMonth? d.toLocaleDateString(undefined,{month:"short"}):"";
    rowMonths.appendChild(cell);
  });
  const rowDays=document.createElement("div");
  rowDays.style.display="grid";
  rowDays.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach(d=>{
    const cell=document.createElement("div");
    cell.className="cell-day";
    if(d.getDate()===1) cell.classList.add("month-start");
    if(toYMD(d) === todayYMD) cell.classList.add("today");
    cell.textContent=String(d.getDate()).padStart(2,"0");
    // Marcar finais de semana com uma classe dedicada e aplicar uma cor
    // de fundo mais clara para distinguir visualmente sábados e domingos.
    const dow = d.getDay();
    const dayCtx = getDayContext(d, null);
    if((dow === 0 || dow === 6) && calendarPrefs.destacarFds !== false){
      cell.classList.add("weekend");
      // Cor de fundo mais suave para cabeçalho de finais de semana
      cell.style.background = "#f8fafc";
    }
    if(dayCtx.isHoliday && calendarPrefs.destacarFeriados !== false){
      cell.classList.add("holiday");
      cell.title = `${dayCtx.holiday.legend || 'Feriado'} (${dayCtx.ymd})`;
    }
    rowDays.appendChild(cell);
  });
  // Append the month row (showing only month names) and then the row of
  // day numbers.  The month row helps users understand which month the
  // numeric day sequence belongs to without showing the year.
  right.appendChild(rowMonths);
  right.appendChild(rowDays);
  header.appendChild(left); header.appendChild(right);
  gantt.appendChild(header);

  const byRes=Object.fromEntries(resources.map(r=>[r.id,[]]));
  filteredActs.forEach(a=>{ if(byRes[a.resourceId]) byRes[a.resourceId].push(a); });
  const dailyLoadIndex = buildDailyLoadIndex(filteredActs, days);
  Object.keys(byRes).forEach(k=>byRes[k].sort((a,b)=>fromYMD(a.inicio)-fromYMD(b.inicio)));

  // Ordena os recursos alfabeticamente (apenas a ordem de exibição).
  // Não altera filtros/regras de exibição.
  const ganttResources = (resources || []).slice().sort((a,b)=>{
    const an = (a && a.nome != null) ? String(a.nome) : '';
    const bn = (b && b.nome != null) ? String(b.nome) : '';
    return an.localeCompare(bn, undefined, { sensitivity: 'base' });
  });

  ganttResources.forEach(r=>{
    // Aplicar filtros básicos: tipo (Interno/Externo), senioridade, ativo e busca por nome do recurso.
    if(filtroTipo && (r.tipo||"").toLowerCase() !== filtroTipo.toLowerCase()) return;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return;
    if(!r.ativo) return;
    if(buscaRecurso && !(r.nome||'').toLowerCase().includes(buscaRecurso)) return;
    const acts=byRes[r.id]||[];
    // Se após aplicar todos os filtros de atividades não restaram tarefas para este recurso,
    // não exibir a linha no Gantt. Isso garante que somente recursos com atividades
    // correspondentes ao filtro apareçam.
    if(acts.length===0) return;

    const row=document.createElement("div");
    row.className="row";

    const info=document.createElement("div");
    info.className="info";
    info.innerHTML=`<div style="font-weight:600">${escHTML(r.nome)}</div>
      <div class="muted" style="font-size:12px">${escHTML(r.tipo)} • ${escHTML(r.senioridade)} • Cap: ${r.capacidade}% • ${getDailyHours(r)}h/dia${(r.inicioAtivo||r.fimAtivo)? " • Janela: " + (escHTML(r.inicioAtivo||"…")) + " → " + (escHTML(r.fimAtivo||"…")) : ""}</div>`;

    const bargrid=document.createElement("div");
    bargrid.className="bargrid";

    const cap=r.capacidade||100;
    days.forEach((d,i)=>{
      const dy=toYMD(d);
      const dayCtx = getDayContext(d, r);
      const dayLoad = (dailyLoadIndex[r.id] && dailyLoadIndex[r.id][dy]) || { activeActs:[], allocatedPct:0, nominalHours:getDailyHours(r), capacityHours:getDailyCapacityHours(r, d), allocatedHours:0, availableHours:getDailyCapacityHours(r, d), excessHours:0 };
      const activeActs = dayLoad.activeActs || [];
      const sum=dayLoad.allocatedPct||0;
      const perc=cap? (sum/cap)*100 : 0;
      const heat=document.createElement("div");
      heat.className="heatcell";
      heat.style.left=`${i*28}px`; heat.style.width="28px";
      // Garante que os blocos de ocupação fiquem acima do plano de fundo da grade
      heat.style.zIndex = "1";
      // A classe "weekend" não é aplicada aos blocos de ocupação (heatcell).
      // A marcação visual de finais de semana é feita apenas no fundo da grade (gridBg).
      if(perc>100) heat.classList.add("heat-over");
      else if(perc>0) heat.classList.add(perc>70?"heat-high":"heat-ok");
      heat.onmouseenter=(ev)=>{
        const sorted = (activeActs || []).slice().sort((a,b)=>((b.alocacao||100)-(a.alocacao||100)));
        const used = sorted.reduce((acc,a)=>acc+(a.alocacao||100),0);
        const freeAbs = (cap - used);
        const over = used > cap;
        const usedPct = cap ? Math.round((used/cap)*100) : 0;
        const freeShown = Math.max(0, Math.round(freeAbs));
        const overShown = Math.max(0, Math.round(used - cap));
        const nominalHours = dayLoad.nominalHours ?? getDailyHours(r);
        const capacityHours = dayLoad.capacityHours ?? getDailyCapacityHours(r, d);
        const allocatedHours = dayLoad.allocatedHours || 0;
        const availableHours = dayLoad.availableHours || 0;
        const excessHours = dayLoad.excessHours || 0;
        const overtimeApproved = dayLoad.overtimeApproved || getApprovedOvertimeHours(r, d);

        const showActs = over ? sorted.slice(0,3) : sorted;
        const rows = showActs.map(a=>`<div class="t-row"><strong>${escHTML(a.titulo)}</strong> — ${a.alocacao||100}% • ${getActivityAllocatedHours(a, r).toFixed(1)}h (${escHTML(a.status)})</div>`).join("");
        const more = over && sorted.length>3 ? `<div class="muted" style="margin-top:6px">+ ${sorted.length-3} outras atividades</div>` : "";
        const extra = over ? ` • ⚠ Excedente: +${overShown}% / ${excessHours.toFixed(1)}h` : "";
        const tipoDia = dayCtx.isHoliday ? `Feriado — ${dayCtx.holiday.legend || 'sem descrição'}` : (dayCtx.isWeekend ? 'Final de semana' : 'Dia útil');
        tooltip.innerHTML = `<div class="t-title">${escHTML(r.nome)} — ${escHTML(dy)}</div><div class="muted">Tipo do dia: ${escHTML(tipoDia)}</div><div class="muted">Jornada do dia: ${nominalHours.toFixed(1)}h • Horas extras aprovadas: ${overtimeApproved.toFixed(1)}h • Capacidade útil: ${capacityHours.toFixed(1)}h</div><div class="muted">Horas alocadas: ${allocatedHours.toFixed(1)}h • ${excessHours>0 ? `Excesso: ${excessHours.toFixed(1)}h` : `Horas disponíveis: ${availableHours.toFixed(1)}h`}</div><div class="muted">Cap: ${cap}% • Usado: ${Math.round(used)}% • Livre: ${freeShown}% • Ocupação: ${usedPct}% • Concorrência: ${activeActs.length}${extra}</div>${rows}${more}`;
        tooltip.classList.remove("hidden");
      };
      heat.onmousemove=(ev)=>{ tooltip.style.left = (ev.clientX+12)+"px"; tooltip.style.top = (ev.clientY+12)+"px"; };
      heat.onmouseleave=()=>{ tooltip.classList.add("hidden"); };
      bargrid.appendChild(heat);
      const c=activeActs.length;
      if(c>1){
        const cc=document.createElement("div");
        cc.className="ccell "+(c>=4?"high":(c>=3?"med":"low"));
        cc.style.left=`${i*28}px`; cc.style.width="28px";
        cc.textContent=String(c);
        cc.title=`${c} atividades simultâneas`;
        // Garante que o contador fique acima do plano de fundo da grade
        cc.style.zIndex = "1";
        bargrid.appendChild(cc);
      }
    });

    const daysLen = days.length;
    const startBase = fromYMD(rangeStart);
    function dayIndex(ymd){
      return Math.max(0, Math.min(daysLen-1, diffDays(fromYMD(ymd), startBase)));
    }
    const intervals = acts.map(a=>{
      const sIdx = dayIndex(a.inicio);
      const eIdx = dayIndex(a.fim);
      return {a, sIdx, eIdx};
    }).filter(iv=>iv.eIdx>=0 && iv.sIdx<=daysLen-1);
    const lanes = [];
    const placed = intervals.sort((x,y)=>x.sIdx - y.sIdx).map(iv=>{
      let lane = 0;
      while(lane < lanes.length && !(lanes[lane] < iv.sIdx - 0)) lane++;
      if(lane === lanes.length) lanes.push(-Infinity);
      lanes[lane] = iv.eIdx;
      return {...iv, lane};
    });

    const busy=acts.map(a=>({inicio:toYMD(new Date(Math.max(fromYMD(a.inicio),fromYMD(rangeStart)))),
                             fim:toYMD(new Date(Math.min(fromYMD(a.fim),fromYMD(rangeEnd))))}))
                  .filter(x=>fromYMD(x.inicio)<=fromYMD(x.fim));
    const gaps=invertIntervals(busy,rangeStart,rangeEnd);
    gaps.forEach(g=>{
      const startIdx=Math.max(0,diffDays(fromYMD(g.inicio),fromYMD(rangeStart)));
      const endIdx=Math.min(days.length-1,diffDays(fromYMD(g.fim),fromYMD(rangeStart)));
      const el=document.createElement("div");
      el.className="gapblock";
      el.style.left=`${startIdx*28}px`;
      el.style.width=`${(endIdx-startIdx+1)*28}px`;
      el.title=`Lacuna: ${g.inicio} → ${g.fim}`;
      // As lacunas devem ficar acima do plano de fundo da grade
      el.style.zIndex = "1";
      bargrid.appendChild(el);
    });

    placed.forEach(p=>{
      const a=p.a;
      if(!selectedStatus.has((a.status||"").trim())) return;
      const aStart=Math.max(0,p.sIdx);
      const aEnd=Math.min(days.length-1,p.eIdx);
      const b=document.createElement("div");
      b.className=`activity ${statusClass(a.status)}`;
      b.style.left=`${aStart*28}px`;
      b.style.width=`${(aEnd-aStart+1)*28}px`;
      b.style.top=`${p.lane*22 + 2}px`;
      b.textContent=a.titulo;
      const linkedTitle = a.linkedOriginId ? activities.find(x=>x.id===a.linkedOriginId) : null;
      const linkedTitleLabel = linkedTitle ? formatActivityLinkOption(linkedTitle) : String(a.linkedOriginCode || '').trim();
      const extrapolated = isActivityExtrapolated(a);
      if(extrapolated) b.classList.add('is-extrapolated');
      b.title=`${escHTML(a.codigoAtividade || "")}${a.codigoAtividade ? " — " : ""}${escHTML(a.titulo)} — ${a.inicio} → ${a.fim} • ${escHTML(a.status)} • ${a.alocacao||100}%${extrapolated ? " • Extrapolada" : ""}${linkedTitleLabel ? " • Vinculada a " + escHTML(linkedTitleLabel) : ""}${a.comentarios ? " • Comentários: " + escHTML(encodeInlineText(a.comentarios)) : ""}`;
      b.onclick=()=>{
        dlgAtividadeTitulo.textContent="Editar Atividade";
        fillRecursoOptions();
        formAtividade.elements["id"].value=a.id;
        if(formAtividade.elements["codigoAtividade"]) formAtividade.elements["codigoAtividade"].value=a.codigoAtividade||"";
        formAtividade.elements["titulo"].value=a.titulo;
        // Preenche campo de recurso com o nome correspondente
        const recObj = resources.find(x=>x.id===a.resourceId);
        formAtividade.elements["resourceName"].value = recObj ? recObj.nome : '';
        formAtividade.elements["inicio"].value=a.inicio;
        formAtividade.elements["fim"].value=a.fim;
        formAtividade.elements["status"].value=a.status;
        formAtividade.elements["alocacao"].value=a.alocacao||100;
        loadCommentDraftIntoForm(a.id, currentUser || '');
        refreshActivityCommentsPanel(a);
        formAtividade.elements["tags"].value = (a.tags || []).join(', ');
        if(formAtividade.elements["linkedOriginId"]) formAtividade.elements["linkedOriginId"].value = a.linkedOriginId || '';
        if(formAtividade.elements["linkedOriginSearch"]) formAtividade.elements["linkedOriginSearch"].value = '';
        fillActivityLinkOptions(a.linkedOriginId || '', '', a.id);
        dlgAtividade.showModal();
      };
      // As atividades devem estar sempre acima do plano de fundo da grade e de outros elementos
      b.style.zIndex = "2";
      bargrid.appendChild(b);
    });

    const lanesH = Math.max(42, lanes.length*22 + 6);
    bargrid.style.height = lanesH + "px";

    const gridBg=document.createElement("div");
    gridBg.style.position="absolute"; gridBg.style.top="0"; gridBg.style.bottom="0"; gridBg.style.left="0"; gridBg.style.right="0";
    gridBg.style.display="grid"; gridBg.style.gridTemplateColumns=`repeat(${days.length}, 28px)`; gridBg.style.pointerEvents="none";
    // Coloca o fundo da grade atrás de outros elementos
    gridBg.style.zIndex = "0";
    for(let i=0;i<days.length;i++){
      const v=document.createElement("div");
      v.style.borderLeft="1px solid #f1f5f9";
      // Aplica uma tonalidade mais clara ao fundo das colunas de finais de semana
      const dt = days[i];
      const dwd = dt.getDay();
      const ctxDay = getDayContext(dt, r);
      v.classList.add('gridBg');
      if((dwd === 0 || dwd === 6) && calendarPrefs.destacarFds !== false){
        v.classList.add('weekend-bg','non-working-bg');
        v.style.background = "#f8fafc";
      }
      if(ctxDay.isHoliday && calendarPrefs.destacarFeriados !== false){
        v.classList.add('holiday-bg','non-working-bg');
        v.title = `${ctxDay.holiday.legend || 'Feriado'} (${ctxDay.ymd})`;
      }
      if(dt.getDate()===1 || i===0){
        v.classList.add("gantt-month-sep");
      }
      gridBg.appendChild(v);
    }
    bargrid.appendChild(gridBg);

    // Linha do "hoje" (se estiver dentro do range atual)
    if(todayIdx >= 0 && todayIdx < days.length){
      const ln = document.createElement("div");
      ln.className = "gantt-today-line";
      ln.style.left = `${todayIdx*28}px`;
      bargrid.appendChild(ln);
    }

    row.appendChild(info); row.appendChild(bargrid);
    gantt.appendChild(row);
  });
}

// ===== Filtragem de atividades =====
function getFilteredActivities(){
  return activities.filter(a=>{
    // Ignorar atividades excluídas
    if(a.deletedAt) return false;
    if(!selectedStatus.has((a.status||"").trim())) return false;
    const r=resources.find(x=>x.id===a.resourceId);
    if(!r) return false;
    // também ignorar se recurso estiver marcado como deletado
    if(r.deletedAt) return false;
    if(filtroTipo && (r.tipo||"").toLowerCase() !== filtroTipo.toLowerCase()) return false;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return false;
    if(buscaRecurso && !(r.nome||"").toLowerCase().includes(buscaRecurso)) return false;
    if(buscaTitulo && !a.titulo.toLowerCase().includes(buscaTitulo)) return false;
    if (buscaTag && (!a.tags || a.tags.length === 0 || !a.tags.some(t => t.toLowerCase().includes(buscaTag)))) {
        return false;
    }
    const s=fromYMD(a.inicio), e=fromYMD(a.fim);
    if(e<fromYMD(rangeStart) || s>fromYMD(rangeEnd)) return false;
    return true;
  });
}

// ===== Helpers para exportações =====
// Retorna lista de recursos que não foram marcados como excluídos.  Útil para
// exportação de dados completos, evitando incluir registros tombados no CSV/XLS.
function getActiveResources() {
  return (resources || []).filter(r => !r.deletedAt);
}

// Retorna lista de atividades que não foram marcadas como excluídas e cujo
// recurso associado também está ativo.  Evita exportar itens que foram
// excluídos ou cujos recursos foram removidos.
function getActiveActivities() {
  const activeResIds = new Set(getActiveResources().map(r => r.id));
  return (activities || []).filter(a => !a.deletedAt && activeResIds.has(a.resourceId));
}

/**
 * Garante que, se houver um banco de dados externo apontado, os dados em
 * memória estejam sincronizados com o arquivo no momento da exportação.
 * Verifica a última modificação do arquivo e, se for mais recente do que
 * __bdLastWrite, recarrega o conteúdo e reaplica snapshot/eventos. Caso
 * contrário, nada faz. Essa função é assíncrona porque a leitura do arquivo
 * utiliza a API File System Access.
 */
async function refreshFromBDIfNeeded() {
  if (!bdHandle) return;
  try {
    const file = await bdHandle.getFile();
    const lm = file.lastModified;
    if (typeof window !== 'undefined' && lm > (window.__bdLastWrite || 0)) {
      let parsed;
      const text = await file.text();
      if (bdFileExt === 'csv') {
      try {
        const existingFile = await bdHandle.getFile();
        const existingText = await existingFile.text();
        const persisted = parseCSVBDUnico(existingText);
        comments = mergeComentarios(persisted.comments || [], comments || []);
        rebuildCommentsIndex();
        syncAllActivityCommentFields();
      } catch(_e){}
        parsed = parseCSVBDUnico(text);
      } else {
        parsed = parseHTMLBDTables(text);
      }
      const newResources = (parsed.recursos || []).map(coerceResource);
      const newActivities = hydrateLoadedActivities(parsed.atividades || []);
      const newComments = hydrateLoadedComments(newActivities, parsed.comments || []);
      const newTrails = {};
      // Reconstrói o histórico vindo do BD apontado.
      // Compatível com:
      // - Modelo antigo: colunas activityId/timestamp/oldInicio/.../justificativa/user
      // - Modelo estendido: metadados em legend (JSON) e/ou colunas tipoEvento/entidadeTipo/entidadeId
      (parsed.historico || []).forEach(h => {
        const entityId = h.entidadeId || h.entityId || h.activityId || h.id;
        if (!entityId) return;

        // Metadados podem vir em colunas próprias OU em legend (JSON)
        let meta = null;
        try {
          if (h.legend && String(h.legend).trim().startsWith('{')) meta = JSON.parse(h.legend);
        } catch(e) { meta = null; }

        const type = h.tipoEvento || h.tipo || (meta && meta.type) || h.type || 'ALTERACAO_DATAS';
        const entityType = h.entidadeTipo || h.entityType || (meta && meta.entityType) || 'atividade';

        const entry = {
          ts: h.timestamp || h.ts || '',
          oldInicio: h.oldInicio || '',
          oldFim: h.oldFim || '',
          newInicio: h.newInicio || '',
          newFim: h.newFim || '',
          justificativa: h.justificativa || '',
          user: h.user || '',
          // Campos extras (delegação/exclusão)
          type,
          entityType,
          oldResourceId: h.recursoAnteriorId || (meta && meta.oldResourceId) || h.oldResourceId || '',
          oldResourceName: h.recursoAnteriorNome || (meta && meta.oldResourceName) || h.oldResourceName || '',
          newResourceId: h.recursoNovoId || (meta && meta.newResourceId) || h.newResourceId || '',
          newResourceName: h.recursoNovoNome || (meta && meta.newResourceName) || h.newResourceName || '',
          legend: h.legend || ''
        };

        if (!newTrails[entityId]) newTrails[entityId] = [];
        newTrails[entityId].push(entry);
      });
      // Atualiza arrays globais e persiste
      resources = newResources;
      activities = newActivities;
      trails = newTrails;
      ingestParsedAccess(parsed);
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      saveLS(LS.trail, trails);
      try {
        adoptCurrentStateAsPersistedBaseline();
      } catch (e) {
        console.error('Erro ao atualizar baseline após refresh BD', e);
      }
      // Atualiza marcador de última escrita
      window.__bdLastWrite = lm;
    }
  } catch (e) {
    console.warn('Falha ao sincronizar com BD antes da exportação', e);
  }
}

// ===== Exportações =====
function download(name, content, type="text/plain"){
  // Excel (Windows) costuma abrir CSV como ANSI/Windows-1252.
  // Para preservar acentuação corretamente, prefixamos CSV com BOM UTF-8 (\uFEFF).
  const isCsv = (type && String(type).toLowerCase().includes("text/csv")) || String(name||"").toLowerCase().endsWith(".csv");
  const normalizedContent = (isCsv && typeof content === "string" && content.charCodeAt(0) !== 0xFEFF)
    ? "\uFEFF" + content
    : content;
  const blob = new Blob([normalizedContent], { type: type || "text/plain" });

  // Em alguns ambientes "instalados" (PWA/Standalone), o download via <a download>
  // pode falhar silenciosamente. Para aumentar a compatibilidade (principalmente
  // no Windows/Chrome app), tentamos primeiro o File Picker (quando disponível).
  const fallbackAnchor = () => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 0);
  };

  // Tenta salvar via File System Access API (mais confiável em PWA)
  const trySavePicker = async () => {
    if (!("showSaveFilePicker" in window)) return false;
    try {
      const ext = (String(name||"").split(".").pop() || "").toLowerCase();
      const accept = type ? { [type]: ext ? [`.${ext}`] : [] } : undefined;

      const handle = await window.showSaveFilePicker({
        suggestedName: name,
        types: accept ? [{ description: "Arquivo", accept }] : undefined
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (e) {
      // AbortError = usuário cancelou; para outros erros, cai no fallback
      if (e && e.name === 'AbortError') return true;
      return false;
    }
  };

  // Tenta Share API (útil em alguns ambientes)
  const tryShare = async () => {
    if (!navigator.share) return false;
    try {
      const file = new File([blob], name, { type });
      if (navigator.canShare && !navigator.canShare({ files: [file] })) return false;
      await navigator.share({ files: [file], title: name });
      return true;
    } catch (e) {
      if (e && e.name === 'AbortError') return true;
      return false;
    }
  };

  (async () => {
    const saved = await trySavePicker();
    if (saved) return;
    const shared = await tryShare();
    if (shared) return;
    fallbackAnchor();
  })();
}

function toCSV(rows, headerOrder){
  const esc=v=>{
    if(v===null||v===undefined) return "";
    const s=String(v).replace(/"/g,'""');
    return /[",\n;]/.test(s)?`"${s}"`:s;
  };
  const header=headerOrder||Object.keys(rows[0]||{});
  const lines=[header.join(";")];
  rows.forEach(r=>{lines.push(header.map(h=>esc(r[h])).join(";"))});
  return lines.join("\n");
}

document.getElementById("btnExportCSV").onclick = async () => {
  // Garante que estamos sincronizados com o BD antes de exportar
  await refreshFromBDIfNeeded();
  // Seleciona apenas recursos e atividades não deletados
  const activeResources = getActiveResources();
  const activeActivities = getActiveActivities();
  const rec = activeResources.map(r => ({
    id: r.id,
    nome: r.nome,
    tipo: r.tipo,
    senioridade: r.senioridade,
    ativo: r.ativo,
    capacidade: r.capacidade || 100,
    cargaHorariaDiaria: getDailyHours(r),
    inicioAtivo: r.inicioAtivo || "",
    fimAtivo: r.fimAtivo || "",
    version: r.version || 1,
    updatedAt: r.updatedAt || 0,
    deletedAt: r.deletedAt || ""
  }));
  const atv = activeActivities.map(a => {
    const linked = a.linkedOriginId ? activeActivities.find(x => x.id === a.linkedOriginId) : null;
    return ({
      id: a.id,
      codigoAtividade: a.codigoAtividade || '',
      titulo: a.titulo,
      resourceId: a.resourceId,
      linkedOriginId: a.linkedOriginId || '',
      linkedOriginCode: linked ? (linked.codigoAtividade || '') : '',
      extrapolada: isActivityExtrapolated(a) ? 'S' : 'N',
      inicio: a.inicio,
      fim: a.fim,
      status: a.status,
      alocacao: a.alocacao || 100,
      tags: (a.tags || []).join(', '),
      horasPlanejadasExecucao: calcExecutionMetrics(a).horasPlanejadas.toFixed(2),
      horasRealizadasExecucao: calcExecutionMetrics(a).horasRealizadas.toFixed(2),
      horasRestantesExecucao: calcExecutionMetrics(a).horasRestantes.toFixed(2),
      previsaoFimExecucao: calcExecutionMetrics(a).previsaoFim || '',
      atrasoDiasExecucao: calcExecutionMetrics(a).atrasoDias || 0,
      version: a.version || 1,
      updatedAt: a.updatedAt || 0,
      deletedAt: a.deletedAt || ""
    });
  });
  download("recursos.csv", toCSV(rec, ["id","nome","tipo","senioridade","ativo","capacidade","cargaHorariaDiaria","inicioAtivo","fimAtivo","version","updatedAt","deletedAt"]), "text/csv;charset=utf-8");
  download("atividades.csv", toCSV(atv, ["id","codigoAtividade","titulo","resourceId","linkedOriginId","linkedOriginCode","extrapolada","inicio","fim","status","alocacao","tags","horasPlanejadasExecucao","horasRealizadasExecucao","horasRestantesExecucao","previsaoFimExecucao","atrasoDiasExecucao","version","updatedAt","deletedAt"]), "text/csv;charset=utf-8");
  alert("Exportados: recursos.csv e atividades.csv");
};

function tableHTML(name, rows, cols){
  const header = cols.map(c=>`<th>${c}</th>`).join("");
  const body = rows.map(r=>`<tr>${cols.map(c=>`<td>${(r[c]??"")}</td>`).join("")}</tr>`).join("");
  return `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns="http://www.w3.org/TR/REC-html40">
  <head><meta charset="utf-8">
  </head><body>
    <table border="1"><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table>
  </body></html>`;
}
document.getElementById("btnExportXLS").onclick = async () => {
  await refreshFromBDIfNeeded();
  const activeResources = getActiveResources();
  const activeActivities = getActiveActivities();
  const rec = activeResources.map(r => ({
    id: r.id,
    nome: r.nome,
    tipo: r.tipo,
    senioridade: r.senioridade,
    ativo: r.ativo,
    capacidade: r.capacidade || 100,
    cargaHorariaDiaria: getDailyHours(r),
    inicioAtivo: r.inicioAtivo || "",
    fimAtivo: r.fimAtivo || ""
  }));
  const atv = activeActivities.map(a => {
    const linked = a.linkedOriginId ? activeActivities.find(x => x.id === a.linkedOriginId) : null;
    return ({
      id: a.id,
      codigoAtividade: a.codigoAtividade || '',
      titulo: a.titulo,
      resourceId: a.resourceId,
      linkedOriginId: a.linkedOriginId || '',
      linkedOriginCode: linked ? (linked.codigoAtividade || '') : '',
      extrapolada: isActivityExtrapolated(a) ? 'S' : 'N',
      inicio: a.inicio,
      fim: a.fim,
      status: a.status,
      alocacao: a.alocacao || 100,
      tags: (a.tags || []).join(', ')
    });
  });
  download("recursos.xls", tableHTML("Recursos", rec, ["id","nome","tipo","senioridade","ativo","capacidade","cargaHorariaDiaria","inicioAtivo","fimAtivo"]), "application/vnd.ms-excel");
  download("atividades.xls", tableHTML("Atividades", atv, ["id","codigoAtividade","titulo","resourceId","linkedOriginId","linkedOriginCode","extrapolada","inicio","fim","status","alocacao", "tags"]), "application/vnd.ms-excel");
  alert("Exportados: recursos.xls e atividades.xls");
};

document.getElementById("btnExportPBI").onclick = async () => {
  await refreshFromBDIfNeeded();
  const rows = [];
  const start = fromYMD(rangeStart), end = fromYMD(rangeEnd);
  // Usa apenas recursos e atividades ativos (sem deletados)
  const activeResources = getActiveResources();
  const byId = Object.fromEntries(activeResources.map(r => [r.id, r]));
  getActiveActivities().forEach(a => {
    const s = fromYMD(a.inicio), e = fromYMD(a.fim);
    const r = byId[a.resourceId];
    if (!r) return;
    for (let d = new Date(Math.max(s, start)); d <= Math.min(e, end); d = addDays(d, 1)) {
      const linked = a.linkedOriginId ? getActiveActivities().find(x => x.id === a.linkedOriginId) : null;
      rows.push({
        data: toYMD(d),
        atividadeId: a.id,
        atividadeCodigo: a.codigoAtividade || "",
        atividadeTitulo: a.titulo,
        atividadeVinculadaA: linked ? (linked.codigoAtividade || linked.id || "") : "",
        atividadeVinculadaId: a.linkedOriginId || "",
        extrapolada: isActivityExtrapolated(a) ? 'S' : 'N',
        status: a.status,
        alocacao: a.alocacao || 100,
        tags: (a.tags || []).join('|'),
        recursoId: r.id,
        recursoNome: r.nome,
        recursoTipo: r.tipo,
        recursoSenioridade: r.senioridade,
        recursoCapacidade: r.capacidade || 100,
        recursoCargaHorariaDiaria: getDailyHours(r)
      });
    }
  });
  download("powerbi_atividades_diarias.csv",
    toCSV(rows, ["data","atividadeId","atividadeCodigo","atividadeTitulo","atividadeVinculadaA","atividadeVinculadaId","extrapolada","status","alocacao","tags","recursoId","recursoNome","recursoTipo","recursoSenioridade","recursoCapacidade","recursoCargaHorariaDiaria"]),
    "text/csv;charset=utf-8");
  alert(`Exportado: powerbi_atividades_diarias.csv (${rows.length} linhas)`);
};



function buildExecutionSubactivityTypeKpiRows(){
  const activeActivities = getActiveActivities();
  const agg = {};
  activeActivities.forEach(a => {
    const data = normalizeExecData(a);
    data.subtasks.forEach(st => {
      const tipo = st.tipoSubatividade || 'Outros';
      const entries = data.entries.filter(en => en.subtaskId === st.id);
      const horasRealizadas = entries.reduce((acc,en)=>acc+(Number(en.horas)||0),0);
      const horasPlanejadas = Number(st.horasPlanejadas)||0;
      if(!agg[tipo]) agg[tipo] = { tipoSubatividade:tipo, subatividades:0, quantidadeExecutada:0, horasPlanejadasTotais:0, horasRealizadasTotais:0, valoresRealizados:[] };
      agg[tipo].subatividades += 1;
      if(horasRealizadas > 0) agg[tipo].quantidadeExecutada += 1;
      agg[tipo].horasPlanejadasTotais += horasPlanejadas;
      agg[tipo].horasRealizadasTotais += horasRealizadas;
      if(horasRealizadas > 0) agg[tipo].valoresRealizados.push(horasRealizadas);
    });
  });
  return Object.values(agg).map(r => {
    const stats = execStats(r.valoresRealizados || []);
    return {
      tipoSubatividade: r.tipoSubatividade,
      subatividades: r.subatividades,
      quantidadeExecutada: r.quantidadeExecutada,
      horasPlanejadasTotais: r.horasPlanejadasTotais.toFixed(2),
      horasRealizadasTotais: r.horasRealizadasTotais.toFixed(2),
      mediaHorasRealizadas: stats.media.toFixed(2),
      medianaHorasRealizadas: stats.mediana.toFixed(2),
      desvioHorasRealizadas: stats.desvio.toFixed(2),
      percentualExecucao: (r.horasPlanejadasTotais > 0 ? (r.horasRealizadasTotais / r.horasPlanejadasTotais) * 100 : 0).toFixed(2)
    };
  }).sort((a,b)=>Number(b.mediaHorasRealizadas)-Number(a.mediaHorasRealizadas));
}

function buildExecutionSubactivityIndicatorRows(){
  const activeResources = getActiveResources();
  const resById = Object.fromEntries(activeResources.map(r => [r.id, r]));
  const activeActivities = getActiveActivities();
  const actById = Object.fromEntries(activeActivities.map(a => [a.id, a]));
  const rows = [];

  activeActivities.forEach(a => {
    const data = normalizeExecData(a);
    if(!data.subtasks.length) return;
    const actMetrics = calcExecutionMetrics(a);
    const mainResource = resById[a.resourceId];
    data.subtasks.forEach(st => {
      const resp = resById[st.responsavelId] || mainResource || {};
      const entries = data.entries.filter(en => en.subtaskId === st.id);
      const horasRealizadas = entries.reduce((acc, en) => acc + (Number(en.horas) || 0), 0);
      const horasPlanejadas = Number(st.horasPlanejadas) || 0;
      const saldoHoras = Math.max(0, horasPlanejadas - horasRealizadas);
      const percentualExecucao = horasPlanejadas > 0 ? Math.min(999, (horasRealizadas / horasPlanejadas) * 100) : 0;
      const ultimoApontamento = entries.slice().sort((x,y)=>String(y.data || '').localeCompare(String(x.data || '')))[0] || null;
      const statusSubatividade = saldoHoras <= 0 ? 'Concluída' : (horasRealizadas > 0 ? 'Em execução' : 'Não iniciada');
      const atrasoPrevisto = st.fimPrevisto && actMetrics.previsaoFim ? (fromYMD(actMetrics.previsaoFim) > fromYMD(st.fimPrevisto) ? 'S' : 'N') : '';
      rows.push({
        atividadeId: a.id,
        atividadeCodigo: a.codigoAtividade || '',
        atividadeTitulo: a.titulo || '',
        atividadeStatus: a.status || '',
        atividadeInicio: a.inicio || '',
        atividadeFimPlanejado: a.fim || '',
        atividadePrevisaoFim: actMetrics.previsaoFim || '',
        atividadeAtrasoDias: actMetrics.atrasoDias || 0,
        recursoPrincipalId: a.resourceId || '',
        recursoPrincipalNome: mainResource?.nome || '',
        subatividadeId: st.id,
        subatividadeTipo: st.tipoSubatividade || 'Outros',
        subatividadeTitulo: st.titulo || '',
        responsavelId: st.responsavelId || '',
        responsavelNome: resp.nome || '',
        fimPrevistoSubatividade: st.fimPrevisto || '',
        statusSubatividade,
        horasPlanejadas: horasPlanejadas.toFixed(2),
        horasRealizadas: horasRealizadas.toFixed(2),
        saldoHoras: saldoHoras.toFixed(2),
        percentualExecucao: percentualExecucao.toFixed(2),
        quantidadeApontamentos: entries.length,
        ultimoApontamentoData: ultimoApontamento?.data || '',
        ultimoApontamentoComentario: ultimoApontamento?.comentario || '',
        atrasoPrevisto,
        criadoEm: st.createdAt ? new Date(st.createdAt).toISOString() : '',
        criadoPor: st.createdBy || ''
      });
    });
  });

  return rows;
}

const btnExportExecSub = document.getElementById('btnExportExecSub');
if(btnExportExecSub){
  btnExportExecSub.onclick = async () => {
    await refreshFromBDIfNeeded();
    const rows = buildExecutionSubactivityTypeKpiRows();
    const cols = [
      'tipoSubatividade','subatividades','quantidadeExecutada','horasPlanejadasTotais','horasRealizadasTotais',
      'mediaHorasRealizadas','medianaHorasRealizadas','desvioHorasRealizadas','percentualExecucao'
    ];
    download('indicadores_execucao_tipos_subatividade.csv', toCSV(rows, cols), 'text/csv;charset=utf-8');
    alert(`Exportado: indicadores_execucao_tipos_subatividade.csv (${rows.length} tipos)`);
  };
}


function getExecutionActivityType(activity){
  const tags = Array.isArray(activity?.tags) ? activity.tags : String(activity?.tags || '').split(',');
  const firstTag = tags.map(t => String(t || '').trim()).filter(Boolean)[0];
  if(firstTag) return normalizeTag(firstTag);
  return String(activity?.status || 'Sem tipo/tag').trim() || 'Sem tipo/tag';
}
function execPct(n,d){
  return d > 0 ? (n / d) * 100 : 0;
}
function execNum(value, decimals=1){
  const n = Number(value) || 0;
  return n.toLocaleString('pt-BR', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
function groupExecutionReportData(){
  const activeActivities = getActiveActivities();
  const resById = Object.fromEntries(getActiveResources().map(r => [r.id, r]));
  const activityRows = [];
  const subRows = [];
  const issueRows = [];

  activeActivities.forEach(a => {
    const data = normalizeExecData(a);
    if(!data.subtasks.length && !data.entries.length && !data.issues.length) return;
    const m = calcExecutionMetrics(a);
    const type = getExecutionActivityType(a);
    activityRows.push({
      id: a.id,
      codigo: a.codigoAtividade || '',
      titulo: a.titulo || '',
      tipo: type,
      status: a.status || '',
      recurso: resById[a.resourceId]?.nome || '',
      inicio: a.inicio || '',
      fimPlanejado: a.fim || '',
      previsaoFim: m.previsaoFim || '',
      atrasoDias: Number(m.atrasoDias) || 0,
      horasPlanejadas: Number(m.horasPlanejadas) || 0,
      horasRealizadas: Number(m.horasRealizadas) || 0,
      horasRestantes: Number(m.horasRestantes) || 0,
      impactoHoras: Number(m.impactoHoras) || 0,
      impactoPrazoDias: Number(m.impactoPrazoDias) || 0,
      percentualReal: Number(m.percentualReal) || 0,
      qtdSubatividades: data.subtasks.length,
      qtdApontamentos: data.entries.length,
      qtdOcorrencias: data.issues.length
    });

    data.subtasks.forEach(st => {
      const entries = data.entries.filter(en => en.subtaskId === st.id);
      const horasRealizadas = entries.reduce((acc,en)=>acc+(Number(en.horas)||0),0);
      const horasPlanejadas = Number(st.horasPlanejadas)||0;
      subRows.push({
        atividadeId: a.id,
        atividadeCodigo: a.codigoAtividade || '',
        atividadeTitulo: a.titulo || '',
        tipo: st.tipoSubatividade || 'Outros',
        subatividade: st.titulo || '',
        responsavel: resById[st.responsavelId]?.nome || resById[a.resourceId]?.nome || '',
        status: st.status || '',
        fimPrevisto: st.fimPrevisto || '',
        horasPlanejadas,
        horasRealizadas,
        saldo: Math.max(0, horasPlanejadas - horasRealizadas),
        percentual: execPct(horasRealizadas, horasPlanejadas),
        apontamentos: entries.length
      });
    });

    data.issues.forEach(is => {
      issueRows.push({
        atividadeId: a.id,
        atividadeCodigo: a.codigoAtividade || '',
        atividadeTitulo: a.titulo || '',
        tipoAtividade: type,
        tipo: is.tipo || 'Ocorrência',
        descricao: is.descricao || '',
        data: is.data || '',
        impactoHoras: Number(is.impactoHoras)||0,
        impactoPrazoDias: Number(is.impactoPrazoDias)||0
      });
    });
  });

  const byType = {};
  activityRows.forEach(r => {
    const key = r.tipo || 'Sem tipo/tag';
    byType[key] ||= { tipo:key, atividades:0, horasPlanejadas:0, horasRealizadas:0, horasRestantes:0, atrasoDias:0, ocorrencias:0 };
    byType[key].atividades += 1;
    byType[key].horasPlanejadas += r.horasPlanejadas;
    byType[key].horasRealizadas += r.horasRealizadas;
    byType[key].horasRestantes += r.horasRestantes;
    byType[key].atrasoDias += r.atrasoDias;
    byType[key].ocorrencias += r.qtdOcorrencias;
  });
  const typeRows = Object.values(byType).map(r => ({
    ...r,
    mediaHorasRealizadas: r.atividades ? r.horasRealizadas / r.atividades : 0,
    mediaHorasPlanejadas: r.atividades ? r.horasPlanejadas / r.atividades : 0,
    mediaAtrasoDias: r.atividades ? r.atrasoDias / r.atividades : 0,
    percentualExecucao: execPct(r.horasRealizadas, r.horasPlanejadas)
  })).sort((a,b)=>b.mediaHorasRealizadas-a.mediaHorasRealizadas);

  const byIssueType = {};
  issueRows.forEach(r => {
    const key = r.tipo || 'Ocorrência';
    byIssueType[key] ||= { tipo:key, quantidade:0, impactoHoras:0, impactoPrazoDias:0 };
    byIssueType[key].quantidade += 1;
    byIssueType[key].impactoHoras += r.impactoHoras;
    byIssueType[key].impactoPrazoDias += r.impactoPrazoDias;
  });
  const issueTypeRows = Object.values(byIssueType).map(r => ({
    ...r,
    mediaImpactoHoras: r.quantidade ? r.impactoHoras / r.quantidade : 0,
    mediaImpactoPrazoDias: r.quantidade ? r.impactoPrazoDias / r.quantidade : 0
  })).sort((a,b)=>b.impactoHoras-a.impactoHoras);

  return { activityRows, subRows, issueRows, typeRows, issueTypeRows };
}
function buildExecBarChart(rows, labelKey, valueKey, title, suffix=''){
  const top = rows.slice(0, 8);
  const max = Math.max(1, ...top.map(r => Number(r[valueKey]) || 0));
  const bars = top.map((r,i)=>{
    const value = Number(r[valueKey]) || 0;
    const w = Math.max(2, (value / max) * 100);
    return `<div class="bar-row"><div class="bar-label" title="${escHTML(r[labelKey] || '')}">${escHTML(r[labelKey] || '')}</div><div class="bar-track"><div class="bar-fill" style="width:${w.toFixed(1)}%"></div></div><div class="bar-value">${execNum(value,1)}${suffix}</div></div>`;
  }).join('') || '<p class="muted">Sem dados para gráfico.</p>';
  return `<section class="chart-card"><h3>${escHTML(title)}</h3>${bars}</section>`;
}
function buildExecReportHtml(data){
  const totals = data.activityRows.reduce((acc,r)=>{
    acc.atividades += 1;
    acc.subatividades += r.qtdSubatividades;
    acc.apontamentos += r.qtdApontamentos;
    acc.ocorrencias += r.qtdOcorrencias;
    acc.planejadas += r.horasPlanejadas;
    acc.realizadas += r.horasRealizadas;
    acc.saldo += r.horasRestantes;
    acc.impacto += r.impactoHoras;
    acc.atrasos += r.atrasoDias > 0 ? 1 : 0;
    return acc;
  }, {atividades:0,subatividades:0,apontamentos:0,ocorrencias:0,planejadas:0,realizadas:0,saldo:0,impacto:0,atrasos:0});
  const pctExec = execPct(totals.realizadas, totals.planejadas + totals.impacto);
  const generatedAt = new Date().toLocaleString('pt-BR');
  const typeRows = data.typeRows.slice(0, 12);
  const issueRows = data.issueTypeRows.slice(0, 12);
  const activityRows = data.activityRows.slice().sort((a,b)=>(b.atrasoDias-a.atrasoDias) || (b.horasRestantes-a.horasRestantes)).slice(0,20);
  const issueDetailRows = data.issueRows.slice().sort((a,b)=>(b.impactoHoras-a.impactoHoras) || String(b.data).localeCompare(String(a.data))).slice(0,30);

  const typeTable = typeRows.map(r=>`<tr><td>${escHTML(r.tipo)}</td><td>${r.atividades}</td><td>${execNum(r.mediaHorasPlanejadas,1)}h</td><td>${execNum(r.mediaHorasRealizadas,1)}h</td><td>${execNum(r.percentualExecucao,1)}%</td><td>${execNum(r.mediaAtrasoDias,1)}</td></tr>`).join('') || '<tr><td colspan="6">Sem dados</td></tr>';
  const issueTable = issueRows.map(r=>`<tr><td>${escHTML(r.tipo)}</td><td>${r.quantidade}</td><td>${execNum(r.impactoHoras,1)}h</td><td>${execNum(r.mediaImpactoHoras,1)}h</td><td>${execNum(r.impactoPrazoDias,1)}d</td><td>${execNum(r.mediaImpactoPrazoDias,1)}d</td></tr>`).join('') || '<tr><td colspan="6">Sem ocorrências</td></tr>';
  const actTable = activityRows.map(r=>`<tr><td>${escHTML(r.codigo)}</td><td>${escHTML(r.titulo)}</td><td>${escHTML(r.tipo)}</td><td>${execNum(r.percentualReal,1)}%</td><td>${execNum(r.horasRealizadas,1)}h</td><td>${execNum(r.horasRestantes,1)}h</td><td>${escHTML(r.fimPlanejado||'')}</td><td>${escHTML(r.previsaoFim||'')}</td><td>${r.atrasoDias}</td></tr>`).join('') || '<tr><td colspan="9">Sem atividades com execução</td></tr>';
  const issueDetailTable = issueDetailRows.map(r=>`<tr><td>${escHTML(r.tipo)}</td><td>${escHTML(r.atividadeCodigo)}</td><td>${escHTML(r.atividadeTitulo)}</td><td>${escHTML(r.data)}</td><td>${execNum(r.impactoHoras,1)}h</td><td>${execNum(r.impactoPrazoDias,1)}d</td><td>${escHTML(r.descricao)}</td></tr>`).join('') || '<tr><td colspan="7">Sem ocorrências</td></tr>';

  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Relatório de Indicadores de Execução</title>
  <style>
    body{font-family:Arial,Helvetica,sans-serif;color:#0f172a;margin:24px;background:#f8fafc} h1{margin:0 0 6px;font-size:24px} h2{margin:22px 0 10px;font-size:18px} h3{margin:0 0 10px;font-size:15px}.muted{color:#64748b}.top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:18px}.print{border:1px solid #2563eb;background:#2563eb;color:white;border-radius:10px;padding:9px 14px;cursor:pointer}.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:16px 0}.kpi{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:14px}.kpi strong{display:block;font-size:24px;margin-bottom:4px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.chart-card{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:14px;margin-bottom:14px}.bar-row{display:grid;grid-template-columns:150px 1fr 70px;gap:10px;align-items:center;margin:8px 0}.bar-label{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.bar-track{height:14px;background:#e2e8f0;border-radius:999px;overflow:hidden}.bar-fill{height:100%;background:#2563eb;border-radius:999px}.bar-value{font-size:12px;text-align:right}table{width:100%;border-collapse:collapse;background:white;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;margin-bottom:16px}th,td{font-size:12px;text-align:left;padding:8px;border-bottom:1px solid #e2e8f0;vertical-align:top}th{background:#eef2ff;font-weight:700}.section{page-break-inside:avoid}@media(max-width:900px){.kpis{grid-template-columns:repeat(2,1fr)}.grid{grid-template-columns:1fr}.top{display:block}.bar-row{grid-template-columns:110px 1fr 60px}}@media print{body{background:white;margin:12mm}.print{display:none}.chart-card,.kpi,table{break-inside:avoid}}
  </style></head><body>
  <div class="top"><div><h1>Relatório de Indicadores de Execução</h1><div class="muted">Gerado em ${escHTML(generatedAt)} • fonte: atividades, subatividades, apontamentos e ocorrências cadastradas no sistema</div></div><button class="print" onclick="window.print()">Exportar / Salvar em PDF</button></div>
  <div class="kpis"><div class="kpi"><strong>${totals.atividades}</strong><span>atividades com execução</span></div><div class="kpi"><strong>${execNum(pctExec,1)}%</strong><span>execução global</span></div><div class="kpi"><strong>${execNum(totals.realizadas,1)}h</strong><span>horas realizadas</span></div><div class="kpi"><strong>${execNum(totals.saldo,1)}h</strong><span>saldo restante</span></div><div class="kpi"><strong>${totals.subatividades}</strong><span>subatividades</span></div><div class="kpi"><strong>${totals.apontamentos}</strong><span>apontamentos</span></div><div class="kpi"><strong>${totals.ocorrencias}</strong><span>ocorrências</span></div><div class="kpi"><strong>${execNum(totals.impacto,1)}h</strong><span>impacto total</span></div></div>
  <div class="grid">${buildExecBarChart(typeRows,'tipo','mediaHorasRealizadas','Tempo médio realizado por tipo','h')}${buildExecBarChart(issueRows,'tipo','impactoHoras','Principais ocorrências por impacto','h')}</div>
  <section class="section"><h2>Tempo médio de execução por tipo</h2><table><thead><tr><th>Tipo</th><th>Atividades</th><th>Média planejada</th><th>Média realizada</th><th>% execução</th><th>Média atraso (dias)</th></tr></thead><tbody>${typeTable}</tbody></table></section>
  <section class="section"><h2>Principais ocorrências e impactos</h2><table><thead><tr><th>Ocorrência</th><th>Qtd.</th><th>Impacto total horas</th><th>Média horas</th><th>Impacto total prazo</th><th>Média prazo</th></tr></thead><tbody>${issueTable}</tbody></table></section>
  <section class="section"><h2>Atividades com maior atenção</h2><table><thead><tr><th>Código</th><th>Atividade</th><th>Tipo</th><th>% real</th><th>Horas realizadas</th><th>Saldo</th><th>Término planejado</th><th>Previsão atual</th><th>Atraso dias</th></tr></thead><tbody>${actTable}</tbody></table></section>
  <section class="section"><h2>Detalhamento das ocorrências críticas</h2><table><thead><tr><th>Tipo</th><th>Código</th><th>Atividade</th><th>Data</th><th>Impacto horas</th><th>Impacto prazo</th><th>Descrição</th></tr></thead><tbody>${issueDetailTable}</tbody></table></section>
  </body></html>`;
}
function exportExecutionIndicatorsPDF(){
  const data = groupExecutionReportData();
  const html = buildExecReportHtml(data);
  const w = window.open('', '_blank');
  if(!w){
    download('relatorio_indicadores_execucao.html', html, 'text/html;charset=utf-8');
    alert('Não foi possível abrir uma nova janela. Foi gerado um HTML para abrir e salvar como PDF.');
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
}


function parseExecReportArraySafe(value){
  try{
    if(Array.isArray(value)) return value;
    if(value == null || value === '') return [];
    let raw = String(value || '').trim();
    if(!raw) return [];
    try{ raw = decodeInlineText(raw); }catch(_e){}
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  }catch(_e){ return []; }
}
function getExecReportDataSafe(){
  const actList = Array.isArray(activities) ? activities.filter(a => !a.deletedAt) : [];
  const resList = Array.isArray(resources) ? resources.filter(r => !r.deletedAt) : [];
  const resById = Object.fromEntries(resList.map(r => [String(r.id || ''), r]));
  const activityRows = [];
  const issueRows = [];
  const issueAgg = {};
  const typeAgg = {};

  actList.forEach(a => {
    try{
      const subtasks = parseExecReportArraySafe(Array.isArray(a.execSubtasks) ? a.execSubtasks : a.execSubtasksJson);
      const entries = parseExecReportArraySafe(Array.isArray(a.execEntries) ? a.execEntries : a.execEntriesJson);
      const issues = parseExecReportArraySafe(Array.isArray(a.execIssues) ? a.execIssues : a.execIssuesJson);
      if(!subtasks.length && !entries.length && !issues.length) return;

      const tipo = 'Consolidado da atividade';
      const horasPlanejadasSub = subtasks.reduce((acc, st) => acc + (Number(String(st.horasPlanejadas ?? st.plannedHours ?? 0).replace(',', '.')) || 0), 0);
      const horasRealizadas = entries.reduce((acc, en) => acc + (Number(String(en.horas ?? en.hours ?? 0).replace(',', '.')) || 0), 0);
      const impactoHoras = issues.reduce((acc, is) => acc + (Number(String(is.impactoHoras ?? is.impactHours ?? 0).replace(',', '.')) || 0), 0);
      const horasPlanejadas = Number(String(a.horasPlanejadas ?? a.execHorasPlanejadas ?? 0).replace(',', '.')) || horasPlanejadasSub || 0;
      const horasRestantes = Math.max(0, horasPlanejadas + impactoHoras - horasRealizadas);
      const percentualReal = horasPlanejadas + impactoHoras > 0 ? (horasRealizadas / (horasPlanejadas + impactoHoras)) * 100 : 0;

      activityRows.push({
        codigo: String(a.codigoAtividade || ''),
        titulo: String(a.titulo || ''),
        tipo,
        recurso: String((resById[String(a.resourceId || '')] || {}).nome || ''),
        fimPlanejado: String(a.fim || ''),
        previsaoFim: String((a.controleExecucao && a.controleExecucao.previsaoFim) || a.fim || ''),
        horasPlanejadas,
        horasRealizadas,
        horasRestantes,
        impactoHoras,
        percentualReal,
        qtdSubatividades: subtasks.length,
        qtdApontamentos: entries.length,
        qtdOcorrencias: issues.length
      });

      subtasks.forEach(st => {
        const subId = String(st.id || st.subtaskId || '');
        const subTipo = getExecSubtaskType(st);
        const subEntries = entries.filter(en => String(en.subtaskId || en.subatividadeId || '') === subId);
        const subHorasRealizadas = subEntries.reduce((acc, en) => acc + (Number(String(en.horas ?? en.hours ?? 0).replace(',', '.')) || 0), 0);
        const subHorasPlanejadas = Number(String(st.horasPlanejadas ?? st.plannedHours ?? 0).replace(',', '.')) || 0;
        if(!typeAgg[subTipo]) typeAgg[subTipo] = { tipo: subTipo, subatividades:0, quantidadeExecutada:0, horasPlanejadas:0, horasRealizadas:0, valoresRealizados:[] };
        typeAgg[subTipo].subatividades += 1;
        if(subHorasRealizadas > 0) typeAgg[subTipo].quantidadeExecutada += 1;
        typeAgg[subTipo].horasPlanejadas += subHorasPlanejadas;
        typeAgg[subTipo].horasRealizadas += subHorasRealizadas;
        if(subHorasRealizadas > 0) typeAgg[subTipo].valoresRealizados.push(subHorasRealizadas);
      });

      issues.forEach(is => {
        const itipo = String(is.tipo || is.type || 'Ocorrência').trim() || 'Ocorrência';
        const ih = Number(String(is.impactoHoras ?? is.impactHours ?? 0).replace(',', '.')) || 0;
        const ip = Number(String(is.impactoPrazoDias ?? is.impactDays ?? 0).replace(',', '.')) || 0;
        issueRows.push({
          tipo: itipo,
          atividadeCodigo: String(a.codigoAtividade || ''),
          atividadeTitulo: String(a.titulo || ''),
          data: String(is.data || ''),
          impactoHoras: ih,
          impactoPrazoDias: ip,
          descricao: String(is.descricao || is.description || '')
        });
        if(!issueAgg[itipo]) issueAgg[itipo] = { tipo: itipo, quantidade:0, impactoHoras:0, impactoPrazoDias:0 };
        issueAgg[itipo].quantidade += 1;
        issueAgg[itipo].impactoHoras += ih;
        issueAgg[itipo].impactoPrazoDias += ip;
      });
    }catch(_rowErr){}
  });

  const typeRows = Object.values(typeAgg).map(r => {
    const stats = execStats(r.valoresRealizados || []);
    return {
      ...r,
      atividades: r.subatividades || 0,
      mediaHorasPlanejadas: r.subatividades ? r.horasPlanejadas / r.subatividades : 0,
      mediaHorasRealizadas: stats.media,
      medianaHorasRealizadas: stats.mediana,
      desvioHorasRealizadas: stats.desvio,
      totalHorasRealizadas: r.horasRealizadas || 0,
      percentualExecucao: r.horasPlanejadas > 0 ? (r.horasRealizadas / r.horasPlanejadas) * 100 : 0
    };
  }).sort((a,b)=>b.mediaHorasRealizadas-a.mediaHorasRealizadas);

  const issueTypeRows = Object.values(issueAgg).map(r => ({
    ...r,
    mediaImpactoHoras: r.quantidade ? r.impactoHoras / r.quantidade : 0,
    mediaImpactoPrazoDias: r.quantidade ? r.impactoPrazoDias / r.quantidade : 0
  })).sort((a,b)=>b.impactoHoras-a.impactoHoras);

  return { activityRows, issueRows, typeRows, issueTypeRows };
}
function buildExecReportHtmlSafe(data){
  data = data || { activityRows:[], issueRows:[], typeRows:[], issueTypeRows:[] };
  const activityRows = Array.isArray(data.activityRows) ? data.activityRows : [];
  const issueRows = Array.isArray(data.issueRows) ? data.issueRows : [];
  const typeRows = Array.isArray(data.typeRows) ? data.typeRows : [];
  const issueTypeRows = Array.isArray(data.issueTypeRows) ? data.issueTypeRows : [];
  const totals = activityRows.reduce((acc,r)=>{
    acc.atividades += 1;
    acc.planejadas += Number(r.horasPlanejadas)||0;
    acc.realizadas += Number(r.horasRealizadas)||0;
    acc.saldo += Number(r.horasRestantes)||0;
    acc.impacto += Number(r.impactoHoras)||0;
    acc.subatividades += Number(r.qtdSubatividades)||0;
    acc.apontamentos += Number(r.qtdApontamentos)||0;
    acc.ocorrencias += Number(r.qtdOcorrencias)||0;
    return acc;
  }, {atividades:0, planejadas:0, realizadas:0, saldo:0, impacto:0, subatividades:0, apontamentos:0, ocorrencias:0});
  const pct = totals.planejadas + totals.impacto > 0 ? (totals.realizadas / (totals.planejadas + totals.impacto)) * 100 : 0;
  const chartBars = (rows, label, value, suffix='h') => {
    const top = rows.slice(0,8);
    const max = Math.max(1, ...top.map(r => Number(r[value]) || 0));
    return top.map(r => {
      const n = Number(r[value]) || 0;
      const w = Math.max(2, (n / max) * 100);
      return `<div class="bar-row"><div class="bar-label">${escHTML(r[label] || '')}</div><div class="bar-track"><div class="bar-fill" style="width:${w.toFixed(1)}%"></div></div><div class="bar-value">${execNum(n,1)}${suffix}</div></div>`;
    }).join('') || '<p class="muted">Sem dados para gráfico.</p>';
  };
  const typeTable = typeRows.map(r => `<tr><td>${escHTML(r.tipo)}</td><td>${r.subatividades||r.atividades||0}</td><td>${r.quantidadeExecutada||0}</td><td>${execNum(r.totalHorasRealizadas,1)}h</td><td>${execNum(r.mediaHorasRealizadas,1)}h</td><td>${execNum(r.medianaHorasRealizadas,1)}h</td><td>${execNum(r.desvioHorasRealizadas,1)}h</td><td>${execNum(r.percentualExecucao,1)}%</td></tr>`).join('') || '<tr><td colspan="8">Sem dados de execução por tipo de subatividade.</td></tr>';
  const issueTable = issueTypeRows.map(r => `<tr><td>${escHTML(r.tipo)}</td><td>${r.quantidade||0}</td><td>${execNum(r.impactoHoras,1)}h</td><td>${execNum(r.mediaImpactoHoras,1)}h</td><td>${execNum(r.impactoPrazoDias,1)}d</td></tr>`).join('') || '<tr><td colspan="5">Sem ocorrências registradas.</td></tr>';
  const actTable = activityRows.slice(0,30).map(r => `<tr><td>${escHTML(r.codigo)}</td><td>${escHTML(r.titulo)}</td><td>${escHTML(r.tipo)}</td><td>${execNum(r.percentualReal,1)}%</td><td>${execNum(r.horasRealizadas,1)}h</td><td>${execNum(r.horasRestantes,1)}h</td><td>${escHTML(r.fimPlanejado)}</td><td>${escHTML(r.previsaoFim)}</td></tr>`).join('') || '<tr><td colspan="8">Sem atividades com dados de execução.</td></tr>';
  const issueDetailTable = issueRows.slice(0,40).map(r => `<tr><td>${escHTML(r.tipo)}</td><td>${escHTML(r.atividadeCodigo)}</td><td>${escHTML(r.atividadeTitulo)}</td><td>${escHTML(r.data)}</td><td>${execNum(r.impactoHoras,1)}h</td><td>${execNum(r.impactoPrazoDias,1)}d</td><td>${escHTML(r.descricao)}</td></tr>`).join('') || '<tr><td colspan="7">Sem ocorrências detalhadas.</td></tr>';
  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Relatório de Indicadores de Execução</title><style>
body{font-family:Arial,Helvetica,sans-serif;margin:24px;background:#f8fafc;color:#0f172a}h1{margin:0 0 6px;font-size:24px}h2{font-size:18px;margin:24px 0 10px}.muted{color:#64748b}.top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}.btn{background:#2563eb;color:white;border:none;border-radius:10px;padding:10px 14px;cursor:pointer}.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:18px 0}.kpi{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:14px}.kpi strong{display:block;font-size:24px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.card{background:white;border:1px solid #e2e8f0;border-radius:14px;padding:14px}.bar-row{display:grid;grid-template-columns:150px 1fr 70px;gap:10px;align-items:center;margin:8px 0}.bar-label{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.bar-track{height:14px;background:#e2e8f0;border-radius:999px;overflow:hidden}.bar-fill{height:100%;background:#2563eb;border-radius:999px}.bar-value{font-size:12px;text-align:right}table{width:100%;border-collapse:collapse;background:white;border:1px solid #e2e8f0;margin-bottom:18px}th,td{font-size:12px;text-align:left;padding:8px;border:1px solid #e2e8f0;vertical-align:top}th{background:#eef2ff}@media(max-width:900px){.grid{grid-template-columns:1fr}.top{display:block}.bar-row{grid-template-columns:110px 1fr 60px}}@media print{body{background:white;margin:12mm}.btn{display:none}.card,.kpi,table{break-inside:avoid}}
</style></head><body><div class="top"><div><h1>Relatório de Indicadores de Execução</h1><div class="muted">Gerado em ${escHTML(new Date().toLocaleString('pt-BR'))}</div></div><button class="btn" onclick="window.print()">Salvar em PDF</button></div>
<div class="kpis"><div class="kpi"><strong>${totals.atividades}</strong>atividades com execução</div><div class="kpi"><strong>${execNum(pct,1)}%</strong>execução global</div><div class="kpi"><strong>${execNum(totals.realizadas,1)}h</strong>horas realizadas</div><div class="kpi"><strong>${execNum(totals.saldo,1)}h</strong>saldo restante</div><div class="kpi"><strong>${totals.subatividades}</strong>subatividades</div><div class="kpi"><strong>${totals.apontamentos}</strong>apontamentos</div><div class="kpi"><strong>${totals.ocorrencias}</strong>ocorrências</div><div class="kpi"><strong>${execNum(totals.impacto,1)}h</strong>impacto total</div></div>
<div class="grid"><section class="card"><h2>Tempo médio realizado por tipo de subatividade</h2>${chartBars(typeRows,'tipo','mediaHorasRealizadas','h')}</section><section class="card"><h2>Principais ocorrências por impacto</h2>${chartBars(issueTypeRows,'tipo','impactoHoras','h')}</section></div>
<h2>Tempo médio de execução por tipo de subatividade</h2><table><thead><tr><th>Tipo de subatividade</th><th>Subatividades</th><th>Qtd. executada</th><th>Horas totais realizadas</th><th>Média realizada</th><th>Mediana</th><th>Desvio</th><th>% execução</th></tr></thead><tbody>${typeTable}</tbody></table>
<h2>Principais ocorrências e impactos</h2><table><thead><tr><th>Ocorrência</th><th>Qtd.</th><th>Impacto total</th><th>Média horas</th><th>Impacto prazo</th></tr></thead><tbody>${issueTable}</tbody></table>
<h2>Atividades acompanhadas</h2><table><thead><tr><th>Código</th><th>Atividade</th><th>Tipo</th><th>% real</th><th>Realizado</th><th>Saldo</th><th>Término planejado</th><th>Previsão atual</th></tr></thead><tbody>${actTable}</tbody></table>
<h2>Detalhamento das ocorrências</h2><table><thead><tr><th>Tipo</th><th>Código</th><th>Atividade</th><th>Data</th><th>Impacto horas</th><th>Impacto prazo</th><th>Descrição</th></tr></thead><tbody>${issueDetailTable}</tbody></table>
</body></html>`;
}

const btnExportExecPdf = document.getElementById('btnExportExecPdf');
if(btnExportExecPdf){
  btnExportExecPdf.type = 'button';
  btnExportExecPdf.onclick = async (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const oldText = btnExportExecPdf.textContent;
    btnExportExecPdf.disabled = true;
    btnExportExecPdf.textContent = 'Gerando PDF...';
    const reportWin = window.open('', '_blank');
    try{
      if(reportWin){
        reportWin.document.open();
        reportWin.document.write('<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Gerando relatório...</title><style>body{font-family:Arial,Helvetica,sans-serif;margin:32px;color:#0f172a}.muted{color:#64748b}</style></head><body><h1>Gerando relatório de indicadores de execução...</h1><p class="muted">Consolidando dados locais e dados sincronizados, quando disponíveis.</p></body></html>');
        reportWin.document.close();
      }
      try{ await refreshFromBDIfNeeded(); }catch(syncErr){ console.warn('Relatório de execução gerado com dados locais; falha ao sincronizar BD antes do relatório.', syncErr); }
      const data = getExecReportDataSafe();
      const html = buildExecReportHtmlSafe(data);
      if(reportWin && !reportWin.closed){
        reportWin.document.open();
        reportWin.document.write(html);
        reportWin.document.close();
        try{ reportWin.focus(); }catch(_e){}
      }else{
        download('relatorio_indicadores_execucao.html', html, 'text/html;charset=utf-8');
        alert('O navegador bloqueou a janela do relatório. Foi gerado um arquivo HTML para abrir e salvar como PDF.');
      }
    }catch(e){
      console.error('Falha ao gerar PDF de indicadores de execução', e);
      const fallback = `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Relatório de Indicadores de Execução</title></head><body><h1>Relatório de Indicadores de Execução</h1><p>Não foi possível montar os gráficos completos, mas o exportador foi estabilizado. Erro técnico: ${escHTML(e && (e.message || e.name) || e)}</p><button onclick="window.print()">Salvar em PDF</button></body></html>`;
      if(reportWin && !reportWin.closed){
        reportWin.document.open();
        reportWin.document.write(fallback);
        reportWin.document.close();
      }else{
        download('relatorio_indicadores_execucao_fallback.html', fallback, 'text/html;charset=utf-8');
        alert('Foi gerado um relatório fallback em HTML.');
      }
    }finally{
      btnExportExecPdf.disabled = false;
      btnExportExecPdf.textContent = oldText || 'Exportar PDF indicadores de execução';
    }
  };
}

// ===== Atrasadas do mês (mês da visão) =====
function monthKeyFromYMD(ymd){
  const d=fromYMD(ymd);
  const y=d.getFullYear();
  const m=String(d.getMonth()+1).padStart(2,"0");
  return `${y}-${m}`;
}

/**
 * Retorna as atividades atrasadas considerando o MÊS do Início (visão):
 * - Intersecta o mês (rangeStart..rangeEnd)
 * - Data fim < hoje
 * - Status != Concluída (e != Cancelada para não inflar indevidamente)
 */
function isActivityExtrapolated(a, today){
  if(!a || a.deletedAt) return false;
  const st=(a.status||"").trim();
  if(st==="Concluída" || st==="Cancelada") return false;
  if(!a.fim) return false;
  const now = today ? clampDate(today) : clampDate(new Date());
  return fromYMD(a.fim) < now;
}

function getExtrapolatedActivities(){
  const activeResById = Object.fromEntries(
    (resources||[]).filter(r=>r.ativo && !r.deletedAt).map(r=>[r.id,r])
  );
  const now=clampDate(new Date());
  return (activities||[]).filter(a=>{
    if(!isActivityExtrapolated(a, now)) return false;
    return !!activeResById[a.resourceId];
  });
}

function getLinkedChainActivities(activityId){
  const active=(activities||[]).filter(a=>a && !a.deletedAt);
  if(!activityId) return [];
  const byId = new Map(active.map(a=>[a.id,a]));
  if(!byId.has(activityId)) return [];
  const neighbors = new Map();
  active.forEach(a=>{
    if(!neighbors.has(a.id)) neighbors.set(a.id, new Set());
    const origin = a.linkedOriginId;
    if(origin && byId.has(origin)){
      neighbors.get(a.id).add(origin);
      if(!neighbors.has(origin)) neighbors.set(origin, new Set());
      neighbors.get(origin).add(a.id);
    }
  });
  const queue=[activityId];
  const seen=new Set([activityId]);
  const out=[];
  while(queue.length){
    const id=queue.shift();
    const item=byId.get(id);
    if(item) out.push(item);
    const adj=neighbors.get(id) || new Set();
    adj.forEach(nid=>{
      if(!seen.has(nid)){
        seen.add(nid);
        queue.push(nid);
      }
    });
  }
  return out;
}

function getActivityChainKey(activityId){
  const chain=getLinkedChainActivities(activityId);
  if(!chain.length) return activityId || '';
  return chain.map(a=>a.id).sort().join('|');
}

function getUniqueChainCount(items){
  const seen=new Set();
  (items||[]).forEach(a=>{
    const key=(a && a.linkedOriginId) ? getActivityChainKey(a.id) : (a && a.id ? a.id : '');
    if(key) seen.add(key);
  });
  return seen.size;
}

function getOverdueInMonthView(){
  const start=fromYMD(rangeStart);
  const end=fromYMD(rangeEnd);
  const now=clampDate(new Date());

  const activeResById = Object.fromEntries(
    (resources||[]).filter(r=>r.ativo && !r.deletedAt).map(r=>[r.id,r])
  );

  const rows=[];
  (activities||[]).forEach(a=>{
    if(!a || a.deletedAt) return;
    const r = activeResById[a.resourceId];
    if(!r) return;

    const st = (a.status||"").trim();
    if(st==="Concluída") return;
    if(st==="Cancelada") return;

    const aIni=fromYMD(a.inicio);
    const aFim=fromYMD(a.fim);

    // Atrasada hoje
    if(aFim>=now) return;

    // Interseção com o mês da visão
    if(aFim<start || aIni>end) return;

    const linked = a.linkedOriginId ? activities.find(x=>x.id===a.linkedOriginId) : null;
    rows.push({
      mesRef: monthKeyFromYMD(rangeStart),
      atividadeId: a.id,
      codigoAtividade: a.codigoAtividade || '',
      titulo: a.titulo,
      atividadeVinculadaA: linked ? (linked.codigoAtividade || linked.id || '') : '',
      responsavel: r.nome,
      tipo: r.tipo,
      senioridade: r.senioridade,
      status: st,
      inicio: a.inicio,
      fim: a.fim,
      diasAtraso: diffDays(now, aFim),
      alocacao: a.alocacao||100,
      tags: (a.tags||[]).join(", "),
    });
  });
  return rows;
}

function exportOverdueMonthCSVs(){
  const rows=getOverdueInMonthView();
  if(!rows.length){
    alert("Nenhuma atividade atrasada no mês selecionado (mês do Início da visão).");
    return;
  }

  const key = monthKeyFromYMD(rangeStart);

  // Detalhado
  download(
    `atrasadas_${key}.csv`,
    toCSV(rows,["mesRef","atividadeId","codigoAtividade","titulo","atividadeVinculadaA","responsavel","tipo","senioridade","status","inicio","fim","diasAtraso","alocacao","tags"]),
    "text/csv;charset=utf-8"
  );

  // Resumo mensal do mês selecionado
  const resumo=[{
    mesRef: key,
    qtdAtrasadas: rows.length,
    responsaveisImpactados: new Set(rows.map(r=>r.responsavel)).size,
    diasAtrasoMedio: Math.round(rows.reduce((acc,r)=>acc+(Number(r.diasAtraso)||0),0)/Math.max(1,rows.length))
  }];
  download(
    `atrasadas_resumo_${key}.csv`,
    toCSV(resumo,["mesRef","qtdAtrasadas","responsaveisImpactados","diasAtrasoMedio"]),
    "text/csv;charset=utf-8"
  );

  alert("Exportados: atrasadas (detalhado) e atrasadas_resumo (mês).");
}

(function(){
  const btn=document.getElementById("btnExportAtrasadasMes");
  if(btn) btn.onclick=async ()=>{
    await refreshFromBDIfNeeded();
    exportOverdueMonthCSVs();
  };
})();


if(btnHistAll) btnHistAll.onclick=async ()=>{
  // Garante sincronização com o BD apontado antes de consolidar o histórico
  await refreshFromBDIfNeeded();
  const rows=[];
  const byId=Object.fromEntries(resources.map(r=>[r.id,r]));
  Object.keys(trails).forEach(entityId=>{
    const a=activities.find(x=>x.id===entityId);
    const r=resources.find(x=>x.id===entityId);
    (trails[entityId]||[]).forEach(it=>{
      // Em alguns modelos de BD, metadados extras podem vir apenas em 'legend'
      // (JSON). Fazemos o parse aqui para garantir consistência na exportação.
      let meta=null;
      try{ if(it.legend && String(it.legend).trim().startsWith('{')) meta=JSON.parse(it.legend); }catch(e){ meta=null; }
      if(meta){
        it.type = it.type || meta.type;
        it.entityType = it.entityType || meta.entityType;
        it.oldResourceId = it.oldResourceId || meta.oldResourceId;
        it.oldResourceName = it.oldResourceName || meta.oldResourceName;
        it.newResourceId = it.newResourceId || meta.newResourceId;
        it.newResourceName = it.newResourceName || meta.newResourceName;
      }
      const tipo = it.type || 'ALTERACAO_DATAS';
      const entidadeTipo = it.entityType || (a ? 'atividade' : (r ? 'recurso' : 'atividade'));
      // Resolve nomes de recurso quando possível
      const oldResName = it.oldResourceName || (it.oldResourceId && byId[it.oldResourceId] ? byId[it.oldResourceId].nome : '');
      const newResName = it.newResourceName || (it.newResourceId && byId[it.newResourceId] ? byId[it.newResourceId].nome : '');
      rows.push({
        tipoEvento: tipo,
        entidadeTipo: entidadeTipo,
        entidadeId: entityId,
        atividadeId: a ? a.id : (entidadeTipo==='atividade'?entityId:''),
        atividadeTitulo: a? a.titulo:"",
        recursoAnteriorId: it.oldResourceId || "",
        recursoAnteriorNome: oldResName || "",
        recursoNovoId: it.newResourceId || "",
        recursoNovoNome: newResName || "",
        dataHora: it.ts,
        oldInicio: it.oldInicio || "",
        oldFim: it.oldFim || "",
        newInicio: it.newInicio || "",
        newFim: it.newFim || "",
        justificativa: it.justificativa||"",
        user: it.user||""
      });
    });
  });
  if(!rows.length){ alert("Sem registros de histórico."); return; }
  download("historico_consolidado.csv", toCSV(rows, [
    "tipoEvento","entidadeTipo","entidadeId",
    "atividadeId","atividadeTitulo",
    "recursoAnteriorId","recursoAnteriorNome","recursoNovoId","recursoNovoNome",
    "dataHora",
    "oldInicio","oldFim","newInicio","newFim",
    "justificativa","user"
  ]), "text/csv;charset=utf-8");
};

if(btnBackup) btnBackup.onclick=()=>{
  const dump={resources, activities, comments, trails, meta:{version:"v3", exportedAt:new Date().toISOString()}};
  download("backup_planejador.json", JSON.stringify(dump,null,2), "application/json;charset=utf-8");
};
if(fileRestore) fileRestore.onchange=(ev)=>{
  const f=ev.target.files[0]; if(!f) return;
  const reader=new FileReader();
  reader.onload=()=>{
    try{
      const dump=JSON.parse(reader.result);
      if(!dump.resources || !dump.activities) throw new Error("Arquivo inválido.");
      resources=(dump.resources||[]).map(coerceResource); activities=hydrateLoadedActivities(dump.activities||[]); comments=hydrateLoadedComments(activities, dump.comments || []); rebuildCommentsIndex(); syncAllActivityCommentFields(); trails=dump.trails||{};
      saveLS(LS.res,resources); saveLS(LS.act,activities); saveLS(LS.comments,comments); saveLS(LS.trail,trails||{});
      renderAll(); alert("Restauração concluída.");
    }catch(err){ alert("Falha ao restaurar: "+err.message); }
  };
  reader.readAsText(f,"utf-8");
};

const __fileImportEl = document.getElementById("fileImport");
if(__fileImportEl) __fileImportEl.onchange=(ev)=>{
  const files=[...ev.target.files];
  if(!files.length) return;
  let pending=files.length;
  files.forEach(file=>{
    const reader=new FileReader();
    reader.onload=()=>{
      const text=reader.result;
      const lines=text.split(/\r?\n/).filter(l=>l.trim().length>0);
      if(!lines.length){ if(--pending===0){renderAll(); alert("Importação concluída.");} return; }
      const sep=lines[0].includes(";")?";":",";
      const headers=lines[0].split(sep).map(h=>h.trim());
      const idx=(name)=>headers.indexOf(name);
      if(headers.includes("nome") && headers.includes("capacidade")){
        const arr=[];
        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(sep);
          if(cols.length!==headers.length) continue;
          const rec={
            id: cols[idx("id")]||uuid(),
            nome: cols[idx("nome")]||"",
            tipo: cols[idx("tipo")]||"Interno",
            senioridade: cols[idx("senioridade")]||"NA",
            ativo: (cols[idx("ativo")]||"true").toLowerCase()!=="false",
            capacidade: Number(cols[idx("capacidade")]||100),
            cargaHorariaDiaria: Math.max(0.5, Number(cols[idx("cargaHorariaDiaria")]||9) || 9),
            inicioAtivo: cols[idx("inicioAtivo")]||"",
            fimAtivo: cols[idx("fimAtivo")]||""
          };
          if(rec.nome) arr.push(rec);
        }
        const ids=new Set(resources.map(r=>r.id));
        resources=[...resources, ...arr.filter(r=>!ids.has(r.id))];
        saveLS(LS.res,resources);
      } else if(headers.includes("titulo") && headers.includes("resourceId")){
        const arr=[];
        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(sep);
          if(cols.length!==headers.length) continue;
          const at={
            id: cols[idx("id")]||uuid(),
            titulo: cols[idx("titulo")]||"",
            resourceId: cols[idx("resourceId")]||"",
            inicio: cols[idx("inicio")]||"",
            fim: cols[idx("fim")]||"",
            status: cols[idx("status")]||"Planejada",
            alocacao: Number(cols[idx("alocacao")]||100),
            tags: (cols[idx("tags")] || "").split(',').map(t => t.trim()).filter(Boolean)
          };
          if(at.titulo && at.resourceId && at.inicio && at.fim) arr.push(at);
        }
        const ids=new Set(activities.map(a=>a.id));
        activities=[...activities, ...arr.filter(a=>!ids.has(a.id))];
        saveLS(LS.act,activities);
      }
      if(--pending===0){ renderAll(); alert("Importação concluída."); }
    };
    reader.readAsText(file, "utf-8");
  });
};

btnHistExport.onclick=(e)=>{
  e.preventDefault();
  if(!histCurrentId){ alert("Abra o histórico de uma atividade."); return; }
  const list = trails[histCurrentId]||[];
  if(list.length===0){ alert("Sem registros para exportar."); return; }
  const s=document.getElementById('histStart').value;
  const e2=document.getElementById('histEnd').value;
  const rows = list.filter(it=>{
    const t=new Date(it.ts);
    return (!s || t>=fromYMD(s)) && (!e2 || t<=addDays(fromYMD(e2),0));
  }).map(it=>({ts:it.ts, oldInicio:it.oldInicio, oldFim:it.oldFim, newInicio:it.newInicio, newFim:it.newFim, justificativa:it.justificativa||"", user:it.user||""}));
  download(`historico_${histCurrentId}.csv`, toCSV(rows, ["ts","oldInicio","oldFim","newInicio","newFim","justificativa","user"]), "text/csv;charset=utf-8");
};

// ===== Agregados =====
function bucketKey(d, gran){
  if(gran==="weekly"){
    const date=new Date(d); const day=(date.getDay()+6)%7; 
    const monday=new Date(date); monday.setDate(date.getDate()-day);
    const y=monday.getFullYear(); const m=String(monday.getMonth()+1).padStart(2,"0"); const day2=String(monday.getDate()).padStart(2,"0");
    return `W ${y}-${m}-${day2}`;
  } else {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
  }
}

// Formata rótulos do eixo X para os gráficos de capacidade agregada
function formatBucketLabel(key, gran){
  if(gran === "weekly"){
    // key: "W YYYY-MM-DD" (segunda-feira da semana)
    const raw = String(key).replace(/^W\s+/, "");
    const dt = fromYMD(raw);
    const dd = String(dt.getDate()).padStart(2,"0");
    const mm = String(dt.getMonth()+1).padStart(2,"0");
    return `${dd}/${mm}`;
  }
  if(gran === "daily"){
    const dt = fromYMD(String(key));
    const dd = String(dt.getDate()).padStart(2,"0");
    const mm = String(dt.getMonth()+1).padStart(2,"0");
    return `${dd}/${mm}`;
  }
  // monthly: "YYYY-MM"
  const m = String(key);
  const parts = m.split("-");
  if(parts.length === 2){
    return `${parts[1]}/${parts[0]}`;
  }
  return String(key);
}
// FUNÇÃO CORRIGIDA
// Agora respeita os mesmos filtros aplicados no Gantt (via lista de atividades filtradas).
function renderAggregates(filteredActsOverride){
  aggCharts.innerHTML="";
  const gran=aggGran.value;
  const mode = aggMode ? aggMode.value : 'percent';
  const days=buildDays();
  const filteredActs = Array.isArray(filteredActsOverride) ? filteredActsOverride : (typeof getFilteredActivities === 'function' ? getFilteredActivities() : []);

  const resIds = new Set((filteredActs || []).map(a => a.resourceId));
  const visibleResources = (resources || [])
    .filter(r => r && r.ativo && !r.deletedAt && resIds.has(r.id))
    .slice()
    .sort((a,b)=>String(a.nome||'').localeCompare(String(b.nome||''), undefined, { sensitivity: 'base' }));

  if(aggGran){
    aggGran.disabled = mode === 'daily_hours';
    aggGran.style.opacity = mode === 'daily_hours' ? '0.6' : '1';
  }

  const dailyIndex = buildDailyLoadIndex(filteredActs, days);

  visibleResources.forEach(r => {
    const card=document.createElement("div");
    card.className="card";
    const h=document.createElement("h3");
    h.textContent=`${escHTML(r.nome)} — ${mode==='daily_hours' ? 'Carga diária (h)' : (gran==="weekly"?"Semanal":"Mensal")}`;
    const canvas=document.createElement("canvas");
    canvas.width=600; canvas.height=180; canvas.className="chart";
    const ctx=canvas.getContext("2d");
    const W=canvas.width, H=canvas.height;
    const marginL = 34, marginT = 12, marginB = 48, axisY = H - marginB;
    ctx.clearRect(0,0,W,H);
    ctx.beginPath();
    ctx.moveTo(marginL, marginT);
    ctx.lineTo(marginL, axisY);
    ctx.lineTo(W-10, axisY);
    ctx.stroke();

    if(mode === 'daily_hours'){
      const entries = days.map(d=>{
        const ymd = toYMD(d);
        const slot = (dailyIndex[r.id] && dailyIndex[r.id][ymd]) || { allocatedHours:0, capacityHours:getDailyCapacityHours(r, ymd) };
        return [ymd, slot];
      });
      const labelStep = entries.length > 20 ? 3 : (entries.length > 12 ? 2 : 1);
      // v1.2.8.26: nesta visão, o valor principal exibido é a capacidade útil do dia.
      // Em feriados/finais de semana, a capacidade base é 0h e somente horas extras aprovadas
      // aumentam a barra. A alocação planejada acima da capacidade é sinalizada por um marcador
      // de excesso, sem transformar o feriado em uma jornada nominal comum.
      const maxHours = Math.max(1, ...entries.map(([ymd,v])=>{
        const ctxDay = getDayContext(ymd, r);
        const displayCap = Math.max(0, v.capacityHours || 0);
        const excess = ctxDay.nonWorking ? Math.max(0, (v.allocatedHours||0) - displayCap) : Math.max(0, v.allocatedHours||0, displayCap);
        return Math.max(displayCap, excess);
      }));
      const barW=Math.max(8, (W - marginL - 20) / Math.max(1, entries.length) - 6);
      entries.forEach((kv,idx)=>{
        const key=kv[0]; const v=kv[1];
        const dayCtx = getDayContext(key, r);
        const label = formatBucketLabel(key, 'daily');
        const x = marginL + 10 + idx*(barW+6);
        const displayHours = Math.max(0, v.capacityHours || 0);
        const hVal = (displayHours / maxHours) * (axisY - marginT - 18);
        const y = axisY - hVal;
        const hasOvertime = Number(v.overtimeApproved || 0) > 0;
        if(dayCtx.isHoliday){
          ctx.fillStyle = hasOvertime ? '#f59e0b' : '#fef3c7';
        } else if(dayCtx.isWeekend){
          ctx.fillStyle = hasOvertime ? '#8b5cf6' : '#e2e8f0';
        } else {
          ctx.fillStyle = (v.allocatedHours||0) > (v.capacityHours||0) ? '#ef4444' : '#2563eb';
        }
        if(displayHours > 0){
          ctx.fillRect(x, y, barW, axisY - y);
        } else {
          ctx.fillRect(x, axisY - 2, barW, 2);
        }
        // Linha/indicador de alocação: quando a atividade planejada excede a capacidade útil,
        // sinaliza o excesso em vermelho, preservando a barra principal como capacidade real.
        const allocHours = Math.max(0, v.allocatedHours || 0);
        if(allocHours > displayHours){
          const allocY = axisY - ((allocHours / maxHours) * (axisY - marginT - 18));
          ctx.strokeStyle = '#ef4444';
          ctx.lineWidth = 2;
          ctx.beginPath();
          ctx.moveTo(x, allocY);
          ctx.lineTo(x + barW, allocY);
          ctx.stroke();
          ctx.lineWidth = 1;
        }
        if(idx % labelStep === 0){
          ctx.save();
          ctx.translate(x + barW/2, axisY + 16);
          ctx.textAlign="center";
          ctx.textBaseline="top";
          ctx.font="9px sans-serif";
          ctx.fillText(label, 0, 0);
          ctx.restore();
        }
        ctx.font="10px sans-serif";
        const suffix = dayCtx.isHoliday ? 'h F' : (dayCtx.isWeekend ? 'h FS' : 'h');
        ctx.fillStyle = dayCtx.isHoliday ? '#92400e' : (dayCtx.isWeekend ? '#5b21b6' : '#0f172a');
        ctx.fillText(displayHours.toFixed(1)+suffix, x, Math.max(marginT+10, y-4));
      });
      const note=document.createElement('div');
      note.className='muted';
      note.style.fontSize='12px';
      note.style.marginTop='6px';
      note.textContent=`Jornada nominal: ${getDailyHours(r)}h/dia • Barras mostram capacidade útil: feriados/finais de semana = 0h + HE aprovada; linha vermelha indica alocação acima da capacidade`;
      card.appendChild(h); card.appendChild(canvas); card.appendChild(note);
    } else {
      const byBucket = {};
      days.forEach(d=>{
        const key = bucketKey(d, gran);
        if(!byBucket[key]) byBucket[key] = { allocatedPct:0, capacityPct:0 };
        const slot = (dailyIndex[r.id] && dailyIndex[r.id][toYMD(d)]) || { allocatedPct:0, capacityPct:Number(r.capacidade||100) };
        byBucket[key].allocatedPct += Number(slot.allocatedPct || 0);
        byBucket[key].capacityPct += Number(slot.capacityPct || 0);
      });
      const entries = Object.entries(byBucket).sort((a,b)=>a[0]>b[0]?1:-1);
      const labelStep = entries.length > 20 ? 3 : (entries.length > 12 ? 2 : 1);
      const barW=Math.max(8, (W - marginL - 20) / Math.max(1, entries.length) - 6);
      entries.forEach((kv,idx)=>{
        const key=kv[0]; const v=kv[1];
        const label = formatBucketLabel(key, gran);
        const perc = v.capacityPct > 0 ? (v.allocatedPct / v.capacityPct) * 100 : 0;
        const x = marginL + 10 + idx*(barW+6);
        const y = axisY - (Math.min(100, perc)/100)*(axisY - marginT - 18);
        ctx.fillStyle = perc > 100 ? '#ef4444' : '#2563eb';
        ctx.fillRect(x, y, barW, axisY - y);
        if(idx % labelStep === 0){
          ctx.save();
          ctx.translate(x + barW/2, axisY + 16);
          ctx.textAlign="center";
          ctx.textBaseline="top";
          ctx.font="9px sans-serif";
          ctx.fillText(label, 0, 0);
          ctx.restore();
        }
        ctx.font="10px sans-serif"; ctx.fillText(Math.round(perc)+"%", x, y-4);
      });
      card.appendChild(h); card.appendChild(canvas);
    }
    aggCharts.appendChild(card);
  });
}

// ===== Auto-sugestão de Tags =====
function updateTagDatalist() {
    const allTags = new Set();
    activities.forEach(a => {
        if (a.tags) {
            a.tags.forEach(t => allTags.add(t));
        }
    });
  
    const datalist = document.getElementById('existingTags');
    if (datalist) {
        datalist.innerHTML = '';
        allTags.forEach(tag => {
            const option = document.createElement('option');
            option.value = tag;
            datalist.appendChild(option);
        });
    }
}

// ===== Render principal =====
function renderAll(){
  const filtered=getFilteredActivities();
  // Persistir filtros para manter contexto entre recarregamentos.
  saveFiltersState();

  // KPI de visão atual (quantos recursos e atividades estão sendo exibidos)
  if(viewKpi){
    try{
      const byRes = Object.fromEntries((resources || []).map(r=>[r.id,0]));
      filtered.forEach(a=>{ if(byRes[a.resourceId] != null) byRes[a.resourceId]++; });
      const visibleRes = (resources || []).filter(r=>{
        if(r.deletedAt) return false;
        if(filtroTipo && (r.tipo||"").toLowerCase() !== filtroTipo.toLowerCase()) return false;
        if(filtroSenioridade && r.senioridade!==filtroSenioridade) return false;
        if(!r.ativo) return false;
        if(buscaRecurso && !(r.nome||"").toLowerCase().includes(buscaRecurso)) return false;
        return (byRes[r.id] || 0) > 0;
      }).length;
      viewKpi.textContent = `${visibleRes} recursos · ${filtered.length} atividades`;
    }catch{ /* ignore */ }
  }
  renderTables(filtered);
  renderGantt(filtered);
  // Capacidade agregada agora respeita os mesmos filtros aplicados no Gantt.
  renderAggregates(filtered);
  if (window.__execDashReady && typeof renderExecutionDashboard === "function") renderExecutionDashboard();
  renderCalendarUI();
}

renderStatusChips();
if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bootAccessGate); else bootAccessGate();
if(aggGran) aggGran.onchange=()=>renderAll();
if(aggMode) aggMode.onchange=()=>renderAll();

// ===== Baseline mensal & KPIs =====
function saveBaselines(){
  const okBase = saveLS(LS.base, baselines);
  const okItems = saveLS(LS.baseItems, baselineItems);
  return !!(okBase && okItems);
}

function fmtDateTimeISO(iso){
  try{
    const d = new Date(iso);
    if(isNaN(d.getTime())) return iso || '';
    const pad = (n)=>String(n).padStart(2,'0');
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }catch{ return iso || ''; }
}

function defaultBaselineName(){
  const now = new Date();
  const ym = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  return `${ym} Baseline`;
}

// ===== Snapshot / aplicação de filtros (evita falha humana) =====
function getCurrentFilterSnapshot(){
  return {
    rangeStart,
    rangeEnd,
    filtroTipo: filtroTipo || "",
    filtroSenioridade: filtroSenioridade || "",
    buscaTitulo: (buscaTituloInput && buscaTituloInput.value ? buscaTituloInput.value : ""),
    buscaTag: (buscaTagInput && buscaTagInput.value ? buscaTagInput.value : ""),
    buscaRecurso: (buscaRecursoInput && buscaRecursoInput.value ? buscaRecursoInput.value : ""),
    selectedStatus: Array.from(selectedStatus || new Set()),
  };
}

function applyFilterSnapshotToUI(snap){
  if(!snap) return;
  // Variáveis
  filtroTipo = snap.filtroTipo || "";
  filtroSenioridade = snap.filtroSenioridade || "";
  buscaTitulo = (snap.buscaTitulo||"").toLowerCase();
  buscaTag = (snap.buscaTag||"").toLowerCase();
  buscaRecurso = (snap.buscaRecurso||"").toLowerCase().trim();
  rangeStart = snap.rangeStart || rangeStart;
  rangeEnd = snap.rangeEnd || rangeEnd;
  // Status
  selectedStatus.clear();
  (snap.selectedStatus || []).forEach(s=>selectedStatus.add(String(s)));
  if(selectedStatus.size===0){ STATUS.forEach(s=>selectedStatus.add(s)); }

  // UI
  if(tipoSel) tipoSel.value = filtroTipo;
  if(senioridadeSel) senioridadeSel.value = filtroSenioridade;
  if(buscaTituloInput) buscaTituloInput.value = snap.buscaTitulo || "";
  if(buscaTagInput) buscaTagInput.value = snap.buscaTag || "";
  if(buscaRecursoInput) buscaRecursoInput.value = snap.buscaRecurso || "";
  if(inicioVisao) inicioVisao.value = rangeStart;
  if(fimVisao) fimVisao.value = rangeEnd;
  renderStatusChips();
  renderAll();
}

function withFilterSnapshot(snap, fn){
  if(!snap || typeof fn !== 'function') return fn();
  // Salva estado atual
  const prev = {
    rangeStart, rangeEnd, filtroTipo, filtroSenioridade, buscaTitulo, buscaTag, buscaRecurso,
    selectedStatus: Array.from(selectedStatus)
  };
  try{
    filtroTipo = snap.filtroTipo || "";
    filtroSenioridade = snap.filtroSenioridade || "";
    buscaTitulo = (snap.buscaTitulo||"").toLowerCase();
    buscaTag = (snap.buscaTag||"").toLowerCase();
    buscaRecurso = (snap.buscaRecurso||"").toLowerCase().trim();
    rangeStart = snap.rangeStart || rangeStart;
    rangeEnd = snap.rangeEnd || rangeEnd;
    selectedStatus.clear();
    (snap.selectedStatus || []).forEach(s=>selectedStatus.add(String(s)));
    if(selectedStatus.size===0){ STATUS.forEach(s=>selectedStatus.add(s)); }
    return fn();
  } finally {
    // Restaura estado (sem mexer na UI)
    rangeStart = prev.rangeStart; rangeEnd = prev.rangeEnd;
    filtroTipo = prev.filtroTipo; filtroSenioridade = prev.filtroSenioridade;
    buscaTitulo = prev.buscaTitulo; buscaTag = prev.buscaTag; buscaRecurso = prev.buscaRecurso;
    selectedStatus.clear();
    (prev.selectedStatus||[]).forEach(s=>selectedStatus.add(s));
  }
}

function getActivitiesBySnapshot(snap){
  // Retorna atividades filtradas como se a UI estivesse com os filtros do baseline,
  // sem depender do estado atual da tela (evita erro humano).
  if(!snap) return getActiveActivities();
  return withFilterSnapshot(snap, ()=> (typeof getFilteredActivities==='function' ? getFilteredActivities() : getActiveActivities()));
}

function renderBaselineSelect(){
  if(!baselineSel) return;
  baselineSel.innerHTML = '';
  const opt0 = document.createElement('option');
  opt0.value = '';
  opt0.textContent = 'Selecione...';
  baselineSel.appendChild(opt0);
  const sorted = [...baselines].sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  sorted.forEach(b=>{
    const opt = document.createElement('option');
    opt.value = b.id;
    opt.textContent = `${b.nome || b.id} · ${fmtDateTimeISO(b.createdAt)}`;
    baselineSel.appendChild(opt);
  });
}

function renderKpiCards(summary){
  if(!kpiCards) return;
  const items = [
    {label:'Total (baseline)', value: summary.totalBaseline},
    {label:'Mantidas', value:`${summary.mantidas} (${summary.totalBaseline?Math.round(summary.mantidas/summary.totalBaseline*100):0}%)`},
    {label:'Modificadas', value:`${summary.modificadas} (${summary.totalBaseline?Math.round(summary.modificadas/summary.totalBaseline*100):0}%)`},
    {label:'Novas', value: summary.novas},
    {label:'Excluídas', value:`${summary.excluidas} (${summary.totalBaseline?Math.round(summary.excluidas/summary.totalBaseline*100):0}%)`},
    {label:'Canceladas', value:`${summary.canceladas} (${summary.totalBaseline?Math.round(summary.canceladas/summary.totalBaseline*100):0}%)`},
    {label:'Bloqueadas', value:`${summary.bloqueadas} (${summary.totalBaseline?Math.round(summary.bloqueadas/summary.totalBaseline*100):0}%)`},
  ];
  kpiCards.innerHTML = items.map(it=>`
    <div class="card" style="padding:12px;">
      <div class="muted" style="font-size:12px;">${escHTML(it.label)}</div>
      <div style="font-size:20px; font-weight:700;">${escHTML(String(it.value))}</div>
    </div>
  `).join('');
}

function getMotivoParaAtividade(activityId, startISO, endISO){
  const list = (trails && trails[activityId]) ? trails[activityId] : [];
  if(!Array.isArray(list) || list.length===0) return {justificativa:'', user:'', ts:''};
  const start = startISO ? new Date(startISO).getTime() : 0;
  const end = endISO ? new Date(endISO).getTime() : Date.now();
  const candidates = list.filter(e=>{
    const t = new Date(e.ts || 0).getTime();
    return t && t>=start && t<=end && (e.justificativa || e.type && e.type!=='ALTERACAO_DATAS');
  }).sort((a,b)=>new Date(b.ts).getTime()-new Date(a.ts).getTime());
  const pick = candidates[0] || list[list.length-1];
  return {justificativa: pick.justificativa || '', user: pick.user || '', ts: pick.ts || ''};
}

function createBaseline(){
  if((baselines || []).length >= MAX_BASELINES){
    const ordered = [...baselines].sort((a,b)=>(a.createdAt||'').localeCompare(b.createdAt||''));
    const names = ordered.map((b, idx)=> `${idx + 1}. ${b.nome || b.id}`).join('\n');
    alert(`Limite de ${MAX_BASELINES} baselines atingido.

Exclua um baseline existente para liberar espaço antes de criar um novo.

Baselines atuais:
${names}`);
    if(baselineSel && ordered[0]) baselineSel.value = ordered[0].id;
    return;
  }
  const nome = (baselineNome && baselineNome.value ? baselineNome.value.trim() : '') || defaultBaselineName();
  const usarFiltros = !!(baselineUsarFiltros && baselineUsarFiltros.checked);
  const filterSnapshot = usarFiltros ? getCurrentFilterSnapshot() : null;
  const acts = usarFiltros && typeof getFilteredActivities==='function'
    ? getActivitiesBySnapshot(filterSnapshot)
    : getActiveActivities();
  const byRes = Object.fromEntries((resources||[]).map(r=>[r.id,r]));
  const id = (window.crypto && window.crypto.randomUUID) ? window.crypto.randomUUID() : String(Date.now());
  const createdAt = new Date().toISOString();
  const newBaseline = {id, nome, createdAt, createdBy: currentUser||'', usarFiltros, filterSnapshot};
  const newItems = (acts||[]).map(a=>{
    const r = byRes[a.resourceId] || null;
    return {
      baselineId: id,
      activityId: a.id,
      titulo: a.titulo,
      resourceId: a.resourceId,
      resourceNome: r? r.nome : '',
      inicio: a.inicio,
      fim: a.fim,
      status: a.status,
      alocacao: a.alocacao||100,
      tags: Array.isArray(a.tags)? a.tags.slice(0) : [],
      codigoAtividade: a.codigoAtividade || '',
      linkedOriginId: a.linkedOriginId || '',
      linkedOriginCode: a.linkedOriginCode || '',
      resourceCapacidade: Number(r && r.capacidade != null ? r.capacidade : 100),
      resourceCargaHorariaDiaria: getDailyHours(r)
    };
  });
  baselines.push(newBaseline);
  baselineItems.push(...newItems);
  if(!saveBaselines()){
    baselines = baselines.filter(b=>b.id !== id);
    baselineItems = baselineItems.filter(i=>i.baselineId !== id);
    return;
  }
  renderBaselineSelect();
  if(baselineSel) baselineSel.value = id;
  alert(`Baseline criado: ${nome} (${acts.length} atividades)`);
}


function compareBaseline(){
  if(!baselineSel || !baselineSel.value){ alert('Selecione um baseline para comparar.'); return; }
  const baseId = baselineSel.value;
  const base = baselines.find(b=>b.id===baseId);
  if(!base){ alert('Baseline não encontrado.'); return; }
  const baseList = baselineItems.filter(i=>i.baselineId===baseId);
  // Para evitar falha humana: quando o baseline foi criado com filtros, a comparação (por padrão)
  // usa o MESMO snapshot de filtros do baseline. Assim, comparar "baseline vs ele mesmo" resulta em 0 diferenças.
  const preferSnapshot = !!(base && base.filterSnapshot);
  const usarFiltrosNaComparacao = baselineCompararFiltros ? !!baselineCompararFiltros.checked : preferSnapshot;
  const currentList = (usarFiltrosNaComparacao && typeof getFilteredActivities==='function')
    ? (preferSnapshot ? getActivitiesBySnapshot(base.filterSnapshot) : getFilteredActivities())
    : getActiveActivities();
  const currentById = Object.fromEntries(currentList.map(a=>[a.id,a]));
  const baseById = Object.fromEntries(baseList.map(i=>[i.activityId,i]));

  const nowISO = new Date().toISOString();
  const rows = [];
  let mantidas=0, modificadas=0, novas=0, excluidas=0, canceladas=0, bloqueadas=0;

  // Baseline -> atual
  baseList.forEach(bi=>{
    const cur = currentById[bi.activityId];
    if(!cur){
      // Pode ter sido excluída logicamente
      const existed = activities.find(a=>a.id===bi.activityId);
      if(existed && existed.deletedAt) excluidas++; else excluidas++;
      const motivo = getMotivoParaAtividade(bi.activityId, base.createdAt, nowISO);
      rows.push({
        categoria:'Excluída',
        activityId: bi.activityId,
        titulo: bi.titulo,
        recursoAntes: bi.resourceNome||'',
        recursoDepois: '',
        inicioAntes: bi.inicio||'',
        fimAntes: bi.fim||'',
        inicioDepois: '',
        fimDepois: '',
        statusAntes: bi.status||'',
        statusDepois: '',
        motivo: motivo.justificativa||'',
        usuario: motivo.user||'',
        dataEvento: motivo.ts||'',
        codigoAtividade: bi.codigoAtividade || '',
        atividadeVinculadaA: bi.linkedOriginCode || '',
        tagsAntes: Array.isArray(bi.tags) ? bi.tags.join(', ') : '',
        tagsDepois: '',
        alocacaoAntes: bi.alocacao ?? '',
        alocacaoDepois: ''
      });
      return;
    }
    const curTags = Array.isArray(cur.tags) ? cur.tags.join(',') : '';
    const baseTags = Array.isArray(bi.tags) ? bi.tags.join(',') : '';
    const curLinkedCode = cur.linkedOriginCode || (cur.linkedOriginId ? (((activities||[]).find(x=>x.id===cur.linkedOriginId)||{}).codigoAtividade || cur.linkedOriginId) : '');
    const mudou = (bi.resourceId||'') !== (cur.resourceId||'') || (bi.inicio||'')!== (cur.inicio||'') || (bi.fim||'')!==(cur.fim||'') || (bi.status||'')!==(cur.status||'') || (bi.alocacao??100)!==(cur.alocacao??100) || (bi.codigoAtividade||'') !== (cur.codigoAtividade||'') || (bi.linkedOriginId||'') !== (cur.linkedOriginId||'') || (bi.linkedOriginCode||'') !== curLinkedCode || baseTags !== curTags;
    if(!mudou){
      mantidas++;
      rows.push({
        categoria:'Mantida',
        activityId: bi.activityId,
        titulo: bi.titulo,
        recursoAntes: bi.resourceNome||'',
        recursoDepois: (resources.find(r=>r.id===cur.resourceId)||{}).nome||'',
        inicioAntes: bi.inicio||'',
        fimAntes: bi.fim||'',
        inicioDepois: cur.inicio||'',
        fimDepois: cur.fim||'',
        statusAntes: bi.status||'',
        statusDepois: cur.status||'',
        motivo: '', usuario:'', dataEvento:''
      });
      return;
    }
    modificadas++;
    if(String(cur.status||'')==='Cancelada') canceladas++;
    if(String(cur.status||'')==='Bloqueada') bloqueadas++;
    const recDepois = (resources.find(r=>r.id===cur.resourceId)||{}).nome||'';
    const motivo = getMotivoParaAtividade(bi.activityId, base.createdAt, nowISO);
    rows.push({
      categoria:'Modificada',
      activityId: bi.activityId,
      titulo: bi.titulo,
      recursoAntes: bi.resourceNome||'',
      recursoDepois: recDepois,
      inicioAntes: bi.inicio||'',
      fimAntes: bi.fim||'',
      inicioDepois: cur.inicio||'',
      fimDepois: cur.fim||'',
      statusAntes: bi.status||'',
      statusDepois: cur.status||'',
      motivo: motivo.justificativa||'',
      usuario: motivo.user||'',
      dataEvento: motivo.ts||'',
      codigoAtividade: cur.codigoAtividade || bi.codigoAtividade || '',
      atividadeVinculadaA: cur.linkedOriginCode || bi.linkedOriginCode || '',
      tagsAntes: Array.isArray(bi.tags) ? bi.tags.join(', ') : '',
      tagsDepois: Array.isArray(cur.tags) ? cur.tags.join(', ') : '',
      alocacaoAntes: bi.alocacao ?? '',
      alocacaoDepois: cur.alocacao ?? ''
    });
  });

  // Novas
  currentList.forEach(cur=>{
    if(baseById[cur.id]) return;
    novas++;
    const recDepois = (resources.find(r=>r.id===cur.resourceId)||{}).nome||'';
    const motivo = getMotivoParaAtividade(cur.id, base.createdAt, nowISO);
    rows.push({
      categoria:'Nova',
      activityId: cur.id,
      titulo: cur.titulo,
      recursoAntes: '',
      recursoDepois: recDepois,
      inicioAntes: '',
      fimAntes: '',
      inicioDepois: cur.inicio||'',
      fimDepois: cur.fim||'',
      statusAntes: '',
      statusDepois: cur.status||'',
      motivo: motivo.justificativa||'',
      usuario: motivo.user||'',
      dataEvento: motivo.ts||'',
      codigoAtividade: cur.codigoAtividade || '',
      atividadeVinculadaA: cur.linkedOriginCode || (cur.linkedOriginId ? (((activities||[]).find(x=>x.id===cur.linkedOriginId)||{}).codigoAtividade || cur.linkedOriginId) : ''),
      tagsAntes: '',
      tagsDepois: Array.isArray(cur.tags) ? cur.tags.join(', ') : '',
      alocacaoAntes: '',
      alocacaoDepois: cur.alocacao ?? ''
    });
  });

  lastComparisonRows = rows;
  renderKpiCards({
    totalBaseline: baseList.length,
    mantidas, modificadas, novas, excluidas, canceladas, bloqueadas
  });
  renderComparisonTable(rows);
  if(btnExportComparacao) btnExportComparacao.disabled = rows.length===0;
}

function renderComparisonTable(rows){
  if(!baselineTabela) return;
  if(!rows || rows.length===0){ baselineTabela.innerHTML = '<div class="muted">Nenhum dado para exibir.</div>'; return; }
  const headers = ['Categoria','ID','Código','Título','Vinculada a','Recurso (antes)','Recurso (depois)','Início (antes)','Fim (antes)','Início (depois)','Fim (depois)','Status (antes)','Status (depois)','Alocação (antes)','Alocação (depois)','Tags (antes)','Tags (depois)','Motivo','Usuário','Data evento'];
  const th = headers.map(h=>`<th style="text-align:left; padding:8px; border-bottom:1px solid #e2e8f0;">${escHTML(h)}</th>`).join('');
  const body = rows.map(r=>`
    <tr>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.categoria||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; font-family:ui-monospace, SFMono-Regular, Menlo, monospace; font-size:12px;">${escHTML(r.activityId||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; font-family:ui-monospace, SFMono-Regular, Menlo, monospace; font-size:12px;">${escHTML(r.codigoAtividade||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.titulo||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; font-family:ui-monospace, SFMono-Regular, Menlo, monospace; font-size:12px;">${escHTML(r.atividadeVinculadaA||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.recursoAntes||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.recursoDepois||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.inicioAntes||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.fimAntes||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.inicioDepois||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.fimDepois||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.statusAntes||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.statusDepois||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(String(r.alocacaoAntes ?? ''))}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(String(r.alocacaoDepois ?? ''))}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; max-width:220px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${escHTML(r.tagsAntes||'')}">${escHTML(r.tagsAntes||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; max-width:220px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${escHTML(r.tagsDepois||'')}">${escHTML(r.tagsDepois||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9; max-width:320px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${escHTML(r.motivo||'')}">${escHTML(r.motivo||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.usuario||'')}</td>
      <td style="padding:8px; border-bottom:1px solid #f1f5f9;">${escHTML(r.dataEvento||'')}</td>
    </tr>
  `).join('');
  baselineTabela.innerHTML = `
    <div style="border:1px solid #e2e8f0; border-radius:14px; overflow:hidden; background:#fff;">
      <table style="width:100%; border-collapse:collapse; font-size:13px;">
        <thead style="background:#f8fafc;"><tr>${th}</tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

function exportComparisonCSV(){
  if(!lastComparisonRows || lastComparisonRows.length===0) return;
  const header = ['categoria','activityId','codigoAtividade','titulo','atividadeVinculadaA','recursoAntes','recursoDepois','inicioAntes','fimAntes','inicioDepois','fimDepois','statusAntes','statusDepois','alocacaoAntes','alocacaoDepois','tagsAntes','tagsDepois','motivo','usuario','dataEvento'];
  const esc = (s)=>{
    s = String(s??'');
    if(/["\n,;]/.test(s)) return '"'+s.replace(/"/g,'""')+'"';
    return s;
  };
  const rows = lastComparisonRows.map(r=>header.map(k=>esc(r[k]||'')).join(';'));
  const csv = header.join(';') + '\n' + rows.join('\n');
  const baseId = baselineSel && baselineSel.value ? baselineSel.value : '';
  const base = baselines.find(b=>b.id===baseId);
  const name = `comparacao_${(base && (base.nome||base.id)) ? (base.nome||base.id).replace(/\s+/g,'_') : 'baseline'}.csv`;
  download(name, csv, 'text/csv');
}

function exportBaselineCSV(){
  if(!baselineSel || !baselineSel.value){ alert('Selecione um baseline para baixar.'); return; }
  const baseId = baselineSel.value;
  const base = baselines.find(b=>b.id===baseId);
  if(!base){ alert('Baseline não encontrado.'); return; }
  const baseList = baselineItems.filter(i=>i.baselineId===baseId);
  const header = ['baselineId','baselineNome','createdAt','createdBy','usarFiltros','filterSnapshot','activityId','codigoAtividade','titulo','resourceId','resourceNome','resourceCapacidade','resourceCargaHorariaDiaria','linkedOriginId','linkedOriginCode','inicio','fim','status','alocacao','tags'];
  const esc = (s)=>{
    s = String(s??'');
    if(/["]|\n|,|;/.test(s)) return '"'+s.replace(/"/g,'""')+'"';
    return s;
  };
  const rows = baseList.map(r=>{
    const row = {
      baselineId: base.id,
      baselineNome: base.nome||'',
      createdAt: base.createdAt||'',
      createdBy: base.createdBy||'',
      usarFiltros: base.usarFiltros ? 'S' : 'N',
      filterSnapshot: base.filterSnapshot ? JSON.stringify(base.filterSnapshot) : '',
      activityId: r.activityId||'',
      codigoAtividade: r.codigoAtividade||'',
      titulo: r.titulo||'',
      resourceId: r.resourceId||'',
      resourceNome: r.resourceNome||'',
      resourceCapacidade: r.resourceCapacidade ?? '',
      resourceCargaHorariaDiaria: r.resourceCargaHorariaDiaria ?? '',
      linkedOriginId: r.linkedOriginId||'',
      linkedOriginCode: r.linkedOriginCode||'',
      inicio: r.inicio||'',
      fim: r.fim||'',
      status: r.status||'',
      alocacao: r.alocacao??'',
      tags: Array.isArray(r.tags)? r.tags.join(',') : ''
    };
    return header.map(k=>esc(row[k])).join(';');
  });
  const csv = header.join(';') + '\n' + rows.join('\n');
  const safeName = (base.nome||base.id).replace(/\s+/g,'_');
  download(`baseline_${safeName}.csv`, csv, 'text/csv');
}

function deleteBaseline(){
  if(!baselineSel || !baselineSel.value){ alert('Selecione um baseline para excluir.'); return; }
  const baseId = baselineSel.value;
  const base = baselines.find(b=>b.id===baseId);
  if(!base){ alert('Baseline não encontrado.'); return; }
  const ok = confirm(`Excluir baseline "${base.nome||base.id}"?\n\nIsso removerá também os itens capturados e não pode ser desfeito.`);
  if(!ok) return;
  baselines = baselines.filter(b=>b.id!==baseId);
  baselineItems = baselineItems.filter(i=>i.baselineId!==baseId);
  saveBaselines();
  renderBaselineSelect();
  // limpa UI
  lastComparisonRows = [];
  if(kpiCards) kpiCards.innerHTML = '';
  if(baselineTabela) baselineTabela.innerHTML = '<div class="muted">Baseline excluído. Selecione outro para comparar.</div>';
  if(btnExportComparacao) btnExportComparacao.disabled = true;
}

function initBaselineUI(){
  if(baselineNome && !baselineNome.value) baselineNome.value = defaultBaselineName();
  renderBaselineSelect();
  if(btnBaselineAplicarFiltros) btnBaselineAplicarFiltros.disabled = true;

  if(baselineSel){
    baselineSel.onchange = ()=>{
      const id = baselineSel.value;
      const base = baselines.find(b=>b.id===id);
      if(btnBaselineAplicarFiltros) btnBaselineAplicarFiltros.disabled = !(base && base.filterSnapshot);
      // Padrão seguro: se o baseline tiver snapshot, compare usando ele.
      if(base && base.filterSnapshot && baselineCompararFiltros) baselineCompararFiltros.checked = true;
      // Ajuda o usuário: ao selecionar, preenche o campo "Nome" para referência.
      if(base && baselineNome) baselineNome.value = base.nome || baselineNome.value;
    };
  }

  if(btnBaselineCriar) btnBaselineCriar.onclick = (ev)=>{ ev.preventDefault(); createBaseline(); };
  if(btnBaselineComparar) btnBaselineComparar.onclick = (ev)=>{ ev.preventDefault(); compareBaseline(); };
  if(btnBaselineAplicarFiltros) btnBaselineAplicarFiltros.onclick = (ev)=>{
    ev.preventDefault();
    if(!baselineSel || !baselineSel.value){ alert('Selecione um baseline para aplicar os filtros.'); return; }
    const base = baselines.find(b=>b.id===baselineSel.value);
    if(!base || !base.filterSnapshot){ alert('Este baseline não possui snapshot de filtros.'); return; }
    applyFilterSnapshotToUI(base.filterSnapshot);
  };
  if(btnBaselineBaixar) btnBaselineBaixar.onclick = (ev)=>{ ev.preventDefault(); exportBaselineCSV(); };
  if(btnBaselineExcluir) btnBaselineExcluir.onclick = (ev)=>{ ev.preventDefault(); deleteBaseline(); };
  if(btnExportComparacao) btnExportComparacao.onclick = (ev)=>{ ev.preventDefault(); exportComparisonCSV(); };
}

initBaselineUI();

// ===== Disponibilidade (% de capacidade livre) =====
(function(){
  const avBtn = document.getElementById('avBtn');
  const avRes = document.getElementById('avResultado');
  if(!avBtn || !avRes) return;

  function isBusinessDay(d){
    const wd = d.getDay();
    return wd>=1 && wd<=5;
  }

  function sumAllocationOn(resourceId, date) {
    const acts = activities.filter(a=>a.resourceId===resourceId && a.status!=='Concluída' && a.status!=='Cancelada' && fromYMD(a.inicio)<=date && date<=fromYMD(a.fim));
    return acts.reduce((acc,a)=>acc+(a.alocacao||100),0);
  }

  function findAvailabilityWindow(resource, startDate, daysNeeded, businessOnly, targetPerc, minAcceptablePerc){
    const cap = Number(resource.capacidade || 100);
    const limit = fromYMD(rangeEnd);
    let d = new Date(startDate);
    const maxSearch = new Date(startDate.getFullYear()+1, startDate.getMonth(), startDate.getDate());
    const hardLimit = limit && limit>startDate ? limit : maxSearch;
    function recIsActiveOn(dt){
      if(resource.inicioAtivo && dt < fromYMD(resource.inicioAtivo)) return false;
      if(resource.fimAtivo && dt > fromYMD(resource.fimAtivo)) return false;
      return true;
    }
    while(d <= hardLimit){
      let cnt = 0;
      let step = new Date(d);
      let actualStart = null;
      let minFree = Infinity;
      let okWindow = true;
      let guard = 0;
      while(cnt < daysNeeded && guard < 4000){
        guard++;
        if(businessOnly && !isBusinessDay(step)){
          step = new Date(step.getFullYear(), step.getMonth(), step.getDate()+1);
          continue;
        }
        if(!recIsActiveOn(step)) { okWindow=false; break; }
        const used = sumAllocationOn(resource.id, step);
        const free = (cap - used);
        if(actualStart===null) actualStart = new Date(step);
        minFree = Math.min(minFree, free);
        if(free < minAcceptablePerc){ okWindow=false; break; }
        cnt++;
        step = new Date(step.getFullYear(), step.getMonth(), step.getDate()+1);
      }
      if(okWindow && cnt>=daysNeeded && actualStart){
        const availablePct = Number.isFinite(minFree) ? Math.max(0, Math.round(minFree * 100) / 100) : 0;
        return {
          inicio: toYMD(actualStart),
          availablePct,
          deficitPct: Math.max(0, Math.round((targetPerc - availablePct) * 100) / 100),
          isApprox: availablePct < targetPerc
        };
      }
      d = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
    }
    return null;
  }

  function buildAvailabilityTable(items, includeApproxColumns){
    const table = document.createElement('table');
    table.className = 'tbl';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    const headers = ['Recurso','Tipo','Senioridade','Data mais próxima'];
    if(includeApproxColumns){
      headers.push('Capacidade livre encontrada (%)', 'Diferença para a meta (%)');
    }
    headers.forEach(label=>{
      const th = document.createElement('th');
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    const tbody = document.createElement('tbody');
    items.forEach(it=>{
      const tr = document.createElement('tr');
      const values = [it.recurso.nome, it.recurso.tipo, it.recurso.senioridade, it.inicio];
      if(includeApproxColumns){
        values.push(it.availablePct, it.deficitPct);
      }
      values.forEach(value=>{
        const td = document.createElement('td');
        td.textContent = String(value ?? '');
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(thead);
    table.appendChild(tbody);
    return table;
  }

  function runAvailability(){
    const dias = Math.max(1, Number(document.getElementById('avDias').value||1));
    const uteis = (document.getElementById('avUteis').value||'1')==='1';
    const percReq = Math.max(1, Number(document.getElementById('avPerc').value||25));
    const tolerancia = Math.max(0, Number(document.getElementById('avPercTol')?.value||0));
    const percMin = Math.max(0, percReq - tolerancia);
    const tipo = document.getElementById('avTipo').value||'';
    const sen = document.getElementById('avSenioridade').value||'';
    const inicioStr = document.getElementById('avInicio').value || toYMD(today);
    const inicio = fromYMD(inicioStr);

    const exact = [];
    const approx = [];
    resources.forEach(r=>{
      if(!r.ativo) return;
      if(tipo && (r.tipo||"").toLowerCase() !== tipo.toLowerCase()) return;
      if(sen && r.senioridade!==sen) return;
      const match = findAvailabilityWindow(r, inicio, dias, uteis, percReq, percMin);
      if(!match) return;
      const item = {recurso:r, ...match};
      if(match.isApprox) approx.push(item);
      else exact.push(item);
    });
    const sorter = (a,b)=> fromYMD(a.inicio) - fromYMD(b.inicio) || (a.deficitPct || 0) - (b.deficitPct || 0) || a.recurso.nome.localeCompare(b.recurso.nome);
    exact.sort(sorter);
    approx.sort(sorter);
    if(!exact.length && !approx.length){
      avRes.innerHTML = '<div class="muted">Nenhum recurso atende nem entra na faixa de tolerância dentro do horizonte de busca.</div>';
      return;
    }

    const wrap = document.createElement('div');
    wrap.style.display = 'grid';
    wrap.style.gap = '14px';

    if(exact.length){
      const sec = document.createElement('div');
      sec.innerHTML = `<div style="font-weight:700;margin-bottom:8px">Atendem integralmente (${exact.length})</div>`;
      sec.appendChild(buildAvailabilityTable(exact, false));
      wrap.appendChild(sec);
    }

    if(approx.length){
      const sec = document.createElement('div');
      sec.innerHTML = `<div style="font-weight:700;margin-bottom:8px">Próximos da meta (${approx.length})</div><div class="muted" style="margin-bottom:8px">Meta solicitada: ${percReq}%. Faixa mínima considerada: ${percMin}%.</div>`;
      sec.appendChild(buildAvailabilityTable(approx, true));
      wrap.appendChild(sec);
    }

    avRes.replaceChildren(wrap);
  }

  avBtn.addEventListener('click', runAvailability);
  const avInicio = document.getElementById('avInicio');
  if(avInicio && !avInicio.value){ avInicio.value = toYMD(today); }
})();
renderAll();

// ===== KPIs (Visão Executiva) =====
function renderKPIs(){
  try{
    const total=activities.length;
    const concluidas=activities.filter(a=>a.status==="Concluída").length;
    const perc=total? Math.round((concluidas/total)*100):0;
    const el1=document.getElementById("kpiExecucao"); if(el1) el1.textContent=perc+"%";
    const el2=document.getElementById("kpiRecursos"); if(el2) el2.textContent=resources.filter(r=>r.ativo).length;
    let sobre=0;
    resources.filter(r=>r.ativo).forEach(r=>{
      const cap=r.capacidade||100;
      const days=buildDays();
      for(const d of days){
        const acts=activities.filter(a=>a.resourceId===r.id && a.status !== 'Concluída' && a.status !== 'Cancelada' && fromYMD(a.inicio)<=d && d<=fromYMD(a.fim));
        const sum=acts.reduce((acc,a)=>acc+(a.alocacao||100),0);
        if(sum>cap){sobre++; break;}
      }
    });
    const el3=document.getElementById("kpiSobrecarga"); if(el3) el3.textContent=sobre;
    const el4=document.getElementById("kpiAtrasadasMes");
    if(el4){
      try{ el4.textContent = getUniqueChainCount(getOverdueInMonthView()); }catch(_){ el4.textContent = "0"; }
    }
    const el5=document.getElementById("kpiExtrapoladasAbertas");
    if(el5){
      try{ el5.textContent = getUniqueChainCount(getExtrapolatedActivities()); }catch(_){ el5.textContent = "0"; }
    }
    try{
      renderOverloadDetails();
    }catch(err){ console.error(err); }
  }catch(e){ /* ignora falhas */ }
}

// ===== Risco por Recurso (Fases 1 & 3) =====
/**
 * Calcula os scores de risco para recursos ativos. Para cada recurso, conta quantos dias do período
 * visualizado o total de alocação ultrapassa a capacidade (dias de sobrecarga). Adiciona uma
 * penalidade fixa de 5 pontos se o recurso for do tipo Externo. Retorna uma lista de objetos
 * {recurso, tipo, score, dias} apenas para recursos com dias>0.
 */
/**
 * Calcula scores de risco por recurso.
 *
 * Algoritmo anterior: O(Recursos × Dias × Atividades) — trava com volume.
 * Novo algoritmo: sweep-line por recurso — O(A log A + D) por recurso.
 *   1. Agrupa atividades ativas por resourceId (índice pré-construído).
 *   2. Para cada recurso, cria eventos de início/fim de alocação e os ordena.
 *   3. Varre os eventos em ordem, mantendo soma corrente de alocação.
 *   4. Conta dias de sobrecarga sem iterar sobre todos os dias do período.
 */
function computeRiskScores(){
  const out = [];
  try {
    const days = buildDays();
    if (!days.length) return out;
    const startDate = days[0];
    const endDate   = days[days.length - 1];

    // Índice: resourceId → atividades ativas no período de visão
    const actsByRes = {};
    (resources || []).filter(r => r.ativo && !r.deletedAt).forEach(r => { actsByRes[r.id] = []; });
    (activities || []).forEach(a => {
      if (a.deletedAt) return;
      if (a.status === 'Concluída' || a.status === 'Cancelada') return;
      if (!actsByRes[a.resourceId]) return;
      // Inclui apenas atividades que interceptam o período de visão
      const s = fromYMD(a.inicio), e = fromYMD(a.fim);
      if (e < startDate || s > endDate) return;
      actsByRes[a.resourceId].push({ s, e, alloc: Number(a.alocacao || 100) });
    });

    (resources || []).filter(r => r.ativo && !r.deletedAt).forEach(r => {
      const cap = Number(r.capacidade || 100);
      const acts = actsByRes[r.id] || [];
      if (!acts.length) return;

      // Sweep-line: eventos [date, delta]
      const events = [];
      acts.forEach(({ s, e, alloc }) => {
        const clampedS = s < startDate ? startDate : s;
        const clampedE = e > endDate   ? endDate   : e;
        events.push([clampedS.getTime(), +alloc]);
        // O evento de saída deve ocorrer no dia seguinte ao fim
        const after = new Date(clampedE); after.setDate(after.getDate() + 1);
        events.push([after.getTime(), -alloc]);
      });
      events.sort((a, b) => a[0] - b[0]);

      // Mapeia dias com sobrecarga usando o sweep
      const overloadDays = new Set();
      let sum = 0;
      let ei = 0;
      const msDay = 86400000;
      days.forEach(d => {
        const ts = d.getTime();
        // Aplica todos os eventos até o início deste dia
        while (ei < events.length && events[ei][0] <= ts) {
          sum += events[ei][1];
          ei++;
        }
        if (sum > cap) overloadDays.add(toYMD(d));
      });

      const overload = overloadDays.size;
      if (overload > 0) {
        let score = overload;
        if ((r.tipo || '').toLowerCase() === 'externo') score += 5;
        out.push({ recurso: r.nome, tipo: r.tipo, score, dias: overload });
      }
    });
  } catch(e) { console.error(e); }
  return out;
}

/**
 * Renderiza a tabela de risco no painel "Risco por Recurso". Caso a caixa de seleção
 * "Mostrar apenas recursos com risco" esteja marcada, lista somente recursos com dias de
 * sobrecarga &gt; 0; caso contrário, lista todos os recursos ativos com score calculado (0 para
 * aqueles sem sobrecarga).
 */
function renderRiskScores(){
  const tbl=document.getElementById('tblRisco');
  if(!tbl) return;
  const only=document.getElementById('riskOnlyToggle')?.checked;
  // Construir lista de objetos de risco para todos os recursos ativos
  const items=[];
  const scoreMap = {};
  const computed = computeRiskScores();
  // computed contém apenas recursos com dias>0
  computed.forEach(s=>{
    scoreMap[s.recurso] = {score: s.score, dias: s.dias, tipo: s.tipo};
  });
  resources.filter(r=>r.ativo).forEach(r=>{
    const found = scoreMap[r.nome];
    const dias = found? found.dias: 0;
    let score = dias;
    const tipoLower = (r.tipo||'').toLowerCase();
    if(tipoLower==='externo' && dias>0) score += 5;
    items.push({ recurso: r.nome, tipo: r.tipo, score: score || 0, dias: dias });
  });
  // se apenas com risco: filtrar itens com dias > 0
  const list = only ? items.filter(i=> i.dias > 0) : items;
  // ordenar por score ascendente, depois por recurso
  list.sort((a,b)=>{
    if(a.score !== b.score) return a.score - b.score;
    if(a.dias !== b.dias) return a.dias - b.dias;
    return a.recurso.localeCompare(b.recurso);
  });
  let html='';
  list.forEach(s=>{
    html += `<tr><td>${s.recurso}</td><td>${s.tipo}</td><td>${s.score}</td><td>${s.dias}</td></tr>`;
  });
  tbl.querySelector('tbody').innerHTML = html;
}

// ===== Exportação filtrada (Fase 4) =====
/**
 * Obtém dados filtrados conforme período, lista de recursos e tipo escolhidos na seção de exportação filtrada.
 */
function getFilteredReport(){
  const start=document.getElementById('expStart')?.value||'';
  const end=document.getElementById('expEnd')?.value||'';
  const names=(document.getElementById('expResources')?.value||'').split(',').map(s=>s.trim().toLowerCase()).filter(Boolean);
  const type=(document.getElementById('expType')?.value||'').toLowerCase();
  // Filtra recursos ativos (ignorando deletados) e aplica filtros de nome e tipo
  const resFilt=getActiveResources().filter(r=>{
    if(names.length && !names.some(n=> (r.nome||'').toLowerCase().includes(n))) return false;
    if(type && type!=='' && (r.tipo||'').toLowerCase()!==type) return false;
    return true;
  });
  // Filtra atividades ativas (ignorando deletadas) e dentro do período, e que pertençam a um recurso filtrado
  const actsFilt=getActiveActivities().filter(a=>{
    const rec=resFilt.find(x=>x.id===a.resourceId);
    if(!rec) return false;
    if(start && fromYMD(a.fim)<fromYMD(start)) return false;
    if(end && fromYMD(a.inicio)>fromYMD(end)) return false;
    return true;
  });
  return {resources: resFilt, activities: actsFilt};
}

/**
 * Exporta os dados filtrados em dois CSVs (recursos e atividades). Inclui cabeçalhos para nome,
 * tipo, senioridade, capacidade e dados das atividades correspondentes.
 */
function exportFilteredCSV(){
  const data=getFilteredReport();
  const recRows=data.resources.map(r=>({nome:r.nome, tipo:r.tipo, senioridade:r.senioridade, capacidade:(r.capacidade||100), cargaHorariaDiaria:getDailyHours(r)}));
  const actRows=data.activities.map(a=>{
    const linked = a.linkedOriginId ? data.activities.find(x=>x.id===a.linkedOriginId) || activities.find(x=>x.id===a.linkedOriginId) : null;
    return ({
      codigoAtividade: a.codigoAtividade || '',
      titulo:a.titulo,
      recurso:(resources.find(x=>x.id===a.resourceId)?.nome)||'',
      atividadeVinculadaA: linked ? (linked.codigoAtividade || linked.id || '') : '',
      extrapolada: isActivityExtrapolated(a) ? 'S' : 'N',
      status:a.status,
      inicio:a.inicio,
      fim:a.fim,
      alocacao:(a.alocacao||100)
    });
  });
  if(recRows.length===0 && actRows.length===0){
    alert('Nenhum dado encontrado para o filtro aplicado.');
    return;
  }
  download('recursos_filtrados.csv', toCSV(recRows, ['nome','tipo','senioridade','capacidade','cargaHorariaDiaria']), 'text/csv;charset=utf-8');
  download('atividades_filtradas.csv', toCSV(actRows, ['codigoAtividade','titulo','recurso','atividadeVinculadaA','extrapolada','status','inicio','fim','alocacao']), 'text/csv;charset=utf-8');
  alert('Exportados: recursos_filtrados.csv e atividades_filtradas.csv');
}

/**
 * Gera um relatório em HTML (em nova aba) dos dados filtrados. Calcula indicadores como % de execução
 * (com base nas atividades filtradas), quantidade de recursos ativos e recursos sobrecarregados. Se
 * possível, usa a janela para exibir as tabelas; caso contrário, faz download de um arquivo HTML.
 */
function exportFilteredPDF(){
  const data=getFilteredReport();
  const win=window.open('', '_blank');
  const now=new Date().toLocaleString();
  const totalActs=data.activities.length;
  const concl=data.activities.filter(a=>a.status==='Concluída').length;
  const execPerc=totalActs?Math.round((concl/totalActs)*100):0;
  const riskScores=computeRiskScores();
  // número de recursos sobrecarregados dentro do filtro
  const overCount=data.resources.filter(r=>riskScores.some(s=>s.recurso===r.nome)).length;
  const htmlParts=[];
  htmlParts.push('<html><head><meta charset="utf-8"><title>Relatório Planejador de Recursos (Filtrado)</title>');
  htmlParts.push('<style>body{font-family:sans-serif;padding:20px;} table{border-collapse:collapse;width:100%;margin-top:12px;} th,td{border:1px solid #ccc;padding:4px;} th{background:#f2f2f2;} .kpi span{font-weight:bold;margin-right:4px;}</style>');
  htmlParts.push('</head><body>');
  htmlParts.push('<h2>Relatório Planejador de Recursos (Filtrado)</h2>');
  htmlParts.push('<p>Gerado em: '+now+'</p>');
  htmlParts.push('<h3>KPIs</h3>');
  htmlParts.push('<p><span class="kpi">% Execução:</span> '+execPerc+'%<br>');
  htmlParts.push('<span class="kpi">Recursos Ativos:</span> '+data.resources.length+'<br>');
  htmlParts.push('<span class="kpi">Recursos Sobrecarregados:</span> '+overCount+'</p>');
  if(data.resources.length>0){
    htmlParts.push('<h3>Recursos</h3>');
    htmlParts.push('<table><thead><tr><th>Nome</th><th>Tipo</th><th>Senioridade</th><th>Capacidade%</th><th>Carga horária diária (h)</th></tr></thead><tbody>');
    data.resources.forEach(r=>{
      htmlParts.push('<tr><td>'+escHTML(r.nome)+'</td><td>'+escHTML(r.tipo)+'</td><td>'+escHTML(r.senioridade)+'</td><td>'+ (r.capacidade||100) +'</td><td>'+ getDailyHours(r) +'</td></tr>');
    });
    htmlParts.push('</tbody></table>');
  }
  if(data.activities.length>0){
    htmlParts.push('<h3>Atividades</h3>');
    htmlParts.push('<table><thead><tr><th>Código</th><th>Título</th><th>Recurso</th><th>Vinculada a</th><th>Extrapolada</th><th>Status</th><th>Início</th><th>Fim</th><th>Alocação%</th></tr></thead><tbody>');
    data.activities.forEach(a=>{
      const rName=(data.resources.find(x=>x.id===a.resourceId)?.nome)||'';
      const linked = a.linkedOriginId ? data.activities.find(x=>x.id===a.linkedOriginId) || activities.find(x=>x.id===a.linkedOriginId) : null;
      htmlParts.push('<tr><td>'+escHTML(a.codigoAtividade || '')+'</td><td>'+escHTML(a.titulo)+'</td><td>'+escHTML(rName)+'</td><td>'+escHTML(linked ? (linked.codigoAtividade || linked.id || '') : '')+'</td><td>'+(isActivityExtrapolated(a) ? 'Sim' : 'Não')+'</td><td>'+escHTML(a.status)+'</td><td>'+a.inicio+'</td><td>'+a.fim+'</td><td>'+ (a.alocacao||100) +'</td></tr>');
    });
    htmlParts.push('</tbody></table>');
  }
  htmlParts.push('</body></html>');
  const html=htmlParts.join('');
  if(win){
    win.document.write(html);
    win.document.close();
  } else {
    download('relatorio_filtrado.html', html, 'text/html;charset=utf-8');
  }
}

// ===== Exportar PDF =====
(function(){
  const btn=document.getElementById("btnExportPDF");
  if(!btn) return;
  btn.onclick = async () => {
    const hasJsPDF = !!(window.jspdf && window.jspdf.jsPDF);
    const hasHtml2Canvas = !!window.html2canvas;
    if(hasJsPDF){
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();
      doc.setFontSize(14);
      doc.text("Relatório Planejador de Recursos", 14, 20);
      doc.setFontSize(10);
      doc.text("Gerado em: "+new Date().toLocaleString(), 14, 28);
      renderKPIs();
      doc.text("KPIs:", 14, 40);
      const k1=document.getElementById("kpiExecucao")?.textContent||"";
      const k2=document.getElementById("kpiRecursos")?.textContent||"";
      const k3=document.getElementById("kpiSobrecarga")?.textContent||"";
      doc.text("% Execução: "+k1, 20, 48);
      doc.text("Recursos Ativos: "+k2, 20, 56);
      doc.text("Recursos Sobrecarregados: "+k3, 20, 64);
      doc.text("Recursos:", 14, 78);
      let y=86;
      resources.forEach(r=>{ doc.text("- "+r.nome+" ("+r.tipo+", "+r.senioridade+", Cap "+(r.capacidade||100)+"%)", 20, y); y+=6; if(y>270){doc.addPage(); y=20;} });
      y+=6; doc.text("Atividades:", 14, y); y+=8;
      activities.forEach(a=>{
        const rec=resources.find(r=>r.id===a.resourceId);
        doc.text("- "+a.titulo+" ("+(rec?rec.nome:"—")+") ["+a.status+"] "+a.inicio+" → "+a.fim, 20, y);
        y+=6; if(y>270){doc.addPage(); y=20;}
      });
      if(hasHtml2Canvas){
        try{
          const canvas = await html2canvas(document.getElementById("gantt"));
          const img = canvas.toDataURL("image/png");
          doc.addPage(); doc.text("Visão Gantt",14,20); doc.addImage(img,"PNG",14,30,180,100);
        }catch(e){ /* ignora */ }
      }
      doc.save("planejador_relatorio.pdf");
    } else {
      const w = window.open("", "_blank");
      const cssCompact = `body{font-family:Arial,sans-serif;padding:16px} h2{margin:16px 0 8px} table{width:100%;border-collapse:collapse;font-size:12px} th,td{border:1px solid #ddd;padding:6px}`;
      w.document.write("<html><head><title>Relatório Planejador</title><style>"+cssCompact+"</style></head><body>");
      w.document.write("<h1>Relatório Planejador de Recursos</h1>");
      w.document.write("<div>Gerado em: "+new Date().toLocaleString()+"</div>");
      renderKPIs();
      w.document.write("<h2>KPIs</h2>");
      w.document.write("<div>% Execução: "+(document.getElementById("kpiExecucao")?.textContent||"0%")+"</div>");
      w.document.write("<div>Recursos Ativos: "+(document.getElementById("kpiRecursos")?.textContent||"0")+"</div>");
      w.document.write("<div>Recursos Sobrecarregados: "+(document.getElementById("kpiSobrecarga")?.textContent||"0")+"</div>");
      w.document.write("<h2>Recursos</h2><table><tr><th>Nome</th><th>Tipo</th><th>Senioridade</th><th>Capacidade%</th></tr>");
      resources.forEach(r=>{ w.document.write("<tr><td>"+r.nome+"</td><td>"+r.tipo+"</td><td>"+r.senioridade+"</td><td>"+(r.capacidade||100)+"</td><td>"+getDailyHours(r)+"</td></tr>"); });
      w.document.write("</table>");
      w.document.write("<h2>Atividades</h2><table><tr><th>Título</th><th>Recurso</th><th>Status</th><th>Início</th><th>Fim</th><th>Alocação%</th></tr>");
      activities.forEach(a=>{
        const rec=resources.find(r=>r.id===a.resourceId);
        w.document.write("<tr><td>"+a.titulo+"</td><td>"+(rec?rec.nome:"—")+"</td><td>"+a.status+"</td><td>"+a.inicio+"</td><td>"+a.fim+"</td><td>"+(a.alocacao||100)+"</td></tr>");
      });
      w.document.write("</table>");
      try{ w.document.close(); w.focus(); w.print(); }catch(e){}
    }
  };
})();

// ===== Hook no renderAll para atualizar KPIs e Tags sem quebrar fluxo =====
(function(){
  const _renderAll = renderAll;
  renderAll = function(){
    _renderAll();
    renderKPIs();
    updateTagDatalist();
    // Atualiza o painel de risco sempre que a tela for re-renderizada
    try { renderRiskScores(); } catch(e) {}
    try { renderAuditTrail(); } catch(e) {}
  };
})();

// === Registro de eventos para painel de risco e exportações filtradas ===
document.addEventListener('DOMContentLoaded', () => {
  try {
    const riskToggle=document.getElementById('riskOnlyToggle');
    if(riskToggle){ riskToggle.addEventListener('change', () => { renderRiskScores(); }); }
    const btnCsv=document.getElementById('btnExpCSV');
    if(btnCsv){ btnCsv.addEventListener('click', exportFilteredCSV); }
    const btnPdf=document.getElementById('btnExpPDF');
    if(btnPdf){ btnPdf.addEventListener('click', exportFilteredPDF); }
    // Preenche datas padrão com o intervalo de visão atual
    const s=document.getElementById('expStart');
    const e=document.getElementById('expEnd');
    if(s && !s.value) s.value=rangeStart;
    if(e && !e.value) e.value=rangeEnd;
    // Renderiza painel de risco na carga inicial
    try { renderRiskScores(); } catch(err) {}

  } catch(err) {
    console.error(err);
  }
});

// === Helpers para detalhamento de sobrecarga na Visão Executiva ===
function computeOverloads(){
  const result=[];
  resources.forEach(r=>{
    const cap=r.capacidade||100;
    const tasks=activities.filter(a=>a.resourceId===r.id && a.status!=='Concluída' && a.status!=='Cancelada');
    if(!tasks.length) return;
    const events=[];
    tasks.forEach(a=>{
      const start=fromYMD(a.inicio);
      const end=fromYMD(a.fim);
      const alloc=a.alocacao||100;
      events.push({date:new Date(start),delta:alloc});
      const after=new Date(end);
      after.setDate(after.getDate()+1);
      events.push({date:after,delta:-alloc});
    });
    events.sort((a,b)=>a.date-b.date);
    let sum=0;
    let openStart=null;
    for(let i=0;i<events.length;i++){
      sum += events[i].delta;
      if(sum>cap && openStart===null){
        openStart=new Date(events[i].date);
      }
      if(sum<=cap && openStart!==null){
        const endDate=new Date(events[i].date);
        endDate.setDate(endDate.getDate()-1);
        const conc=tasks.filter(a=>{
          const s=fromYMD(a.inicio);
          const e=fromYMD(a.fim);
          return s<=endDate && e>=openStart;
        }).map(a=> a.titulo || ("Atividade "+a.id));
        result.push({nome:r.nome, periodo: toYMD(openStart)+" → "+toYMD(endDate), atividades: Array.from(new Set(conc)).join(', ')});
        openStart=null;
      }
    }
    if(openStart!==null){
      let lastEnd=new Date(0);
      tasks.forEach(a=>{
        const e=fromYMD(a.fim);
        if(e>lastEnd) lastEnd=new Date(e);
      });
      const conc=tasks.filter(a=>{
        const s=fromYMD(a.inicio);
        const e=fromYMD(a.fim);
        return s<=lastEnd && e>=openStart;
      }).map(a=> a.titulo || ("Atividade "+a.id));
      result.push({nome:r.nome, periodo: toYMD(openStart)+" → "+toYMD(lastEnd), atividades: Array.from(new Set(conc)).join(', ')});
    }
  });
  return result;
}

function renderOverloadDetails(){
  const tbody=document.querySelector('#overloadDetails tbody');
  const wrap=document.getElementById('overloadDetailsWrap');
  const empty=document.getElementById('overloadEmptyMsg');
  if(!tbody||!wrap||!empty) return;
  const rows=computeOverloads();
  if(!rows.length){
    wrap.style.display='none';
    empty.style.display='';
    return;
  }
  empty.style.display='none';
  wrap.style.display='';
  tbody.innerHTML = rows.map(r=>`<tr><td>${escHTML(r.nome)}</td><td>${escHTML(r.periodo)}</td><td>${escHTML(r.atividades)}</td></tr>`).join('');
}

// ===== BD por Excel/CSV (modelo único) =====
let bdFileHandle = null;

function updateBDStatus(msg){
  const el = document.getElementById('bdStatus');
  if(el) el.textContent = msg||'';
}

function parseHTMLExcelTables(htmlText){
  const doc = new DOMParser().parseFromString(htmlText, 'text/html');
  const tRec = doc.querySelector('#Recursos') || doc.querySelector('table[data-name="Recursos"]') || doc.querySelector('table:nth-of-type(1)');
  const tAtv = doc.querySelector('#Atividades') || doc.querySelector('table[data-name="Atividades"]') || doc.querySelector('table:nth-of-type(2)');
  function tableToObjects(tbl){
    if(!tbl) return [];
    const rows=[...tbl.querySelectorAll('tr')].map(tr=>[...tr.cells].map(td=>td.textContent.trim()));
    if(rows.length===0) return [];
    const headers=rows[0].map(h=>h.trim());
    return rows.slice(1).filter(r=>r.some(v=>v && v.trim().length)).map(r=>{
      const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o;
    });
  }
  return { recursos: tableToObjects(tRec), atividades: tableToObjects(tAtv) };
}

// Normaliza datas vindas de arquivos do BD (CSV/HTML) para o formato ISO esperado (YYYY-MM-DD).
// Aceita datas no formato brasileiro (dd/mm/yyyy), no formato ISO (yyyy-mm-dd) ou já normalizadas. Para outros valores retorna a string original.
function normalizeDateField(val){
  if(!val) return '';
  const s = String(val).trim();
  // separa por barra ou hífen
  const parts = s.split(/[\/\-]/);
  if(parts.length === 3){
    // dd/mm/yyyy ou dd-mm-yyyy
    if(parts[0].length === 2 && parts[1].length === 2 && parts[2].length === 4){
      return `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    // yyyy-mm-dd ou yyyy/mm/dd
    if(parts[0].length === 4 && parts[1].length <= 2 && parts[2].length <= 2){
      return `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
    }
  }
  return s;
}

function getDailyHours(resource){
  return Math.max(0.5, Number(resource?.cargaHorariaDiaria ?? resource?.cargaDiaria ?? 9) || 9);
}

function getDailyCapacityHours(resource, dateOrYmd){
  const ctx = dateOrYmd ? getDayContext(dateOrYmd, resource) : null;
  const overtime = ctx ? Number(ctx.overtimeApproved || 0) : 0;
  const base = (ctx && ctx.nonWorking) ? 0 : getDailyHours(resource) * (Math.max(0, Number(resource?.capacidade ?? 100) || 100) / 100);
  return base + Math.max(0, overtime);
}

function getActivityAllocatedHours(activity, resource){
  return getDailyHours(resource) * ((Number(activity?.alocacao ?? 100) || 100) / 100);
}

function buildDailyLoadIndex(filteredActs, days){
  const index = {};
  (resources || []).forEach(r=>{
    index[r.id] = {};
    days.forEach(d=>{
      const ymd = toYMD(d);
      const nominalHours = getDailyHours(r);
      const capacityHours = getDailyCapacityHours(r, d);
      const capacityPct = nominalHours > 0 ? (capacityHours / nominalHours) * 100 : 0;
      const overtimeApproved = getApprovedOvertimeHours(r, ymd);
      index[r.id][ymd] = { nominalHours, capacityPct, capacityHours, overtimeApproved, allocatedPct: 0, allocatedHours: 0, availableHours: capacityHours, excessHours: 0, activeActs: [] };
    });
  });
  (filteredActs || []).forEach(a=>{
    const r = (resources || []).find(x=>x.id===a.resourceId);
    if(!r) return;
    days.forEach(d=>{
      if(fromYMD(a.inicio)<=d && d<=fromYMD(a.fim)){
        const ymd = toYMD(d);
        const slot = index[r.id] && index[r.id][ymd];
        if(!slot) return;
        // Regra operacional: em finais de semana/feriados, a alocação só entra na capacidade agregada
        // quando houver hora extra aprovada para o recurso naquele dia. Sem HE aprovada, o dia fica
        // como não útil e não deve inflar o gráfico de carga agregada.
        const dayCtx = getDayContext(ymd, r);
        if(dayCtx && dayCtx.nonWorking && Number(dayCtx.overtimeApproved || 0) <= 0) return;
        slot.allocatedPct += Number(a.alocacao || 100);
        slot.allocatedHours += getActivityAllocatedHours(a, r);
        slot.activeActs.push(a);
      }
    });
  });
  Object.values(index).forEach(byDay=>{
    Object.values(byDay).forEach(slot=>{
      slot.availableHours = Math.max(0, slot.capacityHours - slot.allocatedHours);
      slot.excessHours = Math.max(0, slot.allocatedHours - slot.capacityHours);
    });
  });
  return index;
}

function coerceResource(r){
  return {
    id: String(r.id||r.ID||r.Id||''),
    nome: r.nome||r.Nome||r.NOME||'',
    tipo: (r.tipo||'').toLowerCase()||'interno',
    senioridade: (r.senioridade||'NA'),
    capacidade: Number(r.capacidade ?? r.Capacidade ?? 100),
    cargaHorariaDiaria: Math.max(0.5, Number(r.cargaHorariaDiaria ?? r.CargaHorariaDiaria ?? r.cargaDiaria ?? r.HorasDia ?? 9) || 9),
    ativo: String(r.ativo||'S').toUpperCase().startsWith('S'),
    // normaliza datas para ISO; se vierem vazias permanecem strings vazias
    inicioAtivo: normalizeDateField(r.inicioAtivo||r.InicioAtivo||r.inicio||''),
    fimAtivo: normalizeDateField(r.fimAtivo||r.FimAtivo||r.fim||''),
    version: Number(r.version||r.versao||0) || 0,
    updatedAt: r.updatedAt ? Number(r.updatedAt) : 0,
    deletedAt: r.deletedAt ? Number(r.deletedAt) : null
  };
}

function coerceActivity(a){
  const rawTags = Array.isArray(a.tags) ? a.tags : (a.tags || a.Tags || '');
  return {
    id: String(a.id||a.ID||a.Id||''),
    codigoAtividade: String(a.codigoAtividade || a.codigo || a.atividadeCodigo || a['Código Atividade'] || '').trim(),
    titulo: a.titulo||a.Titulo||a['TÍTULO']||'',
    resourceId: String(a.resourceId||a.RecursoID||a.Recurso||a.resource||''),
    linkedOriginId: String(a.linkedOriginId || a.LinkedOriginId || a.atividadeVinculadaId || '').trim(),
    linkedOriginCode: String(a.linkedOriginCode || a.LinkedOriginCode || a.atividadeVinculadaA || '').trim(),
    // normaliza inicio e fim para formato ISO esperado
    inicio: normalizeDateField(a.inicio||a.Inicio||a['Início']||''),
    fim: normalizeDateField(a.fim||a.Fim||''),
    status: (a.status||'planejada'),
    alocacao: Number(a.alocacao ?? a.Alocacao ?? 100),
    comentariosJson: decodeInlineText(a.comentariosJson ?? a.ComentariosJson ?? a.comentarios_json ?? a.comentariosHistorico ?? ''),
    comentarios: decodeInlineText(a.comentarios ?? a.Comentarios ?? a.comentario ?? a.Comentario ?? ''),
    tags: Array.isArray(rawTags) ? rawTags.map(t => String(t).trim()).filter(Boolean) : String(rawTags).split(',').map(t => t.trim()).filter(Boolean),
    execSubtasks: parseExecArray(a.execSubtasks || a.subatividades || a.execSubtasksJson),
    execEntries: parseExecArray(a.execEntries || a.apontamentos || a.execEntriesJson),
    execIssues: parseExecArray(a.execIssues || a.ocorrencias || a.execIssuesJson),
    execSubtasksJson: decodeInlineText(a.execSubtasksJson ?? a.subatividadesJson ?? ''),
    execEntriesJson: decodeInlineText(a.execEntriesJson ?? a.apontamentosJson ?? ''),
    execIssuesJson: decodeInlineText(a.execIssuesJson ?? a.ocorrenciasJson ?? ''),
    version: Number(a.version||a.versao||0) || 0,
    updatedAt: a.updatedAt ? Number(a.updatedAt) : 0,
    deletedAt: a.deletedAt ? Number(a.deletedAt) : null
  };
}

function hydrateLoadedActivities(list){
  const mapped = (list || []).map(coerceActivity);
  const ensured = ensureActivityCodes(mapped);
  const hydrated = ensured.list;
  hydrated.forEach(a => {
    const parsedComments = parseActivityComments(a.comentariosJson, a.comentarios, a.updatedAt);
    a.comentariosLista = parsedComments;
    a.comentariosJson = serializeActivityComments(parsedComments);
    a.comentarios = formatActivityCommentsForDisplay(parsedComments);
    syncExecFields(a);
  });
  const byId = new Map(hydrated.filter(a => a && !a.deletedAt && a.id).map(a => [String(a.id), a]));
  const byCode = new Map(hydrated.filter(a => a && !a.deletedAt && a.codigoAtividade).map(a => [String(a.codigoAtividade).trim().toLowerCase(), a]));
  hydrated.forEach(a => {
    if (!a || a.deletedAt) return;
    const rawId = String(a.linkedOriginId || '').trim();
    const rawCode = String(a.linkedOriginCode || '').trim();
    if (rawId && byId.has(rawId) && rawId !== a.id) {
      a.linkedOriginId = rawId;
      if (!rawCode) a.linkedOriginCode = String((byId.get(rawId) || {}).codigoAtividade || '').trim();
      return;
    }
    if (rawCode) {
      const linked = byCode.get(rawCode.toLowerCase());
      if (linked && linked.id !== a.id) {
        a.linkedOriginId = linked.id;
        a.linkedOriginCode = String(linked.codigoAtividade || rawCode).trim();
        return;
      }
    }
    if (rawId && rawId !== a.id) {
      a.linkedOriginId = rawId;
      a.linkedOriginCode = rawCode || a.linkedOriginCode || '';
    } else if (!rawCode) {
      a.linkedOriginId = '';
      a.linkedOriginCode = '';
    } else {
      a.linkedOriginId = '';
      a.linkedOriginCode = rawCode;
    }
  });
  return hydrated;
}



function parseDelimitedRecords(text, sep){
  const records = [];
  let row = [];
  let cur = '';
  let inQuote = false;
  const src = String(text || '').replace(/^﻿/, '');
  for(let i=0;i<src.length;i++){
    const ch = src[i];
    const next = src[i+1];
    if(ch === '"'){
      if(inQuote && next === '"'){
        cur += '"';
        i += 1;
      }else{
        inQuote = !inQuote;
      }
      continue;
    }
    if(!inQuote && ch === sep){
      row.push(cur);
      cur = '';
      continue;
    }
    if(!inQuote && (ch === '\n' || ch === '\r')){
      if(ch === '\r' && next === '\n') i += 1;
      row.push(cur);
      if(row.some(v => String(v).trim().length)) records.push(row);
      row = [];
      cur = '';
      continue;
    }
    cur += ch;
  }
  row.push(cur);
  if(row.some(v => String(v).trim().length)) records.push(row);
  return records;
}

function parseCSVObjects(text){
  const rawText = String(text || '');
  if(!rawText.trim()) return [];
  const sep = rawText.split(/\r?\n/,1)[0].includes(';') ? ';' : ',';
  const records = parseDelimitedRecords(rawText, sep);
  if(!records.length) return [];
  const headers = records[0].map(h => String(h || '').trim());
  return records.slice(1).map(cols => {
    const o = {};
    headers.forEach((h,i)=>o[h] = String(cols[i] ?? '').trim());
    return o;
  });
}

function parseCSVUnico(text){
  const rows = parseCSVObjects(text);
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  return {recursos, atividades};
}

function parseCSVBDUnico(text){
  const rows = parseCSVObjects(text);
  if(!rows.length) return {recursos:[], atividades:[], horas:[], historico:[], feriados:[], horasExtras:[]};
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  const historico = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('historico'));
  const feriados = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('feriado'));
  const horasExtras = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('hora_extra')).map(r=>normalizeOvertimeEntry({ ...r, horas: r.horas || r.Horas || (r.minutos ? Number(r.minutos)/60 : ''), reason: r.justificativa || r.reason || '', createdBy: r.user || r.usuario || '' })).filter(Boolean);
  const horas = rows.filter(r => {
    const tab = String(r.tabela || '').toLowerCase();
    return tab.startsWith('hora') && !tab.startsWith('hora_cfg') && !tab.startsWith('hora_extra');
  }).map(r => {
    const id = r.id || r.resourceId || r.colaborador || r.Id || r.ID || '';
    const date = r.date || r.Date || r.data || r.Data || '';
    let minutos = r.minutos || r.Minutos || r.horas || r.Horas || '';
    const parseStr = (s) => {
      s = String(s || '').trim();
      if (!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if (m) { return parseInt(m[1], 10) * 60 + parseInt(m[2], 10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if (!isNaN(f)) return Math.round(f * 60);
      }
      const n = parseInt(s, 10);
      return isNaN(n) ? 0 : n;
    };
    minutos = parseStr(minutos);
    const tipo = r.tipo || r.Tipo || '';
    const projeto = r.projeto || r.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const cfg = rows.filter(r => String(r.tabela || '').toLowerCase().startsWith('hora_cfg')).map(r => {
    const rid = r.id || r.Id || r.resourceId || '';
    const horasDia = r.horasDia || r.horasdia || r.horasDiaMin || r.horas_dia || '';
    const dias = r.dias || r.Dias || r.dia || '';
    const projetos = r.projetos || r.Projetos || '';
    return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
  });
  const comments = commentRows.map(normalizeCommentRow).filter(Boolean);
  return { recursos, atividades, horas, cfg, historico, feriados, horasExtras, comments, usuariosSistema, auditEvents };
}

function parseCSVBDUnico(text){
  const rows = parseCSVObjects(text);
  if(!rows.length) return {recursos:[], atividades:[], horas:[], historico:[], feriados:[], horasExtras:[]};
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  const historico = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('historico'));
  const feriados = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('feriado'));
  const horasExtras = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('hora_extra')).map(r=>normalizeOvertimeEntry({ ...r, horas: r.horas || r.Horas || (r.minutos ? Number(r.minutos)/60 : ''), reason: r.justificativa || r.reason || '', createdBy: r.user || r.usuario || '' })).filter(Boolean);
  const horas = rows.filter(r => {
    const tab = String(r.tabela || '').toLowerCase();
    return tab.startsWith('hora') && !tab.startsWith('hora_cfg') && !tab.startsWith('hora_extra');
  }).map(r => {
    const id = r.id || r.resourceId || r.colaborador || r.Id || r.ID || '';
    const date = r.date || r.Date || r.data || r.Data || '';
    let minutos = r.minutos || r.Minutos || r.horas || r.Horas || '';
    const parseStr = (s) => {
      s = String(s || '').trim();
      if (!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if (m) { return parseInt(m[1], 10) * 60 + parseInt(m[2], 10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if (!isNaN(f)) return Math.round(f * 60);
      }
      const n = parseInt(s, 10);
      return isNaN(n) ? 0 : n;
    };
    minutos = parseStr(minutos);
    const tipo = r.tipo || r.Tipo || '';
    const projeto = r.projeto || r.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const cfg = rows.filter(r => String(r.tabela || '').toLowerCase().startsWith('hora_cfg')).map(r => {
    const rid = r.id || r.Id || r.resourceId || '';
    const horasDia = r.horasDia || r.horasdia || r.horasDiaMin || r.horas_dia || '';
    const dias = r.dias || r.Dias || r.dia || '';
    const projetos = r.projetos || r.Projetos || '';
    return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
  });
  const comments = commentRows.map(normalizeCommentRow).filter(Boolean);
  return { recursos, atividades, horas, cfg, historico, feriados, horasExtras, comments, usuariosSistema, auditEvents };
}

function parseHTMLBDTables(htmlText){
  const doc = new DOMParser().parseFromString(htmlText, 'text/html');
  const tRec = doc.querySelector('#Recursos') || doc.querySelector('table[data-name="Recursos"]') || doc.querySelector('table:nth-of-type(1)');
  const tAtv = doc.querySelector('#Atividades') || doc.querySelector('table[data-name="Atividades"]') || doc.querySelector('table:nth-of-type(2)');
  const tHoras = doc.querySelector('#HorasExternos') || doc.querySelector('table[data-name="HorasExternos"]') || doc.querySelector('table:nth-of-type(3)');
  const tComments = doc.querySelector('#Comentarios') || doc.querySelector('table[data-name="Comentarios"]');
  const tHist = doc.querySelector('#HistoricoAtividades') || doc.querySelector('table[data-name="HistoricoAtividades"]');
  const tFeriados = doc.querySelector('#Feriados') || doc.querySelector('table[data-name="Feriados"]');
  const tHorasExtras = doc.querySelector('#HorasExtras') || doc.querySelector('table[data-name="HorasExtras"]');
  const tUsuariosSistema = doc.querySelector('#UsuariosSistema') || doc.querySelector('table[data-name="UsuariosSistema"]');
  function tableToObjects(tbl){
    if(!tbl) return [];
    const rows=[...tbl.querySelectorAll('tr')].map(tr=>[...tr.cells].map(td=>td.textContent.trim()));
    if(rows.length===0) return [];
    const headers=rows[0].map(h=>h.trim());
    return rows.slice(1).filter(r=>r.some(v=>v && v.trim().length)).map(r=>{
      const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o;
    });
  }
  const recursos = tableToObjects(tRec);
  const atividades = tableToObjects(tAtv);
  const horasRows = tableToObjects(tHoras);
  const commentRows = tableToObjects(tComments);
  const historico = tableToObjects(tHist);
  const feriados = tableToObjects(tFeriados);
  const horasExtras = tableToObjects(tHorasExtras).map(r=>normalizeOvertimeEntry(r)).filter(Boolean);
  const usuariosSistema = tableToObjects(tUsuariosSistema).map(normalizeSystemUser).filter(Boolean);
  const tAuditTrail = doc.querySelector('#AuditTrail') || doc.querySelector('table[data-name="AuditTrail"]') || null;
  const auditEvents = tAuditTrail ? tableToObjects(tAuditTrail).map(normalizeAuditEvent).filter(Boolean) : [];
  const horas = horasRows.map(h=>{
    const id = h.id || h.ID || h.resourceId || h.RecursoID || h.colaborador || h.Colaborador || '';
    const date = h.date || h.Date || h.data || h.Data || '';
    let minutos = h.minutos || h.Minutos || h.horas || h.Horas || '';
    const parseStr = (s) => {
      s = String(s||'').trim();
      if(!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if(m){ return parseInt(m[1],10)*60 + parseInt(m[2],10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if(!isNaN(f)) return Math.round(f*60);
      }
      const n = parseInt(s,10);
      return isNaN(n)?0:n;
    };
    minutos = parseStr(minutos);
    const tipo = h.tipo || h.Tipo || '';
    const projeto = h.projeto || h.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const tCfg = doc.querySelector('#HorasExternosCfg') || doc.querySelector('table[data-name="HorasExternosCfg"]') || null;
  let cfg = [];
  if (tCfg) {
    const cfgRows = tableToObjects(tCfg);
    cfg = cfgRows.map(row => {
      const rid = row.id || row.ID || row.resourceId || '';
      const horasDia = row.horasDia || row.horasdia || row.horasDiaMin || row.horas_dia || row.horas_diarias || '';
      const dias = row.dias || row.Dias || row.dia || '';
      const projetos = row.projetos || row.Projetos || row.projeto_cfg || '';
      return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
    });
  }
  const comments = commentRows.map(normalizeCommentRow).filter(Boolean);
  return { recursos, atividades, horas, cfg, historico, feriados, horasExtras, comments, usuariosSistema, auditEvents };
}

function parseCSVBDUnico(text){
  const rows = parseCSVObjects(text);
  if(!rows.length) return {recursos:[], atividades:[], horas:[], historico:[], feriados:[], horasExtras:[]};
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  const historico = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('historico'));
  const feriados = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('feriado'));
  const horasExtras = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('hora_extra')).map(r=>normalizeOvertimeEntry({ ...r, horas: r.horas || r.Horas || (r.minutos ? Number(r.minutos)/60 : ''), reason: r.justificativa || r.reason || '', createdBy: r.user || r.usuario || '' })).filter(Boolean);
  const horas = rows.filter(r => {
    const tab = String(r.tabela || '').toLowerCase();
    return tab.startsWith('hora') && !tab.startsWith('hora_cfg') && !tab.startsWith('hora_extra');
  }).map(r => {
    const id = r.id || r.resourceId || r.colaborador || r.Id || r.ID || '';
    const date = r.date || r.Date || r.data || r.Data || '';
    let minutos = r.minutos || r.Minutos || r.horas || r.Horas || '';
    const parseStr = (s) => {
      s = String(s || '').trim();
      if (!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if (m) { return parseInt(m[1], 10) * 60 + parseInt(m[2], 10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if (!isNaN(f)) return Math.round(f * 60);
      }
      const n = parseInt(s, 10);
      return isNaN(n) ? 0 : n;
    };
    minutos = parseStr(minutos);
    const tipo = r.tipo || r.Tipo || '';
    const projeto = r.projeto || r.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const cfg = rows.filter(r => String(r.tabela || '').toLowerCase().startsWith('hora_cfg')).map(r => {
    const rid = r.id || r.Id || r.resourceId || '';
    const horasDia = r.horasDia || r.horasdia || r.horasDiaMin || r.horas_dia || '';
    const dias = r.dias || r.Dias || r.dia || '';
    const projetos = r.projetos || r.Projetos || '';
    return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
  });
  return { recursos, atividades, horas, cfg, historico, feriados, horasExtras, comments:(rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('coment')).map(normalizeCommentRow).filter(Boolean)), usuariosSistema:(rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('usuario_sistema') || String(r.tabela||'').toLowerCase().startsWith('usuariosistema')).map(normalizeSystemUser).filter(Boolean)), auditEvents:(rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('audit')).map(normalizeAuditEvent).filter(Boolean)) };
}

async function saveBD() {
  if (!bdHandle) return;
  try {
    updateDBStatusBanner('syncing');
    const hasWritePermission = await ensureBDWritePermission({ request: false });
    if (!hasWritePermission) {
      updateBDStatus('Reautorização necessária para gravar no BD');
      updateDBStatusBanner('stale');
      return false;
    }
    // Aguarda se outra sessão estiver salvando recentemente e previne concorrência
    const lockAcquired = await acquireBDLockIfBusy();
    if (!lockAcquired) {
      updateBDStatus('Fila de gravação indisponível');
      updateDBStatusBanner('stale');
      showToast('Fila de gravação indisponível', 'Não foi possível aguardar a janela de gravação.', 'warning', 5000);
      return false;
    }
    // Se o arquivo mudou enquanto aguardávamos a fila, apenas informa e
    // prossegue com a gravação. O watcher já atualiza __bdLastWrite quando
    // detecta alterações externas; aqui evitamos transformar a espera normal
    // em um conflito fatal para o usuário.
    try {
      const fchk = await bdHandle.getFile();
      const currentLm = fchk.lastModified;
      if (typeof window !== 'undefined' && currentLm > (window.__bdLastWrite || 0)) {
        updateBDStatus('Banco atualizado por outra sessão');
        updateDBStatusBanner('stale');
        showToast('Banco atualizado por outra sessão', 'A fila foi respeitada. A gravação vai continuar sem descartar o formulário aberto.', 'info', 4500);
      }
    } catch (e) {
      // falha ao checar modificação; prosseguir com salvamento
    }
    let content = '';
    let mime = '';
    let horasList = [];
    let cfgList = [];
    let feriadosList = [];
    let horasExtrasList = getOvertimeEntries();
    try {
      if (typeof window.getHorasExternosData === 'function') {
        const out = window.getHorasExternosData();
        if (Array.isArray(out)) horasList = out;
      }
    } catch(e){}
    try {
      if (typeof window.getHorasExternosConfig === 'function') {
        const outCfg = window.getHorasExternosConfig();
        if (Array.isArray(outCfg)) cfgList = outCfg;
      }
    } catch(e){}
    try {
        if (typeof window.getFeriados === 'function') {
            const outFeriados = window.getFeriados();
            if(Array.isArray(outFeriados)) feriadosList = outFeriados;
        }
    } catch(e) {}
    if (bdFileExt === 'csv') {
      try {
        const existingFile = await bdHandle.getFile();
        const existingText = await existingFile.text();
        const persisted = parseCSVBDUnico(existingText);
        comments = mergeComentarios(persisted.comments || [], comments || []);
        rebuildCommentsIndex();
        syncAllActivityCommentFields();
      } catch(_e){}
      // Cabeçalho CSV incluindo campos de versionamento e exclusão (version, updatedAt, deletedAt)
      const header = [
        'tabela','id','nome','tipo','senioridade','capacidade','cargaHorariaDiaria','ativo','inicioAtivo','fimAtivo',
        'version','updatedAt','deletedAt',
        'codigoAtividade','titulo','resourceId','linkedOriginId','linkedOriginCode','extrapolada','inicio','fim','status','alocacao','comentarios','comentariosJson','tags','execSubtasksJson','execEntriesJson','execIssuesJson',
        'date','minutos','tipoHora','projeto','horasDia','dias','projetos',
        'activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user','legend','commentId','texto','usuario','ts','createdAt','matricula','perfil','pinHash','trocarPinNoPrimeiroAcesso','administradorInicial','createdBy','eventId','eventType','action','entityType','entityId','entityLabel','reason','nome','beforeJson','afterJson','source'
      ];
      const rows = [];
      // Gera linhas de recursos com campos extras
      resources.forEach(r => {
        rows.push({
          tabela:'recurso',
          id:r.id,
          nome:r.nome,
          tipo:r.tipo,
          senioridade:r.senioridade,
          capacidade:r.capacidade,
          cargaHorariaDiaria:getDailyHours(r),
          ativo:r.ativo?'S':'N',
          inicioAtivo:r.inicioAtivo||'',
          fimAtivo:r.fimAtivo||'',
          version:r.version || 0,
          updatedAt:r.updatedAt || 0,
          deletedAt:r.deletedAt || '',
          codigoAtividade:'', titulo:'', resourceId:'', linkedOriginId:'', linkedOriginCode:'', extrapolada:'', inicio:'', fim:'', status:'', alocacao:'', comentarios:'', comentariosJson:'', tags:'', execSubtasksJson:'', execEntriesJson:'', execIssuesJson:'', execSubtasksJson:'', execEntriesJson:'', execIssuesJson:'',
          date:'', minutos:'', tipoHora:'', projeto:'', horasDia:'', dias:'', projetos:'',
          activityId:'', timestamp:'', oldInicio:'', oldFim:'', newInicio:'', newFim:'', justificativa:'', user:'', legend:''
        });
      });
      // Gera linhas de atividades com campos extras
      activities.forEach(a => {
        const linked = a.linkedOriginId ? activities.find(x=>x.id===a.linkedOriginId) : null;
        rows.push({
          tabela:'atividade',
          id:a.id,
          nome:'', tipo:'', senioridade:'', capacidade:'', cargaHorariaDiaria:'', ativo:'', inicioAtivo:'', fimAtivo:'',
          version:a.version || 0,
          updatedAt:a.updatedAt || 0,
          deletedAt:a.deletedAt || '',
          codigoAtividade:a.codigoAtividade || '',
          titulo:a.titulo,
          resourceId:a.resourceId,
          linkedOriginId:a.linkedOriginId || '',
          linkedOriginCode:(linked ? (linked.codigoAtividade || '') : (a.linkedOriginCode || '')),
          extrapolada:isActivityExtrapolated(a) ? 'S' : 'N',
          inicio:a.inicio,
          fim:a.fim,
          status:a.status,
          alocacao:a.alocacao,
          comentarios:encodeInlineText(a.comentarios || ''),
          comentariosJson:a.comentariosJson || serializeActivityComments(a.comentariosLista || []),
          tags:(a.tags || []).join(', '),
          execSubtasksJson: syncExecFields({...a}).execSubtasksJson || '',
          execEntriesJson: syncExecFields({...a}).execEntriesJson || '',
          execIssuesJson: syncExecFields({...a}).execIssuesJson || '',
          date:'', minutos:'', tipoHora:'', projeto:'', horasDia:'', dias:'', projetos:'',
          activityId:'', timestamp:'', oldInicio:'', oldFim:'', newInicio:'', newFim:'', justificativa:'', user:'', legend:''
        });
      });
      horasList.forEach(h => {
        rows.push({
          tabela:'hora_externo', id:h.id, date:h.date || '', minutos:h.minutos, tipoHora:h.tipo || '', projeto:h.projeto || ''
        });
      });
      cfgList.forEach(cfg => {
        rows.push({
          tabela:'hora_cfg', id:cfg.id, horasDia: cfg.horasDia || '', dias: cfg.dias || '', projetos: cfg.projetos || ''
        });
      });
      (comments || []).filter(c => c && !c.deletedAt).forEach(c => {
        rows.push({
          tabela:'comentario',
          id:'', nome:'', tipo:'', senioridade:'', capacidade:'', cargaHorariaDiaria:'', ativo:'', inicioAtivo:'', fimAtivo:'',
          version:c.version || 1,
          updatedAt:c.updatedAt || c.createdAt || 0,
          deletedAt:c.deletedAt || '',
          codigoAtividade:'', titulo:'', resourceId:'', linkedOriginId:'', linkedOriginCode:'', extrapolada:'', inicio:'', fim:'', status:'', alocacao:'', comentarios:'', comentariosJson:'', tags:'', execSubtasksJson:'', execEntriesJson:'', execIssuesJson:'', execSubtasksJson:'', execEntriesJson:'', execIssuesJson:'',
          date:'', minutos:'', tipoHora:'', projeto:'', horasDia:'', dias:'', projetos:'',
          activityId:c.activityId || '', timestamp:'', oldInicio:'', oldFim:'', newInicio:'', newFim:'', justificativa:'', user:'', legend:'',
          commentId:c.commentId || '', texto:encodeInlineText(c.texto || ''), usuario:c.usuario || '', ts:c.ts || '', createdAt:c.createdAt || 0
        });
      });
      Object.keys(trails).forEach(activityId => {
        (trails[activityId] || []).forEach(entry => {
          // Persistência do histórico: usa 'legend' para carregar metadados extras (ex.: delegação/exclusão)
          let legend = entry.legend;
          if (!legend) {
            const meta = {
              type: entry.type || 'ALTERACAO_DATAS',
              entityType: entry.entityType || 'atividade',
              oldResourceId: entry.oldResourceId || '',
              oldResourceName: entry.oldResourceName || '',
              newResourceId: entry.newResourceId || '',
              newResourceName: entry.newResourceName || '',
              activityTitle: entry.activityTitle || '',
              activityCode: entry.activityCode || '',
              commentSnapshot: entry.commentSnapshot || '',
              deletedByCascade: entry.deletedByCascade ? true : false
            };
            // Só serializa se houver algo além do padrão de datas
            if (meta.type !== 'ALTERACAO_DATAS' || meta.entityType !== 'atividade' ||
                meta.oldResourceId || meta.newResourceId) {
              try { legend = JSON.stringify(meta); } catch(e) { legend = ''; }
            } else {
              legend = '';
            }
          }
          rows.push({
            tabela: 'historico',
            activityId: activityId,
            timestamp: entry.ts,
            oldInicio: entry.oldInicio,
            oldFim: entry.oldFim,
            newInicio: entry.newInicio,
            newFim: entry.newFim,
            justificativa: entry.justificativa,
            user: entry.user,
            legend: legend || ''
          });
        });
      });
      feriadosList.forEach(f => {
          rows.push({
              tabela: 'feriado',
              date: f.date,
              legend: f.legend
          });
      });
      horasExtrasList.forEach(h => {
          rows.push({
              tabela: 'hora_extra',
              id: h.id,
              resourceId: h.resourceId,
              date: h.date,
              minutos: Math.round(Number(h.hours || 0) * 60),
              tipoHora: 'extra',
              status: h.status || 'aprovada',
              justificativa: h.reason || '',
              user: h.createdBy || '',
              createdAt: h.createdAt || ''
          });
      });
      (systemUsers || []).forEach(u => {
          rows.push({
              tabela: 'usuario_sistema',
              matricula: u.matricula || '',
              nome: u.nome || '',
              perfil: u.perfil || 'Executor',
              ativo: u.ativo ? 'S' : 'N',
              pinHash: u.pinHash || '',
              trocarPinNoPrimeiroAcesso: u.trocarPinNoPrimeiroAcesso ? 'S' : 'N',
              administradorInicial: u.administradorInicial ? 'S' : 'N',
              createdAt: u.createdAt || '',
              updatedAt: u.updatedAt || '',
              createdBy: u.createdBy || ''
          });
      });
      normalizeAuditEvents(auditEvents || []).forEach(ev => {
          rows.push({
              tabela:'audit_trail', eventId:ev.eventId || '', ts:ev.ts || '', timestamp:ev.timestamp || '', eventType:ev.eventType || '', action:ev.action || '', entityType:ev.entityType || '', entityId:ev.entityId || '', entityLabel:ev.entityLabel || '', reason:ev.reason || '', matricula:ev.matricula || '', nome:ev.nome || '', perfil:ev.perfil || '', beforeJson:ev.beforeJson || '', afterJson:ev.afterJson || '', source:ev.source || ''
          });
      });
      const csvRows = [];
      csvRows.push(header.join(','));
      rows.forEach(row => {
        const vals = header.map(h => {
          let v = row[h] || '';
          const needsQuote = String(v).includes(',') || String(v).includes(';') || String(v).includes('"') || String(v).includes('\n') || String(v).includes('\r');
          v = String(v).replace(/"/g, '""');
          return needsQuote ? '"'+v+'"' : v;
        });
        csvRows.push(vals.join(','));
      });
      content = csvRows.join('\n');
      mime = 'text/csv;charset=utf-8';
    } else {
      function tableHTML(title, headers, rows) {
        const thead = headers.map(h => `<th>${h}</th>`).join('');
        const tbody = rows.map(r => `<tr>${headers.map(h => `<td>${r[h] ?? ''}</td>`).join('')}</tr>`).join('');
        return `<h3>${title}</h3><table id='${title}' data-name='${title}' border='1'><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
      }
      // Tabelas HTML: adiciona version, updatedAt e deletedAt para recursos e atividades
      const headersRec = ['id','nome','tipo','senioridade','capacidade','cargaHorariaDiaria','ativo','inicioAtivo','fimAtivo','version','updatedAt','deletedAt'];
      const recRows = resources.map(r => ({
        id:r.id,
        nome:r.nome,
        tipo:r.tipo,
        senioridade:r.senioridade,
        capacidade:r.capacidade,
        cargaHorariaDiaria:getDailyHours(r),
        ativo:r.ativo?'S':'N',
        inicioAtivo:r.inicioAtivo||'',
        fimAtivo:r.fimAtivo||'',
        version:r.version || 0,
        updatedAt:r.updatedAt || 0,
        deletedAt:r.deletedAt || ''
      }));
      const headersAtv = ['id','codigoAtividade','titulo','resourceId','linkedOriginId','linkedOriginCode','extrapolada','inicio','fim','status','alocacao','comentarios','comentariosJson','tags','execSubtasksJson','execEntriesJson','execIssuesJson','version','updatedAt','deletedAt'];
      const atvRows = activities.map(a => {
        const linked = a.linkedOriginId ? activities.find(x=>x.id===a.linkedOriginId) : null;
        return ({
          id:a.id,
          codigoAtividade:a.codigoAtividade || '',
          titulo:a.titulo,
          resourceId:a.resourceId,
          linkedOriginId:a.linkedOriginId || '',
          linkedOriginCode:(linked ? (linked.codigoAtividade || '') : (a.linkedOriginCode || '')),
          extrapolada:isActivityExtrapolated(a) ? 'S' : 'N',
          inicio:a.inicio,
          fim:a.fim,
          status:a.status,
          alocacao:a.alocacao,
          comentarios:a.comentarios || '',
          comentariosJson:a.comentariosJson || serializeActivityComments(a.comentariosLista || []),
          tags:(a.tags || []).join(', '),
          execSubtasksJson: syncExecFields({...a}).execSubtasksJson || '',
          execEntriesJson: syncExecFields({...a}).execEntriesJson || '',
          execIssuesJson: syncExecFields({...a}).execIssuesJson || '',
          version:a.version || 0,
          updatedAt:a.updatedAt || 0,
          deletedAt:a.deletedAt || ''
        });
      });
      const headersHoras = ['id','date','minutos','tipo','projeto'];
      const horasRows = horasList.map(h => ({ id:h.id, date:h.date || '', minutos:h.minutos, tipo:h.tipo || '', projeto:h.projeto || '' }));
      const headersCfg = ['id','horasDia','dias','projetos'];
      const cfgRows = cfgList.map(cfg => ({ id: cfg.id, horasDia: cfg.horasDia || '', dias: cfg.dias || '', projetos: cfg.projetos || '' }));
      const headersHist = ['activityId', 'timestamp', 'oldInicio', 'oldFim', 'newInicio', 'newFim', 'justificativa', 'user', 'legend'];
      const headersComments = ['commentId','activityId','texto','usuario','ts','createdAt','updatedAt','deletedAt','version'];
      const commentRows = (comments || []).filter(c => c && !c.deletedAt).map(c => ({
        commentId:c.commentId || '',
        activityId:c.activityId || '',
        texto:c.texto || '',
        usuario:c.usuario || '',
        ts:c.ts || '',
        createdAt:c.createdAt || 0,
        updatedAt:c.updatedAt || c.createdAt || 0,
        deletedAt:c.deletedAt || '',
        version:c.version || 1
      }));
      const histRows = [];
      Object.keys(trails).forEach(activityId => {
          (trails[activityId] || []).forEach(entry => {
              // Persistência do histórico (XLS/HTML): garante 'legend' com metadados extras (delegação/exclusão)
              let legend = entry.legend || '';
              if (!legend) {
                const meta = {
                  type: entry.type || 'ALTERACAO_DATAS',
                  entityType: entry.entityType || 'atividade',
                  oldResourceId: entry.oldResourceId || '',
                  oldResourceName: entry.oldResourceName || '',
                  newResourceId: entry.newResourceId || '',
                  newResourceName: entry.newResourceName || ''
                };
                if (meta.type !== 'ALTERACAO_DATAS' || meta.entityType !== 'atividade' ||
                    meta.oldResourceId || meta.newResourceId) {
                  try { legend = JSON.stringify(meta); } catch(e) { legend = ''; }
                }
              }

              histRows.push({
                  activityId: activityId,
                  timestamp: entry.ts,
                  oldInicio: entry.oldInicio,
                  oldFim: entry.oldFim,
                  newInicio: entry.newInicio,
                  newFim: entry.newFim,
                  justificativa: entry.justificativa,
                  user: entry.user,
                  legend: legend
              });
          });
      });

      const headersFeriados = ['date', 'legend'];
      const feriadosRows = feriadosList.map(f => ({ date: f.date, legend: f.legend || '' }));
      const headersHorasExtras = ['id','resourceId','date','hours','status','reason','createdBy','createdAt'];
      const horasExtrasRows = horasExtrasList.map(h => ({
        id: h.id || '',
        resourceId: h.resourceId || '',
        date: h.date || '',
        hours: Number(h.hours || 0),
        status: h.status || 'aprovada',
        reason: h.reason || '',
        createdBy: h.createdBy || '',
        createdAt: h.createdAt || ''
      }));
      const headersUsuariosSistema = ['matricula','nome','perfil','ativo','pinHash','trocarPinNoPrimeiroAcesso','administradorInicial','createdAt','updatedAt','createdBy'];

      const headersAuditTrail = ['eventId','ts','timestamp','eventType','action','entityType','entityId','entityLabel','reason','matricula','nome','perfil','beforeJson','afterJson','source'];
      const auditRows = normalizeAuditEvents(auditEvents || []).map(ev => ({
        eventId:ev.eventId || '', ts:ev.ts || '', timestamp:ev.timestamp || '', eventType:ev.eventType || '', action:ev.action || '', entityType:ev.entityType || '', entityId:ev.entityId || '', entityLabel:ev.entityLabel || '', reason:ev.reason || '', matricula:ev.matricula || '', nome:ev.nome || '', perfil:ev.perfil || '', beforeJson:ev.beforeJson || '', afterJson:ev.afterJson || '', source:ev.source || ''
      }));
      const usuariosSistemaRows = (systemUsers || []).map(u => ({
        matricula:u.matricula || '',
        nome:u.nome || '',
        perfil:u.perfil || 'Executor',
        ativo:u.ativo ? 'S' : 'N',
        pinHash:u.pinHash || '',
        trocarPinNoPrimeiroAcesso:u.trocarPinNoPrimeiroAcesso ? 'S' : 'N',
        administradorInicial:u.administradorInicial ? 'S' : 'N',
        createdAt:u.createdAt || '',
        updatedAt:u.updatedAt || '',
        createdBy:u.createdBy || ''
      }));
      
      content = `<!doctype html><html><head><meta charset='utf-8'><title>BD</title></head><body>`+
        tableHTML('Recursos', headersRec, recRows) +
        tableHTML('Atividades', headersAtv, atvRows) +
        tableHTML('HorasExternos', headersHoras, horasRows) +
        tableHTML('HorasExternosCfg', headersCfg, cfgRows) +
        tableHTML('Comentarios', headersComments, commentRows) +
        tableHTML('HistoricoAtividades', headersHist, histRows) +
        tableHTML('Feriados', headersFeriados, feriadosRows) +
        tableHTML('HorasExtras', headersHorasExtras, horasExtrasRows) +
        tableHTML('UsuariosSistema', headersUsuariosSistema, usuariosSistemaRows) +
        tableHTML('AuditTrail', headersAuditTrail, auditRows) +
        `</body></html>`;
      mime = 'text/html;charset=utf-8';
    }
    const writable = await bdHandle.createWritable();
    await writable.write(new Blob([content], { type: mime }));
    await writable.close();
    updateBDStatus('Salvo em ' + (bdFileName || 'BD'));
    updateDBStatusBanner('synced');
    // Toast silencioso de confirmação — aparece apenas quando o usuário aguardou o lock
    if (window.__bdShowSaveConfirm) {
      showToast('BD salvo', `Suas alterações foram gravadas em ${bdFileName || 'BD'}.`, 'success', 3500);
      window.__bdShowSaveConfirm = false;
    }
    if (typeof window !== 'undefined') {
      try {
        const savedFile = await bdHandle.getFile();
        window.__bdLastWrite = savedFile.lastModified;
      } catch (e) {
        window.__bdLastWrite = Date.now();
      }
    }
    try {
      adoptCurrentStateAsPersistedBaseline();
    } catch (e) {
      console.error('Erro ao atualizar baseline após salvar BD', e);
    }
    return true;
  } catch (e) {
    console.error('Erro ao salvar BD:', e);
    updateBDStatus('Erro ao salvar BD');
    updateDBStatusBanner('error');
    showToast('Erro ao salvar BD', e?.message || 'Verifique se o arquivo ainda está acessível na pasta de rede.', 'error', 8000);
    return false;
  }
}

function saveBDDebounced() {
  if (!bdHandle) return;
  clearTimeout(_saveBDTimer);
  _saveBDTimer = setTimeout(() => { saveBD(); }, 1000);
}

if (typeof window !== 'undefined') {
  try {
    window.onHorasExternosChange = () => {
      saveBDDebounced();
    };
  } catch(e){}
}

const fileBD = document.getElementById('fileBD');
if(fileBD){
  fileBD.onchange = async (ev)=>{
    const f = ev.target.files && ev.target.files[0];
    if(!f) return;
    try{
      const ext = f.name.toLowerCase().split('.').pop();
      const text = await f.text();
      let parsed;
      if(ext==='csv'){
        parsed = parseCSVBDUnico(text);
        resources = (parsed.recursos || []).map(coerceResource);
        activities = hydrateLoadedActivities(parsed.atividades || []);
        if(parsed.horas && typeof window.setHorasExternosData === 'function') window.setHorasExternosData(parsed.horas);
        if(parsed.cfg && typeof window.setHorasExternosConfig === 'function') window.setHorasExternosConfig(parsed.cfg);
        if(parsed.feriados && typeof window.setFeriados === 'function') window.setFeriados(parsed.feriados);
        if(parsed.horasExtras && typeof window.setHorasExtrasData === 'function') window.setHorasExtrasData(parsed.horasExtras);
      } else {
        parsed = parseHTMLBDTables(text);
        resources = (parsed.recursos || []).map(coerceResource);
        activities = hydrateLoadedActivities(parsed.atividades || []);
        if(parsed.horas && typeof window.setHorasExternosData === 'function') window.setHorasExternosData(parsed.horas);
        if(parsed.cfg && typeof window.setHorasExternosConfig === 'function') window.setHorasExternosConfig(parsed.cfg);
        if(parsed.feriados && typeof window.setFeriados === 'function') window.setFeriados(parsed.feriados);
        if(parsed.horasExtras && typeof window.setHorasExtrasData === 'function') window.setHorasExtrasData(parsed.horasExtras);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push(buildTrailEntryFromBDRow(h));
      });
      trails = newTrails;
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      saveLS(LS.trail, trails);
      ingestParsedAccess(parsed);
      renderAll();
      updateBDStatus('BD carregado: '+ f.name);
    } catch(e){ alert('Erro ao ler arquivo BD: '+ e.message); }
  };
}

const btnSelectDirInBD = document.getElementById('btnSelectDirInBD');
if(btnSelectDirInBD){
  btnSelectDirInBD.onclick = async ()=>{
    try{
      const h = await window.showDirectoryPicker();
      if(!h) return;
      dirHandle = h;
      await idbSet(FSA_DB, FSA_STORE, 'dir', h);
      updateBDStatus('Pasta selecionada ✓ — Salvo');
      try{ await verifyPerm(dirHandle); }catch(e){}
    }catch(e){
      alert('Não foi possível selecionar a pasta.\nDica: abra pelo Chrome/Edge via http(s):// em vez de file://');
      console.warn(e);
    }
  };
}

const btnReauthBD = document.getElementById('btnReauthBD');
if (btnReauthBD) {
  btnReauthBD.onclick = async () => {
    await reauthorizeBDWritePermission();
  };
}

async function selectAndLoadBDFile(){
    if (!('showOpenFilePicker' in window)) {
      alert('Seu navegador não suporta a abertura de arquivos com permissão de gravação. Use o Chrome/Edge via http(s)://');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Arquivos de Banco de Dados',
            accept: {
              'application/vnd.ms-excel': ['.xls', '.html'],
              'text/csv': ['.csv'],
              'text/plain': ['.csv']
            }
          }
        ]
      });
      if (!handle) return;
      bdHandle = handle;
      const file = await handle.getFile();
      bdFileName = file.name || '';
      bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
      const text = await file.text();
      let parsed;
      if (bdFileExt === 'csv') {
      try {
        const existingFile = await bdHandle.getFile();
        const existingText = await existingFile.text();
        const persisted = parseCSVBDUnico(existingText);
        comments = mergeComentarios(persisted.comments || [], comments || []);
        rebuildCommentsIndex();
        syncAllActivityCommentFields();
      } catch(_e){}
        parsed = parseCSVBDUnico(text);
      } else {
        parsed = parseHTMLBDTables(text);
      }
      resources = (parsed.recursos || []).map(coerceResource);
      activities = hydrateLoadedActivities(parsed.atividades || []);
      comments = hydrateLoadedComments(activities, parsed.comments || []);
      rebuildCommentsIndex();
      syncAllActivityCommentFields();
      if (parsed.horas && typeof window.setHorasExternosData === 'function') {
        window.setHorasExternosData(parsed.horas);
      }
      if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
        window.setHorasExternosConfig(parsed.cfg);
      }
      if (parsed.feriados && typeof window.setFeriados === 'function') {
        window.setFeriados(parsed.feriados);
      }
      if (parsed.horasExtras && typeof window.setHorasExtrasData === 'function') {
        window.setHorasExtrasData(parsed.horasExtras);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push(buildTrailEntryFromBDRow(h));
      });
      trails = newTrails;
      ingestParsedAccess(parsed);
      // Ao apontar um novo BD, reinicia o log de eventos e o snapshot para evitar reaplicar
      // eventos antigos (anteriores à escolha do BD) que poderiam trazer dados "fantasmas".
      adoptCurrentStateAsPersistedBaseline();
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      saveLS(LS.trail, trails);
      renderAll();
      updateBDStatus('BD carregado e pronto: ' + bdFileName);
      // registra o lastModified atual como a última modificação conhecida do arquivo
      try {
        const lm = file.lastModified;
        if (typeof window !== 'undefined') {
          window.__bdLastWrite = lm;
        }
      } catch(e){}
      startBDWatcher();
      updateDBStatusBanner(true);
    } catch (e) {
      if (e && e.name !== 'AbortError') {
        alert('Erro ao abrir arquivo BD: ' + e.message);
      }
    }
}
if(typeof window !== 'undefined') window.selectAndLoadBDFile = selectAndLoadBDFile;
const btnPickBDFile = document.getElementById('btnPickBDFile');
if(btnPickBDFile){ btnPickBDFile.onclick = selectAndLoadBDFile; }

// Permitir ao usuário definir um arquivo de BD como padrão para carregamento automático
const btnSetDefaultBD = document.getElementById('btnSetDefaultBD');
if(btnSetDefaultBD){
  btnSetDefaultBD.onclick = async () => {
    if (!('showOpenFilePicker' in window)) {
      alert('Seu navegador não suporta a abertura de arquivos com permissão de gravação. Use o Chrome/Edge via http(s)://');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Arquivos de Banco de Dados',
            accept: {
              'application/vnd.ms-excel': ['.xls', '.html'],
              'text/csv': ['.csv'],
              'text/plain': ['.csv']
            }
          }
        ]
      });
      if(!handle) return;
      bdHandle = handle;
      // Persistir handle como BD padrão
      try { await idbSet(FSA_DB, FSA_STORE, 'bd', handle); } catch(e) { console.warn('Erro ao salvar BD padrão', e); }
      const file = await handle.getFile();
      bdFileName = file.name || '';
      bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
      const text = await file.text();
      let parsed;
      if(bdFileExt === 'csv') parsed = parseCSVBDUnico(text);
      else parsed = parseHTMLBDTables(text);
      resources = (parsed.recursos || []).map(coerceResource);
      activities = hydrateLoadedActivities(parsed.atividades || []);
      comments = hydrateLoadedComments(activities, parsed.comments || []);
      rebuildCommentsIndex();
      syncAllActivityCommentFields();
      if (parsed.horas && typeof window.setHorasExternosData === 'function') {
        window.setHorasExternosData(parsed.horas);
      }
      if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
        window.setHorasExternosConfig(parsed.cfg);
      }
      if (parsed.feriados && typeof window.setFeriados === 'function') {
        window.setFeriados(parsed.feriados);
      }
      if (parsed.horasExtras && typeof window.setHorasExtrasData === 'function') {
        window.setHorasExtrasData(parsed.horasExtras);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push(buildTrailEntryFromBDRow(h));
      });
      trails = newTrails;
      ingestParsedAccess(parsed);
      // Ao definir um BD como padrão, reinicia o log de eventos e o snapshot para
      // evitar reaplicar eventos antigos e trazer dados não pertencentes ao BD.
      adoptCurrentStateAsPersistedBaseline();
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.comments, comments);
      saveLS(LS.trail, trails);
      renderAll();
      updateBDStatus('BD padrão definido: ' + bdFileName);
      // registra o lastModified atual como a última modificação conhecida do arquivo
      try {
        const lm = file.lastModified;
        if (typeof window !== 'undefined') {
          window.__bdLastWrite = lm;
        }
      } catch(e){}
      startBDWatcher();
      updateDBStatusBanner(true);
    } catch(e) {
      if(e && e.name !== 'AbortError'){
        alert('Erro ao definir BD padrão: ' + e.message);
      }
    }
  };
}

async function ensureDirOrAsk(){
  if(dirHandle) return true;
  try{
    const h = await window.showDirectoryPicker();
    if(!h) return false;
    dirHandle = h;
    await idbSet(FSA_DB, FSA_STORE, 'dir', h);
    updateBDStatus('Pasta selecionada ✓ — Salvo');
    return true;
  }catch(e){
    alert('Defina a pasta de dados para salvar o modelo.');
    return false;
  }
}

const btnExportModeloXLS = document.getElementById('btnExportModeloXLS');
if(btnExportModeloXLS){
  btnExportModeloXLS.onclick = () => {
    const headersRec = ['id','nome','tipo','senioridade','capacidade','cargaHorariaDiaria','ativo','inicioAtivo','fimAtivo'];
    const headersAtv = ['id','codigoAtividade','titulo','resourceId','linkedOriginId','linkedOriginCode','inicio','fim','status','alocacao','comentarios','comentariosJson','tags'];
    const headersHoras = ['id','date','minutos','tipo','projeto'];
    const exampleRec = [{id:'R1',nome:'Recurso Exemplo',tipo:'interno',senioridade:'Pl',capacidade:100,cargaHorariaDiaria:9,ativo:'S',inicioAtivo:'2025-01-01',fimAtivo:''}];
    const exampleAtv = [{id:'A1',codigoAtividade:'ATV-0001',titulo:'Atividade Exemplo',resourceId:'R1',linkedOriginId:'',linkedOriginCode:'',inicio:'2025-01-10',fim:'2025-01-20',status:'planejada',alocacao:100, comentarios:'02/04/2026 10:15 • Usuário\nComentário de exemplo', comentariosJson:'[{"id":"C1","ts":"2026-04-02T13:15:00.000Z","user":"Usuário","text":"Comentário de exemplo"}]', tags: 'SAP, Manutenção'}];
    const exampleHoras = [{id:'R1',date:'2025-01-15',minutos:480,tipo:'trabalho',projeto:'Alca Analitico'}];
    const headersCfg = ['id','horasDia','dias','projetos'];
    const exampleCfg = [{id:'R1',horasDia:'08:00',dias:'seg,ter,qua,qui,sex',projetos:'Alca Analitico:300:00'}];
    const headersHist = ['activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user','legend'];
    const exampleHist = [{activityId:'A1', timestamp:new Date().toISOString(), oldInicio:'2025-01-10', oldFim:'2025-01-20', newInicio:'2025-01-11', newFim:'2025-01-22', justificativa:'Ajuste de escopo', user:'usuário'}];
    const headersComments = ['commentId','activityId','texto','usuario','ts','createdAt','updatedAt','deletedAt','version'];
    const nowCommentTs = Date.now();
    const exampleComments = [{commentId:'C1', activityId:'A1', texto:'Comentário de exemplo', usuario:'Usuário', ts:new Date(nowCommentTs).toISOString(), createdAt:nowCommentTs, updatedAt:nowCommentTs, deletedAt:'', version:1}];
    const headersFeriados = ['date', 'legend'];
    const exampleFeriados = [{date: '2025-12-25', legend: 'Natal'}];
    function table(title, headers, rows){
      const thead = headers.map(h=>`<th>${h}</th>`).join('');
      const tbody = rows.map(r=>`<tr>${headers.map(h=>`<td>${(r[h]??'')}</td>`).join('')}</tr>`).join('');
      return `<h3>${title}</h3><table id="${title}" data-name="${title}" border='1'><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
    }
    const html = `<!doctype html><html><head><meta charset='utf-8'><title>Modelo BD</title></head><body>`+
      table('Recursos', headersRec, exampleRec) +
      table('Atividades', headersAtv, exampleAtv) +
      table('HorasExternos', headersHoras, exampleHoras) +
      table('HorasExternosCfg', headersCfg, exampleCfg) +
      table('Comentarios', headersComments, exampleComments) +
      table('HistoricoAtividades', headersHist, exampleHist) +
      table('Feriados', headersFeriados, exampleFeriados) +
      `</body></html>`;
    download('modelo_bd.xls', html, 'application/vnd.ms-excel');
    alert('Modelo de BD (Excel) gerado: modelo_bd.xls');
  };
}

const btnExportModeloCSV = document.getElementById('btnExportModeloCSV');
if(btnExportModeloCSV){
  btnExportModeloCSV.onclick = () => {
    const headers = ['tabela','id','nome','tipo','senioridade','capacidade','cargaHorariaDiaria','ativo','inicioAtivo','fimAtivo','codigoAtividade','titulo','resourceId','linkedOriginId','linkedOriginCode','inicio','fim','status','alocacao','comentarios','comentariosJson','tags','date','minutos','tipoHora','projeto','horasDia','dias','projetos', 'activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user', 'legend','commentId','texto','usuario','ts','createdAt'];
    const sample = [
      {tabela:'recurso',id:'R1',nome:'Recurso Exemplo',tipo:'interno',senioridade:'Pl',capacidade:100,cargaHorariaDiaria:9,ativo:'S',inicioAtivo:'2025-01-01',fimAtivo:''},
      {tabela:'atividade',id:'A1',codigoAtividade:'ATV-0001',titulo:'Atividade Exemplo',resourceId:'R1',linkedOriginId:'',linkedOriginCode:'',inicio:'2025-01-10',fim:'2025-01-20',status:'planejada',alocacao:100, comentarios:'02/04/2026 10:15 • Usuário\nComentário de exemplo', comentariosJson:'[{"id":"C1","ts":"2026-04-02T13:15:00.000Z","user":"Usuário","text":"Comentário de exemplo"}]', tags: 'SAP, Manutenção'},
      {tabela:'hora_externo',id:'R1',date:'2025-01-15',minutos:480,tipoHora:'trabalho',projeto:'Alca Analitico'},
      {tabela:'hora_cfg',id:'R1',horasDia:'08:00',dias:'seg,ter,qua,qui,sex',projetos:'Alca Analitico:300:00'},
      {tabela:'historico', activityId:'A1', timestamp:new Date().toISOString(), oldInicio:'2025-01-10', oldFim:'2025-01-20', newInicio:'2025-01-11', newFim:'2025-01-22', justificativa:'Ajuste de escopo', user:'usuário'},
      {tabela:'feriado', date:'2025-12-25', legend:'Natal'}
    ];
    const rows = [headers.join(',')];
    sample.forEach(obj => {
      const line = headers.map(h => {
        const v = obj[h] ?? '';
        const cleanV = String(v).replace(/"/g, '""');
        return '"' + cleanV + '"';
      }).join(',');
      rows.push(line);
    });
    const csv = rows.join('\n');
    download('modelo_bd.csv', csv, 'text/csv;charset=utf-8');
    alert('Modelo de BD (CSV único) gerado: modelo_bd.csv');
  };
}

(() => {
  if(dirHandle){ updateBDStatus('Pasta selecionada ✓ — Salvo'); }
})();

// ===== Importação em massa de dados =====
const fileImportData = document.getElementById('fileImportData');
if (fileImportData) {
  fileImportData.addEventListener('change', async (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    try {
      const text = await f.text();
      await importCSVData(text);
      // Limpa o valor para permitir importar o mesmo arquivo novamente se necessário
      ev.target.value = '';
    } catch (e) {
      alert('Erro ao importar dados: ' + (e && e.message ? e.message : e));
      console.error(e);
    }
  });
}

const btnExportImportTemplate = document.getElementById('btnExportImportTemplate');
if (btnExportImportTemplate) {
  btnExportImportTemplate.onclick = () => {
    // Cabeçalho e dados de exemplo para o modelo de importação. Utiliza ponto e vírgula
    // como separador para ser compatível com planilhas em português (Excel)
    const headers = ['tipo','nome','tipoRecurso','senioridade','capacidade','cargaHorariaDiaria','ativo','inicioAtivo','fimAtivo','codigoAtividade','titulo','recursoNome','linkedOriginCode','inicio','fim','status','alocacao','tags'];
    const samples = [
      {
        tipo: 'recurso',
        nome: 'Recurso Exemplo',
        tipoRecurso: 'Interno',
        senioridade: 'Pl',
        capacidade: 100,
        cargaHorariaDiaria: 9,
        ativo: 'S',
        inicioAtivo: '2025-01-01',
        fimAtivo: ''
      },
      {
        tipo: 'atividade',
        codigoAtividade: 'ATV-0001',
        titulo: 'Atividade Exemplo',
        recursoNome: 'Recurso Exemplo',
        linkedOriginCode: '',
        inicio: '2025-01-10',
        fim: '2025-01-20',
        status: 'planejada',
        alocacao: 100,
        tags: 'ProjetoX, API'
      }
    ];
    const sep = ';';
    const rows = [headers.join(sep)];
    samples.forEach(obj => {
      const line = headers.map(h => {
        let v = obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
        let s = String(v);
        // Escapa aspas
        s = s.replace(/"/g, '""');
        // Se contiver separador, vírgula ou quebra de linha, envolve em aspas
        if (s.includes(sep) || s.includes(',') || s.includes('\n')) {
          s = '"' + s + '"';
        }
        return s;
      }).join(sep);
      rows.push(line);
    });
    const csv = rows.join('\n');
    download('modelo_importacao.csv', csv, 'text/csv;charset=utf-8');
    alert('Modelo de importação gerado: modelo_importacao.csv');
  };
}

/**
 * Importa dados de um CSV para recursos e atividades. Cada linha deve conter
 * um campo "tipo" definindo se é "recurso" ou "atividade" (pode também
 * utilizar "tabela" com "recurso" ou "atividade"). Para recursos,
 * campos como nome, tipoRecurso, senioridade, capacidade, ativo,
 * inicioAtivo e fimAtivo são considerados. Para atividades, campos
 * titulo, recursoNome, inicio, fim, status, alocacao e tags são
 * utilizados. Recursos são mesclados pelo nome (case-insensitive);
 * atividades são mescladas pela combinação (titulo normalizado,
 * resourceId, inicio e fim). Novos IDs são gerados via uuid().
 * Para cada criação ou atualização, um evento é registrado. Ao final,
 * as listas resources e activities são salvas e a UI atualizada.
 * @param {string} csvText Conteúdo do arquivo CSV
 */
async function importCSVData(csvText) {
  try {
    const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) {
      alert('Arquivo vazio');
      return;
    }
    // Detecta separador (ponto e vírgula ou vírgula)
    const sep = lines[0].includes(';') ? ';' : ',';
    // Função para parsear uma linha com suporte a aspas
    function parseLine(line) {
      const cols = [];
      let cur = '';
      let inQuote = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          inQuote = !inQuote;
          continue;
        }
        if (!inQuote && ch === sep) {
          cols.push(cur);
          cur = '';
          continue;
        }
        cur += ch;
      }
      cols.push(cur);
      return cols.map(c => c.trim());
    }
    const headers = parseLine(lines[0]);
    let createdRes = 0, updatedRes = 0, createdAct = 0, updatedAct = 0, errors = 0;

    // Normaliza datas no formato ISO (YYYY-MM-DD). Aceita entradas no formato
    // brasileiro (dd/mm/yyyy) ou ISO (yyyy-mm-dd) e retorna sempre yyyy-mm-dd.
    // Para outros formatos ou entradas vazias, retorna a string original.
    function normalizeDateStr(d) {
      if (!d) return '';
      // remove aspas e espaços
      const s = String(d).trim().replace(/\"/g, '');
      // detecta separador ( '/' ou '-' )
      const parts = s.split(/[\/\-]/);
      if (parts.length === 3) {
        // dd/mm/yyyy -> yyyy-mm-dd
        if (parts[0].length === 2 && parts[1].length === 2 && parts[2].length === 4) {
          return `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        }
        // yyyy-mm-dd ou yyyy/mm/dd -> yyyy-mm-dd
        if (parts[0].length === 4) {
          return `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
        }
      }
      return s;
    }
    for (let i = 1; i < lines.length; i++) {
      const rowLine = lines[i];
      if (!rowLine.trim()) continue;
      const cols = parseLine(rowLine);
      const row = {};
      headers.forEach((h, idx) => {
        row[h] = cols[idx] !== undefined ? cols[idx] : '';
      });
      let rectype = '';
      if (row.tipo) rectype = String(row.tipo).toLowerCase();
      if (!rectype && row.tabela) rectype = String(row.tabela).toLowerCase();
      if (rectype.startsWith('recurso') || rectype === 'recurso' || rectype === 'resource') {
        // Importar recurso
        const name = row.nome || row.nomeRecurso || row.recursoNome || '';
        if (!name) {
          errors++;
          continue;
        }
        let existing = (resources || []).find(r => normalizeName(r.nome) === normalizeName(name));
        if (existing) {
          let changed = false;
          if (row.tipoRecurso) {
            existing.tipo = row.tipoRecurso;
            changed = true;
          }
          if (row.senioridade) {
            existing.senioridade = row.senioridade;
            changed = true;
          }
          if (row.capacidade) {
            const cap = parseFloat(row.capacidade.replace(',', '.'));
            if (!isNaN(cap)) {
              existing.capacidade = cap;
              changed = true;
            }
          }
          if (row.cargaHorariaDiaria || row.horasDia) {
            const chd = parseFloat(String(row.cargaHorariaDiaria || row.horasDia).replace(',', '.'));
            if (!isNaN(chd)) {
              existing.cargaHorariaDiaria = chd;
              changed = true;
            }
          }
          if (row.ativo) {
            const val = String(row.ativo).toLowerCase();
            existing.ativo = ['s','sim','y','yes','1','true'].includes(val);
            changed = true;
          }
          if (row.inicioAtivo) {
            existing.inicioAtivo = normalizeDateStr(row.inicioAtivo);
            changed = true;
          }
          if (row.fimAtivo) {
            existing.fimAtivo = normalizeDateStr(row.fimAtivo);
            changed = true;
          }
          if (changed) {
            existing.version = (existing.version || 0) + 1;
            existing.updatedAt = Date.now();
            recordEvent('resource', 'update', existing.id, { ...existing });
            updatedRes++;
          }
        } else {
          const newId = uuid();
          const nr = {
            id: newId,
            nome: name,
            tipo: row.tipoRecurso || 'Interno',
            senioridade: row.senioridade || '',
            capacidade: row.capacidade ? parseFloat(String(row.capacidade).replace(',', '.')) || 0 : 0,
            ativo: row.ativo ? ['s','sim','y','yes','1','true'].includes(String(row.ativo).toLowerCase()) : true,
            inicioAtivo: normalizeDateStr(row.inicioAtivo || ''),
            fimAtivo: normalizeDateStr(row.fimAtivo || ''),
            version: 1,
            updatedAt: Date.now(),
            cargaHorariaDiaria: row.cargaHorariaDiaria ? parseFloat(String(row.cargaHorariaDiaria).replace(',', '.')) || 9 : 9,
            deletedAt: null
          };
          (resources = resources || []).push(nr);
          recordEvent('resource', 'create', nr.id, { ...nr });
          createdRes++;
        }
      } else if (rectype.startsWith('atividade') || rectype === 'atividade' || rectype === 'activity') {
        // Importar atividade
        const title = row.titulo || row.title || '';
        const rName = row.recursoNome || row.resourceName || row.nomeRecurso || row.nome || '';
        if (!title || !rName) {
          errors++;
          continue;
        }
        // Garante que o recurso exista
        let res = (resources || []).find(r => normalizeName(r.nome) === normalizeName(rName));
        let resCreated = false;
        if (!res) {
          // Cria novo recurso com valores mínimos
          const newResId = uuid();
          const nr2 = {
            id: newResId,
            nome: rName,
            tipo: 'Interno',
            senioridade: '',
            capacidade: 0,
            cargaHorariaDiaria: row.cargaHorariaDiaria ? parseFloat(String(row.cargaHorariaDiaria).replace(',', '.')) || 9 : 9,
            ativo: true,
            inicioAtivo: '',
            fimAtivo: '',
            version: 1,
            updatedAt: Date.now(),
            deletedAt: null
          };
          (resources = resources || []).push(nr2);
          recordEvent('resource', 'create', nr2.id, { ...nr2 });
          createdRes++;
          res = nr2;
          resCreated = true;
        }
        const start = normalizeDateStr(row.inicio || row.start || '');
        const end = normalizeDateStr(row.fim || row.end || '');
        // Determina o status da atividade. Se não informado, assume 'Planejada'.
        let statusRaw = row.status || row.situacao || '';
        statusRaw = String(statusRaw).trim();
        let status = 'Planejada';
        if (statusRaw) {
          // Procura um status igualando case e acentos. Usa a lista STATUS para validar.
          const foundStatus = STATUS.find(s => s.toLowerCase() === statusRaw.toLowerCase());
          if (foundStatus) {
            status = foundStatus;
          }
        }
        let alloc = 0;
        if (row.alocacao) {
          const num = parseFloat(String(row.alocacao).replace(',', '.'));
          if (!isNaN(num)) alloc = num;
        }
        const tagsStr = row.tags || '';
        const tagsArr = tagsStr ? tagsStr.split(/[,;]/).map(t => t.trim()).filter(Boolean) : [];
        const importedCode = String(row.codigoAtividade || row.codigo || row.atividadeCodigo || '').trim();
        const importedLinkedOriginId = String(row.linkedOriginId || '').trim();
        const importedLinkedOriginCode = String(row.linkedOriginCode || row.atividadeVinculadaA || '').trim();
        // Tenta achar atividade existente por chave composta
        let existingAct = (activities || []).find(a => normalizeName(a.titulo) === normalizeName(title) && a.resourceId === res.id && a.inicio === start && a.fim === end);
        if (existingAct) {
          let changed = false;
          if (status) {
            existingAct.status = status;
            changed = true;
          }
          if (alloc) {
            existingAct.alocacao = alloc;
            changed = true;
          }
          if (tagsArr && tagsArr.length > 0) {
            existingAct.tags = tagsArr;
            changed = true;
          }
          if (importedCode && existingAct.codigoAtividade !== importedCode) {
            existingAct.codigoAtividade = importedCode;
            changed = true;
          }
          if (importedLinkedOriginId && existingAct.linkedOriginId !== importedLinkedOriginId) {
            existingAct.linkedOriginId = importedLinkedOriginId;
            changed = true;
          }
          if (!importedLinkedOriginId && importedLinkedOriginCode) {
            const linkedByCode = (activities || []).find(a => !a.deletedAt && String(a.codigoAtividade || '').trim().toLowerCase() === importedLinkedOriginCode.toLowerCase());
            if (linkedByCode && existingAct.linkedOriginId !== linkedByCode.id) {
              existingAct.linkedOriginId = linkedByCode.id;
              changed = true;
            }
          }
          if (changed) {
            existingAct.version = (existingAct.version || 0) + 1;
            existingAct.updatedAt = Date.now();
            recordEvent('activity', 'update', existingAct.id, { ...existingAct });
            updatedAct++;
          }
        } else {
          const newActId = uuid();
          const na = {
            id: newActId,
            codigoAtividade: importedCode || generateNextActivityCode(activities),
            titulo: title,
            resourceId: res.id,
            linkedOriginId: importedLinkedOriginId || '',
            _pendingLinkedOriginCode: (!importedLinkedOriginId && importedLinkedOriginCode) ? importedLinkedOriginCode : '',
            inicio: start,
            fim: end,
            status: status,
            alocacao: alloc,
            tags: tagsArr,
            version: 1,
            updatedAt: Date.now(),
            deletedAt: null
          };
          (activities = activities || []).push(na);
          recordEvent('activity', 'create', na.id, { ...na });
          createdAct++;
        }
      } else {
        // Tipo não reconhecido
        errors++;
      }
    }
    // Resolve vínculos importados por código após a carga completa, preservando compatibilidade sem perder dados
    (activities || []).forEach(a => {
      if (!a || a.deletedAt || a.linkedOriginId || !a._pendingLinkedOriginCode) return;
      const linked = (activities || []).find(x => !x.deletedAt && String(x.codigoAtividade || '').trim().toLowerCase() === String(a._pendingLinkedOriginCode).trim().toLowerCase());
      if (linked && linked.id !== a.id) a.linkedOriginId = linked.id;
      delete a._pendingLinkedOriginCode;
    });
    // Persiste e atualiza UI
    saveLS(LS.res, resources);
    saveLS(LS.act, activities);
    renderAll();
    // Salva no BD, se houver handle definido
    saveBDDebounced();
    alert(`Importação concluída:\n${createdRes} recursos criados, ${updatedRes} atualizados;\n${createdAct} atividades criadas, ${updatedAct} atualizadas;\n${errors} linhas ignoradas ou inválidas.`);
  } catch (err) {
    console.error('Erro ao processar importação CSV', err);
    alert('Erro ao processar importação: ' + (err && err.message ? err.message : err));
  }
}

function refreshActivityCommentsPanel(activity, resetOffset=true){
  if(!atividadeComentariosBox || !atividadeComentariosLista) return;
  currentCommentsActivityId = activity?.id || null;
  if(resetOffset) currentCommentsOffset = 0;
  const result = currentCommentsActivityId ? getComments(currentCommentsActivityId, { offset:0, limit: currentCommentsOffset + COMMENT_PAGE_SIZE }) : { total:0, items:[] };
  atividadeComentariosLista.innerHTML = renderActivityCommentsHtml((result.items || []).map(c => ({ ts:c.ts, user:c.usuario, text:c.texto })));
  atividadeComentariosBox.style.display = currentCommentsActivityId ? '' : 'none';
  const header = atividadeComentariosBox.querySelector('strong');
  if(header) header.textContent = `Comentários publicados (${result.total})`;
  if(atividadeComentariosVerMaisBtn){
    atividadeComentariosVerMaisBtn.style.display = result.total > result.items.length ? '' : 'none';
    atividadeComentariosVerMaisBtn.onclick = ()=>{ currentCommentsOffset += COMMENT_PAGE_SIZE; refreshActivityCommentsPanel(activity, false); };
  }
}

/* ===== v1.2.8.26 — Indicador anual de capacidade agregada do time ===== */
(function(){
  function $team(id){ return document.getElementById(id); }
  function teamPad(n){ return String(n).padStart(2,'0'); }
  function teamMonthName(idx){
    const names = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    return names[idx] || String(idx+1);
  }
  function teamDaysInMonth(year, monthIndex){
    const days = [];
    const start = new Date(year, monthIndex, 1);
    const end = new Date(year, monthIndex + 1, 0);
    for(let d = new Date(start); d <= end; d = addDays(d, 1)) days.push(new Date(d));
    return days;
  }
  function teamIsActivityOnDay(a, ymd){
    if(!a || a.deletedAt) return false;
    const s = normalizeDateField(a.inicio || a.start || '');
    const e = normalizeDateField(a.fim || a.end || '');
    return !!(s && e && s <= ymd && ymd <= e);
  }
  function teamIsVacationActivity(a){
    if(!a) return false;
    const values = [
      a.tipoAtividade,
      a.tipo_atividade,
      a.activityType,
      a.tipo,
      a.categoria
    ];
    return values.some(v => normalizeName(v) === 'ferias');
  }
  function getTeamCapSelectedIds(){
    return Array.from(document.querySelectorAll('#teamCapResources input[type="checkbox"]:checked')).map(x=>x.value);
  }
  function renderTeamCapResourceSelector(){
    const wrap = $team('teamCapResources');
    if(!wrap) return;
    const active = (resources || [])
      .filter(r=>r && !r.deletedAt && r.ativo)
      .slice()
      .sort((a,b)=>String(a.nome||'').localeCompare(String(b.nome||''), undefined, {sensitivity:'base'}));
    if(!active.length){
      wrap.innerHTML = '<div class="muted">Nenhum recurso ativo cadastrado.</div>';
      return;
    }
    wrap.innerHTML = active.map(r=>`
      <label class="team-cap-check">
        <input type="checkbox" value="${escAttr(r.id)}" checked />
        <span>${escHTML(r.nome || '(sem nome)')}</span>
        <small>${escHTML(r.tipo || '')}${r.senioridade ? ' • ' + escHTML(r.senioridade) : ''}</small>
      </label>
    `).join('');
  }
  function calculateTeamAnnualCapacity(year, selectedIds){
    const idSet = new Set(selectedIds || []);
    const selectedResources = (resources || []).filter(r=>r && !r.deletedAt && r.ativo && idSet.has(r.id));
    const activeActs = (activities || []).filter(a=>a && !a.deletedAt && idSet.has(a.resourceId) && !teamIsVacationActivity(a));
    const rows = [];
    for(let m=0; m<12; m++){
      let capacityHours = 0;
      let allocatedHours = 0;
      let overtimeHours = 0;
      const days = teamDaysInMonth(year, m);
      selectedResources.forEach(r=>{
        days.forEach(d=>{
          const ymd = toYMD(d);
          const cap = Number(getDailyCapacityHours(r, ymd) || 0);
          capacityHours += cap;
          overtimeHours += Number(getApprovedOvertimeHours(r, ymd) || 0);
          activeActs.forEach(a=>{
            if(a.resourceId !== r.id) return;
            if(!teamIsActivityOnDay(a, ymd)) return;
            const ctx = getDayContext(ymd, r);
            if(ctx && ctx.nonWorking && Number(ctx.overtimeApproved || 0) <= 0) return;
            allocatedHours += Number(getActivityAllocatedHours(a, r) || 0);
          });
        });
      });
      const usagePct = capacityHours > 0 ? (allocatedHours / capacityHours) * 100 : (allocatedHours > 0 ? 999 : 0);
      rows.push({month:m, label:teamMonthName(m), capacityHours, allocatedHours, overtimeHours, usagePct});
    }
    return rows;
  }
  function drawTeamAnnualChart(rows, metric){
    const canvas = $team('teamCapCanvas');
    const summary = $team('teamCapSummary');
    const title = $team('teamCapTitle');
    if(!canvas) return;
    const ctx = canvas.getContext('2d');
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0,0,W,H);
    const marginL=54, marginR=24, marginT=26, marginB=54;
    const chartW = W - marginL - marginR;
    const chartH = H - marginT - marginB;
    const axisY = H - marginB;
    const maxHours = Math.max(1, ...rows.map(r=>Math.max(r.capacityHours, r.allocatedHours)));
    const maxPct = Math.max(100, ...rows.map(r=>Math.min(160, r.usagePct || 0)));
    const maxVal = metric === 'percent' ? maxPct : maxHours;

    ctx.strokeStyle = '#94a3b8';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(marginL, marginT);
    ctx.lineTo(marginL, axisY);
    ctx.lineTo(W-marginR, axisY);
    ctx.stroke();

    ctx.fillStyle = '#64748b';
    ctx.font = '12px sans-serif';
    ctx.fillText(metric === 'percent' ? '% uso' : 'Horas', 8, marginT + 4);

    const groupW = chartW / 12;
    const barW = Math.max(10, Math.min(24, groupW * 0.28));

    rows.forEach((r, i)=>{
      const cx = marginL + i*groupW + groupW/2;
      const capVal = metric === 'percent' ? 100 : r.capacityHours;
      const allocVal = metric === 'percent' ? Math.min(160, r.usagePct || 0) : r.allocatedHours;
      const capH = Math.max(0, (capVal / maxVal) * chartH);
      const allocH = Math.max(0, (allocVal / maxVal) * chartH);

      ctx.fillStyle = '#93c5fd';
      ctx.fillRect(cx - barW - 2, axisY - capH, barW, capH);

      ctx.fillStyle = allocVal > capVal ? '#ef4444' : '#2563eb';
      ctx.fillRect(cx + 2, axisY - allocH, barW, allocH);

      ctx.fillStyle = '#0f172a';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(r.label, cx, axisY + 20);

      ctx.font = '10px sans-serif';
      const label = metric === 'percent' ? `${Math.round(r.usagePct || 0)}%` : `${Math.round(r.allocatedHours)}/${Math.round(r.capacityHours)}h`;
      ctx.fillText(label, cx, Math.max(marginT + 10, axisY - Math.max(capH, allocH) - 6));
    });

    ctx.textAlign='left';
    ctx.font='12px sans-serif';
    ctx.fillStyle='#93c5fd';
    ctx.fillRect(marginL, 8, 12, 12);
    ctx.fillStyle='#0f172a';
    ctx.fillText('Capacidade', marginL + 18, 18);
    ctx.fillStyle='#2563eb';
    ctx.fillRect(marginL + 112, 8, 12, 12);
    ctx.fillStyle='#0f172a';
    ctx.fillText('Consolidado', marginL + 130, 18);
    ctx.fillStyle='#ef4444';
    ctx.fillRect(marginL + 235, 8, 12, 12);
    ctx.fillStyle='#0f172a';
    ctx.fillText('Acima da capacidade', marginL + 253, 18);

    const totalCap = rows.reduce((a,r)=>a+r.capacityHours,0);
    const totalAlloc = rows.reduce((a,r)=>a+r.allocatedHours,0);
    const totalHE = rows.reduce((a,r)=>a+r.overtimeHours,0);
    const pct = totalCap > 0 ? (totalAlloc / totalCap) * 100 : 0;
    if(title) title.textContent = `Capacidade x Consolidado — ${$team('teamCapYear')?.value || ''}`;
    if(summary){
      summary.innerHTML = `
        <strong>Capacidade anual:</strong> ${totalCap.toFixed(1)}h ·
        <strong>Consolidado:</strong> ${totalAlloc.toFixed(1)}h ·
        <strong>Uso:</strong> ${pct.toFixed(1)}% ·
        <strong>HE aprovada incluída:</strong> ${totalHE.toFixed(1)}h
      `;
    }
  }
  function renderTeamAnnualCapacity(){
    const yearEl = $team('teamCapYear');
    const metricEl = $team('teamCapMetric');
    if(!yearEl) return;
    const year = Number(yearEl.value || new Date().getFullYear());
    const selectedIds = getTeamCapSelectedIds();
    if(!selectedIds.length){
      const summary = $team('teamCapSummary');
      if(summary) summary.textContent = 'Selecione ao menos um recurso para gerar o indicador.';
      const canvas = $team('teamCapCanvas');
      if(canvas) canvas.getContext('2d').clearRect(0,0,canvas.width,canvas.height);
      return;
    }
    const rows = calculateTeamAnnualCapacity(year, selectedIds);
    drawTeamAnnualChart(rows, metricEl ? metricEl.value : 'hours');
  }
  function bindTeamAnnualCapacityUI(){
    const panel = $team('teamCapacityAnnualPanel');
    if(!panel || panel.__boundTeamCap) return;
    panel.__boundTeamCap = true;
    const yearEl = $team('teamCapYear');
    if(yearEl && !yearEl.value) yearEl.value = String(new Date().getFullYear());
    renderTeamCapResourceSelector();
    const allBtn = $team('btnTeamCapAll');
    const noneBtn = $team('btnTeamCapNone');
    const renderBtn = $team('btnTeamCapRender');
    const metricEl = $team('teamCapMetric');
    const resourcesWrap = $team('teamCapResources');
    if(allBtn) allBtn.onclick = ()=>{ document.querySelectorAll('#teamCapResources input[type="checkbox"]').forEach(x=>x.checked=true); renderTeamAnnualCapacity(); };
    if(noneBtn) noneBtn.onclick = ()=>{ document.querySelectorAll('#teamCapResources input[type="checkbox"]').forEach(x=>x.checked=false); renderTeamAnnualCapacity(); };
    if(renderBtn) renderBtn.onclick = renderTeamAnnualCapacity;
    if(metricEl) metricEl.onchange = renderTeamAnnualCapacity;
    if(yearEl) yearEl.onchange = renderTeamAnnualCapacity;
    if(resourcesWrap) resourcesWrap.addEventListener('change', renderTeamAnnualCapacity);
    renderTeamAnnualCapacity();
  }
  const _renderAllTeamCap = window.renderAll;
  // renderAll é função global por declaração; aqui apenas agenda atualização sem substituir.
  document.addEventListener('DOMContentLoaded', ()=> setTimeout(bindTeamAnnualCapacityUI, 400));
  window.renderTeamAnnualCapacity = renderTeamAnnualCapacity;
  window.bindTeamAnnualCapacityUI = bindTeamAnnualCapacityUI;
  setTimeout(bindTeamAnnualCapacityUI, 800);
})();

// ===== Dashboard de Execução por demanda (v1.2.8.46) =====
const EXEC_DASH_PAGE_SIZE = 12;
let execDashPage = 1;
function execDashNum(v, digits=1){
  const n = Number(v) || 0;
  return n.toLocaleString('pt-BR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
function execDashPct(v){ return Math.max(0, Math.min(100, Number(v)||0)); }
function getExecutionDashboardRows(){
  const activeResources = getActiveResources();
  const resById = Object.fromEntries(activeResources.map(r => [r.id, r]));
  return getActiveActivities().map(a => {
    const data = normalizeExecData(a);
    const m = calcExecutionMetrics(a);
    const subTypes = Array.from(new Set(data.subtasks.map(st => st.tipoSubatividade || 'Outros').filter(Boolean)));
    const lastEntry = data.entries.slice().sort((x,y)=>(y.data||'').localeCompare(x.data||'') || (Number(y.createdAt)||0)-(Number(x.createdAt)||0))[0] || null;
    let situation = 'ok';
    if(!data.subtasks.length && !data.entries.length && !data.issues.length) situation = 'semexec';
    else if((m.atrasoDias||0) > 0) situation = 'atraso';
    else if((m.horasRestantes||0) > 0 && (m.percentualReal||0) < 70) situation = 'acompanhamento';
    const janelaTotal = Math.max(1, diffDays(fromYMD(a.fim), fromYMD(a.inicio)) + 1);
    const janelaPassada = Math.max(0, Math.min(janelaTotal, diffDays(new Date(), fromYMD(a.inicio)) + 1));
    const janelaPct = execDashPct((janelaPassada / janelaTotal) * 100);
    return {
      activity:a,
      data,
      metrics:m,
      resource:resById[a.resourceId] || null,
      subTypes,
      lastEntry,
      situation,
      janelaPct,
      execPct: execDashPct(m.percentualReal || 0)
    };
  });
}
function populateExecutionDashboardFilters(){
  const st = document.getElementById('execDashStatus');
  if(st && st.dataset.ready !== '1'){
    STATUS.forEach(s => { const o=document.createElement('option'); o.value=s; o.textContent=s; st.appendChild(o); });
    st.dataset.ready='1';
  }
  const rs = document.getElementById('execDashResource');
  if(rs){
    const current = rs.value;
    rs.innerHTML = '<option value="">Todos os recursos</option>';
    getActiveResources().slice().sort((a,b)=>(a.nome||'').localeCompare(b.nome||'')).forEach(r=>{ const o=document.createElement('option'); o.value=r.id; o.textContent=r.nome||r.id; rs.appendChild(o); });
    rs.value = current;
  }
  const tp = document.getElementById('execDashSubType');
  if(tp){
    const current = tp.value;
    const types = new Set(EXEC_SUBTASK_TYPES.filter(x=>x !== 'Outros'));
    getActiveActivities().forEach(a => normalizeExecData(a).subtasks.forEach(st => types.add(st.tipoSubatividade || 'Outros')));
    tp.innerHTML = '<option value="">Todos os tipos</option>';
    Array.from(types).sort((a,b)=>a.localeCompare(b)).forEach(t=>{ const o=document.createElement('option'); o.value=t; o.textContent=t; tp.appendChild(o); });
    tp.value = current;
  }
}
function filterExecutionDashboardRows(rows){
  const q = String(document.getElementById('execDashSearch')?.value || '').trim().toLowerCase();
  const status = document.getElementById('execDashStatus')?.value || '';
  const resourceId = document.getElementById('execDashResource')?.value || '';
  const subType = document.getElementById('execDashSubType')?.value || '';
  const risk = document.getElementById('execDashRisk')?.value || '';
  return rows.filter(r => {
    const a = r.activity;
    if(q){
      const hay = [a.titulo, a.codigoAtividade, (a.tags||[]).join(' '), r.resource?.nome, r.subTypes.join(' ')].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    if(status && a.status !== status) return false;
    if(resourceId && a.resourceId !== resourceId) return false;
    if(subType && !r.subTypes.includes(subType)) return false;
    if(risk && r.situation !== risk) return false;
    return true;
  });
}
function renderExecDashKpis(rows){
  const box = document.getElementById('execDashKpis');
  if(!box) return;
  const totals = rows.reduce((acc,r)=>{
    acc.atividades += 1;
    acc.planejadas += Number(r.metrics.horasPlanejadas)||0;
    acc.realizadas += Number(r.metrics.horasRealizadas)||0;
    acc.saldo += Number(r.metrics.horasRestantes)||0;
    acc.impacto += Number(r.metrics.impactoHoras)||0;
    acc.subtasks += r.data.subtasks.length;
    acc.issues += r.data.issues.length;
    if(r.situation === 'atraso') acc.atrasos += 1;
    if(r.situation === 'acompanhamento') acc.atencao += 1;
    if(r.situation === 'semexec') acc.semexec += 1;
    return acc;
  }, {atividades:0, planejadas:0, realizadas:0, saldo:0, impacto:0, subtasks:0, issues:0, atrasos:0, atencao:0, semexec:0});
  const ader = totals.planejadas > 0 ? (totals.realizadas / totals.planejadas) * 100 : 0;
  const items = [
    ['Demandas', totals.atividades, 'no filtro atual'],
    ['Horas planejadas', `${execDashNum(totals.planejadas)}h`, 'base da execução'],
    ['Horas realizadas', `${execDashNum(totals.realizadas)}h`, 'apontamentos'],
    ['Saldo restante', `${execDashNum(totals.saldo)}h`, 'estimado'],
    ['Aderência geral', `${execDashNum(ader,0)}%`, 'realizado ÷ planejado'],
    ['Em atraso', totals.atrasos, 'previsão > prazo'],
    ['Em atenção', totals.atencao, 'baixa execução'],
    ['Impacto ocorrências', `${execDashNum(totals.impacto)}h`, `${totals.issues} ocorrência(s)`]
  ];
  box.innerHTML = items.map(([label,value,hint]) => `<div class="execdash-kpi"><span class="label">${escHTML(label)}</span><span class="value">${escHTML(String(value))}</span><span class="hint">${escHTML(hint)}</span></div>`).join('');
}
function execDashSituationLabel(s){
  if(s === 'atraso') return ['Prazo vencido / atraso previsto','danger'];
  if(s === 'acompanhamento') return ['Em acompanhamento','warn'];
  if(s === 'semexec') return ['Sem execução registrada','empty'];
  return ['Dentro do previsto','ok'];
}
function ensureExecDashPaginationEl(){
  let pg = document.getElementById('execDashPagination');
  const cards = document.getElementById('execDashCards');
  if(!pg && cards){
    pg = document.createElement('div');
    pg.id = 'execDashPagination';
    pg.className = 'execdash-pagination';
    cards.insertAdjacentElement('afterend', pg);
  }
  return pg;
}
function renderExecDashPagination(totalRows, totalPages){
  const pg = ensureExecDashPaginationEl();
  if(!pg) return;
  if(totalRows <= EXEC_DASH_PAGE_SIZE){ pg.innerHTML = ''; return; }
  pg.innerHTML = `<button type="button" class="btn" id="execDashPrev" ${execDashPage<=1?'disabled':''}>Anterior</button>
    <span class="muted small">Página <strong>${execDashPage}</strong> de <strong>${totalPages}</strong> • ${totalRows} demanda(s)</span>
    <button type="button" class="btn" id="execDashNext" ${execDashPage>=totalPages?'disabled':''}>Próxima</button>`;
  const prev = document.getElementById('execDashPrev');
  const next = document.getElementById('execDashNext');
  if(prev) prev.onclick = () => { execDashPage = Math.max(1, execDashPage - 1); renderExecutionDashboard(false); };
  if(next) next.onclick = () => { execDashPage = Math.min(totalPages, execDashPage + 1); renderExecutionDashboard(false); };
}
function renderExecDashCards(rows){
  const box = document.getElementById('execDashCards');
  const count = document.getElementById('execDashCount');
  const totalRows = rows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / EXEC_DASH_PAGE_SIZE));
  execDashPage = Math.max(1, Math.min(execDashPage || 1, totalPages));
  const startIdx = (execDashPage - 1) * EXEC_DASH_PAGE_SIZE;
  const pageRows = rows.slice(startIdx, startIdx + EXEC_DASH_PAGE_SIZE);
  if(count) count.textContent = totalRows ? ` ${totalRows} demanda(s) filtrada(s) • exibindo ${startIdx + 1}-${Math.min(startIdx + pageRows.length, totalRows)} de ${totalRows}` : ' 0 demanda(s)';
  if(!box) return;
  if(!rows.length){
    box.innerHTML = '<div class="execdash-empty">Nenhuma demanda encontrada para os filtros atuais.</div>';
    renderExecDashPagination(0, 1);
    return;
  }
  box.innerHTML = pageRows.map(r => {
    const a=r.activity, m=r.metrics, d=r.data;
    const [lab,cls] = execDashSituationLabel(r.situation);
    const subRows = d.subtasks.length ? `<table class="execdash-detail-table"><thead><tr><th>Tipo</th><th>Subatividade</th><th>Plan.</th><th>Real.</th></tr></thead><tbody>${d.subtasks.map(st=>{
      const done = d.entries.filter(en=>en.subtaskId===st.id).reduce((acc,en)=>acc+(Number(en.horas)||0),0);
      return `<tr><td>${escHTML(st.tipoSubatividade||'Outros')}</td><td>${escHTML(st.titulo||'')}</td><td>${execDashNum(st.horasPlanejadas||0)}h</td><td>${execDashNum(done)}h</td></tr>`;
    }).join('')}</tbody></table>` : '<div class="muted small" style="margin-top:8px;">Sem subatividades cadastradas.</div>';
    const issues = d.issues.length ? `<table class="execdash-detail-table"><thead><tr><th>Ocorrência</th><th>Impacto</th></tr></thead><tbody>${d.issues.slice(0,5).map(is=>`<tr><td>${escHTML(is.tipo||'Ocorrência')}<br><span class="muted">${escHTML(is.descricao||'')}</span></td><td>${execDashNum(is.impactoHoras||0)}h<br>${Number(is.impactoPrazoDias)||0}d</td></tr>`).join('')}</tbody></table>` : '<div class="muted small" style="margin-top:8px;">Sem ocorrências registradas.</div>';
    const title = `${a.codigoAtividade ? a.codigoAtividade + ' - ' : ''}${a.titulo || ''}`;
    return `<article class="execdash-card">
      <div><h3>${escHTML(title)}</h3><div class="meta">Janela: ${escHTML(a.inicio||'—')} até ${escHTML(a.fim||'—')} • ${escHTML(r.resource?.nome || 'Sem recurso')}</div></div>
      <span class="execdash-status ${cls}">${escHTML(lab)}</span>
      <div class="execdash-card-grid">
        <div class="execdash-mini"><span>Horas planejadas</span><strong>${execDashNum(m.horasPlanejadas||0)}h</strong></div>
        <div class="execdash-mini"><span>Horas realizadas</span><strong>${execDashNum(m.horasRealizadas||0)}h</strong></div>
        <div class="execdash-mini good"><span>Saldo</span><strong>${execDashNum(m.horasRestantes||0)}h</strong></div>
        <div class="execdash-mini"><span>Execução real</span><strong>${execDashNum(m.percentualReal||0,0)}%</strong></div>
        <div class="execdash-mini ${m.atrasoDias>0?'bad':''}"><span>Término planejado</span><strong>${escHTML(m.terminoPlanejado||a.fim||'—')}</strong></div>
        <div class="execdash-mini ${m.atrasoDias>0?'bad':''}"><span>Previsão atual</span><strong>${escHTML(m.previsaoFim||'—')}</strong></div>
      </div>
      <div class="execdash-bars">
        <div class="execdash-bar-label"><span>Avanço da janela</span><span>${execDashNum(r.janelaPct,0)}%</span></div><div class="execdash-bar"><span style="width:${r.janelaPct}%"></span></div>
        <div class="execdash-bar-label"><span>Execução real apontada</span><span>${execDashNum(r.execPct,0)}%</span></div><div class="execdash-bar real"><span style="width:${r.execPct}%"></span></div>
      </div>
      <details class="execdash-details"><summary>Detalhar execução</summary>${subRows}<h4 style="margin:12px 0 0 0;">Ocorrências</h4>${issues}</details>
    </article>`;
  }).join('');
  renderExecDashPagination(totalRows, totalPages);
}
function renderExecutionDashboard(resetPage = false){
  if(resetPage) execDashPage = 1;
  const box = document.getElementById('execDashCards');
  if(!box) return;
  populateExecutionDashboardFilters();
  const rows = filterExecutionDashboardRows(getExecutionDashboardRows());
  renderExecDashKpis(rows);
  renderExecDashCards(rows);
}
function bindExecutionDashboard(){
  const ids = ['execDashSearch','execDashStatus','execDashResource','execDashSubType','execDashRisk'];
  ids.forEach(id=>{
    const el = document.getElementById(id);
    if(el && el.dataset.execdashBound !== '1'){
      el.dataset.execdashBound='1';
      el.addEventListener(id==='execDashSearch'?'input':'change', () => renderExecutionDashboard(true));
    }
  });
  const btn = document.getElementById('execDashRefresh');
  if(btn && btn.dataset.execdashBound !== '1'){
    btn.dataset.execdashBound='1';
    btn.addEventListener('click', async ()=>{ await refreshFromBDIfNeeded(); renderExecutionDashboard(true); });
  }
}
window.__execDashReady = true;
if(document.readyState === 'loading'){
  document.addEventListener('DOMContentLoaded', ()=>{ bindExecutionDashboard(); renderExecutionDashboard(true); });
}else{
  bindExecutionDashboard(); renderExecutionDashboard(true);
}
