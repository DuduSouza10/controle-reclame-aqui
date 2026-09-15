const state = {
  lang: localStorage.getItem('raLang') || 'pt',
  regional: 'TODOS', rm: '', start: '', end: '',
  version: null, filters: null, page: 1, pages: 1,
  editorDirty: false, activeTab: 'overview', pollBusy: false,
};

const I18N = {
  pt: {
    title:'Painel Reclame Aqui', subtitle:'Gestão de reclamações, tratativas, SLA e performance por RM', sync:'Sincronizado',
    tabOverview:'Visão Geral', tabRanking:'Ranking por RM', tabEditor:'Editar Base', regional:'Regional', all:'Todos', rm:'RM', allRms:'Todos os RMs', startDate:'Data inicial', endDate:'Data final', clear:'Limpar filtros',
    totalComplaints:'Total de reclamações', currentFilter:'no filtro atual', openOrders:'Pedidos em aberto', openRule:'Solução vazia ou “EM ABERTO”', ownBases:'Bases próprias', ownRule:'Base sem prefixo “F ”', franchises:'Franquias', franchiseRule:'Base iniciada por “F ”', within24:'Dentro de 24h', autoSla:'SLA automático', over24:'Fora de 24h', complaintDateRule:'pela data da reclamação',
    openByRm:'Pedidos em aberto por RM', openByRmSub:'Separação entre base própria, franquia e SLA de 24h', open:'Aberto', own:'Própria', franchise:'Franquia', within:'Dentro', outside:'Fora', reasons:'Motivos', reasonsSub:'Distribuição das reclamações no filtro atual',
    baseComposition:'Composição das bases', baseCompositionSub:'Classificação automática pelo nome da base', slaRuleTitle:'Regra de SLA', slaRuleSub:'O dashboard calcula automaticamente o prazo', withinDeadline:'Dentro do prazo', outsideDeadline:'Fora do prazo', manualDeadlineNote:'A coluna “DENTRO DO PRAZO” continua editável manualmente na aba Editar Base, mas o ranking usa o SLA automático de 24h.',
    recentRecords:'Registros mais recentes', recentRecordsSub:'Últimas reclamações conforme os filtros selecionados', date:'Data', base:'Base', reason:'Motivo', solution:'Solução',
    rankingTitle:'Ranking por RM × Motivo', rankingSub:'Cada motivo vira uma coluna. O ranking também mostra composição e SLA.', onlyOpen:'Somente pedidos em aberto', total:'Total',
    editorTitle:'Editar base de dados', editorSub:'Edite, salve, importe ou exporte. Alterações salvas aparecem para todos os usuários.', importXlsx:'Importar XLSX', exportXlsx:'Exportar XLSX', addRow:'+ Novo registro', allRegions:'Todas as regionais', allBaseTypes:'Todos os tipos', ownBase:'Base própria', searchPlaceholder:'Buscar ID, rastreio, base, RM, motivo...',
    complaintDate:'Data reclamação', owner:'Responsável', tracking:'Rastreio', driver:'Motorista', baseType:'Tipo', directedDate:'Data direcionamento', manualDeadline:'Prazo manual', treatment:'Tratativa', observation:'Observação', actions:'Ações', previous:'Anterior', next:'Próxima', newRecord:'Novo registro', cancel:'Cancelar', save:'Salvar', delete:'Excluir',
    noData:'Nenhum registro encontrado.', semRm:'SEM RM', rows:'registros', page:'Página', of:'de', updatedOther:'Os dados foram atualizados por outro usuário.', updatedOtherDirty:'Há uma atualização de outro usuário. Salve suas edições ou recarregue a base.', saved:'Registro salvo e sincronizado.', created:'Registro criado e sincronizado.', deleted:'Registro excluído e sincronizado.', confirmDelete:'Excluir este registro?', confirmImport:'Importar este XLSX substituirá a base atual. Tratativas/prazo manual são preservados por ID RA quando possível. Continuar?', importing:'Importando planilha...', imported:'Planilha importada e sincronizada.', error:'Ocorreu um erro.',
    auto:'Automático', noDate:'Sem data', inTime:'Dentro do prazo', outTime:'Fora do prazo', status:'Status', loading:'Carregando...',
  },
  zh: {
    title:'Reclame Aqui 投诉看板', subtitle:'投诉、处理方案、SLA 与 RM 绩效管理', sync:'已同步',
    tabOverview:'总览', tabRanking:'RM 排名', tabEditor:'编辑数据', regional:'大区', all:'全部', rm:'RM', allRms:'全部 RM', startDate:'开始日期', endDate:'结束日期', clear:'清除筛选',
    totalComplaints:'投诉总数', currentFilter:'当前筛选范围', openOrders:'待处理订单', openRule:'解决方案为空或“处理中”', ownBases:'直营网点', ownRule:'网点名称不以“F ”开头', franchises:'加盟网点', franchiseRule:'网点名称以“F ”开头', within24:'24小时内', autoSla:'自动 SLA', over24:'超过24小时', complaintDateRule:'按投诉日期计算',
    openByRm:'各 RM 待处理订单', openByRmSub:'区分直营网点、加盟网点与24小时 SLA', open:'待处理', own:'直营', franchise:'加盟', within:'时限内', outside:'超时', reasons:'投诉原因', reasonsSub:'当前筛选范围内的投诉原因分布',
    baseComposition:'网点构成', baseCompositionSub:'根据网点名称自动分类', slaRuleTitle:'SLA 规则', slaRuleSub:'系统自动计算时效', withinDeadline:'时限内', outsideDeadline:'超时', manualDeadlineNote:'“是否在时限内”字段仍可在“编辑数据”页手工填写，但 RM 排名使用自动24小时 SLA。',
    recentRecords:'最新记录', recentRecordsSub:'当前筛选条件下最新的投诉记录', date:'日期', base:'网点', reason:'原因', solution:'解决方案',
    rankingTitle:'RM × 原因排名', rankingSub:'每个投诉原因作为一列，同时显示网点构成与 SLA。', onlyOpen:'仅显示待处理订单', total:'总计',
    editorTitle:'编辑数据库', editorSub:'可编辑、保存、导入或导出。保存后的修改会同步给所有用户。', importXlsx:'导入 XLSX', exportXlsx:'导出 XLSX', addRow:'+ 新增记录', allRegions:'全部大区', allBaseTypes:'全部类型', ownBase:'直营网点', searchPlaceholder:'搜索 ID、运单号、网点、RM、原因…',
    complaintDate:'投诉日期', owner:'负责人', tracking:'运单号', driver:'司机', baseType:'类型', directedDate:'转交日期', manualDeadline:'手工时效', treatment:'处理方案', observation:'备注', actions:'操作', previous:'上一页', next:'下一页', newRecord:'新增记录', cancel:'取消', save:'保存', delete:'删除',
    noData:'未找到记录。', semRm:'未分配 RM', rows:'条记录', page:'第', of:'页 / 共', updatedOther:'其他用户已更新数据。', updatedOtherDirty:'其他用户有新更新。请先保存当前编辑或重新加载数据。', saved:'记录已保存并同步。', created:'记录已创建并同步。', deleted:'记录已删除并同步。', confirmDelete:'确定删除此记录吗？', confirmImport:'导入 XLSX 将替换当前数据。若 ID RA 匹配，将尽量保留处理方案和手工时效。是否继续？', importing:'正在导入表格…', imported:'表格已导入并同步。', error:'发生错误。',
    auto:'自动', noDate:'无日期', inTime:'时限内', outTime:'超时', status:'状态', loading:'加载中…',
  }
};

const VALUE_ZH = {
  'BASE PRÓPRIA':'直营网点', 'FRANQUIA':'加盟网点', 'EM ABERTO':'处理中', 'ENTREGA':'已配送', 'EXTRAVIO':'丢失', 'DEVOLUÇÃO':'退回', 'RESOLVIDO':'已解决', 'ERRO DE TRIAGEM':'分拣错误', 'OUTRO ESTADO':'其他州', 'AVARIA':'破损', 'PAGO':'已赔付', 'RETRATAÇÃO':'致歉处理', 'DENTRO DO PRAZO':'时限内', 'FORA DO PRAZO':'超时', 'FISCALIZAÇÃO TRIBUTÁRIA':'税务检查', 'DESATIVADO':'已停用', 'SEFAZ':'州财政部门',
  'Atraso na entrega':'配送延误', 'Acareação':'核查对质', 'Postura':'服务态度', 'Postura do Motorista':'司机服务态度', 'Trabalhista':'劳动事务', 'Ponto de Coleta':'揽收点', 'Informação':'信息问题', 'Avaria':'破损', 'Extravio/Avaria':'丢失/破损', 'Troca de etiqueta':'标签错换', 'Item faltante':'物品缺失', 'Em rota':'派送中', 'Sem motivo':'无原因', 'SEM RM':'未分配 RM', 'SEM DATA':'无日期',
  'Validar última movimentação, acionar a base responsável e registrar previsão de entrega.':'核查最后一条物流轨迹，联系责任网点并记录预计送达时间。',
  'Realizar acareação com base/motorista, registrar evidências e concluir a devolutiva ao cliente.':'与网点/司机进行核查对质，记录证据并向客户反馈处理结果。',
  'Apurar a conduta, identificar o responsável e registrar orientação/correção aplicada.':'核查服务行为，确认责任人并记录已采取的指导/纠正措施。',
  'Apurar a conduta do motorista, registrar evidências e aplicar orientação/correção.':'核查司机行为，记录证据并执行指导/纠正措施。',
  'Direcionar para a área responsável e registrar o protocolo da tratativa.':'转交责任部门并记录处理单号。',
  'Validar o ponto de coleta, responsável local e evidências do atendimento.':'核查揽收点、现场责任人及服务证据。',
  'Confirmar a informação correta, ajustar o registro e retornar ao cliente.':'确认正确信息，修正记录并回复客户。',
  'Validar evidências da avaria, responsabilidade operacional e fluxo de ressarcimento quando aplicável.':'核查破损证据、运营责任，并在适用时进入赔付流程。',
  'Realizar busca operacional, validar evidências e direcionar o fluxo de extravio/avaria.':'开展运营查找，核验证据并进入丢失/破损处理流程。',
  'Rastrear as etiquetas envolvidas, corrigir a vinculação e validar o destino correto.':'追踪相关标签，修正绑定关系并确认正确目的地。',
  'Conferir quantidade, peso, volumetria e evidências para localizar o item faltante.':'核对数量、重量、体积及证据，查找缺失物品。',
  'Validar rota, motorista e previsão de conclusão da entrega.':'核查路线、司机及预计完成配送时间。',
  'Analisar o motivo, acionar o responsável e registrar evidências e devolutiva da tratativa.':'分析原因，联系责任人，并记录证据与处理反馈。'
};

const $ = sel => document.querySelector(sel);
const $$ = sel => [...document.querySelectorAll(sel)];
const t = key => (I18N[state.lang] && I18N[state.lang][key]) || I18N.pt[key] || key;
const trValue = value => state.lang === 'zh' ? (VALUE_ZH[value] || value || '') : (value || '');
const fmt = n => new Intl.NumberFormat(state.lang === 'zh' ? 'zh-CN' : 'pt-BR').format(Number(n || 0));

function escapeHtml(value='') {
  return String(value).replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
}
function attr(value='') { return escapeHtml(value).replace(/`/g, '&#96;'); }

function setLanguage(lang, refresh=true) {
  state.lang = lang;
  localStorage.setItem('raLang', lang);
  document.documentElement.lang = lang === 'zh' ? 'zh-CN' : 'pt-BR';
  $('#langPt').classList.toggle('active', lang === 'pt');
  $('#langZh').classList.toggle('active', lang === 'zh');
  $$('[data-i18n]').forEach(el => { const key = el.dataset.i18n; if (t(key)) el.textContent = t(key); });
  $$('[data-i18n-placeholder]').forEach(el => { el.placeholder = t(el.dataset.i18nPlaceholder); });
  $$('[data-i18n-label]').forEach(el => {
    const key = el.dataset.i18nLabel;
    const first = [...el.childNodes].find(n => n.nodeType === Node.TEXT_NODE);
    if (first) first.nodeValue = t(key);
  });
  if (state.filters) populateFilterControls(state.filters, true);
  if (refresh) refreshCurrentViews(true);
}

async function api(url, options={}) {
  const res = await fetch(url, options);
  let data = {};
  try { data = await res.json(); } catch (_) {}
  if (!res.ok) throw new Error(data.error || `${res.status} ${res.statusText}`);
  return data;
}

function showToast(message, isError=false) {
  const el = $('#toast');
  el.textContent = message;
  el.classList.toggle('error', isError);
  el.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => el.classList.remove('show'), 3300);
}
function setSyncing(syncing) { $('#syncBadge').classList.toggle('syncing', syncing); }

function getQuery(extra={}) {
  const p = new URLSearchParams({regional: state.regional});
  if (state.rm) p.set('rm', state.rm);
  if (state.start) p.set('start', state.start);
  if (state.end) p.set('end', state.end);
  for (const [k,v] of Object.entries(extra)) if (v !== '' && v != null) p.set(k, v);
  return p.toString();
}

async function loadFilters() {
  state.filters = await api('/api/filters');
  populateFilterControls(state.filters);
  const reasonList = $('#reasonList');
  reasonList.innerHTML = state.filters.reasons.map(v => `<option value="${attr(v)}"></option>`).join('');
  const solutionList = $('#solutionList');
  solutionList.innerHTML = state.filters.solutions.map(v => `<option value="${attr(v)}"></option>`).join('');
}

function populateFilterControls(filters, preserve=true) {
  const rmFilter = $('#rmFilter');
  const rankingRmFilter = $('#rankingRmFilter');
  const editorRm = $('#editorRm');
  const current = preserve ? state.rm : '';
  const baseOptions = [`<option value="">${escapeHtml(t('allRms'))}</option>`];
  if (filters.has_blank_rm) baseOptions.push(`<option value="__SEM_RM__">${escapeHtml(t('semRm'))}</option>`);
  filters.rms.forEach(r => baseOptions.push(`<option value="${attr(r)}">${escapeHtml(r)}</option>`));
  rmFilter.innerHTML = baseOptions.join('');
  rankingRmFilter.innerHTML = baseOptions.join('');
  editorRm.innerHTML = baseOptions.join('');
  rmFilter.value = current;
  rankingRmFilter.value = current;
}

async function loadDashboard() {
  setSyncing(true);
  try {
    const data = await api('/api/dashboard?' + getQuery());
    state.version = data.version;
    $('#kpiTotal').textContent = fmt(data.kpis.total);
    $('#kpiOpen').textContent = fmt(data.kpis.open);
    $('#kpiOwn').textContent = fmt(data.kpis.own);
    $('#kpiFranchise').textContent = fmt(data.kpis.franchise);
    $('#kpiWithin').textContent = fmt(data.kpis.within);
    $('#kpiOverdue').textContent = fmt(data.kpis.overdue);
    renderOpenByRm(data.open_by_rm);
    renderBars('#reasonBars', data.reason_counts);
    renderBaseBars(data.base_counts, data.kpis.total);
    renderRecent(data.recent);
  } finally { setSyncing(false); }
}

function renderOpenByRm(rows) {
  const body = $('#openByRmBody');
  if (!rows.length) { body.innerHTML = `<tr><td colspan="6" class="empty-state">${escapeHtml(t('noData'))}</td></tr>`; return; }
  body.innerHTML = rows.map(r => `<tr>
    <td><strong>${escapeHtml(r.rm === 'SEM RM' ? t('semRm') : r.rm)}</strong></td>
    <td><span class="badge open">${fmt(r.total)}</span></td>
    <td>${fmt(r.own)}</td><td>${fmt(r.franchise)}</td>
    <td><span class="badge good">${fmt(r.within)}</span></td>
    <td><span class="badge bad">${fmt(r.overdue)}</span></td>
  </tr>`).join('');
}

function renderBars(selector, rows) {
  const el = $(selector);
  if (!rows.length) { el.innerHTML = `<div class="empty-state">${escapeHtml(t('noData'))}</div>`; return; }
  const max = Math.max(...rows.map(x => x.value), 1);
  el.innerHTML = rows.map(x => `<div class="bar-row" title="${attr(x.label)}">
    <div class="bar-label">${escapeHtml(trValue(x.label))}</div>
    <div class="bar-track"><div class="bar-fill" style="width:${Math.max(1, x.value/max*100)}%"></div></div>
    <div class="bar-value">${fmt(x.value)}</div>
  </div>`).join('');
}

function renderBaseBars(rows, total) {
  const map = Object.fromEntries(rows.map(x => [x.label, x.value]));
  const own = map['BASE PRÓPRIA'] || 0, fr = map['FRANQUIA'] || 0, denom = total || 1;
  $('#baseBars').innerHTML = `
    <div class="big-bar"><span>${escapeHtml(trValue('BASE PRÓPRIA'))}</span><strong>${fmt(own)}</strong><div class="mini-track"><div class="mini-fill" style="width:${own/denom*100}%"></div></div></div>
    <div class="big-bar franchise"><span>${escapeHtml(trValue('FRANQUIA'))}</span><strong>${fmt(fr)}</strong><div class="mini-track"><div class="mini-fill" style="width:${fr/denom*100}%"></div></div></div>`;
}

function slaBadge(value) {
  if (value === 'DENTRO DO PRAZO') return `<span class="badge good">${escapeHtml(trValue(value))}</span>`;
  if (value === 'FORA DO PRAZO') return `<span class="badge bad">${escapeHtml(trValue(value))}</span>`;
  return `<span class="badge neutral">${escapeHtml(trValue(value || 'SEM DATA'))}</span>`;
}
function typeBadge(value) { return `<span class="badge ${value === 'FRANQUIA' ? 'franchise':'own'}">${escapeHtml(trValue(value))}</span>`; }

function renderRecent(rows) {
  const body = $('#recentBody');
  if (!rows.length) { body.innerHTML = `<tr><td colspan="8" class="empty-state">${escapeHtml(t('noData'))}</td></tr>`; return; }
  body.innerHTML = rows.map(r => `<tr>
    <td>${escapeHtml(r.complaint_date_display)}</td><td>${escapeHtml(r.ra_id)}</td><td>${escapeHtml(r.regional)}</td>
    <td>${escapeHtml(r.base)}<div style="margin-top:5px">${typeBadge(r.base_type)}</div></td><td>${escapeHtml(r.rm || t('semRm'))}</td>
    <td>${escapeHtml(trValue(r.reason))}</td><td>${r.open ? `<span class="badge open">${escapeHtml(trValue(r.solution || 'EM ABERTO'))}</span>` : `<span class="badge closed">${escapeHtml(trValue(r.solution))}</span>`}</td>
    <td>${slaBadge(r.sla_auto)}</td>
  </tr>`).join('');
}

async function loadRanking() {
  setSyncing(true);
  try {
    const data = await api('/api/ranking?' + getQuery({only_open: $('#rankingOnlyOpen').checked ? '1' : '0'}));
    state.version = data.version;
    const head = $('#rankingHead'), body = $('#rankingBody');
    head.innerHTML = `<tr><th>RM</th><th>${escapeHtml(t('total'))}</th><th>${escapeHtml(t('own'))}</th><th>${escapeHtml(t('franchise'))}</th><th>${escapeHtml(t('within'))} 24h</th><th>${escapeHtml(t('outside'))} 24h</th>${data.reasons.map(r => `<th>${escapeHtml(trValue(r))}</th>`).join('')}</tr>`;
    if (!data.rows.length) { body.innerHTML = `<tr><td colspan="${6+data.reasons.length}" class="empty-state">${escapeHtml(t('noData'))}</td></tr>`; return; }
    body.innerHTML = data.rows.map(r => `<tr><td>${escapeHtml(r.rm === 'SEM RM' ? t('semRm') : r.rm)}</td><td><strong>${fmt(r.total)}</strong></td><td>${fmt(r.own)}</td><td>${fmt(r.franchise)}</td><td><span class="badge good">${fmt(r.within)}</span></td><td><span class="badge bad">${fmt(r.overdue)}</span></td>${data.reasons.map(reason => `<td class="rank-num">${fmt(r.reasons[reason] || 0)}</td>`).join('')}</tr>`).join('');
  } finally { setSyncing(false); }
}

function editorQuery() {
  const p = new URLSearchParams({page:String(state.page), per_page:'30'});
  const q = $('#editorSearch').value.trim(); if (q) p.set('q', q);
  const regional = $('#editorRegional').value; if (regional) p.set('regional', regional);
  const rm = $('#editorRm').value; if (rm) p.set('rm', rm);
  const type = $('#editorBaseType').value; if (type) p.set('base_type', type);
  return p.toString();
}

async function loadEditor() {
  setSyncing(true);
  try {
    const data = await api('/api/rows?' + editorQuery());
    state.version = data.version; state.pages = Math.max(data.pages, 1); state.page = Math.min(state.page, state.pages);
    $('#editorCount').textContent = `${fmt(data.total)} ${t('rows')}`;
    $('#pageInfo').textContent = state.lang === 'zh' ? `${t('page')} ${data.page} ${t('of')} ${state.pages}` : `${t('page')} ${data.page} ${t('of')} ${state.pages}`;
    $('#prevPage').disabled = data.page <= 1; $('#nextPage').disabled = data.page >= state.pages;
    renderEditorRows(data.rows); state.editorDirty = false;
  } finally { setSyncing(false); }
}

function editorInput(name, value, cls='') { return `<input class="${cls}" data-field="${name}" value="${attr(value || '')}">`; }
function editorDate(name, value) { return `<input class="w-date" type="datetime-local" data-field="${name}" value="${attr(value || '')}">`; }
function manualDeadlineSelect(value) {
  return `<select data-field="deadline_manual"><option value=""></option><option value="DENTRO DO PRAZO" ${value==='DENTRO DO PRAZO'?'selected':''}>${escapeHtml(trValue('DENTRO DO PRAZO'))}</option><option value="FORA DO PRAZO" ${value==='FORA DO PRAZO'?'selected':''}>${escapeHtml(trValue('FORA DO PRAZO'))}</option></select>`;
}

function renderEditorRows(rows) {
  const body = $('#editorBody');
  if (!rows.length) { body.innerHTML = `<tr><td colspan="17" class="empty-state">${escapeHtml(t('noData'))}</td></tr>`; return; }
  body.innerHTML = rows.map(r => `<tr data-id="${r.id}">
    <td>${editorInput('ra_id', r.ra_id)}</td>
    <td>${editorDate('complaint_date', r.complaint_date)}</td>
    <td><select data-field="regional"><option value="MG" ${r.regional==='MG'?'selected':''}>MG</option><option value="SPN" ${r.regional==='SPN'?'selected':''}>SPN</option><option value="${attr(r.regional)}" ${!['MG','SPN'].includes(r.regional)?'selected':''}>${escapeHtml(r.regional)}</option></select></td>
    <td>${editorInput('owner', r.owner, 'w-text')}</td><td>${editorInput('tracking', r.tracking, 'w-text')}</td><td>${editorInput('driver', r.driver, 'w-text')}</td>
    <td>${editorInput('base', r.base, 'w-text')}</td><td>${typeBadge(r.base_type)}</td><td>${editorInput('rm', r.rm, 'w-text')}</td>
    <td>${editorDate('directed_date', r.directed_date)}</td><td>${editorInput('reason', r.reason, 'w-text reason-input')}</td><td>${editorInput('solution', r.solution, 'w-text')}</td>
    <td>${manualDeadlineSelect(r.deadline_manual)}</td><td>${slaBadge(r.sla_auto)}</td>
    <td><textarea data-field="treatment" class="w-long treatment-input">${escapeHtml(r.treatment)}</textarea></td><td><textarea data-field="observation" class="w-long">${escapeHtml(r.observation)}</textarea></td>
    <td><div class="editor-actions-cell"><button class="mini-btn save" data-action="save">${escapeHtml(t('save'))}</button><button class="mini-btn delete" data-action="delete">${escapeHtml(t('delete'))}</button></div></td>
  </tr>`).join('');
}

function collectRow(tr) {
  const payload = {};
  tr.querySelectorAll('[data-field]').forEach(el => payload[el.dataset.field] = el.value);
  return payload;
}

async function saveRow(tr) {
  const id = tr.dataset.id, btn = tr.querySelector('[data-action="save"]');
  btn.disabled = true;
  try {
    const data = await api(`/api/rows/${id}`, {method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify(collectRow(tr))});
    state.version = data.version; state.editorDirty = false;
    showToast(t('saved')); await loadFilters(); await loadEditor();
  } catch (e) { showToast(e.message || t('error'), true); }
  finally { btn.disabled = false; }
}
async function deleteRow(tr) {
  if (!confirm(t('confirmDelete'))) return;
  try {
    const data = await api(`/api/rows/${tr.dataset.id}`, {method:'DELETE'}); state.version = data.version; showToast(t('deleted')); await loadFilters(); await loadEditor();
  } catch (e) { showToast(e.message || t('error'), true); }
}

async function updateTreatmentSuggestion(tr) {
  const reason = tr.querySelector('[data-field="reason"]').value.trim();
  if (!reason) return;
  try {
    const data = await api('/api/treatment-suggestion?reason=' + encodeURIComponent(reason));
    tr.querySelector('[data-field="treatment"]').value = data.treatment;
    state.editorDirty = true;
  } catch (_) {}
}

async function importFile(file) {
  if (!file) return;
  if (!confirm(t('confirmImport'))) { $('#xlsxUpload').value = ''; return; }
  const fd = new FormData(); fd.append('file', file);
  setSyncing(true); showToast(t('importing'));
  try {
    const data = await api('/api/upload', {method:'POST', body:fd}); state.version = data.version; showToast(`${t('imported')} ${fmt(data.rows)} ${t('rows')}`); state.page=1; await loadFilters(); await refreshAll();
  } catch(e) { showToast(e.message || t('error'), true); }
  finally { setSyncing(false); $('#xlsxUpload').value=''; }
}

async function createRow() {
  const fd = new FormData($('#addForm')); const payload = Object.fromEntries(fd.entries());
  try {
    const data = await api('/api/rows', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)}); state.version = data.version; $('#addDialog').close(); $('#addForm').reset(); showToast(t('created')); state.page=1; await loadFilters(); await loadEditor();
  } catch(e) { showToast(e.message || t('error'), true); }
}

async function refreshCurrentViews(languageOnly=false) {
  if (state.activeTab === 'overview') await loadDashboard();
  if (state.activeTab === 'ranking') await loadRanking();
  if (state.activeTab === 'editor') {
    if (languageOnly && !state.editorDirty) await loadEditor();
    else if (!languageOnly) await loadEditor();
  }
}
async function refreshAll() { await loadDashboard(); await loadRanking(); if (!state.editorDirty) await loadEditor(); }

function syncAnalysisControls() {
  $$('#regionalSegments .seg').forEach(x => x.classList.toggle('active', x.dataset.value === state.regional));
  $$('#rankingRegionalSegments .seg').forEach(x => x.classList.toggle('active', x.dataset.value === state.regional));
  $('#rmFilter').value = state.rm; $('#rankingRmFilter').value = state.rm;
  $('#startDate').value = state.start; $('#rankingStartDate').value = state.start;
  $('#endDate').value = state.end; $('#rankingEndDate').value = state.end;
}

async function clearAnalysisFilters() {
  state.regional='TODOS'; state.rm=''; state.start=''; state.end=''; syncAnalysisControls();
  if (state.activeTab === 'ranking') await loadRanking(); else await loadDashboard();
}

function bindEvents() {
  $('#langToggle').addEventListener('click', () => setLanguage(state.lang === 'pt' ? 'zh' : 'pt'));
  $$('.tab').forEach(btn => btn.addEventListener('click', async () => {
    $$('.tab').forEach(x => x.classList.remove('active')); $$('.tab-panel').forEach(x => x.classList.remove('active'));
    btn.classList.add('active'); $('#' + btn.dataset.tab).classList.add('active'); state.activeTab = btn.dataset.tab;
    await refreshCurrentViews();
  }));
  $$('#regionalSegments .seg').forEach(btn => btn.addEventListener('click', async () => { state.regional = btn.dataset.value; syncAnalysisControls(); await loadDashboard(); }));
  $$('#rankingRegionalSegments .seg').forEach(btn => btn.addEventListener('click', async () => { state.regional = btn.dataset.value; syncAnalysisControls(); await loadRanking(); }));
  $('#rmFilter').addEventListener('change', async e => { state.rm=e.target.value; syncAnalysisControls(); await loadDashboard(); });
  $('#rankingRmFilter').addEventListener('change', async e => { state.rm=e.target.value; syncAnalysisControls(); await loadRanking(); });
  $('#startDate').addEventListener('change', async e => { state.start=e.target.value; syncAnalysisControls(); await loadDashboard(); });
  $('#rankingStartDate').addEventListener('change', async e => { state.start=e.target.value; syncAnalysisControls(); await loadRanking(); });
  $('#endDate').addEventListener('change', async e => { state.end=e.target.value; syncAnalysisControls(); await loadDashboard(); });
  $('#rankingEndDate').addEventListener('change', async e => { state.end=e.target.value; syncAnalysisControls(); await loadRanking(); });
  $('#clearFilters').addEventListener('click', clearAnalysisFilters);
  $('#rankingClearFilters').addEventListener('click', clearAnalysisFilters);
  $('#rankingOnlyOpen').addEventListener('change', loadRanking);
  let searchTimer; $('#editorSearch').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer=setTimeout(()=>{state.page=1; loadEditor();},350); });
  ['editorRegional','editorRm','editorBaseType'].forEach(id => $('#'+id).addEventListener('change',()=>{state.page=1;loadEditor();}));
  $('#prevPage').addEventListener('click',()=>{ if(state.page>1){state.page--;loadEditor();} });
  $('#nextPage').addEventListener('click',()=>{ if(state.page<state.pages){state.page++;loadEditor();} });
  $('#editorBody').addEventListener('input', e => { if(e.target.matches('[data-field]')) state.editorDirty = true; });
  $('#editorBody').addEventListener('change', e => { if(e.target.classList.contains('reason-input')) updateTreatmentSuggestion(e.target.closest('tr')); });
  $('#editorBody').addEventListener('click', e => { const btn=e.target.closest('[data-action]'); if(!btn)return; const tr=btn.closest('tr'); if(btn.dataset.action==='save')saveRow(tr); if(btn.dataset.action==='delete')deleteRow(tr); });
  $('#xlsxUpload').addEventListener('change', e => importFile(e.target.files[0]));
  $('#addRowBtn').addEventListener('click',()=>$('#addDialog').showModal());
  $('#createRowSubmit').addEventListener('click', createRow);
}

async function pollVersion() {
  if (state.pollBusy) return; state.pollBusy = true;
  try {
    const data = await api('/api/version');
    if (state.version == null) state.version = data.version;
    else if (data.version !== state.version) {
      if (state.activeTab === 'editor' && state.editorDirty) {
        showToast(t('updatedOtherDirty')); state.version = data.version;
      } else {
        state.version = data.version; showToast(t('updatedOther')); await loadFilters(); await refreshCurrentViews();
      }
    }
  } catch (_) {} finally { state.pollBusy=false; }
}

async function init() {
  bindEvents(); setLanguage(state.lang, false);
  try { await loadFilters(); syncAnalysisControls(); await loadDashboard(); await loadEditor(); }
  catch(e) { showToast(e.message || t('error'), true); }
  setInterval(pollVersion, 2500);
}

document.addEventListener('DOMContentLoaded', init);
