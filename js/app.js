// Lógica de UI do app Orçamento de Evento

window.onerror = function(msg, src, line, col, err) {
  var div = document.createElement('div');
  div.style.cssText = 'position:fixed;top:0;left:0;right:0;background:#c0483f;color:#fff;padding:12px 16px;z-index:9999;font-size:13px;font-family:monospace;';
  div.textContent = 'ERRO JS [linha ' + line + ']: ' + msg;
  document.body.appendChild(div);
};

// Null-safe $: retorna objeto inerte em vez de null para evitar crashes
var _gepSafe = (function(){
  var noop = function(){ return _gepSafe; };
  var safe = { addEventListener:noop, removeEventListener:noop, click:noop,
    appendChild:noop, removeChild:noop, insertBefore:noop, remove:noop,
    querySelectorAll:function(){ return []; }, querySelector:function(){ return null; },
    classList:{ add:noop, remove:noop, contains:function(){ return false; }, toggle:noop },
    style:{}, options:[], children:[] };
  ['textContent','innerHTML','value','innerText'].forEach(function(p){
    Object.defineProperty(safe, p, { get:function(){ return ''; }, set:noop });
  });
  return safe;
})();
const $ = (id) => document.getElementById(id) || _gepSafe;

var TITULOS_ABA = {
  inicial:      'Planilha Inicial',
  cliente:      'Cliente',
  fechamento:   'Fechamento',
  fornecedores: 'Fornecedores',
  verba:        'Verba de Produção',
  servicos:     'Catálogo de Serviços' // removido
};

var TAB_GRUPOS = {
  inicial:      'Evento',
  cliente:      'Evento',
  fechamento:   'Evento',
  fornecedores: 'Financeiro',
  verba:        'Financeiro',
  cadastro:     'Configuração',
  servicos:     'Configuração'
};

function showTab(name) {
  document.querySelectorAll('.tab-section').forEach(function(el) { el.classList.remove('active'); });
  document.querySelectorAll('.nav-item').forEach(function(el) { el.classList.remove('active'); });
  document.querySelectorAll('.nav-section-label').forEach(function(el) { el.classList.remove('active-grp'); });

  var sec = $('tab-' + name);
  if (sec) sec.classList.add('active');
  var nav = document.querySelector('.nav-item[data-tab="' + name + '"]');
  if (nav) nav.classList.add('active');

  var grupo = TAB_GRUPOS[name] || '';
  var grpLabel = document.querySelector('.nav-section-label[data-group="' + grupo + '"]');
  if (grpLabel) grpLabel.classList.add('active-grp');

  var grpEl = $('topbarGroup');
  if (grpEl) grpEl.textContent = grupo;

  var title = $('topbarTitle');
  if (title) title.textContent = TITULOS_ABA[name] || 'Orçamento de Evento';
  atualizarTopbar();
}

function atualizarTopbar() {
  var num = ($('numEvento') && $('numEvento').value.trim()) || '';
  var el  = $('topbarBadge');
  if (el) el.textContent = num;
}

function novoEvento() {
  if (!confirm('Criar novo orçamento? O atual será apagado.')) return;
  localStorage.removeItem('orcamento_evento_v1');
  location.reload();
}

function fillSelect(sel, options, placeholder) {
  sel.innerHTML = '';
  var opt0 = document.createElement('option');
  opt0.value = '';
  opt0.textContent = placeholder || '-- selecione --';
  sel.appendChild(opt0);
  options.forEach(function(o) {
    var opt = document.createElement('option');
    opt.value = o;
    opt.textContent = o;
    sel.appendChild(opt);
  });
}

fillSelect($('cliente'), LISTAS.secretarias, '-- selecione o cliente/secretaria --');
fillSelect($('produtor'), LISTAS.produtores, '-- selecione o produtor --');
fillSelect($('clienteFech'), LISTAS.secretarias, '-- selecione o cliente/secretaria --');
fillSelect($('produtorFech'), LISTAS.produtores, '-- selecione o produtor --');

// ---------------------------------------------------------------------
// Cadastro de Fornecedores: removido (gerenciado pelo GEP)
// Fornecedores"), carregada inicialmente com a base embutida em LISTAS.
// fornecedorNomes/fornecedorPrazo/fornecedorData/datalist são recalculados
// a partir dessa tabela (refreshFornecedorLookups), então qualquer edição,
// inclusão ou exclusão de fornecedor se reflete imediatamente nas outras
// abas (autocomplete, prazo padrão, dados trazidos para Fornecedores).
// ---------------------------------------------------------------------
var fornecedorNomes = [];
var fornecedorPrazo = {};
var fornecedorData = {};

var dl = document.getElementById('dl-fornecedores');

function addCadastroRow(data) { /* aba Cadastro removida */ }

function refreshFornecedorLookups() {
  // Lê fornecedores do GEP (vtp_fornecedores_v1) em vez do DOM
  fornecedorPrazo = {};
  fornecedorData = {};
  var nomes = [];
  var forns = [];
  try { forns = JSON.parse(localStorage.getItem('vtp_fornecedores_v1')||'[]'); } catch(e) {}
  forns.forEach(function(f) {
    var nome = (f.nome||f.name||'').trim();
    if (!nome) return;
    fornecedorData[nome] = f;
    fornecedorPrazo[nome] = f.prazo || '';
    if(f.telefone) fornecedorData[nome].telefone = f.telefone;
    nomes.push(nome);
  });
  fornecedorNomes = nomes.sort(function(a,b){ return a.localeCompare(b,'pt-BR'); });
  var dl = document.getElementById('dl-fornecedores');
  if (dl) {
    dl.innerHTML = '';
    fornecedorNomes.forEach(function(n){
      var opt = document.createElement('option'); opt.value = n; dl.appendChild(opt);
    });
  }
}

// ─────────────────────────────────────────────────────────────────────
// TABELA GENÉRICA (Planilha Inicial + Fechamento)
// ─────────────────────────────────────────────────────────────────────
function updateValorComMargem() {
  var totalServicos = parseFloat(($('totalGeral').dataset||{}).raw) || 0;
  var q = parseFloat(($('verbaEstQtde')||{}).value) || 0;
  var p = parseFloat(($('verbaEstUnitario')||{}).value) || 0;
  var totalVerba = q * p;
  var totalEvento = totalServicos + totalVerba;
  var margem = parseFloat(($('margemLucro')||{}).value) || 0;
  var valor = totalEvento * (1 + margem / 100);
  $('totalDoEvento').textContent = totalEvento.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
  $('valorComMargem').textContent = valor.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
  $('custoBaseMargem').textContent = totalEvento.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
}
$('margemLucro').addEventListener('input', function() { updateValorComMargem(); });

;['verbaEstQtde','verbaEstUnitario'].forEach(function(id) {
  var el = document.getElementById(id);
  if (el) el.addEventListener('input', function(){
    var q = parseFloat($('verbaEstQtde').value)||0;
    var p = parseFloat($('verbaEstUnitario').value)||0;
    var total = q*p;
    $('totalVerbaEst').textContent = total.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
    updateValorComMargem();
  });
});


// ── Formata campo de preço: mostra R$ ao sair, número ao editar ──
function applyPriceFormat(inp) {
  function toDisplay(v) {
    if (!v && v !== 0) return '';
    return v === 0 ? 'R$\u00a00,00' : v.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
  }
  inp.addEventListener('focus', function() {
    var v = parseFloat(this.dataset.rawVal);
    this.type = 'number';
    this.value = (v && v !== 0) ? v : '';
    this.placeholder = '0';
  });
  inp.addEventListener('blur', function() {
    var v = parseFloat(this.value) || 0;
    this.dataset.rawVal = v;
    this.type = 'text';
    this.value = toDisplay(v);
  });
  inp.addEventListener('input', function() {
    this.dataset.rawVal = parseFloat(this.value) || 0;
  });
  // Inicializar
  var v = parseFloat(inp.value) || 0;
  inp.dataset.rawVal = v;
  inp.type = 'text';
  inp.value = toDisplay(v);
}

// Lê o valor raw de um input de preço (funciona com formato ou sem)
function readPrice(inp) {
  if (!inp) return 0;
  if (inp.dataset.rawVal !== undefined) return parseFloat(inp.dataset.rawVal) || 0;
  var s = inp.value.replace(/[^\d,.-]/g,'').replace(',','.');
  return parseFloat(s) || 0;
}

function createInternalTable(tbodyId, totalSpanId, onTotalChange) {
  var counter = 0;

  function addRow(data) {
    data = data || {};
    var tbody = $(tbodyId);
    var tr = document.createElement('tr');
    tr.dataset.rid = tbodyId + counter++;
    tr.innerHTML =
      '<td><input type="text" list="dl-catalogo-gep" class="f-servico" autocomplete="off" oninput="gepAutoFillServico(this)" value="' + (data.servico || '') + '"></td>' +
      '<td><input type="text" class="f-desc" value="' + (data.desc || '') + '"></td>' +
      '<td><input type="number" class="f-qtde" step="any" value="' + (data.qtde != null ? data.qtde : 1) + '"></td>' +
      '<td><input type="number" class="f-preco" step="any" value="' + (data.preco != null ? data.preco : 0) + '"></td>' +
      '<td><input type="number" class="f-freq" step="any" value="' + (data.freq != null ? data.freq : 1) + '"></td>' +
      '<td><select class="f-unidade"></select></td>' +
      '<td class="valor-final">R$ 0,00</td>' +
      '<td><input list="dl-fornecedores" class="f-fornecedor" value="' + (data.fornecedor || '') + '"></td>' +
      '<td><select class="f-forma"></select></td>' +
      '<td><input type="text" class="f-obs" value="' + (data.obs || '') + '"></td>' +
      '<td><button type="button" class="btn-del-row">\u00d7</button></td>';
    tbody.appendChild(tr);
    fillSelect(tr.querySelector('.f-unidade'), LISTAS.unidades, '--');
    fillSelect(tr.querySelector('.f-forma'), LISTAS.formas, '--');
    if (data.unidade) tr.querySelector('.f-unidade').value = data.unidade;
    if (data.forma)   tr.querySelector('.f-forma').value   = data.forma;
    var precoInp = tr.querySelector('.f-preco');
    applyPriceFormat(precoInp);
    var recalc = function() { updateRowTotal(tr); };
    tr.querySelectorAll('.f-qtde, .f-freq').forEach(function(el) { el.addEventListener('input', recalc); });
    precoInp.addEventListener('blur', recalc);
    precoInp.addEventListener('focus', function(){ setTimeout(recalc, 50); });
    tr.querySelector('.f-fornecedor').addEventListener('change', function(e) {
      var nome = e.target.value.trim();
      if (fornecedorPrazo[nome]) tr.querySelector('.f-forma').value = fornecedorPrazo[nome];
    });
    tr.querySelector('.btn-del-row').addEventListener('click', function() { tr.remove(); updateTotal(); });
    updateRowTotal(tr);
    return tr;
  }

  function updateRowTotal(tr) {
    var q = parseFloat(tr.querySelector('.f-qtde').value) || 0;
    var p = readPrice(tr.querySelector('.f-preco'));
    var f = parseFloat(tr.querySelector('.f-freq').value) || 0;
    var total = q * p * f;
    tr.querySelector('.valor-final').textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
    tr.dataset.total = total;
    updateTotal();
  }

  function updateTotal() {
    var total = 0;
    document.querySelectorAll('#' + tbodyId + ' tr').forEach(function(tr) { total += parseFloat(tr.dataset.total) || 0; });
    $(totalSpanId).textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
    $(totalSpanId).dataset.raw = total;
    if (onTotalChange) onTotalChange(total);
  }

  return { addRow: addRow, updateTotal: updateTotal };
}

var inicialTable    = createInternalTable('itensBody',           'totalGeral',     updateValorComMargem);
var fechamentoTable = createInternalTable('fechamentoItensBody', 'totalFechamento');

$('addRowBtn').addEventListener('click',         function() { inicialTable.addRow(); });
$('addFechamentoRowBtn').addEventListener('click',function() { fechamentoTable.addRow(); });
inicialTable.addRow(); inicialTable.addRow(); inicialTable.addRow();


function updateTotalFornecedores() {
  var total = 0;
  document.querySelectorAll('#fornecedoresBody tr').forEach(function(tr) { total += readPrice(tr.querySelector('.ff-valor')); });
  $('totalFornecedores').textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
}

var fornecedorRowCount = 0;

function getEventoDataInicioIso() {
  return $('dataInicioFech').value || $('dataInicio').value;
}

function vencimentoInfoFor(prazo) {
  var dias = parseDiasPrazo(prazo);
  if (dias != null) return { forma: 'faturado', dias: dias };
  if (isAVistaPrazo(prazo)) return { forma: 'avista', dias: null };
  return { forma: 'parcial', dias: null };
}

function updateFornecedorVencimento(tr) {
  var checked = tr.querySelector('.ff-forma-check:checked');
  var vencCell = tr.querySelector('.ff-vencimento');
  if (!checked) { vencCell.textContent = ''; return; }
  if (checked.value === 'faturado') {
    var dias = parseInt(tr.dataset.dias, 10);
    var dataInicioIso = getEventoDataInicioIso();
    vencCell.textContent = (dataInicioIso && !isNaN(dias))
      ? formatDateBR(addDiasIso(dataInicioIso, dias))
      : 'Condições em OBS';
  } else {
    vencCell.textContent = 'Condições em OBS';
  }
}

function addFornecedorRow(data) {
  data = data || {};
  var tbody = $('fornecedoresBody');
  var tr = document.createElement('tr');
  var rid = 'frow' + (fornecedorRowCount++);
  tr.dataset.rid = rid;
  tr.dataset.fornecedorNome = data.nome || '';

  var info = vencimentoInfoFor(data.prazo);
  var forma = data.forma || info.forma;
  var dias = data.dias != null ? data.dias : info.dias;
  tr.dataset.dias = (dias != null) ? dias : '';

  tr.innerHTML =
    '<td><input list="dl-fornecedores" class="ff-nome" value="' + (data.nome || '') + '"></td>' +
    '<td><input type="text" class="ff-cnpj" value="' + (data.cnpj || '') + '"></td>' +
    '<td><input type="text" class="ff-contato" value="' + (data.contato || '') + '"></td>' +
    '<td><input type="text" class="ff-telefone" value="' + (data.telefone || '') + '"></td>' +
    '<td><input type="number" step="any" class="ff-valor" value="' + (data.valor != null ? data.valor : 0) + '"></td>' +
    '<td class="checkbox-group">' +
      '<label><input type="radio" name="forma' + rid + '" class="ff-forma-check" value="faturado"' + (forma==='faturado'?' checked':'') + '> Faturado</label>' +
      '<label><input type="radio" name="forma' + rid + '" class="ff-forma-check" value="parcial"' + (forma==='parcial'?' checked':'') + '> Parcial</label>' +
      '<label><input type="radio" name="forma' + rid + '" class="ff-forma-check" value="avista"' + (forma==='avista'?' checked':'') + '> À Vista</label>' +
    '</td>' +
    '<td class="vencimento-cell ff-vencimento"></td>' +
    '<td><input type="text" class="ff-obs" value="' + (data.obs || '') + '"></td>' +
    '<td><button type="button" class="btn-del-row">×</button></td>';
  tbody.appendChild(tr);

  applyPriceFormat(tr.querySelector('.ff-valor'));
  tr.querySelector('.ff-valor').addEventListener('blur', updateTotalFornecedores);

  tr.querySelector('.ff-nome').addEventListener('change', function(e) {
    var nome = e.target.value.trim();
    tr.dataset.fornecedorNome = nome;
    var fd = fornecedorData[nome];
    if (fd) {
      tr.querySelector('.ff-cnpj').value = fd.cnpj || '';
      tr.querySelector('.ff-contato').value = fd.contato || '';
      tr.querySelector('.ff-telefone').value = fd.telefone || '';
      var inf = vencimentoInfoFor(fd.prazo);
      tr.dataset.dias = (inf.dias != null) ? inf.dias : '';
      tr.querySelectorAll('.ff-forma-check').forEach(function(cb) { cb.checked = (cb.value === inf.forma); });
      if (inf.forma !== 'faturado' && !tr.querySelector('.ff-obs').value) tr.querySelector('.ff-obs').value = fd.prazo || '';
      updateFornecedorVencimento(tr);
    }
  });

  tr.querySelectorAll('.ff-forma-check').forEach(function(cb) {
    cb.addEventListener('change', function() { updateFornecedorVencimento(tr); });
  });

  tr.querySelector('.ff-valor').addEventListener('input', updateTotalFornecedores);
  tr.querySelector('.btn-del-row').addEventListener('click', function() { tr.remove(); updateTotalFornecedores(); });

  updateFornecedorVencimento(tr);
  updateTotalFornecedores();
  return tr;
}

$('addFornecedorRowBtn').addEventListener('click', function() { addFornecedorRow(); });

// Traz fornecedores únicos usados na Planilha de Fechamento, com dados
// (CNPJ/contato/telefone/prazo) e valor total somado (soma de todos os
// itens de fechamento daquele fornecedor). Idempotente: não duplica
// fornecedor já presente na tabela de pagamento.
function trazerFornecedores() {
  // Atualiza o lookup do cadastro primeiro, para que fornecedores adicionados
  // ou editados pelo usuário sejam reconhecidos (CNPJ, contato, telefone, prazo).
  refreshFornecedorLookups();
  var somaPorFornecedor = {};
  var ordem = [];
  document.querySelectorAll('#fechamentoItensBody tr').forEach(function(tr) {
    var nome = tr.querySelector('.f-fornecedor').value.trim();
    if (!nome) return;
    var total = parseFloat(tr.dataset.total) || 0;
    if (!(nome in somaPorFornecedor)) { somaPorFornecedor[nome] = 0; ordem.push(nome); }
    somaPorFornecedor[nome] += total;
  });

  var jaTem = {};
  document.querySelectorAll('#fornecedoresBody tr').forEach(function(tr) { jaTem[tr.dataset.fornecedorNome] = true; });

  ordem.forEach(function(nome) {
    if (jaTem[nome]) return;
    var fd = fornecedorData[nome] || { cnpj:'', contato:'', telefone:'', prazo:'' };
    var inf = vencimentoInfoFor(fd.prazo);
    addFornecedorRow({
      nome: nome,
      cnpj: fd.cnpj,
      contato: fd.contato,
      telefone: fd.telefone,
      valor: somaPorFornecedor[nome],
      prazo: fd.prazo,
      obs: (inf.forma !== 'faturado') ? (fd.prazo || '') : ''
    });
  });
}
$('trazerFornecedoresBtn').addEventListener('click', trazerFornecedores);


// ─────────────────────────────────────────────────────────────────────
// ABA 2: Cliente
// ─────────────────────────────────────────────────────────────────────
var clienteRowCount = 0;

function addClienteRow(data) {
  data = data || {};
  var tbody = $('clienteItensBody');
  var tr = document.createElement('tr');
  tr.dataset.rid = 'crow' + (clienteRowCount++);
  tr.innerHTML =
    '<td><input type="text" class="cf-servico" value="' + (data.servico || '') + '"></td>' +
    '<td><input type="text" class="cf-desc" value="' + (data.desc || '') + '"></td>' +
    '<td><input type="number" class="cf-qtde" step="any" value="' + (data.qtde != null ? data.qtde : 1) + '"></td>' +
    '<td><input type="number" class="cf-preco" step="any" value="' + (data.preco != null ? data.preco : 0) + '"></td>' +
    '<td><input type="number" class="cf-freq" step="any" value="' + (data.freq != null ? data.freq : 1) + '"></td>' +
    '<td class="valor-final">R$ 0,00</td>' +
    '<td><button type="button" class="btn-del-row">\u00d7</button></td>';
  tbody.appendChild(tr);
  var cfPreco = tr.querySelector('.cf-preco');
  applyPriceFormat(cfPreco);
  var recalc = function() { updateClienteRowTotal(tr); };
  tr.querySelectorAll('.cf-qtde, .cf-freq').forEach(function(el) { el.addEventListener('input', recalc); });
  cfPreco.addEventListener('blur', recalc);
  tr.querySelector('.btn-del-row').addEventListener('click', function() { tr.remove(); updateTotalCliente(); });
  updateClienteRowTotal(tr);
}

function updateClienteRowTotal(tr) {
  var q = parseFloat(tr.querySelector('.cf-qtde').value) || 0;
  var p = readPrice(tr.querySelector('.cf-preco'));
  var f = parseFloat(tr.querySelector('.cf-freq').value) || 0;
  var total = q * p * f;
  tr.querySelector('.valor-final').textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
  tr.dataset.total = total;
  updateTotalCliente();
}

function updateTotalCliente() {
  var total = 0;
  document.querySelectorAll('#clienteItensBody tr').forEach(function(tr) { total += parseFloat(tr.dataset.total) || 0; });
  $('totalCliente').textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
}
$('addClienteRowBtn').addEventListener('click', function() { addClienteRow(); });

// ─────────────────────────────────────────────────────────────────────
// NAVEGAÇÃO E SINCRONIZAÇÃO ENTRE ABAS
// ─────────────────────────────────────────────────────────────────────

function formatDateBR(iso) {
  if (!iso) return '';
  var parts = iso.split('-');
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function isoToDate(iso) {
  if (!iso) return null;
  var p = iso.split('-');
  return new Date(Date.UTC(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10)));
}

function addDiasIso(iso, dias) {
  if (!iso) return '';
  var d = isoToDate(iso);
  d.setUTCDate(d.getUTCDate() + dias);
  return d.toISOString().slice(0,10);
}

function parseDiasPrazo(prazo) {
  if (!prazo) return null;
  var m = String(prazo).trim().match(/^(\d+)\s*d$/i);
  return m ? parseInt(m[1], 10) : null;
}

function isAVistaPrazo(prazo) {
  return /a\s*vista/i.test(String(prazo || ''));
}

// Trazer itens da Planilha Inicial para Cliente
function trazerItensDaInicial() {
  document.querySelectorAll('#itensBody tr').forEach(function(tr) {
    if (tr.dataset.broughtCliente === 'true') return;
    var servico = tr.querySelector('.f-servico').value;
    var desc = tr.querySelector('.f-desc').value.trim();
    var preco = parseFloat(tr.querySelector('.f-preco').value) || 0;
    if ((!servico || servico === '--') && !desc && preco === 0) return;
    addClienteRow({ servico:servico, desc:desc,
      qtde: parseFloat(tr.querySelector('.f-qtde').value) || 0,
      preco: preco, freq: parseFloat(tr.querySelector('.f-freq').value) || 0 });
    tr.dataset.broughtCliente = 'true';
  });
}
$('trazerItensBtn').addEventListener('click', trazerItensDaInicial);

// Trazer itens da Planilha Inicial para Fechamento
function trazerItensParaFechamento() {
  document.querySelectorAll('#itensBody tr').forEach(function(tr) {
    if (tr.dataset.broughtFechamento === 'true') return;
    var servico = tr.querySelector('.f-servico').value;
    var desc = tr.querySelector('.f-desc').value.trim();
    var preco = parseFloat(tr.querySelector('.f-preco').value) || 0;
    if ((!servico || servico === '--') && !desc && preco === 0) return;
    fechamentoTable.addRow({ servico:servico, desc:desc,
      qtde: parseFloat(tr.querySelector('.f-qtde').value) || 0,
      preco: preco, freq: parseFloat(tr.querySelector('.f-freq').value) || 0,
      unidade: tr.querySelector('.f-unidade').value,
      fornecedor: tr.querySelector('.f-fornecedor').value,
      forma: tr.querySelector('.f-forma').value,
      obs: tr.querySelector('.f-obs').value });
    tr.dataset.broughtFechamento = 'true';
  });
}
$('trazerFechamentoBtn').addEventListener('click', trazerItensParaFechamento);

// Sync cabeçalho — Planilha Inicial → Fechamento
function syncHeaderFechamento() {
  $('numEventoFech').value    = $('numEvento').value;
  $('clienteFech').value      = $('cliente').value;
  $('nomeEventoFech').value   = $('nomeEvento').value;
  $('dataInicioFech').value   = $('dataInicio').value;
  $('dataTerminoFech').value  = $('dataTermino').value;
  $('horarioFech').value      = $('horario').value;
  $('localFech').value        = $('local').value;
  $('publicoEstimadoFech').value = $('publicoEstimado').value;
  $('montagemFech').value     = $('montagem').value;
  $('produtorFech').value     = $('produtor').value;
}
$('atualizarHeaderFechBtn').addEventListener('click', syncHeaderFechamento);

// Sync cabeçalho — Planilha Inicial → Cliente
function syncResumoCliente() {
  $('cProjeto').value  = $('numEvento').value;
  $('cEvento').value   = $('nomeEvento').value;
  $('cData').value     = formatDateBR($('dataInicio').value);
  $('cHorario').value  = $('horario').value;
  $('cLocal').value    = $('local').value;
  $('cPax').value      = $('publicoEstimado').value;
}

// Sync cabeçalho — Fechamento/Inicial → Verba
function syncResumoVerba() {
  $('vNumEvento').value = $('numEventoFech').value || $('numEvento').value;
  $('vEvento').value    = $('nomeEventoFech').value || $('nomeEvento').value;
  var ini = formatDateBR($('dataInicioFech').value || $('dataInicio').value);
  var fim = formatDateBR($('dataTerminoFech').value || $('dataTermino').value);
  $('vData').value      = (fim && fim !== ini) ? (ini + ' a ' + fim) : ini;
  $('vProdutor').value  = $('produtorFech').value || $('produtor').value;
}

// Botões de navegação
$('avancarClienteBtn').addEventListener('click', function() {
  syncResumoCliente();
  if (document.querySelectorAll('#clienteItensBody tr').length === 0) trazerItensDaInicial();
  showTab('cliente');
});
$('voltarInicialBtn').addEventListener('click',    function() { showTab('inicial'); });
$('avancarFechamentoBtn').addEventListener('click', function() {
  if (!$('numEventoFech').value) syncHeaderFechamento();
  if (document.querySelectorAll('#fechamentoItensBody tr').length === 0) trazerItensParaFechamento();
  showTab('fechamento');
});
$('voltarClienteBtn').addEventListener('click',    function() { showTab('cliente'); });
// avancarVerbaBtn e voltarFornecedoresBtn definidos abaixo com implementação completa

function syncResumoFornecedores() {
  $('fCliente').value = $('clienteFech').value || $('cliente').value;
  $('fNumEvento').value = $('numEventoFech').value || $('numEvento').value;
  $('fEvento').value = $('nomeEventoFech').value || $('nomeEvento').value;
  $('fDataInicio').value = formatDateBR(getEventoDataInicioIso());
  $('fLocal').value = $('localFech').value || $('local').value;
  $('fProdutor').value = $('produtorFech').value || $('produtor').value;
}

$('avancarFornecedoresBtn').addEventListener('click', function() {
  syncResumoFornecedores();
  if (document.querySelectorAll('#fornecedoresBody tr').length === 0) trazerFornecedores();
  showTab('fornecedores');
});
$('voltarFechamentoBtn').addEventListener('click', function() { showTab('fechamento'); });
$('imprimirFornecedoresBtn').addEventListener('click', function() { window.print(); });

// ---------------------------------------------------------------------
// ABA 5: Verba de Produção (gastos avulsos do produtor, para imprimir e
// entregar ao financeiro)
// ---------------------------------------------------------------------
var verbaRowCount = 0;

function addVerbaRow(data) {
  data = data || {};
  var tbody = $('verbaBody');
  var tr = document.createElement('tr');
  tr.dataset.rid = 'vrow' + (verbaRowCount++);
  tr.innerHTML =
    '<td><select class="vb-produtor"></select></td>' +
    '<td><input type="text" class="vb-item" value="' + (data.item || '') + '"></td>' +
    '<td><input type="number" step="any" class="vb-valor" value="' + (data.valor != null ? data.valor : 0) + '"></td>' +
    '<td><input type="text" class="vb-obs" value="' + (data.obs || '') + '"></td>' +
    '<td><button type="button" class="btn-del-row">×</button></td>';
  tbody.appendChild(tr);
  fillSelect(tr.querySelector('.vb-produtor'), LISTAS.produtores, '-- selecione o produtor --');
  tr.querySelector('.vb-produtor').value = data.produtor || $('produtorFech').value || $('produtor').value || '';
  tr.querySelector('.vb-produtor').addEventListener('change', updateTotalVerba);
  tr.querySelector('.vb-valor').addEventListener('input', updateTotalVerba);
  tr.querySelector('.btn-del-row').addEventListener('click', function() { tr.remove(); updateTotalVerba(); });
  updateTotalVerba();
}

function updateTotalVerba() {
  var total = 0;
  var porProdutor = {};
  var ordem = [];
  document.querySelectorAll('#verbaBody tr').forEach(function(tr) {
    var valor = parseFloat(tr.querySelector('.vb-valor').value) || 0;
    var produtor = tr.querySelector('.vb-produtor').value || '(sem produtor)';
    total += valor;
    if (!(produtor in porProdutor)) { porProdutor[produtor] = 0; ordem.push(produtor); }
    porProdutor[produtor] += valor;
  });
  $('totalVerba').textContent = total.toLocaleString('pt-BR', {style:'currency', currency:'BRL'});

  var html = '<div style="font-weight:bold;color:var(--verba);margin-bottom:6px;">Total por produtor</div>';
  ordem.forEach(function(p) {
    html += '<div style="display:flex;justify-content:space-between;padding:3px 0;"><span>' + p + '</span><span style="font-weight:bold;">' +
      porProdutor[p].toLocaleString('pt-BR', {style:'currency', currency:'BRL'}) + '</span></div>';
  });
  $('totalPorProdutor').innerHTML = ordem.length ? html : '';
}

$('addVerbaRowBtn').addEventListener('click', function() { addVerbaRow(); });

function syncResumoVerba() {
  $('vNumEvento').value = $('numEventoFech').value || $('numEvento').value;
  $('vEvento').value = $('nomeEventoFech').value || $('nomeEvento').value;
  var ini = formatDateBR(getEventoDataInicioIso());
  var fim = formatDateBR($('dataTerminoFech').value || $('dataTermino').value);
  $('vData').value = (fim && fim !== ini) ? (ini + ' a ' + fim) : ini;
  $('vProdutor').value = $('produtorFech').value || $('produtor').value;
}

$('avancarVerbaBtn').addEventListener('click', function() {
  syncResumoVerba();
  showTab('verba');
});
$('voltarFornecedoresBtn').addEventListener('click', function() { showTab('fornecedores'); });
$('imprimirVerbaBtn').addEventListener('click', function() {
  // Coletar dados
  var numEvt   = ($('vNumEvento') && $('vNumEvento').value) || $('numEvento').value || '';
  var dataEvt  = ($('vData') && $('vData').value) || '';
  var evento   = ($('vEvento') && $('vEvento').value) || $('nomeEvento').value || '';
  var produtor = ($('vProdutor') && $('vProdutor').value) || $('produtor').value || '';
  var total    = $('totalVerba').textContent || 'R$ 0,00';

  // Montar linhas da tabela
  var linhas = '';
  var i = 0;
  document.querySelectorAll('#verbaBody tr').forEach(function(tr) {
    var prod  = tr.querySelector('.vb-produtor') ? tr.querySelector('.vb-produtor').value : '';
    var item  = tr.querySelector('.vb-item') ? tr.querySelector('.vb-item').value : '';
    var valor = tr.querySelector('.vb-valor') ? tr.querySelector('.vb-valor').value : '';
    var obs   = tr.querySelector('.vb-obs') ? tr.querySelector('.vb-obs').value : '';
    if (!prod && !item && !valor) return;
    var bg = i % 2 === 0 ? '#fff' : '#f2f2f2';
    linhas += '<tr style="background:' + bg + '">' +
      '<td style="text-align:center">' + (prod||'') + '</td>' +
      '<td>' + (item||'') + '</td>' +
      '<td style="text-align:right">R$ ' + (parseFloat(valor)||0).toLocaleString('pt-BR',{minimumFractionDigits:2}) + '</td>' +
      '<td>' + (obs||'') + '</td>' +
      '</tr>';
    i++;
  });
  if (!linhas) {
    for (var b = 0; b < 8; b++) {
      linhas += '<tr style="background:' + (b%2===0?'#fff':'#f2f2f2') + '"><td>&nbsp;</td><td></td><td></td><td></td></tr>';
    }
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:15mm 15mm 15mm 15mm;color:#000}' +
    'table{border-collapse:collapse;width:100%}' +
    'td,th{padding:6px 8px;font-size:10pt}' +
    '.titulo{background:#1F4E79;color:#fff;text-align:center;font-size:16pt;font-weight:bold;padding:12px;border:2px solid #1F4E79}' +
    '.cab-label{background:#bfbfbf;font-weight:bold;font-size:10pt;padding:6px 8px;border:1px solid #999;width:35%}' +
    '.cab-val{font-size:10pt;padding:6px 8px;border:1px solid #999}' +
    '.hdr{background:#1F4E79;color:#fff;font-weight:bold;text-align:center;font-size:10pt;border:1px solid #1F4E79}' +
    'td{border:1px solid #ccc}' +
    '.total-row{background:#bfbfbf;font-weight:bold;font-size:11pt}' +
    '.data-row{font-size:10pt;border:1px solid #ccc}' +
    '.entrega-label{font-weight:bold;font-size:10pt;border:1px solid #999;width:25%}' +
    '.entrega-val{border:1px solid #999}' +
    'p.obs{font-size:10pt;margin-top:8px}' +
    '@media print{body{margin:10mm}}' +
    '</style></head><body>' +

    // TÍTULO
    '<table><tr><td class="titulo">VERBA DE PRODUÇÃO</td></tr></table>' +
    '<br>' +

    // CABEÇALHO DO EVENTO
    '<table>' +
    '<tr><td class="cab-label">Nº DO EVENTO</td><td class="cab-val">' + numEvt + '</td></tr>' +
    '<tr><td class="cab-label" style="background:#fff">DATA DO EVENTO:</td><td class="cab-val" style="background:#fff">' + dataEvt + '</td></tr>' +
    '<tr><td class="cab-label">NOME DO EVENTO:</td><td class="cab-val">' + evento + '</td></tr>' +
    '<tr><td class="cab-label" style="background:#fff">PRODUTOR:</td><td class="cab-val" style="color:blue">' + produtor + '</td></tr>' +
    '</table>' +
    '<br>' +

    // TABELA DE ITENS
    '<table>' +
    '<tr><th class="hdr" style="width:20%">Produtor</th><th class="hdr" style="width:35%">Item / Descrição</th><th class="hdr" style="width:20%">Valor (R$)</th><th class="hdr" style="width:25%">OBS</th></tr>' +
    linhas +
    '</table>' +
    '<br>' +

    // TOTAL GASTO
    '<table><tr class="total-row">' +
    '<td style="border:1px solid #999;text-align:center;width:20%">TOTAL GASTO</td>' +
    '<td style="border:1px solid #999;text-align:right;width:35%">' + total + '</td>' +
    '<td style="border:1px solid #999;width:20%"></td>' +
    '<td style="border:1px solid #999;width:25%"></td>' +
    '</tr></table>' +
    '<br>' +

    // DATA ENTREGA / REEMBOLSO
    '<table>' +
    '<tr><td class="entrega-label">DATA DA ENTREGA:</td><td class="entrega-val"></td></tr>' +
    '<tr><td class="entrega-label">DATA DO REEMBOLSO:</td><td class="entrega-val"></td></tr>' +
    '</table>' +
    '<p class="obs"><strong>OBS:</strong></p>' +
    '</body></html>';

  var win = window.open('', '_blank', 'width=800,height=900');
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(function(){ win.print(); }, 500);
});

// ---------------------------------------------------------------------
// Geração do Excel com ExcelJS (preserva cores e formatação do template)
// ---------------------------------------------------------------------
// ============================================================================
// GERADOR DE EXCEL — dirigido por template extraído da planilha-máscara
// (PLANILHA_ORÇAMENTAteste.xlsx). Cada célula estática carrega o estilo,
// fórmula e formato numérico exatos da máscara; as linhas repetidas (itens,
// fornecedores, cadastro) usam um protótipo de linha também extraído dela.
// Para atualizar o visual no futuro, basta regerar o GEP_TEMPLATE a partir
// de uma nova máscara — o código não precisa mudar.
// ============================================================================


// =====================================================================
// CATÁLOGO DE SERVIÇOS — persiste separado do orçamento, disponível
// em todos os eventos. Edições aqui atualizam o datalist dl-servicos.
// =====================================================================

// Catálogo: gerenciado pelo GEP (vtp_catalogo_v1)
function carregarCatalogoServicos() { return []; }
function salvarCatalogoServicos(lista) {}

function atualizarDlServicos() {
  // dl-servicos é preenchido pelo GEP via plnlInitGEPData
}
function salvarDoCatalogoServicos() {}

function renderCatalogoServicos() {}
// renderCatalogoServicos: aba removida

// =====================================================================
// AUTO-SAVE com localStorage
// =====================================================================

var _saveTimer = null;

function agendarSave() {
  if (_saveTimer) clearTimeout(_saveTimer);
  var status = document.getElementById('saveStatus');
  if (status) { status.textContent = '● Salvando...'; status.className = 'save-status saving'; }
  _saveTimer = setTimeout(salvarEstado, 500);
}

function coletarEstado() {
  function rows(tbodyId, fields) {
    var result = [];
    document.querySelectorAll('#' + tbodyId + ' tr').forEach(function(tr) {
      var obj = {};
      fields.forEach(function(f) {
        var el = tr.querySelector('.' + f.cls);
        if (el) obj[f.key] = el.tagName === 'SELECT' ? el.value : el.value;
      });
      result.push(obj);
    });
    return result;
  }

  return {
    numEvento:  ($('numEvento')      || {}).value || '',
    nomeEvento: ($('nomeEvento')     || {}).value || '',
    cliente:    ($('cliente')        || {}).value || '',
    dataInicio: ($('dataInicio')     || {}).value || '',
    dataTermino:($('dataTermino')    || {}).value || '',
    horario:    ($('horario')        || {}).value || '',
    local:      ($('local')          || {}).value || '',
    publicoEstimado: ($('publicoEstimado') || {}).value || '',
    montagem:   ($('montagem')       || {}).value || '',
    produtor:   ($('produtor')       || {}).value || '',
    margemLucro:($('margemLucro')    || {}).value || '',
    verbaEstDescritivo: ($('verbaEstDescritivo') || {}).value || '',
    verbaEstQtde:       ($('verbaEstQtde')       || {}).value || '',
    verbaEstUnitario:   ($('verbaEstUnitario')   || {}).value || '',
    nomeEmpresa:        ($('nomeEmpresa')        || {}).value || '',
    verbaEstObs:        ($('verbaEstObs')        || {}).value || '',
    prazoValidade:      ($('prazoValidade')      || {}).value || '',
    observacoes:        ($('observacoes')        || {}).value || '',

    itens: rows('itensBody', [
      {cls:'f-servico', key:'servico'}, {cls:'f-desc', key:'desc'},
      {cls:'f-qtde', key:'qtde'}, {cls:'f-preco', key:'preco'},
      {cls:'f-freq', key:'freq'}, {cls:'f-unidade', key:'unidade'},
      {cls:'f-fornecedor', key:'fornecedor'}, {cls:'f-forma', key:'forma'},
      {cls:'f-obs', key:'obs'}
    ]),

    clienteItens: rows('clienteItensBody', [
      {cls:'cf-servico', key:'servico'}, {cls:'cf-desc', key:'desc'},
      {cls:'cf-qtde', key:'qtde'}, {cls:'cf-preco', key:'preco'},
      {cls:'cf-freq', key:'freq'}
    ]),

    fechamento: {
      numEvento:   ($('numEventoFech')   || {}).value || '',
      cliente:     ($('clienteFech')     || {}).value || '',
      nomeEvento:  ($('nomeEventoFech')  || {}).value || '',
      dataInicio:  ($('dataInicioFech')  || {}).value || '',
      dataTermino: ($('dataTerminoFech') || {}).value || '',
      horario:     ($('horarioFech')     || {}).value || '',
      local:       ($('localFech')       || {}).value || '',
      publicoEstimado: ($('publicoEstimadoFech') || {}).value || '',
      montagem:    ($('montagemFech')    || {}).value || '',
      produtor:    ($('produtorFech')    || {}).value || '',
      itens: rows('fechamentoItensBody', [
        {cls:'f-servico', key:'servico'}, {cls:'f-desc', key:'desc'},
        {cls:'f-qtde', key:'qtde'}, {cls:'f-preco', key:'preco'},
        {cls:'f-freq', key:'freq'}, {cls:'f-unidade', key:'unidade'},
        {cls:'f-fornecedor', key:'fornecedor'}, {cls:'f-forma', key:'forma'},
        {cls:'f-obs', key:'obs'}
      ])
    },

    fornecedores: rows('fornecedoresBody', [
      {cls:'ff-nome', key:'nome'}, {cls:'ff-cnpj', key:'cnpj'},
      {cls:'ff-contato', key:'contato'}, {cls:'ff-telefone', key:'telefone'},
      {cls:'ff-valor', key:'valor'}, {cls:'ff-forma', key:'forma'},
      {cls:'ff-vencimento', key:'vencimento'}, {cls:'ff-obs', key:'obs'}
    ]),

    verba: rows('verbaBody', [
      {cls:'vb-produtor', key:'produtor'}, {cls:'vb-item', key:'item'},
      {cls:'vb-valor', key:'valor'}, {cls:'vb-obs', key:'obs'}
    ]),

    cadastro: rows('cadastroBody', [
      {cls:'cad-nome', key:'nome'}, {cls:'cad-cnpj', key:'cnpj'},
      {cls:'cad-contato', key:'contato'}, {cls:'cad-telefone', key:'telefone'},
      {cls:'cad-prazo', key:'prazo'}, {cls:'cad-servico', key:'servico'}
    ])
  };
}

function salvarEstado() {
  try {
    localStorage.setItem('orcamento_evento_v1', JSON.stringify(coletarEstado()));
    var status = document.getElementById('saveStatus');
    if (status) { status.textContent = '● Salvo'; status.className = 'save-status saved'; }
  } catch(e) { console.warn('Erro ao salvar:', e); }
}

function carregarEstado() {
  var raw = localStorage.getItem('orcamento_evento_v1');
  if (!raw) return;
  try {
    var s = JSON.parse(raw);
    function set(id, val) { var el = $(id); if (el && val !== undefined) el.value = val; }

    set('numEvento', s.numEvento);
    set('nomeEvento', s.nomeEvento);
    set('cliente', s.cliente);
    set('dataInicio', s.dataInicio);
    set('dataTermino', s.dataTermino);
    set('horario', s.horario);
    set('local', s.local);
    set('publicoEstimado', s.publicoEstimado);
    set('montagem', s.montagem);
    set('produtor', s.produtor);
    set('margemLucro', s.margemLucro);
    set('verbaEstDescritivo', s.verbaEstDescritivo);
    set('verbaEstQtde', s.verbaEstQtde);
    set('verbaEstUnitario', s.verbaEstUnitario);
    set('nomeEmpresa', s.nomeEmpresa);
    set('verbaEstObs', s.verbaEstObs);
    set('prazoValidade', s.prazoValidade);
    set('observacoes', s.observacoes);

    if (s.itens && s.itens.length) {
      $('itensBody').innerHTML = '';
      s.itens.forEach(function(r) { inicialTable.addRow(r); });
    }
    if (s.clienteItens && s.clienteItens.length) {
      $('clienteItensBody').innerHTML = '';
      s.clienteItens.forEach(function(r) { addClienteRow(r); });
    }
    if (s.fechamento) {
      var f = s.fechamento;
      set('numEventoFech', f.numEvento);
      set('clienteFech', f.cliente);
      set('nomeEventoFech', f.nomeEvento);
      set('dataInicioFech', f.dataInicio);
      set('dataTerminoFech', f.dataTermino);
      set('horarioFech', f.horario);
      set('localFech', f.local);
      set('publicoEstimadoFech', f.publicoEstimado);
      set('montagemFech', f.montagem);
      set('produtorFech', f.produtor);
      if (f.itens && f.itens.length) {
        $('fechamentoItensBody').innerHTML = '';
        f.itens.forEach(function(r) { fechamentoTable.addRow(r); });
      }
    }
    if (s.fornecedores && s.fornecedores.length) {
      $('fornecedoresBody').innerHTML = '';
      s.fornecedores.forEach(function(r) { addFornecedorRow(r); });
    }
    if (s.verba && s.verba.length) {
      $('verbaBody').innerHTML = '';
      s.verba.forEach(function(r) { addVerbaRow(r); });
    }
    refreshFornecedorLookups(); // s.cadastro: aba removida

    atualizarTopbar();
    updateValorComMargem();
  } catch(e) { console.warn('Erro ao carregar estado:', e); }
}

// Wire auto-save to all inputs
document.addEventListener('input', agendarSave);
document.addEventListener('change', agendarSave);

// Wire topbar badge update
['numEvento','nomeEvento'].forEach(function(id) {
  var el = $(id);
  if (el) el.addEventListener('input', atualizarTopbar);
});

// Load on start
window.addEventListener('load', function() {
  carregarEstado();
  atualizarTopbar();
});
