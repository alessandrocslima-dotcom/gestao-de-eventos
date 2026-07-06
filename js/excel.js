// Geração de Excel — funções GEP e gerarWorkbook

function gepStyle(ix) { return GEP_TEMPLATE.styles[ix] || {}; }

// Quando GEP_VALORES_ONLY é true, as fórmulas do template NÃO são gravadas —
// a célula recebe só o estilo, e o valor calculado é injetado depois por
// gepPreencherValores. Isso deixa a planilha muito mais leve e rápida de
// gerar/abrir (sem VLOOKUP/SUMIF/SUM para o Excel processar), mantendo layout,
// cores, bordas e os números finais idênticos.
var GEP_VALORES_ONLY = true;

function gepApplyCell(ws, ref, def) {
  var c = ws.getCell(ref);
  if (def.fx !== undefined && !GEP_VALORES_ONLY) c.value = { formula: def.fx };
  else if (def.v !== undefined) c.value = def.v;
  var st = (def.s !== undefined) ? gepStyle(def.s) : {};
  var f = st.f || {};
  c.font = {
    name: f.name || FONT_BASE,
    size: f.size || 11,
    bold: !!f.bold,
    italic: !!f.italic,
    color: f.color ? { argb: 'FF' + f.color } : undefined
  };
  if (st.bg) c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF'+st.bg } };
  if (st.a) c.alignment = {
    horizontal: st.a.h || undefined,
    vertical: st.a.v || undefined,
    wrapText: !!st.a.w
  };
  if (st.bd) {
    var b = {};
    if (st.bd.t) b.top = { style: st.bd.t };
    if (st.bd.b) b.bottom = { style: st.bd.b };
    if (st.bd.l) b.left = { style: st.bd.l };
    if (st.bd.r) b.right = { style: st.bd.r };
    c.border = b;
  }
  if (def.n) c.numFmt = def.n;
  return c;
}

// Quantas linhas em branco já formatadas vêm prontas em cada tabela quando o
// usuário preenche pouca coisa. Menos linhas = geração muito mais rápida.
// Se o formulário tiver mais itens que o mínimo, a faixa expande sozinha para
// caber tudo (ver rowLimits). As fórmulas de subtotal continuam somando o
// range cheio da máscara (ex.: H8:H160), então linhas não pintadas somam zero
// normalmente e nada é perdido.
var GEP_MIN_ROWS = {
  'PLANILHA INICIAL': 15,
  'PLANILHA DE FECHAMENTO': 15,
  'PLAN INTERNA FORNECEDORES': 15,
  'CADASTRO DE FORNECEDORES': 15
};
var GEP_ROW_FOLGA = 3; // linhas em branco extras após o último item preenchido

// Guarda, por aba, quantas linhas o rodapé foi deslocado para cima (offset),
// para que gerarWorkbook grave os valores nas posições novas. Ex.: se a faixa
// de Fornecedores foi de 9..100 para 9..20, o rodapé (linha 101+) sobe 80
// linhas e o TOTAL, que era F101, passa a ser F21.
var GEP_FOOTER_OFFSET = {};

// Traduz uma referência do rodapé (ex.: 'F101') para a linha já deslocada,
// usando o offset da aba. Referências fora do rodapé passam inalteradas.
function gepRef(sheetName, ref, footerFromRow) {
  var off = GEP_FOOTER_OFFSET[sheetName] || 0;
  if (!off) return ref;
  var m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) return ref;
  var row = parseInt(m[2], 10);
  if (row < footerFromRow) return ref;   // cabeçalho/faixa: não desloca
  return m[1] + (row - off);
}

// Monta todas as abas do template: larguras, mesclagens, células estáticas,
// linhas repetidas (protótipo) e configuração de página. rowLimits (opcional)
// informa, por aba, até que linha a faixa repetida deve ser formatada. O
// rodapé (tudo abaixo da faixa) é deslocado para encostar na última linha
// formatada — sem vão de linhas vazias e sem depender de ocultar linhas
// (que o ExcelJS não preserva de forma confiável).
function gepBuildFromTemplate(wb, rowLimits) {
  rowLimits = rowLimits || {};
  GEP_FOOTER_OFFSET = {};
  var sheets = {};
  GEP_TEMPLATE.order.forEach(function(name) {
    var t = GEP_TEMPLATE.sheets[name];
    var ws = wb.addWorksheet(name);
    sheets[name] = ws;
    ws.views = [{ showGridLines: !!t.grid }];

    Object.keys(t.widths || {}).forEach(function(col) {
      ws.getColumn(col).width = t.widths[col];
    });

    // Descobre a faixa repetida e o quanto o rodapé sobe (offset).
    var rep = (t.repeat && t.repeat[0]) || null;
    var to = rep ? rep.to : 0;
    var footerFrom = rep ? rep.to + 1 : 1e9;
    if (rep) {
      var lim = rowLimits[name];
      if (typeof lim === 'number') to = Math.min(rep.to, Math.max(rep.from - 1, lim));
    }
    var offset = rep ? (rep.to - to) : 0;
    GEP_FOOTER_OFFSET[name] = offset;

    function shiftRow(row) { return (row >= footerFrom) ? (row - offset) : row; }
    function shiftRef(ref) {
      var m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) return ref;
      return m[1] + shiftRow(parseInt(m[2], 10));
    }
    function shiftMerge(mr) {
      var p = mr.split(':');
      return p.map(shiftRef).join(':');
    }

    // Células estáticas primeiro (com borda/estilo completos), mesclagens
    // depois. A ordem importa: o ExcelJS só preserva corretamente as bordas
    // de cada célula de um range mesclado se o merge acontecer DEPOIS de
    // definir os estilos — mesclando antes, ele descarta as bordas parciais
    // e sobra só a da última célula processada, resultando em caixas com
    // lados faltando (ex.: o quadro do "Fornecedor"/"Contra apresentação").
    Object.keys(t.cells || {}).forEach(function(ref) {
      gepApplyCell(ws, shiftRef(ref), t.cells[ref]);
    });

    // Mesclagens: cabeçalho/faixa iguais; rodapé deslocado.
    (t.merges || []).forEach(function(m) {
      try { ws.mergeCells(shiftMerge(m)); } catch (e) {}
    });

    // Faixa repetida (itens): formata só até o limite enxuto. Estilo de cada
    // célula primeiro, merge K:L da linha por último (mesmo motivo acima).
    if (rep) {
      for (var r = rep.from; r <= to; r++) {
        Object.keys(rep.cols).forEach(function(col) {
          var proto = rep.cols[col];
          var def = { s: proto.s, n: proto.n };
          if (proto.fx !== undefined) def.fx = proto.fx.split('{r}').join(String(r));
          if (proto.v !== undefined) def.v = proto.v;
          gepApplyCell(ws, col + r, def);
        });
        if (rep.mergeKL) { try { ws.mergeCells('K'+r+':L'+r); } catch (e) {} }
        if (rep.h) ws.getRow(r).height = rep.h;
      }
    }

    // Alturas de linha: rodapé deslocado.
    Object.keys(t.heights || {}).forEach(function(r) {
      ws.getRow(shiftRow(parseInt(r, 10))).height = t.heights[r];
    });

    var p = t.print || {};
    ws.pageSetup = {
      orientation: p.orientation || 'portrait',
      margins: { left:0.51, right:0.51, top:0.79, bottom:0.79, header:0.31, footer:0.31 }
    };
    if (p.scale) ws.pageSetup.scale = p.scale;
    if (p.area) {
      // Área de impressão: ajusta a última linha para a nova posição do rodapé
      // e usa a API correta do ExcelJS (pageSetup.printArea).
      var adjustedArea = p.area.replace(/(\d+)(?!.*\d)/, function(n){ return shiftRow(parseInt(n,10)); });
      ws.pageSetup.printArea = adjustedArea;
    }
  });
  return sheets;
}

function gepSetV(ws, ref, val) {
  if (val === undefined || val === null || val === '') return;
  ws.getCell(ref).value = val;
}
function gepNum(v) {
  var n = parseFloat(v);
  return isNaN(n) ? null : n;
}
function gepDateOrNull(iso) {
  if (!iso) return null;
  var d = isoToDate(iso);
  return d || null;
}
function gepMaybeNumber(v) {
  if (v === undefined || v === null || v === '') return '';
  return /^\d+$/.test(String(v).trim()) ? parseInt(v, 10) : v;
}

// Validações de dados (dropdowns dentro do próprio Excel) — iguais à máscara.
function gepAddListValidation(ws, range, formula) {
  // Dropdowns desativados a pedido: sem validações de dados na planilha
  // (deixa o arquivo ainda mais leve e a geração mais rápida).
}

var GEP_ITEM_COLS = { servico:'B', desc:'C', qtde:'D', preco:'E', freq:'F', unidade:'G', fornecedor:'I', obs:'K' };

// Injeta os itens de serviço (Inicial/Fechamento) nas linhas 8..160.
// Linhas vazias mantêm D/E realmente vazios (nunca '') para não travar a
// fórmula H=D*E*F com #VALOR!; Freq. recebe 1 como na máscara.
// Injeta itens de serviço e, no modo só-valores, já calcula o Valor final
// (H = Qtde × Unit × Freq) e a Forma de pagamento (J = prazo do fornecedor no
// cadastro). Recebe cadIndex (nome→dados do cadastro) para o "VLOOKUP" em JS.
// Retorna a soma dos valores finais (subtotal dos serviços).
function gepInjectItems(ws, trs, selPrefix, cadIndex, maxRow) {
  var first = 8, last = 160, soma = 0;
  for (var i = 0; i < trs.length && (first + i) <= last; i++) {
    var tr = trs[i], r = first + i;
    gepSetV(ws, 'B'+r, tr.querySelector('.'+selPrefix+'servico').value);
    gepSetV(ws, 'C'+r, tr.querySelector('.'+selPrefix+'desc').value);
    var q = gepNum(tr.querySelector('.'+selPrefix+'qtde').value);
    var p = gepNum(tr.querySelector('.'+selPrefix+'preco').value);
    if (q !== null) ws.getCell('D'+r).value = q;
    if (p !== null) ws.getCell('E'+r).value = p;
    var fq = gepNum(tr.querySelector('.'+selPrefix+'freq').value);
    var freq = (fq !== null && fq !== 0) ? fq : 1;
    ws.getCell('F'+r).value = freq;
    gepSetV(ws, 'G'+r, tr.querySelector('.'+selPrefix+'unidade').value);
    var fornecedor = tr.querySelector('.'+selPrefix+'fornecedor').value;
    gepSetV(ws, 'I'+r, fornecedor);
    gepSetV(ws, 'K'+r, tr.querySelector('.'+selPrefix+'obs').value);
    // Valor final = D×E×F (H). Na máscara era fórmula; aqui grava o número.
    var valor = (q || 0) * (p || 0) * freq;
    if (q !== null || p !== null) {
      ws.getCell('H'+r).value = valor;
      soma += valor;
    }
    // Forma de pagamento = prazo do fornecedor no cadastro (era VLOOKUP col 5).
    if (GEP_VALORES_ONLY && cadIndex && fornecedor && cadIndex[fornecedor]) {
      var prazo = cadIndex[fornecedor].prazo;
      if (prazo) ws.getCell('J'+r).value = prazo;
    }
  }
  // Freq. = 1 só nas linhas em branco DENTRO da faixa visível (limite). Não
  // preenche até 160, senão sobrariam células "1" soltas embaixo do rodapé
  // que agora sobe. maxRow é o fim visível da faixa desta aba.
  var fimVisivel = (typeof maxRow === 'number') ? maxRow : last;
  for (var r2 = first + trs.length; r2 <= fimVisivel; r2++) {
    ws.getCell('F'+r2).value = 1;
  }
  return soma;
}

$('gerarBtn').addEventListener('click', function() {
  var btn = this;
  $('msg').textContent = 'Gerando planilha...';
  btn.disabled = true;
  // Espera o navegador pintar a mensagem antes do trabalho pesado (senão a
  // tela "congela" sem mostrar nada). Envolve tudo em try/catch para que um
  // erro nunca deixe o botão preso em "Gerando planilha...".
  setTimeout(function() {
    var p;
    try {
      p = gerarWorkbook();
    } catch (err) {
      console.error(err);
      $('msg').textContent = 'Erro ao gerar planilha: ' + (err && err.message ? err.message : err);
      btn.disabled = false;
      return;
    }
    Promise.resolve(p).then(function() {
      $('msg').textContent = 'Planilha gerada e baixada!';
      btn.disabled = false;
      setTimeout(function() { $('msg').textContent = ''; }, 5000);
    }).catch(function(err) {
      console.error(err);
      $('msg').textContent = 'Erro ao gerar planilha: ' + (err && err.message ? err.message : err);
      btn.disabled = false;
    });
  }, 50);
});

// Monta um índice do Cadastro de Fornecedores (nome -> dados) a partir das
// linhas do formulário, para substituir os VLOOKUPs por consulta direta em JS.
function gepIndexarCadastro() {
  var idx = {};
  document.querySelectorAll('#cadastroBody tr').forEach(function(tr) {
    var nome = (tr.querySelector('.cad-nome').value || '').trim();
    if (!nome) return;
    idx[nome] = {
      cnpj:     tr.querySelector('.cad-cnpj').value || '',
      contato:  tr.querySelector('.cad-contato').value || '',
      telefone: tr.querySelector('.cad-telefone').value || '',
      prazo:    tr.querySelector('.cad-prazo').value || '',
      servico:  tr.querySelector('.cad-servico').value || ''
    };
  });
  return idx;
}

// Converte um prazo ("30D", "45D", "À VISTA"...) no número de dias.
function gepPrazoDias(prazo) {
  if (!prazo) return 0;
  var m = String(prazo).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

// Data (Date) -> "dd/mm/aaaa" para textos que na máscara usavam TEXT().
function gepDataBR(d) {
  if (!d) return '';
  var dd = ('0' + d.getUTCDate()).slice(-2);
  var mm = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + d.getUTCFullYear();
}

// Antes de gerar, garante que os itens da Planilha Inicial estejam propagados
// para as abas Cliente, Fechamento e Fornecedores (caso o usuário não tenha
// feito isso manualmente com os botões "Trazer..."). Sem isso, a chamada no
// início de gerarWorkbook quebrava com "garantirPropagacaoDeItens is not
// defined" e o botão dava erro ao gerar.
function garantirPropagacaoDeItens() {
  if (!$('numEventoFech').value) syncHeaderFechamento();
  if (document.querySelectorAll('#clienteItensBody tr').length === 0) trazerItensDaInicial();
  if (document.querySelectorAll('#fechamentoItensBody tr').length === 0) trazerItensParaFechamento();
  if (document.querySelectorAll('#fornecedoresBody tr').length === 0) trazerFornecedores();
}

function gerarWorkbook() {
  garantirPropagacaoDeItens();
  var wb = new ExcelJS.Workbook();

  // Descobre quantas linhas cada tabela do formulário realmente tem e define
  // até onde formatar a grade em cada aba: pelo menos o mínimo, ou o número de
  // itens + uma folga, o que for maior (nunca além do limite da máscara). Isso
  // evita pintar centenas de linhas em branco e deixa a geração bem mais rápida.
  function gepLimite(sel, first, min) {
    var n = document.querySelectorAll(sel).length;
    return (first - 1) + Math.max(min, n + GEP_ROW_FOLGA);
  }
  var rowLimits = {
    'PLANILHA INICIAL':          gepLimite('#itensBody tr', 8, GEP_MIN_ROWS['PLANILHA INICIAL']),
    'PLANILHA DE FECHAMENTO':    gepLimite('#fechamentoItensBody tr', 8, GEP_MIN_ROWS['PLANILHA DE FECHAMENTO']),
    'PLAN INTERNA FORNECEDORES': gepLimite('#fornecedoresBody tr', 9, GEP_MIN_ROWS['PLAN INTERNA FORNECEDORES']),
    'CADASTRO DE FORNECEDORES':  gepLimite('#cadastroBody tr', 3, GEP_MIN_ROWS['CADASTRO DE FORNECEDORES'])
  };

  var S = gepBuildFromTemplate(wb, rowLimits);

  var wsI = S['PLANILHA INICIAL'];
  var wsC = S['CLIENTE'];
  var wsF = S['PLANILHA DE FECHAMENTO'];
  var wsForn = S['PLAN INTERNA FORNECEDORES'];
  var wsV = S['VERBA DE PRODUÇÃO'];

  // Helpers que gravam já aplicando o deslocamento do rodapé de cada aba.
  // A partir da linha "footerFrom", a referência é traduzida para a posição
  // nova (rodapé que subiu). Antes disso, fica igual.
  function cellI(ref) { return wsI.getCell(gepRef('PLANILHA INICIAL', ref, 161)); }
  function cellC(ref) { return wsC.getCell(gepRef('CLIENTE', ref, 47)); }
  function cellF(ref) { return wsF.getCell(gepRef('PLANILHA DE FECHAMENTO', ref, 161)); }
  function cellForn(ref) { return wsForn.getCell(gepRef('PLAN INTERNA FORNECEDORES', ref, 101)); }
  // Verba não tem faixa repetida (offset 0), mas mantemos por consistência.
  function cellV(ref) { return wsV.getCell(ref); }
  var wsCad = S['CADASTRO DE FORNECEDORES'];

  // Índice do cadastro (nome->dados) para os "VLOOKUP" feitos em JS, e alguns
  // valores de cabeçalho reaproveitados no espelhamento das outras abas.
  var cadIndex = gepIndexarCadastro();
  var vNumEvento = gepMaybeNumber($('numEvento').value);
  var vCliente = $('cliente').value;
  var vNomeEvento = $('nomeEvento').value;
  var vDataIni = gepDateOrNull($('dataInicio').value);
  var vDataFim = gepDateOrNull($('dataTermino').value);
  var vHorario = $('horario').value;
  var vLocal = $('local').value;
  var vPublico = gepMaybeNumber($('publicoEstimado').value);
  var vMontagem = $('montagem').value;
  var vProdutor = $('produtor').value;

  // ---- PLANILHA INICIAL: cabeçalho ----
  if (vNumEvento !== '') wsI.getCell('B3').value = vNumEvento;
  gepSetV(wsI, 'C3', vCliente);
  gepSetV(wsI, 'D3', vNomeEvento);
  if (vDataIni) wsI.getCell('I3').value = vDataIni;
  if (vDataFim) wsI.getCell('J3').value = vDataFim;
  gepSetV(wsI, 'K3', vHorario);
  gepSetV(wsI, 'C4', vLocal);
  if (vPublico !== '') wsI.getCell('J4').value = vPublico;
  gepSetV(wsI, 'C5', vMontagem);

  // ---- PLANILHA INICIAL: itens (retorna o subtotal já somado em JS) ----
  var fimI = 160 - (GEP_FOOTER_OFFSET['PLANILHA INICIAL'] || 0);
  var subtotalServicos = gepInjectItems(wsI, document.querySelectorAll('#itensBody tr'), 'f-', cadIndex, fimI);
  cellI('H161').value = subtotalServicos; // SUBTOTAL DOS SERVIÇOS

  // ---- PLANILHA INICIAL: rodapé (Despesas/Verba, total, multiplicador) ----
  cellI('J163').value = $('verbaEstObs').value || 'Contra apresentação somente com notas fiscais';
  cellI('B164').value = 'Verba de Produção';
  var vDescr = $('verbaEstDescritivo').value; if (vDescr) cellI('C164').value = vDescr;
  var vq = gepNum($('verbaEstQtde').value);
  var vu = gepNum($('verbaEstUnitario').value);
  if (vq !== null) cellI('D164').value = vq;
  if (vu !== null) cellI('E164').value = vu;
  var vEmpresa = $('nomeEmpresa') ? $('nomeEmpresa').value : ''; if (vEmpresa) cellI('I164').value = vEmpresa;
  var verbaProd = (vq || 0) * (vu || 0);          // H164 = D164*E164
  cellI('H164').value = verbaProd;
  cellI('H165').value = verbaProd;          // SUBTOTAL VERBA = H164
  var totalEvento = subtotalServicos + verbaProd; // H167 = H161+H165
  cellI('H167').value = totalEvento;
  var margem = parseFloat($('margemLucro').value) || 0;
  var mult = 1 + margem / 100;
  cellI('G168').value = mult;
  cellI('H168').value = totalEvento * mult; // H168 = H167*G168
  var vProd = vProdutor; if (vProd) cellI('K168').value = vProd;

  // Dropdowns iguais à máscara (listas na aba CONFIGURAÇÕES / CADASTRO)
  gepAddListValidation(wsI, 'B8:B160', "'CONFIGURAÇÕES'!$C$3:$C$25");
  gepAddListValidation(wsI, 'G8:G160', "'CONFIGURAÇÕES'!$F$3:$F$10");
  gepAddListValidation(wsI, 'I8:I160', "'CADASTRO DE FORNECEDORES'!$B$3:$B$300");
  gepAddListValidation(wsI, 'C3', "'CONFIGURAÇÕES'!$E$3:$E$25");

  // ---- CLIENTE: layout novo (logo + ficha laranja + dados + tabela) ----
  // Montada inteiramente aqui, sem depender do template JSON (que mantém o
  // layout antigo da máscara para as demais abas). Esse layout foi aprovado
  // pelo Alessandro na prévia PREVIA_CLIENTE.xlsx.

  // -- Estilos reutilizáveis --
  var BRAND = 'FFFF5437';
  var brandFill = { type:'pattern', pattern:'solid', fgColor:{ argb:BRAND } };
  var whiteFill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFFFFF' } };
  var thinBd = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
  var medBd  = { top:{style:'medium'}, bottom:{style:'medium'}, left:{style:'medium'}, right:{style:'medium'} };
  var ctrAl  = { horizontal:'center', vertical:'center' };
  var leftAl = { horizontal:'left', vertical:'center' };
  var fontLbl = { name:FONT_BASE, size:10, bold:true, color:{argb:'FFFFFFFF'} };
  var fontVal = { name:FONT_BASE, size:11 };
  var fontValB = { name:FONT_BASE, size:11, bold:true };
  var fontHdr = { name:FONT_BASE, size:10, bold:true, color:{argb:'FFFFFFFF'} };
  var fontTitle = { name:FONT_BASE, size:14, bold:true };
  var CUR_FMT = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-';

  function cSet(ref, val, font, fill, align, border, numFmt) {
    var c = wsC.getCell(ref);
    if (val !== undefined && val !== null) c.value = val;
    if (font) c.font = font;
    if (fill) c.fill = fill;
    if (align) c.alignment = align;
    if (border) c.border = border;
    if (numFmt) c.numFmt = numFmt;
    return c;
  }

  // Larguras
  var cWidths = {A:4, B:24, C:28, D:12, E:14, F:9, G:16};
  Object.keys(cWidths).forEach(function(k){ wsC.getColumn(k).width = cWidths[k]; });
  wsC.views = [{ showGridLines: false }];

  // -- Logo --
  try {
    var logoId = wb.addImage({ base64: GEP_LOGO_B64, extension: 'png' });
    wsC.addImage(logoId, { tl: { col: 1.3, row: 1.1 }, ext: { width: 320, height: 143 } });
  } catch (e) { console.warn('logo:', e); }

  // -- Ficha laranja (D2:G7) --
  var fichaRows = [
    [2, 'Projeto Nº',      vNumEvento !== '' ? vNumEvento : ''],
    [3, 'Evento',           vNomeEvento],
    [4, 'Data',             vDataIni ? gepDataBR(vDataIni) : ''],
    [5, 'Local',            vLocal],
    [6, 'Horário',          vHorario],
    [7, 'Público Estimado', vPublico !== '' ? vPublico : '']
  ];
  fichaRows.forEach(function(f) {
    var r = f[0];
    cSet('D'+r, f[1], fontLbl, brandFill, ctrAl, thinBd);
    try { wsC.mergeCells('E'+r+':G'+r); } catch(e){}
    cSet('E'+r, f[2], fontVal, null, ctrAl, thinBd);
    cSet('F'+r, null, null, null, null, thinBd);
    cSet('G'+r, null, null, null, null, thinBd);
    wsC.getRow(r).height = 22;
  });

  // Separadores
  wsC.getRow(8).height = 6;
  wsC.getRow(9).height = 8;

  // -- Título --
  try { wsC.mergeCells('B10:G10'); } catch(e){}
  cSet('B10', 'PLANILHA ORÇAMENTO N.  ' + (vNumEvento !== '' ? vNumEvento : ''), fontTitle, null, leftAl);
  wsC.getRow(11).height = 8;

  // -- Dados resumidos (B12:G16) --
  var dadosResumo = [
    [12, 'NUMERO EVENTO', vNumEvento !== '' ? vNumEvento : '', null, null],
    [13, 'EVENTO:', vNomeEvento, null, null],
    [14, 'DATA EVENTO:', vDataIni ? gepDataBR(vDataIni) : '', 'HORÁRIO', vHorario],
    [15, 'LOCAL :', vLocal, null, null],
    [16, 'PAX:', vPublico !== '' ? vPublico : '', null, null]
  ];
  dadosResumo.forEach(function(d) {
    var r = d[0];
    cSet('B'+r, d[1], fontValB, null, leftAl, thinBd);
    if (d[3]) {
      try { wsC.mergeCells('C'+r+':D'+r); } catch(e){}
      cSet('C'+r, d[2], fontVal, null, ctrAl, thinBd);
      cSet('D'+r, null, null, null, null, thinBd);
      cSet('E'+r, d[3], fontValB, null, ctrAl, thinBd);
      try { wsC.mergeCells('F'+r+':G'+r); } catch(e){}
      cSet('F'+r, d[4], fontVal, null, ctrAl, thinBd);
      cSet('G'+r, null, null, null, null, thinBd);
    } else {
      try { wsC.mergeCells('C'+r+':G'+r); } catch(e){}
      cSet('C'+r, d[2], fontVal, null, ctrAl, thinBd);
      ['D','E','F','G'].forEach(function(cc){ cSet(cc+r, null, null, null, null, thinBd); });
    }
    wsC.getRow(r).height = 22;
  });

  // -- Tabela de itens (linha 18+) --
  wsC.getRow(17).height = 8;
  var cHeaders = ['SERVIÇO','DESCRIÇÃO DO SERVIÇO','Qtde','R$ Unitário','Freq.','Valor final'];
  var cCols = ['B','C','D','E','F','G'];
  cCols.forEach(function(col, i) {
    cSet(col+'18', cHeaders[i], fontHdr, brandFill, ctrAl, { top:{style:'medium'}, bottom:{style:'medium'}, left:{style:'thin'}, right:{style:'thin'} });
  });
  wsC.getRow(18).height = 24;

  // Itens do formulário
  var cRows = document.querySelectorAll('#clienteItensBody tr');
  var firstC = 19, totalCliente = 0;
  var minLinhasC = 15;
  var limC = Math.max(minLinhasC, cRows.length + 3);
  for (var ci = 0; ci < cRows.length && ci < limC; ci++) {
    var tr = cRows[ci], rc = firstC + ci;
    gepSetV(wsC, 'B'+rc, tr.querySelector('.cf-servico').value);
    gepSetV(wsC, 'C'+rc, tr.querySelector('.cf-desc').value);
    var q = gepNum(tr.querySelector('.cf-qtde').value);
    var p = gepNum(tr.querySelector('.cf-preco').value);
    if (q !== null) cSet('D'+rc, q, fontVal, null, ctrAl, thinBd);
    if (p !== null) cSet('E'+rc, p, fontVal, null, ctrAl, thinBd, CUR_FMT);
    var fq = gepNum(tr.querySelector('.cf-freq').value);
    var freq = (fq !== null && fq !== 0) ? fq : 1;
    cSet('F'+rc, freq, fontVal, null, ctrAl, thinBd);
    var g = (q || 0) * (p || 0) * freq;
    if (q !== null || p !== null) { cSet('G'+rc, g, fontVal, null, ctrAl, thinBd, CUR_FMT); totalCliente += g; }
    cCols.forEach(function(col) { cSet(col+rc, undefined, fontVal, null, ctrAl, thinBd); });
    wsC.getRow(rc).height = 22;
  }
  // Linhas em branco restantes
  for (var rb = firstC + cRows.length; rb < firstC + limC; rb++) {
    cCols.forEach(function(col) { cSet(col+rb, col==='F'?1:undefined, fontVal, null, ctrAl, thinBd); });
    wsC.getRow(rb).height = 22;
  }

  // VALOR TOTAL
  var rTotal = firstC + limC + 1;
  wsC.getRow(firstC + limC).height = 8;
  try { wsC.mergeCells('B'+rTotal+':F'+rTotal); } catch(e){}
  cSet('B'+rTotal, 'VALOR TOTAL EVENTO', { name:FONT_BASE, size:12, bold:true, color:{argb:'FFFFFFFF'} }, brandFill, ctrAl, medBd);
  ['C','D','E','F'].forEach(function(cc){ cSet(cc+rTotal, null, null, null, null, medBd); });
  cSet('G'+rTotal, totalCliente, { name:FONT_BASE, size:12, bold:true }, null, ctrAl, medBd, CUR_FMT);
  wsC.getRow(rTotal).height = 28;

  // Observações
  var rObs = rTotal + 2;
  wsC.getRow(rTotal+1).height = 8;
  cSet('B'+rObs, 'OBSERVAÇÕES: ' + ($('observacoes').value || ''), { name:FONT_BASE, size:10, bold:true });
  cSet('B'+(rObs+1), 'PRAZO DE VALIDADE DO ORÇAMENTO: ' + ($('prazoValidade').value || ''), { name:FONT_BASE, size:10 });

  wsC.pageSetup.orientation = 'portrait';
  wsC.pageSetup.fitToPage = true;
  wsC.pageSetup.fitToWidth = 1;
  wsC.pageSetup.fitToHeight = 0;
  wsC.pageSetup.margins = {left:0.51,right:0.51,top:0.79,bottom:0.79,header:0.31,footer:0.31};
  wsC.pageSetup.printArea = 'A1:G' + (rObs+1);

  // ---- PLANILHA DE FECHAMENTO ----
  // Cabeçalho espelha a Inicial (mesmos valores). Itens vêm do fechamento.
  if (vNumEvento !== '') wsF.getCell('B3').value = vNumEvento;
  gepSetV(wsF, 'C3', vCliente);
  gepSetV(wsF, 'D3', vNomeEvento);
  if (vDataIni) wsF.getCell('I3').value = vDataIni;
  if (vDataFim) wsF.getCell('J3').value = vDataFim;
  gepSetV(wsF, 'K3', vHorario);
  gepSetV(wsF, 'C4', vLocal);
  if (vPublico !== '') wsF.getCell('J4').value = vPublico;
  gepSetV(wsF, 'C5', vMontagem);
  var fechTrs = document.querySelectorAll('#fechamentoItensBody tr');
  var fimF = 160 - (GEP_FOOTER_OFFSET['PLANILHA DE FECHAMENTO'] || 0);
  var subtotalFech = gepInjectItems(wsF, fechTrs, 'f-', cadIndex, fimF);
  cellF('H161').value = subtotalFech;   // SUBTOTAL DOS SERVIÇOS (fech)
  if (vProdutor) cellF('J164').value = vProdutor;   // Produtor (era link p/ K168)
  gepAddListValidation(wsF, 'B8:B160', "'CONFIGURAÇÕES'!$C$3:$C$25");
  gepAddListValidation(wsF, 'G8:G160', "'CONFIGURAÇÕES'!$F$3:$F$10");
  gepAddListValidation(wsF, 'I8:I160', "'CADASTRO DE FORNECEDORES'!$B$3:$B$300");

  // Soma dos valores do fechamento por fornecedor (equivale ao SUMIF da aba
  // Fornecedores). Percorre os itens do fechamento uma vez.
  var somaPorFornecedor = {};
  for (var fi = 0; fi < fechTrs.length; fi++) {
    var nomeF = (fechTrs[fi].querySelector('.f-fornecedor').value || '').trim();
    if (!nomeF) continue;
    var qF = gepNum(fechTrs[fi].querySelector('.f-qtde').value) || 0;
    var pF = gepNum(fechTrs[fi].querySelector('.f-preco').value) || 0;
    var fqF = gepNum(fechTrs[fi].querySelector('.f-freq').value);
    var frF = (fqF !== null && fqF !== 0) ? fqF : 1;
    somaPorFornecedor[nomeF] = (somaPorFornecedor[nomeF] || 0) + qF * pF * frF;
  }

  // ---- PLAN INTERNA FORNECEDORES ----
  // Cabeçalho espelhado + por linha: CNPJ/contato/telefone (cadastro),
  // valor final (soma do fechamento), forma (X) e vencimento (data início +
  // prazo, quando Faturado) — tudo calculado em JS e gravado como valor.
  gepSetV(wsForn, 'B3', vCliente);
  if (vNumEvento !== '') wsForn.getCell('D3').value = vNumEvento;
  gepSetV(wsForn, 'E3', vNomeEvento);
  if (vDataIni) wsForn.getCell('J3').value = vDataIni;
  if (vDataFim) wsForn.getCell('K3').value = vDataFim;
  gepSetV(wsForn, 'L3', vHorario);
  gepSetV(wsForn, 'C4', vLocal);
  gepSetV(wsForn, 'C5', vProdutor);

  var fRows = document.querySelectorAll('#fornecedoresBody tr');
  var firstF = 9, lastF = 100, totalFornecedores = 0;
  for (var j = 0; j < fRows.length && (firstF + j) <= lastF; j++) {
    var trF = fRows[j], rF = firstF + j;
    var nome = (trF.querySelector('.ff-nome').value || '').trim();
    gepSetV(wsForn, 'B'+rF, nome);
    var checked = trF.querySelector('.ff-forma-check:checked');
    var forma = checked ? checked.value : '';
    if (forma === 'faturado') wsForn.getCell('G'+rF).value = 'X';
    if (forma === 'parcial')  wsForn.getCell('H'+rF).value = 'X';
    if (forma === 'avista')   wsForn.getCell('I'+rF).value = 'X';
    gepSetV(wsForn, 'K'+rF, trF.querySelector('.ff-obs').value);
    // CNPJ / contato / telefone (VLOOKUP no cadastro)
    var cad = cadIndex[nome];
    if (cad) {
      gepSetV(wsForn, 'C'+rF, cad.cnpj);
      gepSetV(wsForn, 'D'+rF, cad.contato);
      gepSetV(wsForn, 'E'+rF, cad.telefone);
    }
    // Valor final / fornecedor (SUMIF do fechamento)
    var valF = somaPorFornecedor[nome] || 0;
    if (nome) { wsForn.getCell('F'+rF).value = valF; totalFornecedores += valF; }
    // Vencimento: Faturado -> data início + prazo; Parcial/À vista -> texto
    if (forma === 'faturado' && vDataIni && cad) {
      var venc = new Date(vDataIni.getTime());
      venc.setUTCDate(venc.getUTCDate() + gepPrazoDias(cad.prazo));
      wsForn.getCell('J'+rF).value = venc;
    } else if (forma === 'parcial' || forma === 'avista') {
      wsForn.getCell('J'+rF).value = 'Condições em OBS';
    }
  }
  // ---- FORNECEDORES: rodapé melhorado (aprovado na prévia) ----
  // Montado inteiramente pelo JS com 3 blocos visuais: Total, Produtor e Conferência.

  var AZUL_FILL = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1F4E79'} };
  var CINZA_FILL = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
  var VERDE_FILL = { type:'pattern', pattern:'solid', fgColor:{argb:'FFC6EFCE'} };
  var VERM_FILL  = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFC7CE'} };
  var fBox  = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
  var fBoxM = { top:{style:'medium'}, bottom:{style:'medium'}, left:{style:'medium'}, right:{style:'medium'} };
  var fCtr  = { horizontal:'center', vertical:'center', wrapText:true };
  var fLeft = { horizontal:'left', vertical:'center' };
  var fBranco = { name:FONT_BASE, size:12, bold:true, color:{argb:'FFFFFFFF'} };
  var fBold  = { name:FONT_BASE, size:11, bold:true };
  var fNorm  = { name:FONT_BASE, size:11 };
  var fSmB   = { name:FONT_BASE, size:10, bold:true };
  var fSm    = { name:FONT_BASE, size:10 };
  var fVerm  = { name:FONT_BASE, size:11, bold:true, color:{argb:'FFFF0000'} };
  var fVerde = { name:FONT_BASE, size:11, bold:true, color:{argb:'FF006100'} };

  function fSet(ref, val, font, fill, align, border, numFmt) {
    var c = wsForn.getCell(ref);
    if (val !== undefined && val !== null) c.value = val;
    if (font) c.font = font;
    if (fill) c.fill = fill;
    if (align) c.alignment = align;
    if (border) c.border = border;
    if (numFmt) c.numFmt = numFmt;
  }

  // Calcular a linha inicial do rodapé: logo após a última linha formatada
  var rFoot = 9 + Math.max(15, fRows.length + 3) + 1; // +1 linha de espaço

  // BLOCO 1: TOTAL DO EVENTO (barra azul)
  var rT = rFoot;
  try { wsForn.mergeCells('B'+rT+':E'+rT); } catch(e){}
  fSet('B'+rT, 'TOTAL DO EVENTO', fBranco, AZUL_FILL, fCtr, fBoxM);
  ['C','D','E'].forEach(function(cc){ fSet(cc+rT, null, null, null, null, fBoxM); });
  fSet('F'+rT, totalFornecedores, fBranco, AZUL_FILL, fCtr, fBoxM, CUR_FMT);
  wsForn.getRow(rT).height = 30;

  // BLOCO 2: PRODUTOR (fundo cinza claro)
  var rP = rT + 2;
  fSet('B'+rP, 'Produtor extra?', fSmB, CINZA_FILL, fLeft, fBox);
  try { wsForn.mergeCells('C'+rP+':F'+rP); } catch(e){}
  fSet('C'+rP, '(   ) SIM   ou   ( X ) NÃO', fSm, CINZA_FILL, fCtr, fBox);
  ['D','E','F'].forEach(function(cc){ fSet(cc+rP, null, null, CINZA_FILL, fCtr, fBox); });
  wsForn.getRow(rP).height = 24;

  var rP2 = rP + 1;
  fSet('B'+rP2, 'Produtor responsável:', fSmB, CINZA_FILL, fLeft, fBox);
  try { wsForn.mergeCells('C'+rP2+':F'+rP2); } catch(e){}
  fSet('C'+rP2, vProdutor || '', fVerm, CINZA_FILL, fCtr, fBox);
  ['D','E','F'].forEach(function(cc){ fSet(cc+rP2, null, null, CINZA_FILL, fCtr, fBox); });
  wsForn.getRow(rP2).height = 24;

  // BLOCO 3: CONFERÊNCIA DE VALORES
  var rC = rP2 + 2;
  // Título
  try { wsForn.mergeCells('B'+rC+':F'+rC); } catch(e){}
  fSet('B'+rC, 'CONFERÊNCIA DE VALORES', { name:FONT_BASE, size:11, bold:true, color:{argb:'FF1F4E79'} }, null, fLeft,
    { bottom:{style:'medium', color:{argb:'FF1F4E79'}} });
  wsForn.getRow(rC).height = 26;

  // Planilha de Fechamento
  var rC1 = rC + 1;
  fSet('B'+rC1, 'Planilha de Fechamento', fSmB, null, fLeft, fBox);
  try { wsForn.mergeCells('C'+rC1+':F'+rC1); } catch(e){}
  fSet('C'+rC1, subtotalFech, fBold, null, fCtr, fBox, CUR_FMT);
  ['D','E','F'].forEach(function(cc){ fSet(cc+rC1, null, null, null, fCtr, fBox); });
  wsForn.getRow(rC1).height = 24;

  // Planilha Interna Fornecedores
  var rC2 = rC1 + 1;
  fSet('B'+rC2, 'Planilha Interna Fornecedores', fSmB, null, fLeft, fBox);
  try { wsForn.mergeCells('C'+rC2+':F'+rC2); } catch(e){}
  fSet('C'+rC2, totalFornecedores, fBold, null, fCtr, fBox, CUR_FMT);
  ['D','E','F'].forEach(function(cc){ fSet(cc+rC2, null, null, null, fCtr, fBox); });
  wsForn.getRow(rC2).height = 24;

  // Status: confere ou não?
  var rC3 = rC2 + 1;
  var confere = (subtotalFech === totalFornecedores);
  fSet('B'+rC3, 'Status:', fSmB, null, fLeft, fBox);
  try { wsForn.mergeCells('C'+rC3+':F'+rC3); } catch(e){}
  if (confere) {
    fSet('C'+rC3, '✅ VALORES CONFEREM', fVerde, VERDE_FILL, fCtr, fBox);
  } else {
    fSet('C'+rC3, '⚠️ VALORES NÃO CONFEREM', fVerm, VERM_FILL, fCtr, fBox);
  }
  ['D','E','F'].forEach(function(cc){ fSet(cc+rC3, null, null, confere ? VERDE_FILL : VERM_FILL, fCtr, fBox); });
  wsForn.getRow(rC3).height = 28;

  // Diferença
  var rC4 = rC3 + 1;
  var diff = subtotalFech - totalFornecedores;
  fSet('B'+rC4, 'Diferença:', fSmB, null, fLeft, fBox);
  try { wsForn.mergeCells('C'+rC4+':F'+rC4); } catch(e){}
  fSet('C'+rC4, diff, diff !== 0 ? fVerm : fVerde, null, fCtr, fBox, CUR_FMT);
  ['D','E','F'].forEach(function(cc){ fSet(cc+rC4, null, null, null, fCtr, fBox); });
  wsForn.getRow(rC4).height = 24;

  // Atualizar área de impressão
  wsForn.pageSetup.printArea = 'B1:L' + (rC4 + 1);
  gepAddListValidation(wsForn, 'B9:B100', "'PLANILHA DE FECHAMENTO'!$I$8:$I$160");

  // ---- VERBA DE PRODUÇÃO — novo layout aprovado ----
  // Montado inteiramente pelo JS: bordas fechadas, coluna Produtor, tabela
  // limpa sem blocos separados, total único no rodapé.
  wsV.views = [{ showGridLines: false }];

  // Larguras
  var vWidths = {A:3, B:20, C:35, D:16, E:14};
  Object.keys(vWidths).forEach(function(k){ wsV.getColumn(k).width = vWidths[k]; });

  var AZUL_V = 'FF1F4E79';
  var CINZA_V = 'FFBFBFBF';
  var CINZA2_V = 'FFF2F2F2';
  var azulV_fill = { type:'pattern', pattern:'solid', fgColor:{argb:AZUL_V} };
  var cinzaV_fill = { type:'pattern', pattern:'solid', fgColor:{argb:CINZA_V} };
  var cinza2V_fill = { type:'pattern', pattern:'solid', fgColor:{argb:CINZA2_V} };
  var vBox  = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
  var vBoxM = { top:{style:'medium'}, bottom:{style:'medium'}, left:{style:'medium'}, right:{style:'medium'} };
  var vCtr   = { horizontal:'center', vertical:'center', wrapText:false };
  var vLeft  = { horizontal:'left',   vertical:'center', wrapText:false };
  var vRight = { horizontal:'right',  vertical:'center', wrapText:false };
  var WHITE_V = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFFFFF'} };
  var CUR_FMT = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-';

  function vSet(ref, val, font, fill, align, border, numFmt) {
    var c = wsV.getCell(ref);
    if (val !== undefined && val !== null) c.value = val;
    if (font) c.font = font;
    if (fill) c.fill = fill;
    if (align) c.alignment = align;
    if (border) c.border = border;
    if (numFmt) c.numFmt = numFmt;
  }
  function vMerge(range) { try { wsV.mergeCells(range); } catch(e){} }
  var vFontWhiteB = { name:FONT_BASE, size:16, bold:true, color:{argb:'FFFFFFFF'} };
  var vFontHdr    = { name:FONT_BASE, size:10, bold:true, color:{argb:'FFFFFFFF'} };
  var vFontLbl    = { name:FONT_BASE, size:10, bold:true };
  var vFontVal    = { name:FONT_BASE, size:11 };
  var vFontTot    = { name:FONT_BASE, size:12, bold:true };

  // TÍTULO
  vMerge('B1:E1');
  vSet('B1', 'VERBA DE PRODUÇÃO', vFontWhiteB, azulV_fill, vCtr, vBoxM);
  ['C1','D1','E1'].forEach(function(c){ vSet(c, null, null, azulV_fill, null, vBoxM); });
  wsV.getRow(1).height = 40;

  // CABEÇALHO DE DADOS
  var vDados = [
    ['B2','Nº DO EVENTO',   'C2:E2', vNumEvento !== '' ? vNumEvento : '', true],
    ['B3','DATA DO EVENTO:','C3:E3', vDataIni ? (vDataFim ? gepDataBR(vDataIni)+' a '+gepDataBR(vDataFim) : gepDataBR(vDataIni)) : '', false],
    ['B4','NOME DO EVENTO:','C4:E4', vNomeEvento, true],
    ['B5','PRODUTOR:',      'C5:E5', vProdutor, false]
  ];
  vDados.forEach(function(d, i) {
    var r = i + 2;
    var altFill = d[4] ? cinzaV_fill : WHITE_V;
    vSet(d[0], d[1], vFontLbl, altFill, vLeft, {top:{style:'thin'},bottom:{style:'thin'},left:{style:'medium'},right:{style:'thin'}});
    vMerge(d[2]);
    var lbl = d[2].split(':')[0];
    vSet(lbl, d[3], d[0]==='B5' ? { name:FONT_BASE, size:11, color:{argb:'FF0000FF'} } : vFontVal, altFill, vLeft,
      {top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'medium'}});
    var others = d[2].split(':')[0]==='C' ? ['D','E'] : [];
    others.forEach(function(cc){ vSet(cc+r, null, null, altFill, null, {top:{style:'thin'},bottom:{style:'thin'},right:{style:'medium'}}); });
    wsV.getRow(r).height = 24;
  });
  wsV.getRow(6).height = 8;

  // CABEÇALHO DA TABELA (com coluna Produtor)
  var vHdrs = [['B7','Produtor'],['C7','Item / Descrição'],['D7','Valor (R$)'],['E7','OBS']];
  vHdrs.forEach(function(h, i) {
    var lBd = i===0 ? {style:'medium'} : {style:'thin'};
    var rBd = i===3 ? {style:'medium'} : {style:'thin'};
    vSet(h[0], h[1], vFontHdr, azulV_fill, vCtr, {top:{style:'medium'},bottom:{style:'medium'},left:lBd,right:rBd});
  });
  wsV.getRow(7).height = 28;

  // LINHAS DE ITENS
  var vRows = document.querySelectorAll('#verbaBody tr');
  var minV = 8, totalVerba = 0;
  var limV = Math.max(minV, vRows.length + 3);
  var firstV = 8;
  for (var k = 0; k < limV; k++) {
    var rv = firstV + k;
    var altFill2 = (k % 2 === 1) ? cinza2V_fill : WHITE_V;
    var trV = vRows[k] || null;
    var prod = trV ? trV.querySelector('.vb-produtor').value : '';
    var item = trV ? trV.querySelector('.vb-item').value : '';
    var vv   = trV ? (gepNum(trV.querySelector('.vb-valor').value) || null) : null;
    var obs  = trV ? trV.querySelector('.vb-obs').value : '';
    if (vv !== null) totalVerba += vv;
    vSet('B'+rv, prod||'', vFontVal, altFill2, vCtr, {top:{style:'thin'},bottom:{style:'thin'},left:{style:'medium'},right:{style:'thin'}});
    vSet('C'+rv, item||'', vFontVal, altFill2, vLeft, vBox);
    vSet('D'+rv, vv, vFontVal, altFill2, vRight, vBox, vv!==null?CUR_FMT:null);
    vSet('E'+rv, obs||'', vFontVal, altFill2, vLeft, {top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'medium'}});
    wsV.getRow(rv).height = 22;
  }

  // TOTAL GASTO
  var rTotV = firstV + limV + 1;
  wsV.getRow(firstV + limV).height = 8;
  vSet('B'+rTotV, 'TOTAL GASTO', vFontTot, cinzaV_fill, vCtr,
    {top:{style:'medium'},bottom:{style:'medium'},left:{style:'medium'},right:{style:'thin'}});
  vSet('C'+rTotV, totalVerba, vFontTot, cinzaV_fill, vRight,
    {top:{style:'medium'},bottom:{style:'medium'},left:{style:'thin'},right:{style:'thin'}}, CUR_FMT);
  vSet('D'+rTotV, null, null, cinzaV_fill, null,
    {top:{style:'medium'},bottom:{style:'medium'},left:{style:'thin'},right:{style:'thin'}});
  vSet('E'+rTotV, null, null, cinzaV_fill, null,
    {top:{style:'medium'},bottom:{style:'medium'},left:{style:'thin'},right:{style:'medium'}});
  wsV.getRow(rTotV).height = 30;

  // DATA ENTREGA / REEMBOLSO
  wsV.getRow(rTotV+1).height = 8;
  [['DATA DA ENTREGA:', rTotV+2], ['DATA DO REEMBOLSO:', rTotV+3]].forEach(function(d) {
    vSet('B'+d[1], d[0], vFontLbl, null, vCtr,
      {top:{style:'thin'},bottom:{style:'thin'},left:{style:'medium'},right:{style:'thin'}});
    vMerge('C'+d[1]+':E'+d[1]);
    vSet('C'+d[1], null, null, null, null,
      {top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'medium'}});
    ['D','E'].forEach(function(cc){ vSet(cc+d[1], null, null, null, null,
      {top:{style:'thin'},bottom:{style:'thin'},right:{style:'medium'}}); });
    wsV.getRow(d[1]).height = 24;
  });
  vSet('B'+(rTotV+4), 'OBS:', vFontLbl);

  wsV.pageSetup.orientation = 'portrait';
  wsV.pageSetup.fitToPage = true;
  wsV.pageSetup.fitToWidth = 1;
  wsV.pageSetup.fitToHeight = 0;
  wsV.pageSetup.margins = {left:0.51, right:0.51, top:0.79, bottom:0.79, header:0.31, footer:0.31};
  wsV.pageSetup.printArea = 'B1:E' + (rTotV + 4);

  // ---- CADASTRO DE FORNECEDORES ----
  var cadRows = document.querySelectorAll('#cadastroBody tr');
  for (var m = 0; m < cadRows.length && m < 298; m++) {
    var trC2 = cadRows[m], rC = 3 + m;
    gepSetV(wsCad, 'B'+rC, trC2.querySelector('.cad-nome').value);
    gepSetV(wsCad, 'C'+rC, trC2.querySelector('.cad-cnpj').value);
    gepSetV(wsCad, 'D'+rC, trC2.querySelector('.cad-contato').value);
    gepSetV(wsCad, 'E'+rC, trC2.querySelector('.cad-telefone').value);
    gepSetV(wsCad, 'F'+rC, trC2.querySelector('.cad-prazo').value);
    gepSetV(wsCad, 'G'+rC, trC2.querySelector('.cad-servico').value);
  }
  gepAddListValidation(wsCad, 'F3:F300', '"À VISTA,15D,30D,45D,60D,90D"');

  return wb.xlsx.writeBuffer().then(function(buffer) {
    var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    var nEvt  = ($('numEvento')  ? $('numEvento').value.trim()  : '') || 'evento';
    var nNome = ($('nomeEvento') ? $('nomeEvento').value.trim() : '');
    var nomeArq = nEvt + (nNome ? ' - ' + nNome : '');
    a.href = url;
    a.download = nomeArq + '.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });
}

