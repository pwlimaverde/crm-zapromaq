/* previa.js - desenha uma tela do CRM a partir de fonte\layout (via dados-layout.js).
   Sem dependências e sem internet: roda no Edge que já vem no Windows.
   Unidades: pt (1 pt do UserForm = 1 pt aqui). */
(function () {
  'use strict';
  var D = window.CRM_LAYOUT;
  var q = new URLSearchParams(location.search);
  var raiz = document.getElementById('tela');
  var barra = document.getElementById('barra');
  var saida = document.getElementById('relatorio');
  if (!D) {
    barra.textContent = 'dados-layout.js não encontrado: rode build\\GERAR-VISUAL.bat (gerar-visual.ps1).';
    return;
  }
  var nomes = Object.keys(D.telas);
  var nomeTela = q.get('tela') || nomes[0];
  var T = D.telas[nomeTela];
  var modo = q.get('modo') || 'completo';
  document.body.classList.add('modo-' + modo);
  if (q.get('captura')) document.body.classList.add('captura');
  if (q.get('contornos')) document.body.classList.add('contornos');

  // ------------------------------------------------------------ tema
  function valor(v) {
    if (typeof v !== 'string' || v.charAt(0) !== '@') return v;
    var n = v.substring(1);
    if (D.tema.cores[n]) return D.tema.cores[n];
    if (D.tema.fontes[n]) return D.tema.fontes[n];
    throw new Error('valor de tema desconhecido: ' + v);
  }
  function glifo(nome) {
    var i = D.icones[nome];
    if (!i) throw new Error('ícone desconhecido: ' + nome);
    return String.fromCharCode(parseInt(i[0], 16));
  }
  function hexRgba(hex, a) {
    var h = hex.replace('#', '');
    return 'rgba(' + parseInt(h.substr(0, 2), 16) + ',' + parseInt(h.substr(2, 2), 16) + ',' + parseInt(h.substr(4, 2), 16) + ',' + a + ')';
  }
  function caixa(pos, classe, nome) {
    var d = document.createElement('div');
    d.className = 'el ' + (classe || '');
    d.style.left = pos[0] + 'pt'; d.style.top = pos[1] + 'pt';
    d.style.width = pos[2] + 'pt'; d.style.height = pos[3] + 'pt';
    if (nome) d.setAttribute('data-nome', nome);
    raiz.appendChild(d);
    return d;
  }

  raiz.style.width = T.canvas.largura + 'pt';
  raiz.style.height = T.canvas.altura + 'pt';
  raiz.style.background = valor((T.propriedades && T.propriedades.BackColor) || '@fundo');

  // ------------------------------------------------------------ fundo (vira o BMP)
  (T.fundo || []).forEach(function (f) {
    var d;
    if (f.tipo === 'retangulo') {
      d = caixa(f.pos, '');
      if (f.gradiente) d.style.background = 'linear-gradient(90deg,' + valor(f.gradiente[0]) + ',' + valor(f.gradiente[1]) + ')';
      else d.style.background = valor(f.cor);
      if (f.raio) d.style.borderRadius = f.raio + 'pt';
      if (f.borda) d.style.border = '1px solid ' + valor(f.borda);
      if (f.sombra) d.style.boxShadow = '0 1.5pt 6pt ' + hexRgba(valor('@azul'), 0.10);
    } else if (f.tipo === 'linha') {
      d = caixa(f.pos, '');
      d.style.background = valor(f.cor);
      if (f.opacidade !== undefined) d.style.opacity = f.opacidade;
    } else if (f.tipo === 'texto') {
      d = caixa(f.pos, 'txt');
      d.textContent = f.texto;
      d.style.color = valor(f.cor || '@texto');
      d.style.font = (f.peso || 400) + ' ' + (f.tamanho || 9) + 'pt "' + valor('@padrao') + '"';
      d.style.lineHeight = f.pos[3] + 'pt';
    } else if (f.tipo === 'imagem') {
      d = caixa(f.pos, '');
      var img = document.createElement('img');
      img.src = '../../fonte/' + f.arquivo;
      img.style.width = '100%'; img.style.height = '100%'; img.style.objectFit = 'contain'; img.style.objectPosition = 'left center';
      if (f.filtro === 'branco') img.style.filter = 'brightness(0) invert(1)';
      d.appendChild(img);
    } else if (f.tipo === 'icone') {
      d = caixa(f.pos, 'ico');
      d.textContent = glifo(f.icone);
      d.style.color = valor(f.cor || '@azul');
      d.style.fontSize = (f.tamanho || 12) + 'pt';
    } else {
      throw new Error('tipo de fundo desconhecido: ' + f.tipo);
    }
  });

  var itens = [];   // o que entra na conferência (controles e componentes)
  if (modo !== 'fundo') {
    // ---------------------------------------------------------- controles nativos
    (T.controles || []).forEach(function (c) {
      var p = {};
      [].concat(c.estilo || []).forEach(function (e) { Object.assign(p, (T.estilos || {})[e] || {}); });
      Object.assign(p, c.props || {});
      var d = caixa(c.pos, 'ctl c-' + c.tipo + (p.MultiLine ? ' multi' : ''), c.nome);
      var texto = c.texto || c.amostra || '';
      if (c.tipo === 'Frame') {
        var cap = document.createElement('span'); cap.className = 'cap'; cap.textContent = c.texto || ''; d.appendChild(cap);
        if (c.amostra === 'ficha') desenharFicha(d, c);
      } else if (c.tipo === 'ListBox') {
        d.textContent = c.amostra || '';
      } else if (c.tipo !== 'ScrollBar') {
        d.textContent = texto;
      }
      if (p['Font.Name']) d.style.fontFamily = '"' + valor(p['Font.Name']) + '"';
      if (p['Font.Size']) d.style.fontSize = p['Font.Size'] + 'pt';
      if (p['Font.Bold']) d.style.fontWeight = 600;
      if (p.ForeColor) d.style.color = valor(p.ForeColor);
      if (p.BackStyle === 0) d.style.background = 'transparent';
      else if (p.BackColor) d.style.background = valor(p.BackColor);
      if (p.BorderColor) d.style.borderColor = valor(p.BorderColor);
      if (p.BorderStyle === 0) d.style.border = 'none';
      if (c.tipo === 'Frame') d.querySelector('.cap').style.background = d.style.background || valor('@fundo');
      itens.push({ nome: c.nome, pos: c.pos, texto: (c.tipo === 'Label' || c.tipo === 'CheckBox') ? texto : '', fonte: p,
                   dentro: c.dentro, permiteSobrepor: c.permiteSobrepor, el: d, tipo: c.tipo });
    });

    // ---------------------------------------------------------- componentes
    (T.componentes || []).forEach(function (c) {
      var d;
      if (c.tipo === 'botao') {
        var v = D.tema.botoes[c.variante || 'primario'];
        d = caixa(c.pos, 'ctl botao', c.nome);
        d.style.background = c.desabilitado ? valor('@desabilitado') : v.fundo;
        d.style.color = c.desabilitado ? valor('@desabilitadoTexto') : v.texto;
        if (v.borda) d.style.border = '1px solid ' + v.borda;
        if (c.icone) { var ic = document.createElement('span'); ic.className = 'ico'; ic.textContent = glifo(c.icone); d.appendChild(ic); }
        if (c.texto) { var tx = document.createElement('span'); tx.textContent = c.texto; d.appendChild(tx); }
        itens.push({ nome: c.nome, pos: c.pos, texto: c.texto || '', fonte: { 'Font.Size': 9, 'Font.Bold': true }, extra: c.icone ? 20 : 0,
                     dentro: c.dentro, permiteSobrepor: c.permiteSobrepor, el: d, tipo: 'botao' });
      } else if (c.tipo === 'icone') {
        d = caixa(c.pos, 'ctl ico', c.nome);
        d.textContent = glifo(c.icone);
        d.style.color = valor(c.cor || '@azul');
        d.style.fontSize = (c.tamanho || 12) + 'pt';
        itens.push({ nome: c.nome, pos: c.pos, dentro: c.dentro, permiteSobrepor: c.permiteSobrepor, el: d, tipo: 'icone' });
      } else if (c.tipo === 'selo') {
        var s = D.tema.situacoes[c.situacao] || { texto: valor('@azul'), fundo: valor('@azulClaro') };
        d = caixa(c.pos, 'ctl selo', c.nome);
        d.style.background = s.fundo; d.style.color = s.texto;
        d.textContent = c.texto || c.situacao || '';
        itens.push({ nome: c.nome, pos: c.pos, texto: d.textContent, fonte: { 'Font.Size': 8, 'Font.Bold': true },
                     dentro: c.dentro, permiteSobrepor: c.permiteSobrepor, el: d, tipo: 'selo' });
      } else if (c.tipo === 'grade') {
        d = caixa(c.pos, 'ctl grade', c.nome);
        desenharGrade(d, c);
        itens.push({ nome: c.nome, pos: c.pos, dentro: c.dentro, permiteSobrepor: c.permiteSobrepor, el: d, tipo: 'grade' });
      } else {
        throw new Error('componente desconhecido: ' + c.tipo);
      }
    });
    conferir();
  }

  // Ficha de exemplo dentro do Frame, com a MESMA geometria de
  // frmCRM.MontarCampos (que cria os campos em execução a partir de
  // modSchema): mudou lá, muda aqui.
  function desenharFicha(d, c) {
    var F = D.amostras.ficha || { secoes: [] };
    var util = c.pos[2] - 18, X0 = 4, ROT = 86, PASSO = 24;
    var meia = Math.floor(util / 2), CAMPO = meia - ROT - 12, X1 = X0 + meia, LARGO = util - ROT - 4;
    var y = 6;
    d.style.overflow = 'hidden';
    function el(cls, x, yy, w, h, txt) {
      var e = document.createElement('div'); e.className = 'ctl ' + cls; e.style.position = 'absolute';
      e.style.left = x + 'pt'; e.style.top = yy + 'pt'; e.style.width = w + 'pt'; e.style.height = h + 'pt';
      if (txt !== undefined) e.textContent = txt;
      d.appendChild(e); return e;
    }
    F.secoes.forEach(function (sec, k) {
      if (k > 0) y += 14;
      var t = el('c-Label', X0, y, util, 15, sec.titulo);
      t.style.color = valor('@azul'); t.style.fontWeight = 600; t.style.background = 'transparent';
      var l = el('', X0, y + 16, util, 1); l.style.background = valor('@bordaSuave'); l.style.border = 'none';
      y += 22;
      sec.campos.forEach(function (r) {
        var lar = !r[2];
        var a = el('c-Label', X0, y + 3, ROT, 14, r[0]);
        a.style.color = valor('@texto2'); a.style.background = 'transparent';
        var ca = el('c-TextBox', X0 + ROT + 4, y, lar ? LARGO : CAMPO, r[4] ? 54 : 18, r[1]);
        ca.style.background = valor('@zebra'); ca.style.borderColor = valor('@borda'); ca.style.color = valor('@texto');
        if (!lar) {
          var b = el('c-Label', X1, y + 3, ROT, 14, r[2]);
          b.style.color = valor('@texto2'); b.style.background = 'transparent';
          var cb = el('c-TextBox', X1 + ROT + 4, y, CAMPO, 18, r[3]);
          cb.style.background = valor('@zebra'); cb.style.borderColor = valor('@borda'); cb.style.color = valor('@texto');
        }
        y += r[4] ? 60 : PASSO;
      });
    });
  }

  // Grade própria (substitui a ListBox): cabeçalho azul, linhas alternadas,
  // 1ª linha selecionada, coluna de situação com ponto colorido.
  function desenharGrade(d, c) {
    var A = D.amostras[c.amostra] || { colunas: [], linhas: [] };
    var ALT_CAB = 18, ALT_LIN = 16, ROLAGEM = 12;
    var largUtil = c.pos[2] - ROLAGEM - 2;
    var soma = A.colunas.reduce(function (s, k) { return s + k.largura; }, 0) || 1;
    var fator = Math.min(1, largUtil / soma);
    var cab = document.createElement('div'); cab.className = 'cab';
    cab.style.top = 0; cab.style.height = ALT_CAB + 'pt'; cab.style.width = largUtil + 'pt';
    cab.style.background = valor('@azul'); cab.style.color = valor('@branco');
    A.colunas.forEach(function (k) {
      var e = document.createElement('div'); e.className = 'cel'; e.style.width = (k.largura * fator) + 'pt';
      e.textContent = k.titulo; if (k.alinhamento === 'direita') e.style.textAlign = 'right';
      cab.appendChild(e);
    });
    d.appendChild(cab);
    var n = Math.min(A.linhas.length, c.linhas || Math.floor((c.pos[3] - ALT_CAB) / ALT_LIN));
    for (var i = 0; i < n; i++) {
      var lin = document.createElement('div'); lin.className = 'lin';
      lin.style.top = (ALT_CAB + i * ALT_LIN) + 'pt'; lin.style.height = ALT_LIN + 'pt'; lin.style.width = largUtil + 'pt';
      lin.style.background = i === 0 ? valor('@selecao') : (i % 2 ? valor('@zebra') : valor('@branco'));
      lin.style.color = i === 0 ? valor('@selecaoTexto') : valor('@texto');
      A.colunas.forEach(function (k, j) {
        var e = document.createElement('div'); e.className = 'cel'; e.style.width = (k.largura * fator) + 'pt';
        var t = A.linhas[i][j] || '';
        if (k.alinhamento === 'direita') e.style.textAlign = 'right';
        if (k.selo && D.tema.situacoes[t]) {
          e.style.color = D.tema.situacoes[t].texto; e.style.fontWeight = 600;
          var p = document.createElement('span'); p.className = 'ico'; p.style.display = 'inline'; p.style.fontSize = '5pt';
          p.textContent = glifo('ponto') + ' '; e.appendChild(p); e.appendChild(document.createTextNode(t));
        } else { e.textContent = t; }
        lin.appendChild(e);
      });
      d.appendChild(lin);
    }
    var r = document.createElement('div'); r.className = 'rolagem'; r.style.width = ROLAGEM + 'pt'; d.appendChild(r);
  }

  // ------------------------------------------------------------ conferência
  function conferir() {
    var problemas = [];
    var W = T.canvas.largura, H = T.canvas.altura;
    var ctx = document.createElement('canvas').getContext('2d');
    itens.forEach(function (it) {
      var p = it.pos;
      if (p[0] < 0 || p[1] < 0 || p[0] + p[2] > W + 0.01 || p[1] + p[3] > H + 0.01) {
        problemas.push({ tipo: 'fora', nome: it.nome, msg: it.nome + ' passa da área da tela (' + W + ' x ' + H + ' pt)' });
        it.el.classList.add('problema');
      }
      if (it.texto) {
        var f = it.fonte || {};
        ctx.font = (f['Font.Bold'] ? '600 ' : '400 ') + (f['Font.Size'] || 9) + 'pt "' + valor(f['Font.Name'] || '@padrao') + '"';
        var larg = ctx.measureText(it.texto).width * 72 / 96 + (it.extra || 0) + 6;
        // Label com WordWrap (padrão do MSForms) e altura para mais de uma
        // linha quebra o texto: vale a largura vezes as linhas, com folga
        // para a quebra por palavra
        var linhas = 1;
        if (it.tipo === 'Label' && f.WordWrap !== false) linhas = Math.max(1, Math.floor(p[3] / ((f['Font.Size'] || 9) * 1.35)));
        if (larg > p[2] * linhas * (linhas > 1 ? 0.9 : 1)) {
          problemas.push({ tipo: 'texto', nome: it.nome, msg: it.nome + ': "' + it.texto + '" precisa de ' + Math.ceil(larg) + ' pt e tem ' + p[2] });
          it.el.classList.add('problema');
        }
      }
    });
    for (var i = 0; i < itens.length; i++) {
      for (var j = i + 1; j < itens.length; j++) {
        var a = itens[i], b = itens[j];
        if (a.permiteSobrepor || b.permiteSobrepor || a.dentro === b.nome || b.dentro === a.nome) continue;
        var x = Math.min(a.pos[0] + a.pos[2], b.pos[0] + b.pos[2]) - Math.max(a.pos[0], b.pos[0]);
        var y = Math.min(a.pos[1] + a.pos[3], b.pos[1] + b.pos[3]) - Math.max(a.pos[1], b.pos[1]);
        if (x > 0.5 && y > 0.5) {
          problemas.push({ tipo: 'sobreposicao', nome: a.nome + '/' + b.nome, msg: a.nome + ' e ' + b.nome + ' se sobrepõem (' + Math.round(x) + ' x ' + Math.round(y) + ' pt)' });
          a.el.classList.add('problema'); b.el.classList.add('problema');
        }
      }
    }
    saida.textContent = JSON.stringify({ tela: nomeTela, canvas: T.canvas, itens: itens.length, problemas: problemas });
    barra.innerHTML = '';
    nomes.forEach(function (n) {
      var a = document.createElement('a'); a.href = '?tela=' + n + '&contornos=' + (q.get('contornos') || '');
      a.textContent = n; a.style.marginRight = '10pt'; if (n === nomeTela) a.style.fontWeight = 600; barra.appendChild(a);
    });
    var st = document.createElement('span');
    st.textContent = ' | ' + T.canvas.largura + ' x ' + T.canvas.altura + ' pt | ' + (problemas.length ? problemas.length + ' problema(s) - veja abaixo' : 'sem problemas');
    st.style.color = problemas.length ? '#C00' : '#2C6B3A';
    barra.appendChild(st);
  }
})();
