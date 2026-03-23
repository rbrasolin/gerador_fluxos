let ultimoNomeArquivo = "fluxograma_processo";

const CONFIG = {
  boxWidth: 250,
  boxHeight: 110,
  colGap: 140,
  rowGap: 70,
  marginX: 120,
  marginY: 60,
  fontFamily: "Arial, sans-serif",
  fontSize: 16,
  smallFontSize: 14,
  stroke: "#111111",
  lineWidth: 2.2,
  cornerRadius: 12,
  decisionWidth: 180,
  decisionTextWidthFactor: 0.52,
  decisionHeightFactor: 1.3,
  routeGap: 28,
  entryExitGap: 40,
  laneGap: 36,
  sameRowTolerance: 1,
  sameColTolerance: 1,
  sharedMergeGap: 34,
  obstaclePadding: 10,
  textLineHeight: 18,
  textPaddingVertical: 20,
  rectTextPaddingHorizontal: 16
};

function limpar(txt) {
  if (txt === null || txt === undefined) return "";
  return String(txt).replace(/"/g, "").replace(/\(/g, "").replace(/\)/g, "").trim();
}

function escaparHTML(txt) {
  return String(txt || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function obterValorCampo(id) {
  const el = document.getElementById(id);
  return el ? limpar(el.value) : "";
}

function limparCampo(id) {
  const el = document.getElementById(id);
  if (el) el.value = "";
}

function normalizarCor(cor) {
  const c = limpar(cor).toLowerCase();
  const mapa = {
    blue: "blue",
    azul: "blue",
    yellow: "yellow",
    amarelo: "yellow",
    green: "green",
    verde: "green",
    white: "white",
    branco: "white",
    red: "red",
    vermelho: "red"
  };
  return mapa[c] || "white";
}

function corHex(cor) {
  const mapa = {
    white: "#ffffff",
    blue: "#8ecae6",
    yellow: "#ffd166",
    green: "#95d5b2",
    red: "#ef476f"
  };
  return mapa[cor] || "#ffffff";
}

function ehCabecalho(colunas) {
  if (!colunas || colunas.length === 0) return false;
  return limpar(colunas[0]).toLowerCase() === "ordem";
}

function tempoParaSegundos(tempo) {
  if (!tempo) return 0;
  tempo = String(tempo).trim();

  if (!tempo.includes(":")) {
    return (Number(tempo.replace(",", ".")) || 0) * 3600;
  }

  const partes = tempo.split(":");
  const h = Number(partes[0]) || 0;
  const m = Number(partes[1]) || 0;
  const s = Number(partes[2]) || 0;

  return h * 3600 + m * 60 + s;
}

function segundosParaTempo(seg) {
  seg = Math.round(Number(seg) || 0);
  const h = Math.floor(seg / 3600);
  const m = Math.floor((seg % 3600) / 60);
  const s = seg % 60;

  return (
    String(h).padStart(2, "0") + ":" +
    String(m).padStart(2, "0") + ":" +
    String(s).padStart(2, "0")
  );
}

function formatarTempo(seg) {
  return segundosParaTempo(seg);
}

function formatarPercentual(valor) {
  return (Number(valor) || 0).toFixed(1).replace(".", ",");
}

function formatarOrdem(ordem, fallbackIndex = 0) {
  if (ordem > 0) return String(ordem);
  const n = fallbackIndex + 1;
  let resultado = "";
  let valor = n;

  while (valor > 0) {
    valor--;
    resultado = String.fromCharCode(65 + (valor % 26)) + resultado;
    valor = Math.floor(valor / 26);
  }

  return resultado;
}

function parsePercentual(valor) {
  if (valor === null || valor === undefined || valor === "") return 0;

  let texto = String(valor).trim();
  const tinhaPercentual = texto.includes("%");
  texto = texto.replace(/%/g, "").replace(/\s/g, "").replace(/\./g, "").replace(/,/g, ".");

  let numero = Number(texto);
  if (!Number.isFinite(numero)) return 0;

  if (!tinhaPercentual && numero > 0 && numero <= 1) {
    numero = numero * 100;
  }

  if (numero < 0) numero = 0;
  if (numero > 100) numero = 100;
  return numero;
}

function normalizarCategoriaOportunidade(valor) {
  const texto = limpar(valor).toLowerCase();
  const mapa = {
    "automação": "Automação",
    "automacao": "Automação",
    "melhoria de processo": "Melhoria de processo",
    "já automatizado": "Já automatizado",
    "ja automatizado": "Já automatizado",
    "sem oportunidade": "Sem oportunidade",
    "sem oportunidade clara": "Sem oportunidade"
  };
  return mapa[texto] || (limpar(valor) || "Sem oportunidade");
}

function calcularTempoToBe(tempo, percentualReducao) {
  const ganhoPotencial = tempo * ((Number(percentualReducao) || 0) / 100);
  return {
    ganhoPotencial,
    tempoToBe: Math.max(0, tempo - ganhoPotencial)
  };
}

function enriquecerEtapaComMelhoria(etapa) {
  const percentualReducao = parsePercentual(etapa.potencialReducao ?? etapa.percentualReducao ?? 0);
  const categoriaOportunidade = normalizarCategoriaOportunidade(etapa.categoriaOportunidade ?? etapa.categoria ?? "");
  const observacao = limpar(etapa.observacao);
  const calculo = calcularTempoToBe(etapa.tempo || 0, percentualReducao);

  return {
    ...etapa,
    percentualReducao,
    potencialReducao: percentualReducao,
    categoriaOportunidade,
    observacao,
    ganhoPotencial: calculo.ganhoPotencial,
    tempoToBe: calculo.tempoToBe
  };
}

function montarDadosSimulacaoMelhoria(etapas) {
  const rows = [...etapas]
    .sort((a, b) => (Number(a.ordem) || 0) - (Number(b.ordem) || 0))
    .map((etapa, index) => {
      const etapaEnriquecida = enriquecerEtapaComMelhoria(etapa);
      return {
        ...etapaEnriquecida,
        ordemFmt: formatarOrdem(etapaEnriquecida.ordem, index)
      };
    });

  const tempoTotalAsIs = rows.reduce((acc, item) => acc + (item.tempo || 0), 0);
  const ganhoPotencialHoras = rows.reduce((acc, item) => acc + (item.ganhoPotencial || 0), 0);
  const tempoTotalToBe = rows.reduce((acc, item) => acc + (item.tempoToBe || 0), 0);
  const eficienciaPotencial = tempoTotalAsIs ? (ganhoPotencialHoras / tempoTotalAsIs) * 100 : 0;

  const ranking = [...rows]
    .filter(item => (item.ganhoPotencial || 0) > 0)
    .sort((a, b) => b.ganhoPotencial - a.ganhoPotencial)
    .slice(0, 5);

  return {
    rows,
    ranking,
    quantidadeAtividades: rows.length,
    tempoTotalAsIs,
    ganhoPotencialHoras,
    tempoTotalToBe,
    eficienciaPotencial
  };
}

function quebrarListaIds(valor) {
  return String(valor || "")
    .split(",")
    .map(item => limpar(item))
    .filter(Boolean);
}

function destinoEhValido(destinoId, idsValidos) {
  return !!destinoId && idsValidos.has(destinoId);
}

function criarElementoSVG(tag) {
  return document.createElementNS("http://www.w3.org/2000/svg", tag);
}

function isPergunta(texto) {
  return limpar(texto).endsWith("?");
}

function medirLarguraTexto(texto, fontSize = CONFIG.fontSize, fontWeight = "normal") {
  const svgMedicao = criarElementoSVG("svg");
  svgMedicao.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svgMedicao.setAttribute("width", "0");
  svgMedicao.setAttribute("height", "0");
  svgMedicao.setAttribute(
    "style",
    "position:absolute;left:-9999px;top:-9999px;visibility:hidden;overflow:hidden;"
  );

  const text = criarElementoSVG("text");
  text.setAttribute("font-family", CONFIG.fontFamily);
  text.setAttribute("font-size", String(fontSize));
  text.setAttribute("font-weight", fontWeight);
  text.textContent = texto || "";
  svgMedicao.appendChild(text);

  document.body.appendChild(svgMedicao);
  const largura = text.getComputedTextLength();
  document.body.removeChild(svgMedicao);

  return largura;
}

function quebrarTextoPorLargura(texto, larguraMaxima, fontSize = CONFIG.fontSize, fontWeight = "normal") {
  const textoLimpo = String(texto || "").trim();
  if (!textoLimpo) return [];

  const palavras = textoLimpo.split(/\s+/).filter(Boolean);
  if (!palavras.length) return [];

  const linhas = [];
  let linhaAtual = "";

  for (const palavra of palavras) {
    const tentativa = linhaAtual ? `${linhaAtual} ${palavra}` : palavra;

    if (medirLarguraTexto(tentativa, fontSize, fontWeight) <= larguraMaxima) {
      linhaAtual = tentativa;
      continue;
    }

    if (linhaAtual) {
      linhas.push(linhaAtual);
      linhaAtual = "";
    }

    if (medirLarguraTexto(palavra, fontSize, fontWeight) <= larguraMaxima) {
      linhaAtual = palavra;
      continue;
    }

    let parteAtual = "";

    for (const caractere of palavra) {
      const tentativaParte = parteAtual + caractere;

      if (medirLarguraTexto(tentativaParte, fontSize, fontWeight) <= larguraMaxima) {
        parteAtual = tentativaParte;
      } else {
        if (parteAtual) linhas.push(parteAtual);
        parteAtual = caractere;
      }
    }

    if (parteAtual) {
      linhaAtual = parteAtual;
    }
  }

  if (linhaAtual) linhas.push(linhaAtual);
  return linhas;
}

function obterLarguraNo(etapa) {
  return isPergunta(etapa.atividade) ? CONFIG.decisionWidth : CONFIG.boxWidth;
}

function obterLarguraUtilTexto(etapa, larguraCaixa) {
  if (isPergunta(etapa.atividade)) {
    return Math.max(40, larguraCaixa * CONFIG.decisionTextWidthFactor);
  }
  return Math.max(60, larguraCaixa - CONFIG.rectTextPaddingHorizontal * 2);
}

function obterLinhasEtapa(etapa, larguraCaixa) {
  const larguraTexto = obterLarguraUtilTexto(etapa, larguraCaixa);

  const linhasAtividade = quebrarTextoPorLargura(etapa.atividade, larguraTexto, CONFIG.fontSize);
  const linhasSistema = quebrarTextoPorLargura(etapa.sistema || "Sem sistema informado", larguraTexto, CONFIG.fontSize);

  return [
    ...linhasAtividade,
    ...linhasSistema,
    formatarTempo(etapa.tempo)
  ];
}

function obterAlturaNo(etapa, alturaPadraoBase) {
  if (isPergunta(etapa.atividade)) {
    return Math.ceil(alturaPadraoBase * CONFIG.decisionHeightFactor);
  }
  return alturaPadraoBase;
}

function calcularAlturaNecessariaEtapa(etapa) {
  const largura = obterLarguraNo(etapa);
  const linhas = obterLinhasEtapa(etapa, largura);

  const alturaTexto = linhas.length * CONFIG.textLineHeight;
  const alturaNecessaria = alturaTexto + CONFIG.textPaddingVertical * 2;

  if (isPergunta(etapa.atividade)) {
    return Math.max(CONFIG.boxHeight, Math.ceil(alturaNecessaria / CONFIG.decisionHeightFactor));
  }

  return Math.max(CONFIG.boxHeight, alturaNecessaria);
}

function calcularAlturaPadraoNos(etapas) {
  let maiorAltura = CONFIG.boxHeight;

  etapas.forEach((etapa) => {
    const altura = calcularAlturaNecessariaEtapa(etapa);
    if (altura > maiorAltura) maiorAltura = altura;
  });

  return maiorAltura;
}

function gerarNomeArquivo() {
  const processo = limpar(document.getElementById("processo").value) || "fluxograma";
  return processo
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase() || "fluxograma_processo";
}

function limparTudo() {
  limparCampo("desenho");
  limparCampo("processo");
  limparCampo("analista");
  limparCampo("negocio");
  limparCampo("area");
  limparCampo("gestor");
  limparCampo("entrada");

  const diagram = document.getElementById("diagram");
  const infoProcesso = document.getElementById("infoProcesso");
  const metricas = document.getElementById("metricas");

  if (diagram) diagram.innerHTML = "";
  if (infoProcesso) infoProcesso.innerHTML = "";
  if (metricas) metricas.innerHTML = "";

  ultimoNomeArquivo = "fluxograma_processo";
}

function adicionarEtapasImpactadasPorRetorno(origemId, destinoId, etapaPorId, etapasOrdenadas, etapasImpactadasRef) {
  const origem = etapaPorId[origemId];
  const destino = etapaPorId[destinoId];

  if (!origem || !destino) return;
  if (destino.ordem >= origem.ordem) return;

  etapasOrdenadas.forEach((etapa) => {
    if (etapa.ordem >= destino.ordem && etapa.ordem <= origem.ordem) {
      etapasImpactadasRef.add(etapa.id);
    }
  });
}

function desenharCapsula(svg, texto, x, y, width = 60, height = 36) {
  const g = criarElementoSVG("g");

  const rect = criarElementoSVG("rect");
  rect.setAttribute("x", x);
  rect.setAttribute("y", y);
  rect.setAttribute("width", width);
  rect.setAttribute("height", height);
  rect.setAttribute("rx", height / 2);
  rect.setAttribute("ry", height / 2);
  rect.setAttribute("fill", "#ffffff");
  rect.setAttribute("stroke", CONFIG.stroke);
  rect.setAttribute("stroke-width", CONFIG.lineWidth);
  g.appendChild(rect);

  const t = criarElementoSVG("text");
  t.setAttribute("x", x + width / 2);
  t.setAttribute("y", y + height / 2 + 5);
  t.setAttribute("text-anchor", "middle");
  t.setAttribute("font-family", CONFIG.fontFamily);
  t.setAttribute("font-size", CONFIG.fontSize);
  t.textContent = texto;
  g.appendChild(t);

  svg.appendChild(g);
}

function desenharNo(svg, etapa, x, y, width, height, fill) {
  const group = criarElementoSVG("g");
  group.setAttribute("class", "node");

  const isDecision = isPergunta(etapa.atividade);

  if (isDecision) {
    const points = [
      [x + width / 2, y],
      [x + width, y + height / 2],
      [x + width / 2, y + height],
      [x, y + height / 2]
    ]
      .map(p => p.join(","))
      .join(" ");

    const poly = criarElementoSVG("polygon");
    poly.setAttribute("points", points);
    poly.setAttribute("fill", fill);
    poly.setAttribute("stroke", CONFIG.stroke);
    poly.setAttribute("stroke-width", CONFIG.lineWidth);
    group.appendChild(poly);
  } else {
    const rect = criarElementoSVG("rect");
    rect.setAttribute("x", x);
    rect.setAttribute("y", y);
    rect.setAttribute("width", width);
    rect.setAttribute("height", height);
    rect.setAttribute("rx", CONFIG.cornerRadius);
    rect.setAttribute("ry", CONFIG.cornerRadius);
    rect.setAttribute("fill", fill);
    rect.setAttribute("stroke", CONFIG.stroke);
    rect.setAttribute("stroke-width", CONFIG.lineWidth);
    group.appendChild(rect);
  }

  const linhas = obterLinhasEtapa(etapa, width);
  const alturaTotalTexto = linhas.length * CONFIG.textLineHeight;
  const inicioY = y + (height - alturaTotalTexto) / 2 + CONFIG.fontSize - 2;

  linhas.forEach((linha, idx) => {
    const text = criarElementoSVG("text");
    text.setAttribute("x", x + width / 2);
    text.setAttribute("y", inicioY + idx * CONFIG.textLineHeight);
    text.setAttribute("text-anchor", "middle");
    text.setAttribute("font-family", CONFIG.fontFamily);
    text.setAttribute("font-size", CONFIG.fontSize);
    if (idx === linhas.length - 1) text.setAttribute("font-weight", "bold");
    text.textContent = linha;
    group.appendChild(text);
  });

  svg.appendChild(group);
}

function criarMarkerArrow(svg) {
  const defs = criarElementoSVG("defs");
  const marker = criarElementoSVG("marker");
  marker.setAttribute("id", "arrowhead");
  marker.setAttribute("markerWidth", "10");
  marker.setAttribute("markerHeight", "7");
  marker.setAttribute("refX", "9");
  marker.setAttribute("refY", "3.5");
  marker.setAttribute("orient", "auto");
  marker.setAttribute("markerUnits", "strokeWidth");

  const path = criarElementoSVG("path");
  path.setAttribute("d", "M0,0 L10,3.5 L0,7 z");
  path.setAttribute("fill", CONFIG.stroke);

  marker.appendChild(path);
  defs.appendChild(marker);
  svg.appendChild(defs);
}

function desenharTextoSobreLinha(svg, texto, x, y) {
  if (!texto) return;

  const paddingX = 10;
  const width = Math.max(34, medirLarguraTexto(texto, CONFIG.smallFontSize, "bold") + paddingX * 2);
  const height = 22;

  const group = criarElementoSVG("g");

  const rect = criarElementoSVG("rect");
  rect.setAttribute("x", x - width / 2);
  rect.setAttribute("y", y - height / 2);
  rect.setAttribute("width", width);
  rect.setAttribute("height", height);
  rect.setAttribute("rx", "11");
  rect.setAttribute("ry", "11");
  rect.setAttribute("fill", "#ffffff");
  rect.setAttribute("stroke", CONFIG.stroke);
  rect.setAttribute("stroke-width", "1.2");
  group.appendChild(rect);

  const text = criarElementoSVG("text");
  text.setAttribute("x", x);
  text.setAttribute("y", y + 4);
  text.setAttribute("text-anchor", "middle");
  text.setAttribute("font-family", CONFIG.fontFamily);
  text.setAttribute("font-size", CONFIG.smallFontSize);
  text.setAttribute("font-weight", "bold");
  text.textContent = texto;
  group.appendChild(text);

  svg.appendChild(group);
}

function desenharPolyline(svg, points, className = "") {
  const polyline = criarElementoSVG("polyline");
  polyline.setAttribute("points", points.map((p) => `${p.x},${p.y}`).join(" "));
  polyline.setAttribute("fill", "none");
  polyline.setAttribute("stroke", CONFIG.stroke);
  polyline.setAttribute("stroke-width", CONFIG.lineWidth);
  polyline.setAttribute("stroke-linejoin", "round");
  polyline.setAttribute("stroke-linecap", "round");
  polyline.setAttribute("marker-end", "url(#arrowhead)");
  if (className) polyline.setAttribute("class", className);
  svg.appendChild(polyline);
  return polyline;
}

function normalizarPontos(points) {
  if (!points || points.length === 0) return [];

  const resultado = [points[0]];
  for (let i = 1; i < points.length; i++) {
    const prev = resultado[resultado.length - 1];
    const curr = points[i];
    if (!curr) continue;
    if (prev.x === curr.x && prev.y === curr.y) continue;
    resultado.push(curr);
  }

  let mudou = true;
  while (mudou) {
    mudou = false;
    for (let i = 1; i < resultado.length - 1; i++) {
      const a = resultado[i - 1];
      const b = resultado[i];
      const c = resultado[i + 1];
      const mesmoX = a.x === b.x && b.x === c.x;
      const mesmoY = a.y === b.y && b.y === c.y;
      if (mesmoX || mesmoY) {
        resultado.splice(i, 1);
        mudou = true;
        break;
      }
    }
  }

  return resultado;
}

function calcularComprimento(points) {
  let total = 0;
  for (let i = 1; i < points.length; i++) {
    total += Math.abs(points[i].x - points[i - 1].x) + Math.abs(points[i].y - points[i - 1].y);
  }
  return total;
}

function getAnchorPoint(node, side) {
  switch (side) {
    case "left": return { x: node.x, y: node.y + node.height / 2 };
    case "right": return { x: node.x + node.width, y: node.y + node.height / 2 };
    case "top": return { x: node.x + node.width / 2, y: node.y };
    case "bottom": return { x: node.x + node.width / 2, y: node.y + node.height };
    default: return { x: node.x + node.width, y: node.y + node.height / 2 };
  }
}

function getNodeRect(node) {
  return {
    left: node.x,
    right: node.x + node.width,
    top: node.y,
    bottom: node.y + node.height
  };
}

function segmentCruzaRetangulo(p1, p2, rect) {
  if (p1.x === p2.x) {
    const x = p1.x;
    const minY = Math.min(p1.y, p2.y);
    const maxY = Math.max(p1.y, p2.y);
    if (x > rect.left && x < rect.right) return maxY > rect.top && minY < rect.bottom;
    return false;
  }

  if (p1.y === p2.y) {
    const y = p1.y;
    const minX = Math.min(p1.x, p2.x);
    const maxX = Math.max(p1.x, p2.x);
    if (y > rect.top && y < rect.bottom) return maxX > rect.left && minX < rect.right;
    return false;
  }

  return false;
}

function pathCruzaCaixas(points, posicoes = {}, excludeIds = []) {
  const exclude = new Set(excludeIds || []);
  const ids = Object.keys(posicoes);

  for (let i = 1; i < points.length; i++) {
    const p1 = points[i - 1];
    const p2 = points[i];

    for (const id of ids) {
      if (exclude.has(id)) continue;
      const rect = getNodeRect(posicoes[id]);
      const expandido = {
        left: rect.left - CONFIG.obstaclePadding,
        right: rect.right + CONFIG.obstaclePadding,
        top: rect.top - CONFIG.obstaclePadding,
        bottom: rect.bottom + CONFIG.obstaclePadding
      };

      if (segmentCruzaRetangulo(p1, p2, expandido)) {
        return true;
      }
    }
  }

  return false;
}

function detectarLadoSaida(points) {
  if (!points || points.length < 2) return "right";
  const p1 = points[0];
  const p2 = points[1];
  if (p2.x > p1.x) return "right";
  if (p2.x < p1.x) return "left";
  if (p2.y > p1.y) return "bottom";
  return "top";
}

function detectarLadoEntrada(points) {
  if (!points || points.length < 2) return "left";
  const p1 = points[points.length - 2];
  const p2 = points[points.length - 1];
  if (p2.x > p1.x) return "left";
  if (p2.x < p1.x) return "right";
  if (p2.y > p1.y) return "top";
  return "bottom";
}

function ajustarPrimeiroTrechoParaLado(points, start, side) {
  if (!points || points.length < 2) return points;
  const segundo = points[1];
  const novo = [start];

  if (side === "right" || side === "left") {
    if (segundo.y !== start.y) novo.push({ x: segundo.x, y: start.y });
  } else {
    if (segundo.x !== start.x) novo.push({ x: start.x, y: segundo.y });
  }

  for (let i = 1; i < points.length; i++) novo.push(points[i]);
  return normalizarPontos(novo);
}

function ajustarUltimoTrechoParaLado(points, end, side, destinoNode = null) {
  if (!points || points.length < 2) return points;

  let resultado = [...points];
  const penultimo = resultado[resultado.length - 2];
  const ultimo = resultado[resultado.length - 1];

  if (ultimo.x === end.x && ultimo.y === end.y) {
    let precisaAjuste = false;
    if ((side === "left" || side === "right") && penultimo.y !== end.y) precisaAjuste = true;
    if ((side === "top" || side === "bottom") && penultimo.x !== end.x) precisaAjuste = true;
    if (!precisaAjuste) return normalizarPontos(resultado);
  }

  const escape = (() => {
    switch (side) {
      case "left": return { x: end.x - CONFIG.routeGap, y: end.y };
      case "right": return { x: end.x + CONFIG.routeGap, y: end.y };
      case "top": return { x: end.x, y: end.y - CONFIG.routeGap };
      case "bottom": return { x: end.x, y: end.y + CONFIG.routeGap };
      default: return { x: end.x - CONFIG.routeGap, y: end.y };
    }
  })();

  if (destinoNode) {
    const nodeRect = getNodeRect(destinoNode);
    if (side === "left") escape.x = Math.min(escape.x, nodeRect.left - CONFIG.routeGap);
    else if (side === "right") escape.x = Math.max(escape.x, nodeRect.right + CONFIG.routeGap);
    else if (side === "top") escape.y = Math.min(escape.y, nodeRect.top - CONFIG.routeGap);
    else if (side === "bottom") escape.y = Math.max(escape.y, nodeRect.bottom + CONFIG.routeGap);
  }

  resultado[resultado.length - 1] = escape;
  resultado.push(end);

  if (resultado.length >= 3) {
    const terceiro = resultado[resultado.length - 3];
    const novo = [];

    for (let i = 0; i < resultado.length - 2; i++) novo.push(resultado[i]);

    if (side === "left" || side === "right") {
      if (escape.x !== terceiro.x && escape.y !== terceiro.y) novo.push({ x: escape.x, y: terceiro.y });
    } else {
      if (escape.x !== terceiro.x && escape.y !== terceiro.y) novo.push({ x: terceiro.x, y: escape.y });
    }

    for (let i = 1; i < resultado.length; i++) novo.push(resultado[i]);
    resultado = novo;
  }

  return normalizarPontos(resultado);
}

function gerarCandidatosRotas(start, end) {
  const candidates = [];
  const mids = [
    [{ x: end.x, y: start.y }],
    [{ x: start.x, y: end.y }],
    [{ x: start.x + CONFIG.routeGap, y: start.y }, { x: start.x + CONFIG.routeGap, y: end.y }],
    [{ x: start.x - CONFIG.routeGap, y: start.y }, { x: start.x - CONFIG.routeGap, y: end.y }],
    [{ x: start.x, y: start.y + CONFIG.routeGap }, { x: end.x, y: start.y + CONFIG.routeGap }],
    [{ x: start.x, y: start.y - CONFIG.routeGap }, { x: end.x, y: start.y - CONFIG.routeGap }]
  ];

  mids.forEach(midPoints => {
    candidates.push(normalizarPontos([start, ...midPoints, end]));
  });

  return candidates;
}

function encontrarRotaSegura(start, end, posicoes = {}, excludeIds = [], preferredEndSide = null, preferredStartSide = null, destinoNode = null) {
  const candidates = gerarCandidatosRotas(start, end);

  const avaliadas = candidates.map(points => {
    let ajustado = [...points];
    const startSide = preferredStartSide || detectarLadoSaida(ajustado);
    const endSide = preferredEndSide || detectarLadoEntrada(ajustado);

    ajustado = ajustarPrimeiroTrechoParaLado(ajustado, start, startSide);
    ajustado = ajustarUltimoTrechoParaLado(ajustado, end, endSide, destinoNode);

    return {
      points: ajustado,
      startSide,
      endSide,
      safe: !pathCruzaCaixas(ajustado, posicoes, excludeIds)
    };
  });

  const validos = avaliadas.filter(r => r.safe);

  if (validos.length) {
    validos.sort((a, b) => {
      if (a.points.length !== b.points.length) return a.points.length - b.points.length;
      return calcularComprimento(a.points) - calcularComprimento(b.points);
    });

    const melhor = validos[0];
    let ajustado = melhor.points;
    ajustado = ajustarPrimeiroTrechoParaLado(ajustado, start, preferredStartSide || melhor.startSide);
    ajustado = ajustarUltimoTrechoParaLado(ajustado, end, preferredEndSide || melhor.endSide, destinoNode);

    return {
      points: ajustado,
      startSide: preferredStartSide || melhor.startSide,
      endSide: preferredEndSide || melhor.endSide,
      safe: !pathCruzaCaixas(ajustado, posicoes, excludeIds)
    };
  }

  let fallback = normalizarPontos([start, { x: end.x, y: start.y }, end]);
  const startSide = preferredStartSide || detectarLadoSaida(fallback);
  const endSide = preferredEndSide || detectarLadoEntrada(fallback);

  fallback = ajustarPrimeiroTrechoParaLado(fallback, start, startSide);
  fallback = ajustarUltimoTrechoParaLado(fallback, end, endSide, destinoNode);

  return {
    points: fallback,
    startSide,
    endSide,
    safe: !pathCruzaCaixas(fallback, posicoes, excludeIds)
  };
}

function buildOrthogonalToMerge(start, mergePoint, end, endSide, posicoes = {}, excludeIds = [], preferredStartSide = null) {
  const ateMergeObj = encontrarRotaSegura(start, mergePoint, posicoes, excludeIds, null, preferredStartSide, null);
  return normalizarPontos([...ateMergeObj.points, end]);
}

function podeCompartilharDestino(origem, sharedInfo) {
  if (!sharedInfo) return false;
  return origem.gridCol === sharedInfo.sourceGridCol;
}

function escolherParesCandidatos(origem, destino, rotulo = "") {
  if (origem.isDecision && rotulo === "Sim") return [{ startSide: "right", endSide: "left" }];

  if (origem.isDecision && rotulo === "Não") {
    if (destino.gridRow <= origem.gridRow) return [{ startSide: "bottom", endSide: "bottom" }];
    return [{ startSide: "bottom", endSide: "top" }];
  }

  const dx = destino.gridCol - origem.gridCol;
  const dy = destino.gridRow - origem.gridRow;

  if (dx === 0 && dy > 0) return [{ startSide: "bottom", endSide: "top" }];
  if (dx === 0 && dy < 0) return [{ startSide: "top", endSide: "bottom" }];
  if (dy === 0 && dx > 0) return [{ startSide: "right", endSide: "left" }];
  if (dy === 0 && dx < 0) return [{ startSide: "left", endSide: "right" }];

  if (dx > 0 && dy > 0) {
    return [
      { startSide: "right", endSide: "left" },
      { startSide: "right", endSide: "top" },
      { startSide: "bottom", endSide: "left" },
      { startSide: "bottom", endSide: "top" }
    ];
  }

  if (dx > 0 && dy < 0) {
    return [
      { startSide: "right", endSide: "left" },
      { startSide: "right", endSide: "bottom" },
      { startSide: "top", endSide: "left" },
      { startSide: "top", endSide: "bottom" }
    ];
  }

  if (dx < 0 && dy > 0) {
    return [
      { startSide: "left", endSide: "right" },
      { startSide: "left", endSide: "top" },
      { startSide: "bottom", endSide: "right" },
      { startSide: "bottom", endSide: "top" }
    ];
  }

  if (dx < 0 && dy < 0) {
    return [
      { startSide: "left", endSide: "right" },
      { startSide: "left", endSide: "bottom" },
      { startSide: "top", endSide: "right" },
      { startSide: "top", endSide: "bottom" }
    ];
  }

  return [{ startSide: "right", endSide: "left" }];
}

function montarRotaOrtogonal(points, label, startSide = "", endSide = "") {
  return {
    points: normalizarPontos(points),
    label,
    startSide,
    endSide
  };
}

function getMergePoint(end, side, gap = CONFIG.sharedMergeGap) {
  switch (side) {
    case "left": return { x: end.x - gap, y: end.y };
    case "right": return { x: end.x + gap, y: end.y };
    case "top": return { x: end.x, y: end.y - gap };
    case "bottom": return { x: end.x, y: end.y + gap };
    default: return { x: end.x - gap, y: end.y };
  }
}

function escolherRota(origem, destino, contexto = {}) {
  const rotulo = contexto.rotulo || "";
  const posicoes = contexto.posicoes || {};
  const excludeIds = [origem.id, destino.id, "__INICIO__", "__FIM__"];

  if (origem.id === "__INICIO__") {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    const rota = encontrarRotaSegura(start, end, posicoes, excludeIds, "left", "right", destino);
    return montarRotaOrtogonal(
      rota.points,
      { x: (start.x + end.x) / 2, y: start.y - 10 },
      rota.startSide,
      rota.endSide
    );
  }

  if (destino.id === "__FIM__") {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    const rota = encontrarRotaSegura(start, end, posicoes, excludeIds, "left", "right", destino);
    return montarRotaOrtogonal(
      rota.points,
      { x: (start.x + end.x) / 2, y: start.y - 10 },
      rota.startSide,
      rota.endSide
    );
  }

  const pares = escolherParesCandidatos(origem, destino, rotulo);
  const tentativas = [];

  for (const par of pares) {
    const start = getAnchorPoint(origem, par.startSide);
    const end = getAnchorPoint(destino, par.endSide);
    const rota = encontrarRotaSegura(start, end, posicoes, excludeIds, par.endSide, par.startSide, destino);

    const pontosRota = rota.points || [];
    let comprimentoTotal = 0;
    const segmentos = [];

    for (let i = 0; i < pontosRota.length - 1; i++) {
      const p1 = pontosRota[i];
      const p2 = pontosRota[i + 1];
      const comprimento = Math.abs(p2.x - p1.x) + Math.abs(p2.y - p1.y);

      if (comprimento > 0) {
        segmentos.push({
          p1,
          p2,
          comprimento,
          inicio: comprimentoTotal,
          fim: comprimentoTotal + comprimento
        });
        comprimentoTotal += comprimento;
      }
    }

    let labelPoint = { x: start.x + 18, y: start.y - 10 };

    if (segmentos.length > 0 && comprimentoTotal > 0) {
      const alvo = comprimentoTotal / 2;

      for (const segmento of segmentos) {
        if (alvo >= segmento.inicio && alvo <= segmento.fim) {
          const deslocamento = alvo - segmento.inicio;

          if (segmento.p1.y === segmento.p2.y) {
            const direcao = segmento.p2.x >= segmento.p1.x ? 1 : -1;
            labelPoint = {
              x: segmento.p1.x + deslocamento * direcao,
              y: segmento.p1.y
            };
          } else if (segmento.p1.x === segmento.p2.x) {
            const direcao = segmento.p2.y >= segmento.p1.y ? 1 : -1;
            labelPoint = {
              x: segmento.p1.x,
              y: segmento.p1.y + deslocamento * direcao
            };
          }
          break;
        }
      }
    }

    tentativas.push({
      ...rota,
      label: labelPoint
    });
  }

  tentativas.sort((a, b) => {
    if (a.safe !== b.safe) return a.safe ? -1 : 1;
    if (a.points.length !== b.points.length) return a.points.length - b.points.length;
    return calcularComprimento(a.points) - calcularComprimento(b.points);
  });

  const melhor = tentativas[0];

  return montarRotaOrtogonal(
    melhor.points,
    melhor.label,
    melhor.startSide,
    melhor.endSide
  );
}

function construirRotaCompartilhada(start, sharedInfo, posicoes = {}, excludeIds = [], preferredStartSide = null) {
  const mergePoint = { x: sharedInfo.mergePoint.x, y: sharedInfo.mergePoint.y };
  const end = { x: sharedInfo.end.x, y: sharedInfo.end.y };
  const endSide = sharedInfo.endSide;

  return {
    points: buildOrthogonalToMerge(start, mergePoint, end, endSide, posicoes, excludeIds, preferredStartSide),
    label: sharedInfo.label,
    startSide: preferredStartSide,
    endSide
  };
}

function desenharConexao(svg, origem, destino, rotulo = "", ordemConexao = 0, posicoes = {}, sharedRegistry = {}) {
  let rota = escolherRota(origem, destino, { rotulo, ordemConexao, posicoes });

  const sharedKey = `${destino.id}__${rota.endSide || "auto"}`;
  const sharedInfo = sharedRegistry[sharedKey];

  if (
    destino.id !== "__FIM__" &&
    destino.id !== "__INICIO__" &&
    sharedInfo &&
    origem.id !== sharedInfo.origemId &&
    podeCompartilharDestino(origem, sharedInfo)
  ) {
    const parPreferido = escolherParesCandidatos(origem, destino, origem.isDecision ? rotulo : "")[0];
    const startReal = getAnchorPoint(origem, parPreferido.startSide);

    rota = construirRotaCompartilhada(
      startReal,
      sharedInfo,
      posicoes,
      [origem.id, destino.id, "__INICIO__", "__FIM__"],
      parPreferido.startSide
    );
  } else if (
    destino.id !== "__FIM__" &&
    destino.id !== "__INICIO__" &&
    !sharedInfo &&
    rota.endSide
  ) {
    const end = getAnchorPoint(destino, rota.endSide);
    const mergePoint = getMergePoint(end, rota.endSide);

    sharedRegistry[sharedKey] = {
      origemId: origem.id,
      sourceGridCol: origem.gridCol,
      mergePoint,
      end,
      endSide: rota.endSide,
      label: rota.label
    };
  }

  desenharPolyline(svg, rota.points);

  if (rotulo && rota.label) {
    desenharTextoSobreLinha(svg, rotulo, rota.label.x, rota.label.y - 14);
  }
}

function renderInformacoesProcessoExecutivas(dados) {
  const quantidadeItens = [
    ["Desenho", dados.desenho || "-"],
    ["Processo", dados.processo || "-"],
    ["Analista", dados.analista || "-"],
    ["Negócio", dados.negocio || "-"],
    ["Área", dados.area || "-"],
    ["Gestor", dados.gestor || "-"]
  ];

  return `
    <div class="exec-card">
      <div class="exec-card-title">Informações do Processo</div>
      <div class="exec-info-grid">
        ${quantidadeItens.map(item => `
          <div class="exec-info-item">
            <div class="exec-info-label">${escaparHTML(item[0])}</div>
            <div class="exec-info-value">${escaparHTML(item[1])}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderTabelaAnaliseHTML({ titulo, columns, rows }) {
  const thead = `
    <thead>
      <tr>
        ${columns.map(col => `
          <th class="${col.align === "center" ? "th-center" : ""}">
            ${escaparHTML(col.header)}
          </th>
        `).join("")}
      </tr>
    </thead>
  `;

  const tbody = `
    <tbody>
      ${rows.map(row => `
        <tr class="${escaparHTML(row.rowClass || "")}">
          ${columns.map(col => `
            <td class="${col.align === "center" ? "td-center" : ""}">
              ${escaparHTML(row[col.key] ?? "")}
            </td>
          `).join("")}
        </tr>
      `).join("")}
    </tbody>
  `;

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">${escaparHTML(titulo)}</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          ${thead}
          ${tbody}
        </table>
      </div>
    </div>
  `;
}

function renderResumoAnaliseExecutivo(dados) {
  return `
    <div class="exec-summary-grid">
      <div class="exec-summary-item">
        <div class="exec-summary-label">Tempo total do processo</div>
        <div class="exec-summary-value">${formatarTempo(dados.tempoTotal)}</div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Loops detectados</div>
        <div class="exec-summary-value">${dados.loops}</div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Potencial retrabalho</div>
        <div class="exec-summary-value">${formatarTempo(dados.tempoPotencialRetrabalho)} | ${formatarPercentual(dados.impactoPotencialRetrabalho)}%</div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Taxa de decisão</div>
        <div class="exec-summary-value">${dados.decisoes} etapa(s) | ${formatarPercentual(dados.taxaDecisao)}%</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Ganho potencial em horas</div>
        <div class="exec-summary-value">${formatarTempo(dados.simulacaoMelhoria.ganhoPotencialHoras)}</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Tempo total To Be</div>
        <div class="exec-summary-value">${formatarTempo(dados.simulacaoMelhoria.tempoTotalToBe)}</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Eficiência potencial</div>
        <div class="exec-summary-value">${formatarPercentual(dados.simulacaoMelhoria.eficienciaPotencial)}%</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Atividades na simulação</div>
        <div class="exec-summary-value">${dados.simulacaoMelhoria.quantidadeAtividades}</div>
      </div>
    </div>
  `;
}

function renderTabelaSimulacaoMelhoria(dados) {
  const maiorGanho = Math.max(...dados.rows.map(item => item.ganhoPotencial || 0), 0);

  const rows = dados.rows.map(item => ({
    atividade: item.atividade,
    tempoAsIsFmt: formatarTempo(item.tempo),
    reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
    categoriaFmt: item.categoriaOportunidade || "Sem oportunidade",
    ganhoFmt: formatarTempo(item.ganhoPotencial),
    tempoToBeFmt: formatarTempo(item.tempoToBe),
    observacao: item.observacao || "-"
  }));

  rows.push({
    atividade: `${dados.quantidadeAtividades} atividade(s)`,
    tempoAsIsFmt: formatarTempo(dados.tempoTotalAsIs),
    reducaoFmt: `${formatarPercentual(dados.eficienciaPotencial)}%`,
    categoriaFmt: "-",
    ganhoFmt: formatarTempo(dados.ganhoPotencialHoras),
    tempoToBeFmt: formatarTempo(dados.tempoTotalToBe),
    observacao: `Eficiência total: ${formatarPercentual(dados.eficienciaPotencial)}%`,
    rowClass: "sim-total-row"
  });

  const thead = `
    <thead>
      <tr>
        <th>Atividade</th>
        <th class="th-center">Tempo As Is</th>
        <th class="th-center">% Redução</th>
        <th>Categoria</th>
        <th class="th-center">Ganho Potencial</th>
        <th class="th-center">Tempo To Be</th>
        <th>Observação</th>
      </tr>
    </thead>
  `;

  const tbody = `
    <tbody>
      ${rows.map((row, idx) => {
        const isMaiorOportunidade = idx < dados.rows.length && maiorGanho > 0 && dados.rows[idx].ganhoPotencial === maiorGanho;
        const rowClass = row.rowClass || (isMaiorOportunidade ? "row-highlight-opportunity" : "");
        return `
          <tr class="${rowClass}">
            <td>${escaparHTML(row.atividade)}</td>
            <td class="td-center">${escaparHTML(row.tempoAsIsFmt)}</td>
            <td class="td-center">${escaparHTML(row.reducaoFmt)}</td>
            <td>${escaparHTML(row.categoriaFmt)}</td>
            <td class="td-center">${escaparHTML(row.ganhoFmt)}</td>
            <td class="td-center">${escaparHTML(row.tempoToBeFmt)}</td>
            <td>${escaparHTML(row.observacao)}</td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">Simulação de Melhoria (As Is vs To Be)</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          ${thead}
          ${tbody}
        </table>
      </div>
    </div>
  `;
}

function renderRankingOportunidades(dados) {
  if (!dados.ranking.length) return "";

  return renderTabelaAnaliseHTML({
    titulo: "Ranking das Top Atividades por Ganho Potencial",
    columns: [
      { header: "#", key: "ranking", align: "center" },
      { header: "Atividade", key: "atividade", align: "left" },
      { header: "Ganho Potencial", key: "ganhoFmt", align: "center" },
      { header: "% Redução", key: "reducaoFmt", align: "center" },
      { header: "Categoria", key: "categoria", align: "left" }
    ],
    rows: dados.ranking.map((item, index) => ({
      ranking: String(index + 1),
      atividade: item.atividade,
      ganhoFmt: formatarTempo(item.ganhoPotencial),
      reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
      categoria: item.categoriaOportunidade || "Sem oportunidade"
    }))
  });
}

function renderizarAnaliseExecutiva(dados) {
  const tipoRows = dados.tempoPorTipo.map(item => ({
    tipo: item.tipo,
    tempoFmt: formatarTempo(item.tempo),
    percentualFmt: `${formatarPercentual(item.percentual)}%`
  }));

  const sistemaRows = dados.tempoPorSistema.map(item => ({
    sistema: item.sistema,
    tempoFmt: formatarTempo(item.tempo),
    percentualFmt: `${formatarPercentual(item.percentual)}%`
  }));

  const paretoRows = dados.pareto.map(item => ({
    atividade: item.atividade,
    tempoFmt: formatarTempo(item.tempo),
    percentualFmt: `${formatarPercentual(item.percentual)}%`,
    paretoFmt: `${formatarPercentual(item.pareto)}%`
  }));

  return `
    <div class="exec-card">
      <div class="exec-card-title">Análise do Processo</div>
      ${renderResumoAnaliseExecutivo(dados)}
      ${renderTabelaAnaliseHTML({
        titulo: "Tempo por Tipo",
        columns: [
          { header: "Tipo", key: "tipo", align: "left" },
          { header: "Tempo (horas)", key: "tempoFmt", align: "center" },
          { header: "%", key: "percentualFmt", align: "center" }
        ],
        rows: tipoRows
      })}
      ${renderTabelaAnaliseHTML({
        titulo: "Tempo por Sistema",
        columns: [
          { header: "Sistema", key: "sistema", align: "left" },
          { header: "Tempo (horas)", key: "tempoFmt", align: "center" },
          { header: "%", key: "percentualFmt", align: "center" }
        ],
        rows: sistemaRows
      })}
      ${renderTabelaAnaliseHTML({
        titulo: "Pareto de Tempo",
        columns: [
          { header: "Atividade", key: "atividade", align: "left" },
          { header: "Tempo (horas)", key: "tempoFmt", align: "center" },
          { header: "%", key: "percentualFmt", align: "center" },
          { header: "% Acumulado", key: "paretoFmt", align: "center" }
        ],
        rows: paretoRows
      })}
      ${renderTabelaSimulacaoMelhoria(dados.simulacaoMelhoria)}
      ${renderRankingOportunidades(dados.simulacaoMelhoria)}
    </div>
  `;
}

function gerarFluxo() {
  const texto = document.getElementById("entrada").value;

  if (!texto.trim()) {
    alert("Cole a tabela do Excel primeiro.");
    return;
  }

  const desenho = obterValorCampo("desenho");
  const processo = obterValorCampo("processo");
  const analista = obterValorCampo("analista");
  const negocio = obterValorCampo("negocio");
  const area = obterValorCampo("area");
  const gestor = obterValorCampo("gestor");

  const linhasBrutas = texto.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(l => l.trim() !== "");
  let linhas = linhasBrutas.map(l => l.split("\t"));

  if (linhas.length && ehCabecalho(linhas[0])) {
    linhas.shift();
  }

  const etapas = [];
  const idsValidos = new Set();

  linhas.forEach((col) => {
    while (col.length < 15) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]) || "Não informado";
    const sistema = limpar(col[4]) || "Sem sistema informado";
    const tempo = tempoParaSegundos(limpar(col[5]));
    const proxSim = limpar(col[6]);
    const proxNao = limpar(col[7]);
    const conexoesExtras = limpar(col[8]);
    const coluna = Number(limpar(col[9])) || 1;
    const linha = Number(limpar(col[10])) || 1;
    const categoriaOportunidade = normalizarCategoriaOportunidade(col[11]);
    const percentualReducao = parsePercentual(col[12]);
    const observacao = limpar(col[13]);
    const cor = normalizarCor(col[14]);

    if (!id || !atividade) return;

    etapas.push({
      ordem,
      id,
      atividade,
      tipo,
      sistema,
      tempo,
      proxSim,
      proxNao,
      conexoesExtras,
      coluna,
      linha,
      categoriaOportunidade,
      percentualReducao,
      potencialReducao: percentualReducao,
      observacao,
      cor
    });

    idsValidos.add(id);
  });

  etapas.sort((a, b) => a.ordem - b.ordem);

  if (!etapas.length) {
    alert("Nenhuma etapa válida foi encontrada na tabela.");
    return;
  }

  ultimoNomeArquivo = gerarNomeArquivo();

  const alturaPadraoNos = calcularAlturaPadraoNos(etapas);
  const maiorAlturaLosango = Math.ceil(alturaPadraoNos * CONFIG.decisionHeightFactor);
  const rowSlotHeight = Math.max(alturaPadraoNos, maiorAlturaLosango);

  const etapaPorId = {};
  etapas.forEach((e) => {
    etapaPorId[e.id] = e;
  });

  let maxCol = 1;
  let maxRow = 1;

  etapas.forEach((etapa) => {
    if (etapa.coluna > maxCol) maxCol = etapa.coluna;
    if (etapa.linha > maxRow) maxRow = etapa.linha;
  });

  const svgWidth = CONFIG.marginX * 2 + maxCol * CONFIG.boxWidth + (maxCol - 1) * CONFIG.colGap + 300;
  const svgHeight = CONFIG.marginY * 2 + maxRow * rowSlotHeight + (maxRow - 1) * CONFIG.rowGap + 220;

  const svg = criarElementoSVG("svg");
  svg.setAttribute("width", svgWidth);
  svg.setAttribute("height", svgHeight);
  svg.setAttribute("viewBox", `0 0 ${svgWidth} ${svgHeight}`);
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("id", "fluxogramaSVG");

  criarMarkerArrow(svg);

  const posicoes = {};

  etapas.forEach((etapa) => {
    const width = obterLarguraNo(etapa);
    const height = obterAlturaNo(etapa, alturaPadraoNos);

    const x = CONFIG.marginX + (etapa.coluna - 1) * (CONFIG.boxWidth + CONFIG.colGap);
    const yBase = CONFIG.marginY + (etapa.linha - 1) * (rowSlotHeight + CONFIG.rowGap);
    const y = yBase + (rowSlotHeight - height) / 2;

    posicoes[etapa.id] = {
      id: etapa.id,
      x,
      y,
      width,
      height,
      gridCol: etapa.coluna,
      gridRow: etapa.linha,
      isDecision: isPergunta(etapa.atividade)
    };
  });

  const primeiraEtapa = etapas[0];
  const ultimaEtapa = etapas[etapas.length - 1];

  posicoes["__INICIO__"] = {
    id: "__INICIO__",
    x: posicoes[primeiraEtapa.id].x - 150,
    y: posicoes[primeiraEtapa.id].y + (posicoes[primeiraEtapa.id].height / 2) - 18,
    width: 80,
    height: 36,
    gridCol: primeiraEtapa.coluna - 1,
    gridRow: primeiraEtapa.linha,
    isDecision: false
  };

  posicoes["__FIM__"] = {
    id: "__FIM__",
    x: posicoes[ultimaEtapa.id].x + posicoes[ultimaEtapa.id].width + 70,
    y: posicoes[ultimaEtapa.id].y + (posicoes[ultimaEtapa.id].height / 2) - 18,
    width: 60,
    height: 36,
    gridCol: ultimaEtapa.coluna + 1,
    gridRow: ultimaEtapa.linha,
    isDecision: false
  };

  desenharCapsula(svg, "Início", posicoes["__INICIO__"].x, posicoes["__INICIO__"].y, posicoes["__INICIO__"].width, posicoes["__INICIO__"].height);
  desenharCapsula(svg, "Fim", posicoes["__FIM__"].x, posicoes["__FIM__"].y, posicoes["__FIM__"].width, posicoes["__FIM__"].height);

  etapas.forEach((etapa) => {
    const pos = posicoes[etapa.id];
    desenharNo(svg, etapa, pos.x, pos.y, pos.width, pos.height, corHex(etapa.cor));
  });

  const sharedRegistry = {};
  let loops = 0;
  let conexoesExtrasCount = 0;
  let decisoes = 0;
  const etapasImpactadasRetrabalho = new Set();
  const tiposTempo = {};
  const sistemasTempo = {};
  let tempoTotal = 0;

  const desenharListaConexoes = (origemEtapa, destinoIds, rotulo) => {
    destinoIds.forEach((destinoId, index) => {
      if (!destinoEhValido(destinoId, idsValidos)) return;

      const origem = posicoes[origemEtapa.id];
      const destino = posicoes[destinoId];
      if (!origem || !destino) return;

      desenharConexao(svg, origem, destino, rotulo, index, posicoes, sharedRegistry);

      const destinoEtapa = etapaPorId[destinoId];
      if (destinoEtapa && destinoEtapa.ordem < origemEtapa.ordem) {
        loops++;
        adicionarEtapasImpactadasPorRetorno(origemEtapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
      }
    });
  };

  desenharConexao(svg, posicoes["__INICIO__"], posicoes[primeiraEtapa.id], "", 0, posicoes, sharedRegistry);

  etapas.forEach((etapa) => {
    tempoTotal += etapa.tempo;
    if (isPergunta(etapa.atividade)) decisoes++;

    if (!tiposTempo[etapa.tipo]) tiposTempo[etapa.tipo] = 0;
    tiposTempo[etapa.tipo] += etapa.tempo;

    if (!sistemasTempo[etapa.sistema]) sistemasTempo[etapa.sistema] = 0;
    sistemasTempo[etapa.sistema] += etapa.tempo;

    const destinosSim = quebrarListaIds(etapa.proxSim);
    const destinosNao = quebrarListaIds(etapa.proxNao);
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras);

    desenharListaConexoes(etapa, destinosSim, isPergunta(etapa.atividade) ? "Sim" : "");
    desenharListaConexoes(etapa, destinosNao, isPergunta(etapa.atividade) ? "Não" : "");
    desenharListaConexoes(etapa, destinosExtras, "");

    conexoesExtrasCount += destinosExtras.length;
  });

  desenharConexao(svg, posicoes[ultimaEtapa.id], posicoes["__FIM__"], "", 0, posicoes, sharedRegistry);

  document.getElementById("diagram").innerHTML = "";
  document.getElementById("diagram").appendChild(svg);

  const atividadesTempo = etapas.map((etapa) => ({
    atividade: etapa.atividade,
    tempo: etapa.tempo
  })).sort((a, b) => b.tempo - a.tempo);

  const tiposOrdenados = Object.entries(tiposTempo).map(([nome, tempo]) => ({ nome, tempo })).sort((a, b) => b.tempo - a.tempo);
  const sistemasOrdenados = Object.entries(sistemasTempo).map(([nome, tempo]) => ({ nome, tempo })).sort((a, b) => b.tempo - a.tempo);

  let tempoPotencialRetrabalho = 0;
  etapas.forEach((etapa) => {
    if (etapasImpactadasRetrabalho.has(etapa.id)) tempoPotencialRetrabalho += etapa.tempo;
  });

  const impactoPotencialRetrabalhoNum = tempoTotal ? (tempoPotencialRetrabalho / tempoTotal) * 100 : 0;
  const taxaDecisaoNum = etapas.length ? (decisoes / etapas.length) * 100 : 0;

  const infoProcessoData = { desenho, processo, analista, negocio, area, gestor };
  document.getElementById("infoProcesso").innerHTML = renderInformacoesProcessoExecutivas(infoProcessoData);

  const dadosAnalise = {
    tempoTotal,
    loops,
    conexoesExtrasCount,
    tempoPotencialRetrabalho,
    impactoPotencialRetrabalho: impactoPotencialRetrabalhoNum,
    decisoes,
    taxaDecisao: taxaDecisaoNum,
    tempoPorTipo: tiposOrdenados.map(item => ({
      tipo: item.nome,
      tempo: item.tempo,
      percentual: tempoTotal ? (item.tempo / tempoTotal) * 100 : 0
    })),
    tempoPorSistema: sistemasOrdenados.map(item => ({
      sistema: item.nome,
      tempo: item.tempo,
      percentual: tempoTotal ? (item.tempo / tempoTotal) * 100 : 0
    })),
    pareto: (() => {
      let acumulado = 0;
      return atividadesTempo.map(item => {
        const percentual = tempoTotal ? (item.tempo / tempoTotal) * 100 : 0;
        acumulado += percentual;
        return {
          atividade: item.atividade,
          tempo: item.tempo,
          percentual,
          pareto: acumulado
        };
      });
    })(),
    simulacaoMelhoria: montarDadosSimulacaoMelhoria(etapas)
  };

  document.getElementById("metricas").innerHTML = renderizarAnaliseExecutiva(dadosAnalise);
}

function obterSVGPronto() {
  const svgOriginal = document.querySelector("#diagram svg");
  if (!svgOriginal) {
    alert("Gere o fluxo primeiro.");
    return null;
  }
  return svgOriginal.cloneNode(true);
}

function coletarDadosAnaliseEstruturados() {
  const etapas = [];
  const texto = document.getElementById("entrada").value;

  const linhasBrutas = texto.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(l => l.trim() !== "");
  let linhas = linhasBrutas.map(l => l.split("\t"));

  if (linhas.length && ehCabecalho(linhas[0])) linhas.shift();

  linhas.forEach((col) => {
    while (col.length < 15) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]) || "Não informado";
    const sistema = limpar(col[4]) || "Sem sistema informado";
    const tempo = tempoParaSegundos(limpar(col[5]));
    const proxSim = limpar(col[6]);
    const proxNao = limpar(col[7]);
    const conexoesExtras = limpar(col[8]);
    const categoriaOportunidade = normalizarCategoriaOportunidade(col[11]);
    const percentualReducao = parsePercentual(col[12]);
    const observacao = limpar(col[13]);

    if (!id || !atividade) return;

    etapas.push({
      ordem,
      id,
      atividade,
      tipo,
      sistema,
      tempo,
      proxSim,
      proxNao,
      conexoesExtras,
      categoriaOportunidade,
      percentualReducao,
      potencialReducao: percentualReducao,
      observacao
    });
  });

  etapas.sort((a, b) => a.ordem - b.ordem);

  const etapaPorId = {};
  etapas.forEach(e => { etapaPorId[e.id] = e; });

  let tempoTotal = 0;
  let loops = 0;
  let conexoesExtrasCount = 0;
  let decisoes = 0;
  const etapasImpactadasRetrabalho = new Set();
  const tiposTempo = {};
  const sistemasTempo = {};

  etapas.forEach((etapa) => {
    tempoTotal += etapa.tempo;
    if (isPergunta(etapa.atividade)) decisoes++;

    if (!tiposTempo[etapa.tipo]) tiposTempo[etapa.tipo] = 0;
    tiposTempo[etapa.tipo] += etapa.tempo;

    if (!sistemasTempo[etapa.sistema]) sistemasTempo[etapa.sistema] = 0;
    sistemasTempo[etapa.sistema] += etapa.tempo;

    const destinosSim = quebrarListaIds(etapa.proxSim);
    const destinosNao = quebrarListaIds(etapa.proxNao);
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras);

    destinosSim.forEach((destinoId) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
      }
    });

    destinosNao.forEach((destinoId) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
      }
    });

    conexoesExtrasCount += destinosExtras.length;

    destinosExtras.forEach((destinoId) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
      }
    });
  });

  let tempoPotencialRetrabalho = 0;
  etapas.forEach((etapa) => {
    if (etapasImpactadasRetrabalho.has(etapa.id)) tempoPotencialRetrabalho += etapa.tempo;
  });

  const impactoPotencialRetrabalho = tempoTotal ? (tempoPotencialRetrabalho / tempoTotal) * 100 : 0;
  const taxaDecisao = etapas.length ? (decisoes / etapas.length) * 100 : 0;

  const tempoPorTipo = Object.entries(tiposTempo)
    .map(([tipo, tempo]) => ({
      tipo,
      tempo,
      percentual: tempoTotal ? (tempo / tempoTotal) * 100 : 0
    }))
    .sort((a, b) => b.tempo - a.tempo);

  const tempoPorSistema = Object.entries(sistemasTempo)
    .map(([sistema, tempo]) => ({
      sistema,
      tempo,
      percentual: tempoTotal ? (tempo / tempoTotal) * 100 : 0
    }))
    .sort((a, b) => b.tempo - a.tempo);

  const pareto = [...etapas]
    .sort((a, b) => b.tempo - a.tempo)
    .map((e) => ({ atividade: e.atividade, tempo: e.tempo }));

  let acumulado = 0;
  pareto.forEach((item) => {
    item.percentual = tempoTotal ? (item.tempo / tempoTotal) * 100 : 0;
    acumulado += item.percentual;
    item.pareto = acumulado;
  });

  return {
    tempoTotal,
    loops,
    conexoesExtrasCount,
    tempoPotencialRetrabalho,
    impactoPotencialRetrabalho,
    decisoes,
    taxaDecisao,
    tempoPorTipo,
    tempoPorSistema,
    pareto,
    simulacaoMelhoria: montarDadosSimulacaoMelhoria(etapas)
  };
}

function extrairLinhasInfoProcesso() {
  const el = document.getElementById("infoProcesso");
  if (!el) return [];

  const labels = el.querySelectorAll(".exec-info-item");
  if (labels.length) {
    return Array.from(labels).map(item => {
      const label = item.querySelector(".exec-info-label")?.innerText?.trim() || "";
      const value = item.querySelector(".exec-info-value")?.innerText?.trim() || "";
      return `${label}: ${value}`;
    });
  }

  return el.innerText.split("\n").map(l => l.trim()).filter(Boolean);
}

function adicionarTextoQuebrado(doc, texto, x, y, maxWidth, lineHeight = 14, options = {}) {
  const linhas = doc.splitTextToSize(String(texto || ""), maxWidth);
  if (!linhas.length) return y;
  doc.text(linhas, x, y, options);
  return y + linhas.length * lineHeight;
}

function limparTextoPDF(txt) {
  return String(txt || "")
    .replace(/⏱/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getPageSpec(doc) {
  return {
    pageWidth: doc.internal.pageSize.getWidth(),
    pageHeight: doc.internal.pageSize.getHeight()
  };
}

function garantirEspacoPagina(doc, yAtual, alturaNecessaria, margem, onNewPage = null) {
  const { pageHeight } = getPageSpec(doc);
  if (yAtual + alturaNecessaria > pageHeight - margem) {
    doc.addPage();
    let novoY = margem;
    if (typeof onNewPage === "function") {
      novoY = onNewPage(novoY);
    }
    return novoY;
  }
  return yAtual;
}

function iniciarNovaPaginaPaisagem(doc, titulo = "") {
  doc.addPage("a4", "landscape");
  const margem = 36;
  let y = margem;
  const { pageWidth } = getPageSpec(doc);

  if (titulo) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(17, 24, 39);
    doc.text(titulo, margem, y);
    y += 16;
  }

  return {
    margem,
    y,
    larguraUtil: pageWidth - margem * 2
  };
}

function desenharLegendaPDF(doc, x, y) {
  const itens = [
    { cor: "#95d5b2", texto: "Verde: Oportunidade de automação" },
    { cor: "#8ecae6", texto: "Azul: Oportunidade de melhoria de processo" },
    { cor: "#ffd166", texto: "Amarelo: Já automatizado" },
    { cor: "#ffffff", texto: "Branco: Sem oportunidade clara" }
  ];

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(17, 24, 39);
  doc.text("Legenda", x, y);

  let cursorY = y + 14;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);

  itens.forEach((item) => {
    const rgb = hexToRgb(item.cor);
    doc.setDrawColor(120);
    doc.setFillColor(rgb.r, rgb.g, rgb.b);
    doc.rect(x, cursorY - 8, 10, 10, "FD");
    doc.setTextColor(31, 41, 55);
    doc.text(item.texto, x + 18, cursorY);
    cursorY += 16;
  });

  return cursorY + 6;
}

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16)
  };
}

function desenharCardsResumoPDF(doc, dados, x, y, larguraUtil, margem) {
  const cards = [
    { label: "Tempo total do processo", value: formatarTempo(dados.tempoTotal) },
    { label: "Loops detectados", value: String(dados.loops) },
    { label: "Potencial retrabalho", value: `${formatarTempo(dados.tempoPotencialRetrabalho)} | ${formatarPercentual(dados.impactoPotencialRetrabalho)}%` },
    { label: "Taxa de decisão", value: `${dados.decisoes} etapa(s) | ${formatarPercentual(dados.taxaDecisao)}%` },
    { label: "Ganho potencial em horas", value: formatarTempo(dados.simulacaoMelhoria.ganhoPotencialHoras) },
    { label: "Tempo total To Be", value: formatarTempo(dados.simulacaoMelhoria.tempoTotalToBe) },
    { label: "Eficiência potencial", value: `${formatarPercentual(dados.simulacaoMelhoria.eficienciaPotencial)}%` },
    { label: "Atividades na simulação", value: String(dados.simulacaoMelhoria.quantidadeAtividades) }
  ];

  const gap = 12;
  const cols = 2;
  const cardWidth = (larguraUtil - gap) / cols;
  const cardHeight = 54;
  let cursorY = y;

  for (let i = 0; i < cards.length; i += cols) {
    cursorY = garantirEspacoPagina(doc, cursorY, cardHeight + 4, margem);

    for (let c = 0; c < cols; c++) {
      const item = cards[i + c];
      if (!item) continue;

      const cardX = x + c * (cardWidth + gap);
      doc.setDrawColor(219, 227, 238);
      doc.setFillColor(248, 250, 252);
      doc.roundedRect(cardX, cursorY, cardWidth, cardHeight, 8, 8, "FD");

      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      doc.setTextColor(100, 116, 139);
      doc.text(item.label, cardX + 10, cursorY + 16);

      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(17, 24, 39);
      const linhas = doc.splitTextToSize(item.value, cardWidth - 20);
      doc.text(linhas, cardX + 10, cursorY + 34);
    }

    cursorY += cardHeight + gap;
  }

  return cursorY + 4;
}

function desenharTabelaPDF(doc, config) {
  const {
    titulo,
    columns,
    rows,
    x,
    yInicial,
    larguraTotal,
    margem,
    pageTitle = "",
    fontSize = 9
  } = config;

  const borderWidth = 0.5;
  const cellPaddingX = 5;
  const lineHeight = fontSize + 2;
  const minRowHeight = 20;
  const gapAntesTitulo = 14;
  const gapTituloCabecalho = 6;
  const gapDepoisTabela = 20;

  let y = yInicial + gapAntesTitulo;

  const weights = columns.map(col => col.weight || 1);
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  const colWidths = weights.map(w => (larguraTotal * w) / totalWeight);

  const buildHeaderHeight = () => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(fontSize);

    let maxLines = 1;
    columns.forEach((col, i) => {
      const linhas = doc.splitTextToSize(col.header, colWidths[i] - cellPaddingX * 2);
      if (linhas.length > maxLines) maxLines = linhas.length;
    });

    return Math.max(minRowHeight, maxLines * lineHeight + 10);
  };

  const headerHeight = buildHeaderHeight();

  const drawTitleAndHeader = (yStart) => {
    if (pageTitle) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.setTextColor(17, 24, 39);
      doc.text(pageTitle, margem, margem);
    }

    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(17, 24, 39);
    doc.text(titulo, x, yStart);

    let yHeader = yStart + gapTituloCabecalho;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(fontSize);
    doc.setFillColor(234, 241, 251);
    doc.setDrawColor(214, 224, 238);
    doc.setLineWidth(borderWidth);

    let currentX = x;
    columns.forEach((col, i) => {
      const width = colWidths[i];
      const linhas = doc.splitTextToSize(col.header, width - cellPaddingX * 2);

      doc.rect(currentX, yHeader, width, headerHeight, "FD");

      const totalTextHeight = linhas.length * lineHeight;
      const startY = yHeader + (headerHeight - totalTextHeight) / 2 + fontSize;

      linhas.forEach((linha, idx) => {
        doc.text(linha, currentX + width / 2, startY + idx * lineHeight, { align: "center" });
      });

      currentX += width;
    });

    return yHeader + headerHeight;
  };

  y = garantirEspacoPagina(doc, y, 18 + gapTituloCabecalho + headerHeight, margem, (novoY) => {
    if (pageTitle) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.setTextColor(17, 24, 39);
      doc.text(pageTitle, margem, margem);
      return margem + 16;
    }
    return novoY;
  });

  y = drawTitleAndHeader(y);

  rows.forEach((row) => {
    doc.setFont(row.isTotal ? "helvetica" : "helvetica", row.isTotal ? "bold" : "normal");
    doc.setFontSize(fontSize);
    doc.setTextColor(31, 41, 55);

    let maxLines = 1;
    const rowLineCache = columns.map((col, i) => {
      const raw = row[col.key] !== undefined && row[col.key] !== null ? String(row[col.key]) : "";
      const valor = limparTextoPDF(raw);
      const align = col.align || "left";
      const linhas = align === "left"
        ? doc.splitTextToSize(valor, colWidths[i] - cellPaddingX * 2)
        : [valor];

      if (linhas.length > maxLines) maxLines = linhas.length;
      return linhas;
    });

    const rowHeight = Math.max(minRowHeight, maxLines * lineHeight + 10);

    y = garantirEspacoPagina(doc, y, rowHeight, margem, (novoY) => {
      let baseY = novoY;
      if (pageTitle) {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(13);
        doc.setTextColor(17, 24, 39);
        doc.text(pageTitle, margem, margem);
        baseY = margem + 16;
      }
      return drawTitleAndHeader(baseY);
    });

    doc.setLineWidth(borderWidth);
    doc.setDrawColor(230, 236, 243);

    if (row.isTotal) doc.setFillColor(238, 247, 240);
    else doc.setFillColor(255, 255, 255);

    let currentX = x;

    columns.forEach((col, i) => {
      const width = colWidths[i];
      const align = col.align || "left";
      const linhas = rowLineCache[i];

      doc.rect(currentX, y, width, rowHeight, "FD");

      const totalTextHeight = linhas.length * lineHeight;
      const startY = y + (rowHeight - totalTextHeight) / 2 + fontSize;

      if (align === "center") {
        linhas.forEach((linha, idx) => {
          doc.text(linha, currentX + width / 2, startY + idx * lineHeight, { align: "center" });
        });
      } else if (align === "right") {
        linhas.forEach((linha, idx) => {
          doc.text(linha, currentX + width - cellPaddingX, startY + idx * lineHeight, { align: "right" });
        });
      } else {
        linhas.forEach((linha, idx) => {
          doc.text(linha, currentX + cellPaddingX, startY + idx * lineHeight);
        });
      }

      currentX += width;
    });

    y += rowHeight;
  });

  return y + gapDepoisTabela;
}

async function baixarAnalisePDF() {
  const svg = obterSVGPronto();
  if (!svg) return;

  const infoLinhas = extrairLinhasInfoProcesso();
  const dados = coletarDadosAnaliseEstruturados();
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF({
    orientation: "p",
    unit: "pt",
    format: "a4"
  });

  const margem = 40;
  let { pageWidth } = getPageSpec(doc);
  let larguraUtil = pageWidth - margem * 2;
  let y = margem;

  const processoNome = obterValorCampo("processo") || "Fluxograma do Processo";

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  y = adicionarTextoQuebrado(doc, processoNome, margem, y, larguraUtil, 18);
  y += 8;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(17, 24, 39);
  y = garantirEspacoPagina(doc, y, 24, margem);
  doc.text("Informações do Processo", margem, y);
  y += 16;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(31, 41, 55);

  infoLinhas.forEach((linha) => {
    y = garantirEspacoPagina(doc, y, 18, margem);
    y = adicionarTextoQuebrado(doc, linha, margem, y, larguraUtil, 13);
  });

  y += 10;
  y = garantirEspacoPagina(doc, y, 90, margem);
  y = desenharLegendaPDF(doc, margem, y);

  const svgWidth = Number(svg.getAttribute("width")) || 1200;
  const svgHeight = Number(svg.getAttribute("height")) || 800;
  const escala = Math.min(larguraUtil / svgWidth, 1);
  const larguraSvgPdf = svgWidth * escala;
  const alturaSvgPdf = svgHeight * escala;

  y = garantirEspacoPagina(doc, y, alturaSvgPdf + 20, margem);

  await doc.svg(svg, {
    x: margem + (larguraUtil - larguraSvgPdf) / 2,
    y,
    width: larguraSvgPdf,
    height: alturaSvgPdf
  });

  y += alturaSvgPdf + 24;
  y = garantirEspacoPagina(doc, y, 28, margem);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(17, 24, 39);
  doc.text("Análise do Processo", margem, y);
  y += 16;

  y = desenharCardsResumoPDF(doc, dados, margem, y, larguraUtil, margem);

  y = desenharTabelaPDF(doc, {
    titulo: "Tempo por Tipo",
    columns: [
      { header: "Tipo", key: "tipo", weight: 2.8, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.2, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.0, align: "center" }
    ],
    rows: dados.tempoPorTipo.map(item => ({
      tipo: item.tipo,
      tempoFmt: formatarTempo(item.tempo),
      percentualFmt: `${formatarPercentual(item.percentual)}%`
    })),
    x: margem,
    yInicial: y,
    larguraTotal: larguraUtil,
    margem,
    pageTitle: "Análise do Processo",
    fontSize: 9
  });

  y = desenharTabelaPDF(doc, {
    titulo: "Tempo por Sistema",
    columns: [
      { header: "Sistema", key: "sistema", weight: 2.8, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.2, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.0, align: "center" }
    ],
    rows: dados.tempoPorSistema.map(item => ({
      sistema: item.sistema,
      tempoFmt: formatarTempo(item.tempo),
      percentualFmt: `${formatarPercentual(item.percentual)}%`
    })),
    x: margem,
    yInicial: y,
    larguraTotal: larguraUtil,
    margem,
    pageTitle: "Análise do Processo",
    fontSize: 9
  });

  // Pareto em landscape para respirar melhor
  let landscape = iniciarNovaPaginaPaisagem(doc, "Análise do Processo");
  let margemL = landscape.margem;
  let yL = landscape.y;
  let larguraUtilL = landscape.larguraUtil;

  yL = desenharTabelaPDF(doc, {
    titulo: "Pareto de Tempo",
    columns: [
      { header: "Atividade", key: "atividade", weight: 4.2, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.2, align: "center" },
      { header: "%", key: "percentualFmt", weight: 0.9, align: "center" },
      { header: "% Acumulado", key: "paretoFmt", weight: 1.2, align: "center" }
    ],
    rows: dados.pareto.map(item => ({
      atividade: item.atividade,
      tempoFmt: formatarTempo(item.tempo),
      percentualFmt: `${formatarPercentual(item.percentual)}%`,
      paretoFmt: `${formatarPercentual(item.pareto)}%`
    })),
    x: margemL,
    yInicial: yL,
    larguraTotal: larguraUtilL,
    margem: margemL,
    pageTitle: "Análise do Processo",
    fontSize: 8.5
  });

  // Simulação em landscape com colunas mais largas
  landscape = iniciarNovaPaginaPaisagem(doc, "Análise do Processo");
  margemL = landscape.margem;
  yL = landscape.y;
  larguraUtilL = landscape.larguraUtil;

  yL = desenharTabelaPDF(doc, {
    titulo: "Simulação de Melhoria (As Is vs To Be)",
    columns: [
      { header: "Atividade", key: "atividade", weight: 4.4, align: "left" },
      { header: "As Is", key: "tempoAsIsFmt", weight: 1.0, align: "center" },
      { header: "% Red.", key: "reducaoFmt", weight: 0.9, align: "center" },
      { header: "Categoria", key: "categoriaFmt", weight: 2.1, align: "left" },
      { header: "Ganho", key: "ganhoFmt", weight: 1.0, align: "center" },
      { header: "To Be", key: "tempoToBeFmt", weight: 1.0, align: "center" },
      { header: "Observação", key: "observacao", weight: 3.2, align: "left" }
    ],
    rows: [
      ...dados.simulacaoMelhoria.rows.map(item => ({
        atividade: item.atividade,
        tempoAsIsFmt: formatarTempo(item.tempo),
        reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
        categoriaFmt: item.categoriaOportunidade || "Sem oportunidade",
        ganhoFmt: formatarTempo(item.ganhoPotencial),
        tempoToBeFmt: formatarTempo(item.tempoToBe),
        observacao: item.observacao || "-"
      })),
      {
        atividade: `${dados.simulacaoMelhoria.quantidadeAtividades} atividade(s)`,
        tempoAsIsFmt: formatarTempo(dados.simulacaoMelhoria.tempoTotalAsIs),
        reducaoFmt: `${formatarPercentual(dados.simulacaoMelhoria.eficienciaPotencial)}%`,
        categoriaFmt: "-",
        ganhoFmt: formatarTempo(dados.simulacaoMelhoria.ganhoPotencialHoras),
        tempoToBeFmt: formatarTempo(dados.simulacaoMelhoria.tempoTotalToBe),
        observacao: `Eficiência total: ${formatarPercentual(dados.simulacaoMelhoria.eficienciaPotencial)}%`,
        isTotal: true
      }
    ],
    x: margemL,
    yInicial: yL,
    larguraTotal: larguraUtilL,
    margem: margemL,
    pageTitle: "Análise do Processo",
    fontSize: 8.3
  });

  if (dados.simulacaoMelhoria.ranking.length) {
    landscape = iniciarNovaPaginaPaisagem(doc, "Análise do Processo");
    margemL = landscape.margem;
    yL = landscape.y;
    larguraUtilL = landscape.larguraUtil;

    yL = desenharTabelaPDF(doc, {
      titulo: "Ranking das Top Atividades por Ganho Potencial",
      columns: [
        { header: "#", key: "ranking", weight: 0.6, align: "center" },
        { header: "Atividade", key: "atividade", weight: 4.0, align: "left" },
        { header: "Ganho", key: "ganhoFmt", weight: 1.2, align: "center" },
        { header: "% Red.", key: "reducaoFmt", weight: 1.0, align: "center" },
        { header: "Categoria", key: "categoria", weight: 2.2, align: "left" }
      ],
      rows: dados.simulacaoMelhoria.ranking.map((item, index) => ({
        ranking: String(index + 1),
        atividade: item.atividade,
        ganhoFmt: formatarTempo(item.ganhoPotencial),
        reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
        categoria: item.categoriaOportunidade || "Sem oportunidade"
      })),
      x: margemL,
      yInicial: yL,
      larguraTotal: larguraUtilL,
      margem: margemL,
      pageTitle: "Análise do Processo",
      fontSize: 8.8
    });
  }

  doc.save(`${ultimoNomeArquivo}_analise.pdf`);
}

function baixarSVG() {
  const svg = obterSVGPronto();
  if (!svg) return;

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const blob = new Blob([source], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = `${ultimoNomeArquivo}.svg`;
  link.click();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function baixarFluxo() {
  baixarSVG();
}