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
  decisionHeightFactor: 1.30,
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
  const permitidas = ["blue", "yellow", "green", "red", "white"];
  return permitidas.includes(c) ? c : "white";
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

  const linhasAtividade = quebrarTextoPorLargura(
    etapa.atividade,
    larguraTexto,
    CONFIG.fontSize
  );

  const linhasSistema = quebrarTextoPorLargura(
    etapa.sistema || "Sem sistema informado",
    larguraTexto,
    CONFIG.fontSize
  );

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
    if (altura > maiorAltura) {
      maiorAltura = altura;
    }
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
  limparCampo("processo");
  limparCampo("analista");
  limparCampo("negocio");
  limparCampo("area");
  limparCampo("coordenador");
  limparCampo("entrada");

  const diagram = document.getElementById("diagram");
  const infoProcesso = document.getElementById("infoProcesso");
  const metricas = document.getElementById("metricas");

  if (diagram) diagram.innerHTML = "";
  if (infoProcesso) infoProcesso.innerHTML = "";
  if (metricas) metricas.innerHTML = "";

  ultimoNomeArquivo = "fluxograma_processo";
}

function adicionarLoopSeNecessario(origemId, destinoId, etapaPorId, etapaAtual, etapasOrigemComRetornoRef) {
  const destino = etapaPorId[destinoId];
  if (destino && destino.ordem < etapaAtual.ordem) {
    etapasOrigemComRetornoRef.add(origemId);
    return 1;
  }
  return 0;
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
  rect.setAttribute("stroke-width", "2");
  g.appendChild(rect);

  const t = criarElementoSVG("text");
  t.setAttribute("x", x + width / 2);
  t.setAttribute("y", y + height / 2 + 5);
  t.setAttribute("text-anchor", "middle");
  t.setAttribute("font-family", CONFIG.fontFamily);
  t.setAttribute("font-size", "14");
  t.setAttribute("fill", "#111111");
  t.textContent = texto;
  g.appendChild(t);

  svg.appendChild(g);
}

function desenharNo(svg, etapa, pos) {
  const g = criarElementoSVG("g");
  const fill = corHex(etapa.cor);
  const pergunta = isPergunta(etapa.atividade);

  if (pergunta) {
    const cx = pos.x + pos.w / 2;
    const cy = pos.y + pos.h / 2;

    const polygon = criarElementoSVG("polygon");
    polygon.setAttribute(
      "points",
      `${cx},${pos.y} ${pos.x + pos.w},${cy} ${cx},${pos.y + pos.h} ${pos.x},${cy}`
    );
    polygon.setAttribute("fill", fill);
    polygon.setAttribute("stroke", CONFIG.stroke);
    polygon.setAttribute("stroke-width", "2");
    g.appendChild(polygon);
  } else {
    const rect = criarElementoSVG("rect");
    rect.setAttribute("x", pos.x);
    rect.setAttribute("y", pos.y);
    rect.setAttribute("width", pos.w);
    rect.setAttribute("height", pos.h);
    rect.setAttribute("fill", fill);
    rect.setAttribute("stroke", CONFIG.stroke);
    rect.setAttribute("stroke-width", "2");
    rect.setAttribute("rx", CONFIG.cornerRadius);
    rect.setAttribute("ry", CONFIG.cornerRadius);
    g.appendChild(rect);
  }

  const linhas = obterLinhasEtapa(etapa, pos.w);
  const totalAlturaTexto = linhas.length * CONFIG.textLineHeight;
  const inicioYTexto = pos.y + (pos.h - totalAlturaTexto) / 2 + 14;

  linhas.forEach((linha, i) => {
    const text = criarElementoSVG("text");
    text.setAttribute("x", pos.x + pos.w / 2);
    text.setAttribute("y", inicioYTexto + i * CONFIG.textLineHeight);
    text.setAttribute("text-anchor", "middle");
    text.setAttribute("font-family", CONFIG.fontFamily);
    text.setAttribute("font-size", i === linhas.length - 1 ? CONFIG.smallFontSize : CONFIG.fontSize);
    text.setAttribute("fill", "#111111");
    text.textContent = linha;
    g.appendChild(text);
  });

  svg.appendChild(g);
}

function getAnchorPoint(node, side) {
  switch (side) {
    case "right":
      return { x: node.x + node.w, y: node.y + node.h / 2 };
    case "left":
      return { x: node.x, y: node.y + node.h / 2 };
    case "top":
      return { x: node.x + node.w / 2, y: node.y };
    case "bottom":
      return { x: node.x + node.w / 2, y: node.y + node.h };
    default:
      return { x: node.x + node.w, y: node.y + node.h / 2 };
  }
}

function createPolylinePath(points) {
  if (!points.length) return "";
  let d = `M ${points[0].x} ${points[0].y}`;
  for (let i = 1; i < points.length; i++) {
    d += ` L ${points[i].x} ${points[i].y}`;
  }
  return d;
}

function normalizarPontos(points) {
  const resultado = [];

  points.forEach((point) => {
    if (!point) return;

    const p = {
      x: Number(point.x.toFixed(1)),
      y: Number(point.y.toFixed(1))
    };

    const ultimo = resultado[resultado.length - 1];
    if (ultimo && ultimo.x === p.x && ultimo.y === p.y) return;
    resultado.push(p);
  });

  if (resultado.length <= 2) return resultado;

  const simplificado = [resultado[0]];

  for (let i = 1; i < resultado.length - 1; i++) {
    const anterior = simplificado[simplificado.length - 1];
    const atual = resultado[i];
    const proximo = resultado[i + 1];

    const mesmoX = anterior.x === atual.x && atual.x === proximo.x;
    const mesmoY = anterior.y === atual.y && atual.y === proximo.y;

    if (!mesmoX && !mesmoY) simplificado.push(atual);
  }

  simplificado.push(resultado[resultado.length - 1]);
  return simplificado;
}

function segmentoInterceptaAreaExpandida(p1, p2, left, top, right, bottom) {
  const minX = Math.min(p1.x, p2.x);
  const maxX = Math.max(p1.x, p2.x);
  const minY = Math.min(p1.y, p2.y);
  const maxY = Math.max(p1.y, p2.y);

  const horizontal = p1.y === p2.y;
  const vertical = p1.x === p2.x;

  if (horizontal) {
    return p1.y >= top && p1.y <= bottom && maxX >= left && minX <= right;
  }

  if (vertical) {
    return p1.x >= left && p1.x <= right && maxY >= top && minY <= bottom;
  }

  return false;
}

function pathCruzaCaixas(points, posicoes = {}, excludeIds = []) {
  const nodes = Object.values(posicoes).filter(node => !excludeIds.includes(node.id));

  for (let i = 0; i < points.length - 1; i++) {
    const p1 = points[i];
    const p2 = points[i + 1];

    for (const node of nodes) {
      const padX = CONFIG.obstaclePadding;
      const padY = CONFIG.obstaclePadding;
      const left = node.x - padX;
      const right = node.x + node.w + padX;
      const top = node.y - padY;
      const bottom = node.y + node.h + padY;

      if (segmentoInterceptaAreaExpandida(p1, p2, left, top, right, bottom)) {
        return true;
      }
    }
  }

  return false;
}

function calcularComprimento(points) {
  let total = 0;
  for (let i = 0; i < points.length - 1; i++) {
    total += Math.abs(points[i + 1].x - points[i].x) + Math.abs(points[i + 1].y - points[i].y);
  }
  return total;
}

function detectarLadoEntrada(points) {
  if (!points || points.length < 2) return "left";
  const prev = points[points.length - 2];
  const end = points[points.length - 1];

  if (prev.x < end.x) return "left";
  if (prev.x > end.x) return "right";
  if (prev.y < end.y) return "top";
  return "bottom";
}

function detectarLadoSaida(points) {
  if (!points || points.length < 2) return "right";
  const start = points[0];
  const next = points[1];

  if (next.x > start.x) return "right";
  if (next.x < start.x) return "left";
  if (next.y > start.y) return "bottom";
  return "top";
}

function ajustarUltimoTrechoParaLado(points, end, side, destinoNode = null) {
  if (!points || points.length < 2) return points;

  const resultado = [...points];
  const prev = resultado[resultado.length - 2];
  const destino = { x: end.x, y: end.y };

  if (side === "left" && Math.abs(prev.y - end.y) <= CONFIG.sameRowTolerance && prev.x <= (destinoNode ? destinoNode.x : end.x)) {
    resultado[resultado.length - 1] = destino;
    return normalizarPontos(resultado);
  }

  if (side === "right" && Math.abs(prev.y - end.y) <= CONFIG.sameRowTolerance && prev.x >= (destinoNode ? destinoNode.x + destinoNode.w : end.x)) {
    resultado[resultado.length - 1] = destino;
    return normalizarPontos(resultado);
  }

  if (side === "top" && Math.abs(prev.x - end.x) <= CONFIG.sameColTolerance && prev.y <= (destinoNode ? destinoNode.y : end.y)) {
    resultado[resultado.length - 1] = destino;
    return normalizarPontos(resultado);
  }

  if (side === "bottom" && Math.abs(prev.x - end.x) <= CONFIG.sameColTolerance && prev.y >= (destinoNode ? destinoNode.y + destinoNode.h : end.y)) {
    resultado[resultado.length - 1] = destino;
    return normalizarPontos(resultado);
  }

  let pivot = null;

  if (side === "left") {
    pivot = { x: end.x - CONFIG.routeGap, y: end.y };
  } else if (side === "right") {
    pivot = { x: end.x + CONFIG.routeGap, y: end.y };
  } else if (side === "top") {
    pivot = { x: end.x, y: end.y - CONFIG.routeGap };
  } else if (side === "bottom") {
    pivot = { x: end.x, y: end.y + CONFIG.routeGap };
  }

  if (!pivot) return normalizarPontos(resultado);

  const novo = [...resultado.slice(0, -1)];
  const ultimoAntes = novo[novo.length - 1];

  if (ultimoAntes.x !== pivot.x && ultimoAntes.y !== pivot.y) {
    if (side === "left" || side === "right") {
      novo.push({ x: pivot.x, y: ultimoAntes.y });
    } else {
      novo.push({ x: ultimoAntes.x, y: pivot.y });
    }
  }

  novo.push(pivot);
  novo.push(destino);
  return normalizarPontos(novo);
}

function ajustarPrimeiroTrechoParaLado(points, start, side) {
  if (!points || points.length < 2) return points;

  const resultado = [...points];
  const next = resultado[1];

  if (side === "right" && Math.abs(next.y - start.y) <= CONFIG.sameRowTolerance && next.x > start.x) {
    resultado[0] = { x: start.x, y: start.y };
    return normalizarPontos(resultado);
  }

  if (side === "left" && Math.abs(next.y - start.y) <= CONFIG.sameRowTolerance && next.x < start.x) {
    resultado[0] = { x: start.x, y: start.y };
    return normalizarPontos(resultado);
  }

  if (side === "top" && Math.abs(next.x - start.x) <= CONFIG.sameColTolerance && next.y < start.y) {
    resultado[0] = { x: start.x, y: start.y };
    return normalizarPontos(resultado);
  }

  if (side === "bottom" && Math.abs(next.x - start.x) <= CONFIG.sameColTolerance && next.y > start.y) {
    resultado[0] = { x: start.x, y: start.y };
    return normalizarPontos(resultado);
  }

  let escape = null;

  if (side === "right") escape = { x: start.x + CONFIG.routeGap, y: start.y };
  else if (side === "left") escape = { x: start.x - CONFIG.routeGap, y: start.y };
  else if (side === "top") escape = { x: start.x, y: start.y - CONFIG.routeGap };
  else if (side === "bottom") escape = { x: start.x, y: start.y + CONFIG.routeGap };

  if (!escape) return normalizarPontos(resultado);

  const novo = [{ x: start.x, y: start.y }, escape];

  if (resultado.length > 1) {
    const terceiro = resultado[1];

    if (side === "left" || side === "right") {
      if (escape.y !== terceiro.y && escape.x !== terceiro.x) {
        novo.push({ x: escape.x, y: terceiro.y });
      }
    } else {
      if (escape.x !== terceiro.x && escape.y !== terceiro.y) {
        novo.push({ x: terceiro.x, y: escape.y });
      }
    }

    for (let i = 1; i < resultado.length; i++) {
      novo.push(resultado[i]);
    }
  }

  return normalizarPontos(novo);
}

function gerarCandidatosRotas(start, end) {
  const candidates = [];

  const mids = [
    [{ x: end.x, y: start.y }],
    [{ x: start.x, y: end.y }],
    [
      { x: start.x + CONFIG.routeGap, y: start.y },
      { x: start.x + CONFIG.routeGap, y: end.y }
    ],
    [
      { x: start.x - CONFIG.routeGap, y: start.y },
      { x: start.x - CONFIG.routeGap, y: end.y }
    ],
    [
      { x: start.x, y: start.y + CONFIG.routeGap },
      { x: end.x, y: start.y + CONFIG.routeGap }
    ],
    [
      { x: start.x, y: start.y - CONFIG.routeGap },
      { x: end.x, y: start.y - CONFIG.routeGap }
    ]
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

function escolherParesCandidatos(origem, destino, rotulo = "") {
  if (origem.isDecision && rotulo === "Sim") {
    return [{ startSide: "right", endSide: "left" }];
  }

  if (origem.isDecision && rotulo === "Não") {
    if (destino.gridRow <= origem.gridRow) {
      return [{ startSide: "bottom", endSide: "bottom" }];
    }
    return [{ startSide: "bottom", endSide: "top" }];
  }

  const dx