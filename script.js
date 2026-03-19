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

  decisionSize: 170,

  routeGap: 28,
  entryExitGap: 40,
  laneGap: 36,
  sameRowTolerance: 1,
  sameColTolerance: 1,
  sharedMergeGap: 34
};

function limpar(txt) {
  if (txt === null || txt === undefined) return "";
  return String(txt).replace(/"/g, "").replace(/\(/g, "").replace(/\)/g, "").trim();
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

function quebrarTextoAutomatico(texto, maxPalavrasPorLinha = 3, maxCaracteresPorLinha = 18) {
  const palavras = String(texto || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (!palavras.length) return [];

  const linhas = [];
  let linhaAtual = [];

  for (const palavra of palavras) {
    const testeLinha = [...linhaAtual, palavra].join(" ");
    const excedeuPalavras = linhaAtual.length >= maxPalavrasPorLinha;
    const excedeuCaracteres = testeLinha.length > maxCaracteresPorLinha;

    if (linhaAtual.length > 0 && (excedeuPalavras || excedeuCaracteres)) {
      linhas.push(linhaAtual.join(" "));
      linhaAtual = [palavra];
    } else {
      linhaAtual.push(palavra);
    }
  }

  if (linhaAtual.length) {
    linhas.push(linhaAtual.join(" "));
  }

  return linhas;
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
  document.getElementById("processo").value = "";
  document.getElementById("analista").value = "";
  document.getElementById("entrada").value = "";
  document.getElementById("diagram").innerHTML = "";
  document.getElementById("infoProcesso").innerHTML = "";
  document.getElementById("metricas").innerHTML = "";
  ultimoNomeArquivo = "fluxograma_processo";
}

function criarElementoSVG(tag) {
  return document.createElementNS("http://www.w3.org/2000/svg", tag);
}

function isPergunta(texto) {
  return limpar(texto).endsWith("?");
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

  const linhas = [
    ...quebrarTextoAutomatico(etapa.atividade, 3, pergunta ? 14 : 18),
    ...quebrarTextoAutomatico(etapa.sistema || "Sem sistema informado", 3, 20),
    formatarTempo(etapa.tempo)
  ];

  const lineHeight = 18;
  const totalAlturaTexto = linhas.length * lineHeight;
  const inicioYTexto = pos.y + (pos.h - totalAlturaTexto) / 2 + 14;

  linhas.forEach((linha, i) => {
    const text = criarElementoSVG("text");
    text.setAttribute("x", pos.x + pos.w / 2);
    text.setAttribute("y", inicioYTexto + i * lineHeight);
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
    if (ultimo && ultimo.x === p.x && ultimo.y === p.y) {
      return;
    }

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

    if (!mesmoX && !mesmoY) {
      simplificado.push(atual);
    }
  }

  simplificado.push(resultado[resultado.length - 1]);
  return simplificado;
}

function estaAlinhadoHorizontalmente(a, b) {
  return Math.abs((a.y + a.h / 2) - (b.y + b.h / 2)) <= CONFIG.sameRowTolerance;
}

function estaAlinhadoVerticalmente(a, b) {
  return Math.abs((a.x + a.w / 2) - (b.x + b.w / 2)) <= CONFIG.sameColTolerance;
}

function mesmaLinhaY(a, b) {
  return Math.abs(a.y - b.y) <= CONFIG.sameRowTolerance;
}

function mesmaColunaX(a, b) {
  return Math.abs(a.x - b.x) <= CONFIG.sameColTolerance;
}

function existeNoEntreMesmaLinha(origem, destino, posicoes) {
  const yCentro = origem.y + origem.h / 2;
  const xMin = Math.min(origem.x + origem.w, destino.x + destino.w);
  const xMax = Math.max(origem.x, destino.x);

  return Object.values(posicoes).some((node) => {
    if (!node || !node.id) return false;
    if (node.id === origem.id || node.id === destino.id) return false;
    if (node.id === "__INICIO__" || node.id === "__FIM__") return false;

    const nodeCentroY = node.y + node.h / 2;
    const alinhado = Math.abs(nodeCentroY - yCentro) <= CONFIG.sameRowTolerance;
    if (!alinhado) return false;

    return node.x < xMax && (node.x + node.w) > xMin;
  });
}

function existeNoEntreMesmaColuna(origem, destino, posicoes) {
  const xCentro = origem.x + origem.w / 2;
  const yMin = Math.min(origem.y + origem.h, destino.y + destino.h);
  const yMax = Math.max(origem.y, destino.y);

  return Object.values(posicoes).some((node) => {
    if (!node || !node.id) return false;
    if (node.id === origem.id || node.id === destino.id) return false;
    if (node.id === "__INICIO__" || node.id === "__FIM__") return false;

    const nodeCentroX = node.x + node.w / 2;
    const alinhado = Math.abs(nodeCentroX - xCentro) <= CONFIG.sameColTolerance;
    if (!alinhado) return false;

    return node.y < yMax && (node.y + node.h) > yMin;
  });
}

function montarRotaReta(start, end, label, endSide = "") {
  return {
    points: normalizarPontos([start, end]),
    label,
    endSide
  };
}

function montarRotaOrtogonal(points, label, endSide = "") {
  return {
    points: normalizarPontos(points),
    label,
    endSide
  };
}

function getMergePoint(end, side, gap = CONFIG.sharedMergeGap) {
  switch (side) {
    case "left":
      return { x: end.x - gap, y: end.y };
    case "right":
      return { x: end.x + gap, y: end.y };
    case "top":
      return { x: end.x, y: end.y - gap };
    case "bottom":
      return { x: end.x, y: end.y + gap };
    default:
      return { x: end.x - gap, y: end.y };
  }
}

function buildOrthogonalToMerge(start, mergePoint, end, endSide) {
  if (mesmaLinhaY(start, mergePoint)) {
    return normalizarPontos([start, mergePoint, end]);
  }

  if (mesmaColunaX(start, mergePoint)) {
    return normalizarPontos([start, mergePoint, end]);
  }

  if (endSide === "left" || endSide === "right") {
    const elbow = { x: mergePoint.x, y: start.y };
    return normalizarPontos([start, elbow, mergePoint, end]);
  }

  const elbow = { x: start.x, y: mergePoint.y };
  return normalizarPontos([start, elbow, mergePoint, end]);
}

function escolherRota(origem, destino, contexto = {}) {
  const rotulo = contexto.rotulo || "";
  const ordemConexao = contexto.ordemConexao || 0;
  const posicoes = contexto.posicoes || {};
  const offset = ordemConexao * 14;

  const origemPergunta = origem.isDecision;
  const mesmaLinha = estaAlinhadoHorizontalmente(origem, destino);
  const mesmaColuna = estaAlinhadoVerticalmente(origem, destino);
  const temBloqueioHorizontal = mesmaLinha && existeNoEntreMesmaLinha(origem, destino, posicoes);
  const temBloqueioVertical = mesmaColuna && existeNoEntreMesmaColuna(origem, destino, posicoes);

  if (origem.id === "__INICIO__" && mesmaLinha) {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    return montarRotaReta(start, end, { x: (start.x + end.x) / 2, y: start.y - 10 }, "left");
  }

  if (destino.id === "__FIM__" && mesmaLinha) {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    return montarRotaReta(start, end, { x: (start.x + end.x) / 2, y: start.y - 10 }, "left");
  }

  if (origemPergunta && rotulo === "Sim") {
    const start = getAnchorPoint(origem, "right");

    if (mesmaLinha && destino.x >= origem.x && !temBloqueioHorizontal) {
      const end = getAnchorPoint(destino, "left");
      return montarRotaReta(start, end, { x: (start.x + end.x) / 2, y: start.y - 10 }, "left");
    }

    const endSide = destino.x >= origem.x ? "left" : "right";
    const end = getAnchorPoint(destino, endSide);

    const laneX = Math.max(
      start.x + CONFIG.routeGap + offset,
      end.x + (endSide === "left" ? -CONFIG.routeGap : CONFIG.routeGap) + offset
    );

    return montarRotaOrtogonal([
      start,
      { x: laneX, y: start.y },
      { x: laneX, y: end.y },
      end
    ], { x: start.x + 18, y: start.y - 10 }, endSide);
  }

  if (origemPergunta && rotulo === "Não") {
    const start = getAnchorPoint(origem, "bottom");

    if (mesmaColuna && destino.y >= origem.y && !temBloqueioVertical) {
      const end = getAnchorPoint(destino, "top");
      return montarRotaReta(start, end, { x: start.x + 18, y: (start.y + end.y) / 2 - 8 }, "top");
    }

    const endSide = destino.y >= origem.y ? "top" : "bottom";
    const end = getAnchorPoint(destino, endSide);

    const laneY = Math.max(
      start.y + CONFIG.routeGap + offset,
      end.y + (endSide === "top" ? -CONFIG.routeGap : CONFIG.routeGap) + offset
    );

    return montarRotaOrtogonal([
      start,
      { x: start.x, y: laneY },
      { x: end.x, y: laneY },
      end
    ], { x: start.x + 18, y: start.y + 18 }, endSide);
  }

  if (mesmaLinha && destino.x > origem.x && !temBloqueioHorizontal) {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    return montarRotaReta(start, end, { x: (start.x + end.x) / 2, y: start.y - 10 }, "left");
  }

  if (mesmaLinha && destino.x < origem.x && !temBloqueioHorizontal) {
    const start = getAnchorPoint(origem, "left");
    const end = getAnchorPoint(destino, "right");
    return montarRotaReta(start, end, { x: (start.x + end.x) / 2, y: start.y - 10 }, "right");
  }

  if (mesmaColuna && destino.y > origem.y && !temBloqueioVertical) {
    const start = getAnchorPoint(origem, "bottom");
    const end = getAnchorPoint(destino, "top");
    return montarRotaReta(start, end, { x: start.x + 18, y: (start.y + end.y) / 2 - 8 }, "top");
  }

  if (mesmaColuna && destino.y < origem.y && !temBloqueioVertical) {
    const start = getAnchorPoint(origem, "top");
    const end = getAnchorPoint(destino, "bottom");
    return montarRotaReta(start, end, { x: start.x + 18, y: (start.y + end.y) / 2 - 8 }, "bottom");
  }

  if (mesmaColuna && temBloqueioVertical) {
    const usarEsquerda = (origem.gridCol || 1) <= (destino.gridCol || 1);
    const endSide = usarEsquerda ? "left" : "right";
    const start = getAnchorPoint(origem, endSide);
    const end = getAnchorPoint(destino, endSide);

    const laneX = usarEsquerda
      ? Math.min(origem.x, destino.x) - CONFIG.laneGap - offset
      : Math.max(origem.x + origem.w, destino.x + destino.w) + CONFIG.laneGap + offset;

    return montarRotaOrtogonal([
      start,
      { x: laneX, y: start.y },
      { x: laneX, y: end.y },
      end
    ], { x: laneX, y: start.y - 10 }, endSide);
  }

  if (mesmaLinha && temBloqueioHorizontal) {
    const usarTopo = (origem.gridRow || 1) <= (destino.gridRow || 1);
    const endSide = usarTopo ? "top" : "bottom";
    const start = getAnchorPoint(origem, endSide);
    const end = getAnchorPoint(destino, endSide);

    const laneY = usarTopo
      ? Math.min(origem.y, destino.y) - CONFIG.laneGap - offset
      : Math.max(origem.y + origem.h, destino.y + destino.h) + CONFIG.laneGap + offset;

    return montarRotaOrtogonal([
      start,
      { x: start.x, y: laneY },
      { x: end.x, y: laneY },
      end
    ], { x: start.x + 18, y: laneY - 8 }, endSide);
  }

  if (destino.x >= origem.x + origem.w) {
    const start = getAnchorPoint(origem, "right");
    const end = getAnchorPoint(destino, "left");
    const midX = (start.x + end.x) / 2 + offset;

    return montarRotaOrtogonal([
      start,
      { x: midX, y: start.y },
      { x: midX, y: end.y },
      end
    ], { x: midX, y: start.y - 10 }, "left");
  }

  if (destino.y >= origem.y + origem.h) {
    const start = getAnchorPoint(origem, "bottom");
    const end = getAnchorPoint(destino, "top");
    const midY = (start.y + end.y) / 2 + offset;

    return montarRotaOrtogonal([
      start,
      { x: start.x, y: midY },
      { x: end.x, y: midY },
      end
    ], { x: start.x + 18, y: midY - 8 }, "top");
  }

  if (destino.y + destino.h <= origem.y) {
    const start = getAnchorPoint(origem, "top");
    const end = getAnchorPoint(destino, "bottom");
    const midY = (start.y + end.y) / 2 - offset;

    return montarRotaOrtogonal([
      start,
      { x: start.x, y: midY },
      { x: end.x, y: midY },
      end
    ], { x: start.x + 18, y: midY - 8 }, "bottom");
  }

  const start = getAnchorPoint(origem, "left");
  const end = getAnchorPoint(destino, "right");
  const midX = (start.x + end.x) / 2 - offset;

  return montarRotaOrtogonal([
    start,
    { x: midX, y: start.y },
    { x: midX, y: end.y },
    end
  ], { x: midX, y: start.y - 10 }, "right");
}

function construirRotaCompartilhada(start, sharedInfo) {
  const mergePoint = { x: sharedInfo.mergePoint.x, y: sharedInfo.mergePoint.y };
  const end = { x: sharedInfo.end.x, y: sharedInfo.end.y };
  const endSide = sharedInfo.endSide;

  return {
    points: buildOrthogonalToMerge(start, mergePoint, end, endSide),
    label: sharedInfo.label,
    endSide
  };
}

function desenharConexao(
  svg,
  origem,
  destino,
  rotulo = "",
  ordemConexao = 0,
  posicoes = {},
  sharedRegistry = {}
) {
  let rota = escolherRota(origem, destino, {
    rotulo,
    ordemConexao,
    posicoes
  });

  const sharedKey = `${destino.id}__${rota.endSide || "auto"}`;
  const startReal = { ...rota.points[0] };

  if (
    destino.id !== "__FIM__" &&
    destino.id !== "__INICIO__" &&
    sharedRegistry[sharedKey] &&
    origem.id !== sharedRegistry[sharedKey].origemId
  ) {
    rota = construirRotaCompartilhada(startReal, sharedRegistry[sharedKey]);
  } else if (destino.id !== "__FIM__" && destino.id !== "__INICIO__") {
    const end = rota.points[rota.points.length - 1];
    sharedRegistry[sharedKey] = {
      origemId: origem.id,
      endSide: rota.endSide || "left",
      end: { x: end.x, y: end.y },
      mergePoint: getMergePoint(end, rota.endSide || "left"),
      label: rota.label
    };
  }

  const path = criarElementoSVG("path");
  path.setAttribute("d", createPolylinePath(rota.points));
  path.setAttribute("fill", "none");
  path.setAttribute("stroke", "#111111");
  path.setAttribute("stroke-width", CONFIG.lineWidth);
  path.setAttribute("marker-end", "url(#arrow)");
  path.setAttribute("stroke-linejoin", "round");
  path.setAttribute("stroke-linecap", "round");
  svg.appendChild(path);

  if (rotulo) {
    const tx = criarElementoSVG("text");
    tx.setAttribute("x", rota.label.x);
    tx.setAttribute("y", rota.label.y);
    tx.setAttribute("text-anchor", "middle");
    tx.setAttribute("font-family", CONFIG.fontFamily);
    tx.setAttribute("font-size", "12");
    tx.setAttribute("fill", "#111111");
    tx.textContent = rotulo;
    svg.appendChild(tx);
  }
}

function gerarFluxo() {
  const texto = document.getElementById("entrada").value;

  if (!texto.trim()) {
    alert("Cole a tabela do Excel primeiro.");
    return;
  }

  const processo = limpar(document.getElementById("processo").value);
  const analista = limpar(document.getElementById("analista").value);

  const linhasBrutas = texto
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter(l => l.trim() !== "");

  let linhas = linhasBrutas.map(l => l.split("\t"));

  if (linhas.length && ehCabecalho(linhas[0])) {
    linhas.shift();
  }

  const etapas = [];
  const idsValidos = new Set();

  linhas.forEach((col) => {
    while (col.length < 12) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]) || "Não informado";
    const sistema = limpar(col[4]);
    const tempo = tempoParaSegundos(limpar(col[5]));
    const proxSim = limpar(col[6]);
    const proxNao = limpar(col[7]);
    const conexoesExtras = limpar(col[8]);
    const coluna = Number(limpar(col[9])) || 1;
    const linha = Number(limpar(col[10])) || 1;
    const cor = normalizarCor(col[11]);

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
      cor
    });

    idsValidos.add(id);
  });

  if (!etapas.length) {
    alert("Nenhuma etapa válida foi encontrada na tabela.");
    return;
  }

  etapas.sort((a, b) => a.ordem - b.ordem);

  const etapaPorId = {};
  etapas.forEach(e => {
    etapaPorId[e.id] = e;
  });

  const rowSlotHeight = Math.max(CONFIG.boxHeight, CONFIG.decisionSize);
  const colSlotWidth = Math.max(CONFIG.boxWidth, CONFIG.decisionSize);

  const maxColuna = Math.max(...etapas.map(e => e.coluna), 1) + 1;
  const maxLinha = Math.max(...etapas.map(e => e.linha), 1) + 1;

  const svgWidth =
    CONFIG.marginX * 2 +
    maxColuna * colSlotWidth +
    (maxColuna - 1) * CONFIG.colGap;

  const svgHeight =
    CONFIG.marginY * 2 +
    maxLinha * rowSlotHeight +
    (maxLinha - 1) * CONFIG.rowGap;

  const svg = criarElementoSVG("svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", svgWidth);
  svg.setAttribute("height", svgHeight);
  svg.setAttribute("viewBox", `0 0 ${svgWidth} ${svgHeight}`);
  svg.setAttribute("style", "background:#ffffff");

  const defs = criarElementoSVG("defs");
  const marker = criarElementoSVG("marker");
  marker.setAttribute("id", "arrow");
  marker.setAttribute("markerWidth", "10");
  marker.setAttribute("markerHeight", "10");
  marker.setAttribute("refX", "9");
  marker.setAttribute("refY", "3");
  marker.setAttribute("orient", "auto");
  marker.setAttribute("markerUnits", "strokeWidth");

  const markerPath = criarElementoSVG("path");
  markerPath.setAttribute("d", "M0,0 L10,3 L0,6 Z");
  markerPath.setAttribute("fill", "#111111");
  marker.appendChild(markerPath);
  defs.appendChild(marker);
  svg.appendChild(defs);

  const posicoes = {};

  etapas.forEach((e) => {
    const pergunta = isPergunta(e.atividade);
    const w = pergunta ? CONFIG.decisionSize : CONFIG.boxWidth;
    const h = pergunta ? CONFIG.decisionSize : CONFIG.boxHeight;

    const slotX = CONFIG.marginX + (e.coluna - 1) * (colSlotWidth + CONFIG.colGap);
    const slotY = CONFIG.marginY + (e.linha - 1) * (rowSlotHeight + CONFIG.rowGap);

    const x = slotX + (colSlotWidth - w) / 2;
    const y = slotY + (rowSlotHeight - h) / 2;

    posicoes[e.id] = {
      id: e.id,
      x,
      y,
      w,
      h,
      isDecision: pergunta,
      gridCol: e.coluna,
      gridRow: e.linha
    };
  });

  const primeiraEtapa = etapas[0];
  const ultimaEtapa = etapas[etapas.length - 1];
  const primeiraPos = posicoes[primeiraEtapa.id];
  const ultimaPos = posicoes[ultimaEtapa.id];

  posicoes["__INICIO__"] = {
    id: "__INICIO__",
    x: primeiraPos.x - 60 - CONFIG.entryExitGap,
    y: primeiraPos.y + (primeiraPos.h - 36) / 2,
    w: 60,
    h: 36,
    isDecision: false,
    gridCol: Math.max(0, primeiraPos.gridCol - 1),
    gridRow: primeiraPos.gridRow
  };

  posicoes["__FIM__"] = {
    id: "__FIM__",
    x: ultimaPos.x + ultimaPos.w + CONFIG.entryExitGap,
    y: ultimaPos.y + (ultimaPos.h - 36) / 2,
    w: 60,
    h: 36,
    isDecision: false,
    gridCol: ultimaPos.gridCol + 1,
    gridRow: ultimaPos.gridRow
  };

  desenharCapsula(svg, "Início", posicoes["__INICIO__"].x, posicoes["__INICIO__"].y, 60, 36);
  desenharCapsula(svg, "Fim", posicoes["__FIM__"].x, posicoes["__FIM__"].y, 60, 36);

  etapas.forEach((etapa) => {
    desenharNo(svg, etapa, posicoes[etapa.id]);
  });

  let loops = 0;
  let decisoes = 0;
  let conexoesExtrasCount = 0;
  const etapasOrigemComRetorno = new Set();
  const sharedRegistry = {};

  desenharConexao(svg, posicoes["__INICIO__"], posicoes[primeiraEtapa.id], "", 0, posicoes, sharedRegistry);

  etapas.forEach((etapa) => {
    const origem = posicoes[etapa.id];
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    const pergunta = isPergunta(etapa.atividade);
    if (pergunta) {
      decisoes++;
    }

    destinosSim.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa, etapasOrigemComRetorno);
      }

      desenharConexao(
        svg,
        origem,
        posicoes[destinoId],
        pergunta ? (indice === 0 ? "Sim" : `Sim ${indice + 1}`) : "",
        indice,
        posicoes,
        sharedRegistry
      );
    });

    destinosNao.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa, etapasOrigemComRetorno);
      }

      desenharConexao(
        svg,
        origem,
        posicoes[destinoId],
        pergunta ? (indice === 0 ? "Não" : `Não ${indice + 1}`) : (indice === 0 ? "Não" : `Não ${indice + 1}`),
        indice,
        posicoes,
        sharedRegistry
      );
    });

    destinosExtras.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa, etapasOrigemComRetorno);
      }

      conexoesExtrasCount++;
      desenharConexao(svg, origem, posicoes[destinoId], "", indice + 1, posicoes, sharedRegistry);
    });
  });

  etapas.forEach((etapa) => {
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    if (destinosSim.length === 0 && destinosNao.length === 0 && destinosExtras.length === 0) {
      desenharConexao(svg, posicoes[etapa.id], posicoes["__FIM__"], "", 0, posicoes, sharedRegistry);
    }
  });

  document.getElementById("diagram").innerHTML = "";
  document.getElementById("diagram").appendChild(svg);

  ultimoNomeArquivo = gerarNomeArquivo();

  document.getElementById("infoProcesso").innerHTML =
    "<b>Processo:</b> " + (processo || "Não informado") + "<br>" +
    "<b>Analista:</b> " + (analista || "Não informado");

  let tempoTotal = 0;
  const atividadesTempo = [];
  const tiposTempo = {};

  etapas.forEach((etapa) => {
    tempoTotal += etapa.tempo;
    atividadesTempo.push({ atividade: etapa.atividade, tempo: etapa.tempo });

    if (!tiposTempo[etapa.tipo]) {
      tiposTempo[etapa.tipo] = 0;
    }
    tiposTempo[etapa.tipo] += etapa.tempo;
  });

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const top3HTML = atividadesTempo
    .slice(0, 3)
    .map(a => {
      const pct = tempoTotal ? formatarPercentual((a.tempo / tempoTotal) * 100) : "0,0";
      return '<div class="analytics-item">' +
        a.atividade +
        ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
        ' <span class="icon-pct">%</span>' + pct + '%' +
      '</div>';
    })
    .join("");

  let paretoHTML = "";
  atividadesTempo.forEach(a => {
    const pct = tempoTotal ? formatarPercentual((a.tempo / tempoTotal) * 100) : "0,0";
    paretoHTML += '<div class="analytics-item">' +
      a.atividade +
      ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
      ' <span class="icon-pct">%</span>' + pct + '%' +
    '</div>';
  });

  const tiposOrdenados = Object.entries(tiposTempo).sort((a, b) => b[1] - a[1]);

  let tiposHTML = "";
  tiposOrdenados.forEach(([tipo, tempo]) => {
    const pct = tempoTotal ? formatarPercentual((tempo / tempoTotal) * 100) : "0,0";
    tiposHTML += '<div class="analytics-item">' +
      tipo +
      ' — <span class="icon-time">⏱</span>' + formatarTempo(tempo) +
      ' <span class="icon-pct">%</span>' + pct + '%' +
    '</div>';
  });

  let tempoPotencialRetrabalho = 0;
  etapas.forEach((etapa) => {
    if (etapasOrigemComRetorno.has(etapa.id)) {
      tempoPotencialRetrabalho += etapa.tempo;
    }
  });

  const impactoPotencialRetrabalho = tempoTotal
    ? formatarPercentual((tempoPotencialRetrabalho / tempoTotal) * 100)
    : "0,0";

  const taxaDecisao = etapas.length
    ? formatarPercentual((decisoes / etapas.length) * 100)
    : "0,0";

  document.getElementById("metricas").innerHTML =
    '<div class="analytics-grid">' +
      '<div class="analytics-col">' +
        '<div class="analytics-section">' +
          '<div class="metric-highlight"><b>Tempo total do processo:</b> <span class="icon-time">⏱</span>' + formatarTempo(tempoTotal) + '</div>' +
          '<div class="analytics-title">Top 3 gargalos</div>' +
          top3HTML +
          '<br>' +
          '<div class="analytics-item"><b>Loops detectados:</b> ' + loops + '</div>' +
          '<div class="analytics-item"><b>Conexões extras:</b> ' + conexoesExtrasCount + '</div>' +
          '<div class="analytics-item"><b>Impacto potencial de retrabalho:</b> <span class="icon-time">⏱</span>' + formatarTempo(tempoPotencialRetrabalho) + ' <span class="icon-pct">%</span>' + impactoPotencialRetrabalho + '%</div>' +
          '<div class="analytics-item"><b>Taxa de decisão:</b> ' + decisoes + ' etapa(s) <span class="icon-pct">%</span>' + taxaDecisao + '%</div>' +
        '</div>' +
        '<div class="analytics-section">' +
          '<div class="analytics-title">Tempo por tipo</div>' +
          tiposHTML +
        '</div>' +
      '</div>' +
      '<div class="analytics-col">' +
        '<div class="analytics-section">' +
          '<div class="analytics-title">Pareto de tempo</div>' +
          paretoHTML +
        '</div>' +
      '</div>' +
    '</div>';
}

function obterSVGPronto() {
  const svgOriginal = document.querySelector("#diagram svg");
  if (!svgOriginal) {
    alert("Gere o fluxo primeiro.");
    return null;
  }

  return svgOriginal.cloneNode(true);
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

function baixarPNG() {
  const svg = obterSVGPronto();
  if (!svg) return;

  const larguraReal = Number(svg.getAttribute("width")) || 2000;
  const alturaReal = Number(svg.getAttribute("height")) || 1200;

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const svg64 = btoa(unescape(encodeURIComponent(source)));
  const image64 = "data:image/svg+xml;base64," + svg64;

  const img = new Image();

  img.onload = function () {
    const canvas = document.createElement("canvas");
    canvas.width = larguraReal;
    canvas.height = alturaReal;

    const ctx = canvas.getContext("2d");
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    const link = document.createElement("a");
    link.download = `${ultimoNomeArquivo}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
  };

  img.onerror = function (erro) {
    console.error("Erro ao gerar PNG:", erro);
    alert("Não foi possível gerar o PNG.");
  };

  img.src = image64;
}