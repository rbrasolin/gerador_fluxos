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
  laneGap: 28,
  lanePaddingTop: 26,
  lanePaddingBottom: 26,
  laneLabelWidth: 70,
  laneEntryWidth: 110,
  laneHeaderFontSize: 16,
  laneHeaderFill: "#f7f7f7",
  laneBorder: "#bdbdbd",
  laneSeparator: "#d6d6d6",
  sameRowTolerance: 1,
  sameColTolerance: 1,
  sharedMergeGap: 34,
  obstaclePadding: 10,
  textLineHeight: 18,
  textPaddingVertical: 20,
  rectTextPaddingHorizontal: 16
};

const EXCEL_EXPORT_SCALE = 0.65;
// 0.85 = redução leve
// 0.72 = redução boa
// 0.65 = redução forte

const EXCEL_LAYOUT = {
  colGap: 70, //Espaço horizontal entre as colunas (atividades)
  rowGap: 36, //Espaço vertical entre atividades dentro da mesma raia
  laneGap: 50, //Espaço entre uma raia e outra
  lanePaddingTop: 18, //Espaço interno no topo da raia, Se estiver baixo → caixa “grudada” no topo
  lanePaddingBottom: 18, //Espaço interno na parte de baixo da raia, Evita que última atividade fique colada na borda
  laneLabelWidth: 22, //“Reserva” horizontal para a área da raia (estrutura no desenho da raia, não aplica no excel)
  laneEntryWidth: 12, //Espaço entre: área do nome da raia e início do fluxo (Muito pequeno → fluxo começa “em cima” da raia)
  startGap: 16, //Distância entre: o “Início” e a primeira atividade
  endGap: 26, //Distância entre: última atividade e o “Fim”
  laneTextOffsetLeft: 120, //Distância do nome da raia para a esquerda (Aumenta → nome vai mais para esquerda Diminui → nome se aproxima do fluxo)
  extraLeftPadding: 50 //Empurra TODO o fluxo para a direita
};

function aplicarEscalaSVGExcel(svgOriginal, escala = EXCEL_EXPORT_SCALE) {
  const svg = svgOriginal.cloneNode(true);

  const larguraOriginal = Number(svg.getAttribute("width")) || 1000;
  const alturaOriginal = Number(svg.getAttribute("height")) || 800;

  const filhos = Array.from(svg.childNodes);

  const grupoEscalado = criarElementoSVG("g");
  grupoEscalado.setAttribute("transform", `scale(${escala})`);

  filhos.forEach((filho) => {
    grupoEscalado.appendChild(filho);
  });

  svg.appendChild(grupoEscalado);

  const novaLargura = Math.round(larguraOriginal * escala);
  const novaAltura = Math.round(alturaOriginal * escala);

  svg.setAttribute("width", novaLargura);
  svg.setAttribute("height", novaAltura);
  svg.setAttribute("viewBox", `0 0 ${novaLargura} ${novaAltura}`);

  return svg;
}

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

  return [
    ...linhasAtividade,
    formatarTempo(etapa.tempo)
  ];
}

function desenharSistemaSeparado(svg, etapa, pos) {
  const altura = 24;
  const y = pos.y + pos.h - altura;

  const g = criarElementoSVG("g");

  const rect = criarElementoSVG("rect");
  rect.setAttribute("x", pos.x);
  rect.setAttribute("y", y);
  rect.setAttribute("width", pos.w);
  rect.setAttribute("height", altura);
  rect.setAttribute("class", "system");
  g.appendChild(rect);

  const text = criarElementoSVG("text");
  text.setAttribute("x", pos.x + pos.w / 2);
  text.setAttribute("y", y + 16);
  text.setAttribute("text-anchor", "middle");
  text.setAttribute("font-family", CONFIG.fontFamily);
  text.setAttribute("font-size", CONFIG.smallFontSize);
  text.setAttribute("fill", "#111111");
  text.textContent = etapa.sistema || "Sem sistema";

  g.appendChild(text);
  svg.appendChild(g);
}

function desenharTextoMultilinhaSVG(g, linhas, x, yInicial, fontSize, lineHeight, fill = "#111111", fontWeight = "normal") {
  const text = criarElementoSVG("text");
  text.setAttribute("x", x);
  text.setAttribute("y", yInicial);
  text.setAttribute("text-anchor", "middle");
  text.setAttribute("font-family", CONFIG.fontFamily);
  text.setAttribute("font-size", String(fontSize));
  text.setAttribute("font-weight", fontWeight);
  text.setAttribute("fill", fill);
  text.setAttribute("xml:space", "preserve");

  linhas.forEach((linha, i) => {
    const tspan = criarElementoSVG("tspan");
    tspan.setAttribute("x", x);
    if (i === 0) {
      tspan.setAttribute("dy", "0");
    } else {
      tspan.setAttribute("dy", String(lineHeight));
    }
    tspan.textContent = linha;
    text.appendChild(tspan);
  });

  g.appendChild(text);
}

function desenharNoExcel(svg, etapa, pos) {
  const g = criarElementoSVG("g");
  const pergunta = isPergunta(etapa.atividade);

  if (pergunta) {
    const cx = pos.x + pos.w / 2;
    const cy = pos.y + pos.h / 2;

    const polygon = criarElementoSVG("polygon");
    polygon.setAttribute(
      "points",
      `${cx},${pos.y} ${pos.x + pos.w},${cy} ${cx},${pos.y + pos.h} ${pos.x},${cy}`
    );
    polygon.setAttribute("fill", corHex(etapa.cor));
    polygon.setAttribute("stroke", "#111111");
    polygon.setAttribute("stroke-width", "2");
    g.appendChild(polygon);

    const linhas = obterLinhasEtapa(etapa, pos.w);
    const totalAlturaTexto = linhas.length * CONFIG.textLineHeight;
    const inicioYTexto = pos.y + (pos.h - totalAlturaTexto) / 2 + 14;

    desenharTextoMultilinhaSVG(
      g,
      linhas,
      pos.x + pos.w / 2,
      inicioYTexto,
      CONFIG.fontSize,
      CONFIG.textLineHeight,
      "#111111"
    );
  } else {
    const alturaSistema = 24;
    const ySeparador = pos.y + pos.h - alturaSistema;

    const rectPrincipal = criarElementoSVG("rect");
    rectPrincipal.setAttribute("x", pos.x);
    rectPrincipal.setAttribute("y", pos.y);
    rectPrincipal.setAttribute("width", pos.w);
    rectPrincipal.setAttribute("height", pos.h);
    rectPrincipal.setAttribute("fill", corHex(etapa.cor));
    rectPrincipal.setAttribute("stroke", "#111111");
    rectPrincipal.setAttribute("stroke-width", "2");
    g.appendChild(rectPrincipal);

    const linhaSeparadora = criarElementoSVG("line");
    linhaSeparadora.setAttribute("x1", pos.x);
    linhaSeparadora.setAttribute("y1", ySeparador);
    linhaSeparadora.setAttribute("x2", pos.x + pos.w);
    linhaSeparadora.setAttribute("y2", ySeparador);
    linhaSeparadora.setAttribute("stroke", "#111111");
    linhaSeparadora.setAttribute("stroke-width", "1.5");
    g.appendChild(linhaSeparadora);

    const linhasAtividade = quebrarTextoPorLargura(
  etapa.atividade,
  obterLarguraUtilTexto(etapa, pos.w),
  CONFIG.fontSize
);

const linhaTempo = formatarTempo(etapa.tempo);

const areaUtil = pos.h - alturaSistema - 6;
const totalAlturaTexto = (linhasAtividade.length * CONFIG.textLineHeight) + CONFIG.textLineHeight;
const inicioYTexto = pos.y + (areaUtil - totalAlturaTexto) / 2 + 14;

// atividade: uma forma por linha
linhasAtividade.forEach((linha, i) => {
  const text = criarElementoSVG("text");
  text.setAttribute("x", pos.x + pos.w / 2);
  text.setAttribute("y", inicioYTexto + i * CONFIG.textLineHeight);
  text.setAttribute("text-anchor", "middle");
  text.setAttribute("font-family", CONFIG.fontFamily);
  text.setAttribute("font-size", String(CONFIG.fontSize));
  text.setAttribute("fill", "#111111");
  text.setAttribute("xml:space", "preserve");
  text.textContent = linha;
  g.appendChild(text);
});

// tempo: uma forma separada
const textTempo = criarElementoSVG("text");
textTempo.setAttribute("x", pos.x + pos.w / 2);
textTempo.setAttribute("y", inicioYTexto + linhasAtividade.length * CONFIG.textLineHeight);
textTempo.setAttribute("text-anchor", "middle");
textTempo.setAttribute("font-family", CONFIG.fontFamily);
textTempo.setAttribute("font-size", String(CONFIG.smallFontSize));
textTempo.setAttribute("fill", "#111111");
textTempo.setAttribute("xml:space", "preserve");
textTempo.textContent = linhaTempo;
g.appendChild(textTempo);

    const textoSistema = criarElementoSVG("text");
    textoSistema.setAttribute("x", pos.x + pos.w / 2);
    textoSistema.setAttribute("y", ySeparador + 16);
    textoSistema.setAttribute("text-anchor", "middle");
    textoSistema.setAttribute("font-family", CONFIG.fontFamily);
    textoSistema.setAttribute("font-size", String(CONFIG.smallFontSize));
    textoSistema.setAttribute("fill", "#111111");
    textoSistema.setAttribute("xml:space", "preserve");
    textoSistema.textContent = etapa.sistema || "Sem sistema";
    g.appendChild(textoSistema);
  }

  svg.appendChild(g);
}

function obterAlturaNo(etapa, alturaPadraoBase) {
  const alturaSistema = 24;

  if (isPergunta(etapa.atividade)) {
    return Math.ceil(alturaPadraoBase * CONFIG.decisionHeightFactor) + alturaSistema;
  }

  return alturaPadraoBase + alturaSistema;
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

function adicionarLoopSeNecessario(origemId, destinoId, etapaPorId, etapaAtual) {
  const destino = etapaPorId[destinoId];
  if (destino && destino.ordem < etapaAtual.ordem) {
    return 1;
  }
  return 0;
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
  const pergunta = isPergunta(etapa.atividade);

  if (pergunta) {
    const cx = pos.x + pos.w / 2;
    const cy = pos.y + pos.h / 2;

    const polygon = criarElementoSVG("polygon");
polygon.setAttribute(
  "points",
  `${cx},${pos.y} ${pos.x + pos.w},${cy} ${cx},${pos.y + pos.h} ${pos.x},${cy}`
);
polygon.setAttribute("class", `box ${etapa.cor}`);
g.appendChild(polygon);
  } else {
  const rect = criarElementoSVG("rect");
  rect.setAttribute("x", pos.x);
  rect.setAttribute("y", pos.y);
  rect.setAttribute("width", pos.w);
  rect.setAttribute("height", pos.h);
  rect.setAttribute("class", `box ${etapa.cor}`);
  g.appendChild(rect);
  }
  const linhas = obterLinhasEtapa(etapa, pos.w);
  const alturaSistema = 24;
  const areaUtil = pos.h - alturaSistema;

  const totalAlturaTexto = linhas.length * CONFIG.textLineHeight;
  const inicioYTexto = pos.y + (areaUtil - totalAlturaTexto) / 2 + 14;

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

if (!pergunta) {
  desenharSistemaSeparado(svg, etapa, pos);
}
}

function desenharRaias(svg, areasOrdenadas, lanes, svgWidth) {
  areasOrdenadas.forEach((area) => {
    const lane = lanes[area];
    if (!lane) return;

    const g = criarElementoSVG("g");

    const rectExterno = criarElementoSVG("rect");
    rectExterno.setAttribute("x", lane.x);
    rectExterno.setAttribute("y", lane.y);
    rectExterno.setAttribute("width", lane.width);
    rectExterno.setAttribute("height", lane.height);
    rectExterno.setAttribute("fill", "#ffffff");
    rectExterno.setAttribute("stroke", CONFIG.laneBorder);
    rectExterno.setAttribute("stroke-width", "1.2");
    g.appendChild(rectExterno);

    const header = criarElementoSVG("rect");
    header.setAttribute("x", lane.x);
    header.setAttribute("y", lane.y);
    header.setAttribute("width", CONFIG.laneLabelWidth);
    header.setAttribute("height", lane.height);
    header.setAttribute("fill", CONFIG.laneHeaderFill);
    header.setAttribute("stroke", CONFIG.laneBorder);
    header.setAttribute("stroke-width", "1");
    g.appendChild(header);

    const separator = criarElementoSVG("line");
    separator.setAttribute("x1", lane.x + CONFIG.laneLabelWidth);
    separator.setAttribute("y1", lane.y);
    separator.setAttribute("x2", lane.x + CONFIG.laneLabelWidth);
    separator.setAttribute("y2", lane.y + lane.height);
    separator.setAttribute("stroke", CONFIG.laneSeparator);
    separator.setAttribute("stroke-width", "1");
    g.appendChild(separator);

    const texto = criarElementoSVG("text");
    const tx = lane.x + CONFIG.laneLabelWidth / 2;
    const ty = lane.y + lane.height / 2;
    texto.setAttribute("x", tx);
    texto.setAttribute("y", ty);
    texto.setAttribute("text-anchor", "middle");
    texto.setAttribute("font-family", CONFIG.fontFamily);
    texto.setAttribute("font-size", String(CONFIG.laneHeaderFontSize));
    texto.setAttribute("font-weight", "bold");
    texto.setAttribute("fill", "#333333");
    texto.setAttribute("transform", `rotate(-90 ${tx} ${ty})`);
    texto.textContent = area;
    g.appendChild(texto);

    svg.appendChild(g);
  });
}

function desenharRaiasExcel(svg, lanes) {
  lanes.forEach((lane) => {
    const textoArea = criarElementoSVG("text");

    const areaCenterX = Math.max(
      14,
      lane.x - EXCEL_LAYOUT.laneTextOffsetLeft
    );
    const areaCenterY = lane.y + lane.height / 2;

    textoArea.setAttribute(
      "transform",
      `rotate(-90 ${areaCenterX} ${areaCenterY})`
    );
    textoArea.setAttribute("x", areaCenterX);
    textoArea.setAttribute("y", areaCenterY);
    textoArea.setAttribute("text-anchor", "middle");
    textoArea.setAttribute("dominant-baseline", "middle");
    textoArea.setAttribute("font-family", CONFIG.fontFamily);
    textoArea.setAttribute("font-size", CONFIG.laneHeaderFontSize);
    textoArea.setAttribute("font-weight", "bold");
    textoArea.setAttribute("fill", "#333333");
    textoArea.textContent = lane.area;

    svg.appendChild(textoArea);
  });
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

  const dx = destino.gridCol - origem.gridCol;
  const dy = destino.gridRowGlobal - origem.gridRowGlobal;

  if (dx === 0 && dy > 0) {
    return [{ startSide: "bottom", endSide: "top" }];
  }

  if (dx === 0 && dy < 0) {
    return [{ startSide: "top", endSide: "bottom" }];
  }

  if (dy === 0 && dx > 0) {
    return [{ startSide: "right", endSide: "left" }];
  }

  if (dy === 0 && dx < 0) {
    return [{ startSide: "left", endSide: "right" }];
  }

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

function podeCompartilharDestino(origem, sharedInfo) {
  if (!sharedInfo) return false;
  return origem.gridCol === sharedInfo.sourceGridCol;
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
  } else if (destino.id !== "__FIM__" && destino.id !== "__INICIO__") {
    const end = rota.points[rota.points.length - 1];
    sharedRegistry[sharedKey] = {
      origemId: origem.id,
      sourceGridCol: origem.gridCol,
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
    const larguraTexto = medirLarguraTexto(rotulo, 12, "bold");
    const larguraCapsula = Math.max(42, larguraTexto + 20);
    const alturaCapsula = 24;
    const xCapsula = rota.label.x - larguraCapsula / 2;
    const yCapsula = rota.label.y - alturaCapsula / 2 - 1;

    const bg = criarElementoSVG("rect");
    bg.setAttribute("x", xCapsula);
    bg.setAttribute("y", yCapsula);
    bg.setAttribute("width", larguraCapsula);
    bg.setAttribute("height", alturaCapsula);
    bg.setAttribute("rx", alturaCapsula / 2);
    bg.setAttribute("ry", alturaCapsula / 2);
    bg.setAttribute("fill", "#ffffff");
    bg.setAttribute("stroke", "#111111");
    bg.setAttribute("stroke-width", "1");
    svg.appendChild(bg);

    const tx = criarElementoSVG("text");
    tx.setAttribute("x", rota.label.x);
    tx.setAttribute("y", rota.label.y + 4);
    tx.setAttribute("text-anchor", "middle");
    tx.setAttribute("font-family", CONFIG.fontFamily);
    tx.setAttribute("font-size", "12");
    tx.setAttribute("font-weight", "bold");
    tx.setAttribute("fill", "#111111");
    tx.textContent = rotulo;
    svg.appendChild(tx);
  }
}

function desenharConexaoExcel(
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
  } else if (destino.id !== "__FIM__" && destino.id !== "__INICIO__") {
    const end = rota.points[rota.points.length - 1];
    sharedRegistry[sharedKey] = {
      origemId: origem.id,
      sourceGridCol: origem.gridCol,
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
  path.setAttribute("stroke-linejoin", "round");
  path.setAttribute("stroke-linecap", "round");
  svg.appendChild(path);

  const p1 = rota.points[rota.points.length - 2];
  const p2 = rota.points[rota.points.length - 1];

  const arrow = criarElementoSVG("path");
  let dArrow = "";

  if (p2.x > p1.x) {
    dArrow = `M ${p2.x - 10} ${p2.y - 4} L ${p2.x} ${p2.y} L ${p2.x - 10} ${p2.y + 4}`;
  } else if (p2.x < p1.x) {
    dArrow = `M ${p2.x + 10} ${p2.y - 4} L ${p2.x} ${p2.y} L ${p2.x + 10} ${p2.y + 4}`;
  } else if (p2.y > p1.y) {
    dArrow = `M ${p2.x - 4} ${p2.y - 10} L ${p2.x} ${p2.y} L ${p2.x + 4} ${p2.y - 10}`;
  } else {
    dArrow = `M ${p2.x - 4} ${p2.y + 10} L ${p2.x} ${p2.y} L ${p2.x + 4} ${p2.y + 10}`;
  }

  arrow.setAttribute("d", dArrow);
  arrow.setAttribute("fill", "none");
  arrow.setAttribute("stroke", "#111111");
  arrow.setAttribute("stroke-width", CONFIG.lineWidth);
  arrow.setAttribute("stroke-linejoin", "round");
  arrow.setAttribute("stroke-linecap", "round");
  svg.appendChild(arrow);

  if (rotulo) {
    const larguraTexto = medirLarguraTexto(rotulo, 12, "bold");
    const larguraCapsula = Math.max(42, larguraTexto + 20);
    const alturaCapsula = 24;
    const xCapsula = rota.label.x - larguraCapsula / 2;
    const yCapsula = rota.label.y - alturaCapsula / 2 - 1;

    const bg = criarElementoSVG("rect");
    bg.setAttribute("x", xCapsula);
    bg.setAttribute("y", yCapsula);
    bg.setAttribute("width", larguraCapsula);
    bg.setAttribute("height", alturaCapsula);
    bg.setAttribute("rx", alturaCapsula / 2);
    bg.setAttribute("ry", alturaCapsula / 2);
    bg.setAttribute("fill", "#ffffff");
    bg.setAttribute("stroke", "#111111");
    bg.setAttribute("stroke-width", "1");
    svg.appendChild(bg);

    const tx = criarElementoSVG("text");
    tx.setAttribute("x", rota.label.x);
    tx.setAttribute("y", rota.label.y + 4);
    tx.setAttribute("text-anchor", "middle");
    tx.setAttribute("font-family", CONFIG.fontFamily);
    tx.setAttribute("font-size", "12");
    tx.setAttribute("font-weight", "bold");
    tx.setAttribute("fill", "#111111");
    tx.textContent = rotulo;
    svg.appendChild(tx);
  }
}

function gerarHTMLResumoTempo(lista, tempoTotal) {
  return lista
    .map(item => {
      const pct = tempoTotal ? formatarPercentual((item.tempo / tempoTotal) * 100) : "0,0";
      return '<div class="analytics-item">' +
        escaparHTML(item.nome) +
        ' — <span class="icon-time">⏱</span>' + formatarTempo(item.tempo) +
        ' <span class="icon-pct">%</span>' + pct + '%' +
      '</div>';
    })
    .join("");
}

function gerarTabelaPareto(atividadesTempo, tempoTotal) {
  let acumulado = 0;

  const linhas = atividadesTempo.map((item) => {
    const percentual = tempoTotal ? (item.tempo / tempoTotal) * 100 : 0;
    acumulado += percentual;

    return (
      '<tr>' +
        '<td style="padding:8px;border:1px solid #d9d9d9;vertical-align:top;">' + escaparHTML(item.atividade) + '</td>' +
        '<td style="padding:8px;border:1px solid #d9d9d9;white-space:nowrap;text-align:center;">⏱ ' + formatarTempo(item.tempo) + '</td>' +
        '<td style="padding:8px;border:1px solid #d9d9d9;white-space:nowrap;text-align:center;">' + formatarPercentual(percentual) + '%</td>' +
        '<td style="padding:8px;border:1px solid #d9d9d9;white-space:nowrap;text-align:center;">' + formatarPercentual(acumulado) + '%</td>' +
      '</tr>'
    );
  }).join("");

  return (
    '<div style="overflow-x:auto;">' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
        '<thead>' +
          '<tr>' +
            '<th style="padding:10px;border:1px solid #d9d9d9;background:#f5f5f5;text-align:left;">Atividade</th>' +
            '<th style="padding:10px;border:1px solid #d9d9d9;background:#f5f5f5;text-align:center;">Tempo</th>' +
            '<th style="padding:10px;border:1px solid #d9d9d9;background:#f5f5f5;text-align:center;">%</th>' +
            '<th style="padding:10px;border:1px solid #d9d9d9;background:#f5f5f5;text-align:center;">Pareto</th>' +
          '</tr>' +
        '</thead>' +
        '<tbody>' + linhas + '</tbody>' +
      '</table>' +
    '</div>'
  );
}

function renderInformacoesProcessoExecutivas(info) {
  return `
    <div class="exec-card">
      <div class="exec-card-title">Informações do Processo</div>
      <div class="exec-info-grid">
        <div class="exec-info-item">
          <div class="exec-info-label">Desenho</div>
          <div class="exec-info-value">${escaparHTML(info.desenho || "Não informado")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Processo</div>
          <div class="exec-info-value">${escaparHTML(info.processo || "Não informado")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Analista</div>
          <div class="exec-info-value">${escaparHTML(info.analista || "Não informado")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Negócio</div>
          <div class="exec-info-value">${escaparHTML(info.negocio || "Não informado")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Área</div>
          <div class="exec-info-value">${escaparHTML(info.area || "Não informado")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Gestor</div>
          <div class="exec-info-value">${escaparHTML(info.gestor || "Não informado")}</div>
        </div>
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
        <tr>
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
    </div>
  `;
}

function renderizarAnaliseExecutiva(dados) {
  const top3Rows = dados.top3Gargalos.map(item => ({
    atividade: item.atividade,
    tempoFmt: formatarTempo(item.tempo),
    percentualFmt: `${formatarPercentual(item.percentual)}%`
  }));

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
        titulo: "Top 3 Gargalos",
        columns: [
          { header: "Atividade", key: "atividade", align: "left" },
          { header: "Tempo (horas)", key: "tempoFmt", align: "center" },
          { header: "%", key: "percentualFmt", align: "center" }
        ],
        rows: top3Rows
      })}

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
          { header: "Pareto", key: "paretoFmt", align: "center" }
        ],
        rows: paretoRows
      })}
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
  while (col.length < 13) col.push("");

  const ordem = Number(limpar(col[0])) || 0;
  const id = limpar(col[1]);
  const areaEtapa = limpar(col[2]) || "Sem Área";
  const atividade = limpar(col[3]);
  const tipo = limpar(col[4]) || "Não informado";
  const sistema = limpar(col[5]) || "Sem sistema informado";
  const tempo = tempoParaSegundos(limpar(col[6]));
  const proxSim = limpar(col[7]);
  const proxNao = limpar(col[8]);
  const conexoesExtras = limpar(col[9]);
  const coluna = Number(limpar(col[10])) || 1;
  const linha = Number(limpar(col[11])) || 1;
  const cor = normalizarCor(col[12]);

    if (!id || !atividade) return;

    etapas.push({
      ordem,
      id,
      atividade,
      area: areaEtapa,
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

  const alturaPadraoNos = calcularAlturaPadraoNos(etapas);
  const maiorAlturaLosango = Math.ceil(alturaPadraoNos * CONFIG.decisionHeightFactor);
  const rowSlotHeight = Math.max(alturaPadraoNos, maiorAlturaLosango);
  const colSlotWidth = Math.max(CONFIG.boxWidth, CONFIG.decisionWidth);

  const maxColuna = Math.max(...etapas.map(e => e.coluna), 1) + 1;

  const areasOrdenadas = [];
  const areaJaExiste = new Set();
  etapas.forEach((e) => {
    const nome = e.area || "Sem Área";
    if (!areaJaExiste.has(nome)) {
      areaJaExiste.add(nome);
      areasOrdenadas.push(nome);
    }
  });

  const linhasPorArea = {};
  areasOrdenadas.forEach((nome) => {
    linhasPorArea[nome] = 1;
  });

  etapas.forEach((e) => {
    const nome = e.area || "Sem Área";
    linhasPorArea[nome] = Math.max(linhasPorArea[nome] || 1, e.linha || 1);
  });

  const laneContentWidth =
    maxColuna * colSlotWidth +
    (maxColuna - 1) * CONFIG.colGap;

  const lanes = {};
  let cursorY = CONFIG.marginY;
  let rowOffsetGlobal = 0;

  areasOrdenadas.forEach((nome) => {
    const qtdLinhas = linhasPorArea[nome] || 1;

    const contentHeight =
      qtdLinhas * rowSlotHeight +
      (qtdLinhas - 1) * CONFIG.rowGap;

    const laneHeight =
      CONFIG.lanePaddingTop +
      contentHeight +
      CONFIG.lanePaddingBottom;

    lanes[nome] = {
      x: CONFIG.marginX,
      y: cursorY,
      width: CONFIG.laneLabelWidth + CONFIG.laneEntryWidth + laneContentWidth,
      height: laneHeight,
      contentX: CONFIG.marginX + CONFIG.laneLabelWidth + CONFIG.laneEntryWidth,
      contentY: cursorY + CONFIG.lanePaddingTop,
      rows: qtdLinhas,
      rowOffsetGlobalStart: rowOffsetGlobal
    };

    rowOffsetGlobal += qtdLinhas + 2;
    cursorY += laneHeight + CONFIG.laneGap;
  });

  const svgWidth =
  CONFIG.marginX * 2 +
  CONFIG.laneLabelWidth +
  CONFIG.laneEntryWidth +
  laneContentWidth;

  const svgHeight = cursorY;

  const svg = criarElementoSVG("svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", svgWidth);
  svg.setAttribute("height", svgHeight);
  svg.setAttribute("viewBox", `0 0 ${svgWidth} ${svgHeight}`);
  svg.setAttribute("style", "background:#ffffff");

  const style = criarElementoSVG("style");
style.textContent = `
  .box {
    stroke: #111111;
    stroke-width: 2;
  }

  .green { fill: #95d5b2; }
  .blue { fill: #8ecae6; }
  .yellow { fill: #ffd166; }
  .red { fill: #ef476f; }
  .white { fill: #ffffff; }

  .system {
    fill: #f5f5f5;
    stroke: #111111;
    stroke-width: 1.5;
  }

  .text {
    font-family: Arial, sans-serif;
    fill: #111111;
  }
`;
svg.appendChild(style);

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

  desenharRaias(svg, areasOrdenadas, lanes, svgWidth);

  const posicoes = {};

  etapas.forEach((e) => {
    const pergunta = isPergunta(e.atividade);
    const w = pergunta ? CONFIG.decisionWidth : CONFIG.boxWidth;
    const h = obterAlturaNo(e, alturaPadraoNos);
    const lane = lanes[e.area || "Sem Área"];

    const slotX = lane.contentX + (e.coluna - 1) * (colSlotWidth + CONFIG.colGap);
    const slotY = lane.contentY + (e.linha - 1) * (rowSlotHeight + CONFIG.rowGap);

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
      gridRow: e.linha,
      gridRowGlobal: lane.rowOffsetGlobalStart + e.linha,
      area: e.area || "Sem Área"
    };
  });

  const primeiraEtapa = etapas[0];
  const ultimaEtapa = etapas[etapas.length - 1];
  const primeiraPos = posicoes[primeiraEtapa.id];
  const ultimaPos = posicoes[ultimaEtapa.id];

  const primeiraLane = lanes[primeiraEtapa.area];

posicoes["__INICIO__"] = {
  id: "__INICIO__",
  x: primeiraLane.x + CONFIG.laneLabelWidth + 20,
  y: primeiraPos.y + (primeiraPos.h - 36) / 2,
  w: 60,
  h: 36,
  isDecision: false,
  gridCol: Math.max(0, primeiraPos.gridCol - 1),
  gridRow: primeiraPos.gridRow,
  gridRowGlobal: primeiraPos.gridRowGlobal,
  area: primeiraPos.area
};

  posicoes["__FIM__"] = {
    id: "__FIM__",
    x: ultimaPos.x + ultimaPos.w + CONFIG.entryExitGap,
    y: ultimaPos.y + (ultimaPos.h - 36) / 2,
    w: 60,
    h: 36,
    isDecision: false,
    gridCol: ultimaPos.gridCol + 1,
    gridRow: ultimaPos.gridRow,
    gridRowGlobal: ultimaPos.gridRowGlobal,
    area: ultimaPos.area
  };

  desenharCapsula(svg, "Início", posicoes["__INICIO__"].x, posicoes["__INICIO__"].y, 60, 36);
  desenharCapsula(svg, "Fim", posicoes["__FIM__"].x, posicoes["__FIM__"].y, 60, 36);

  etapas.forEach((etapa) => {
    desenharNo(svg, etapa, posicoes[etapa.id]);
  });

  let loops = 0;
  let decisoes = 0;
  let conexoesExtrasCount = 0;
  const etapasImpactadasRetrabalho = new Set();
  const sharedRegistry = {};

  desenharConexao(svg, posicoes["__INICIO__"], posicoes[primeiraEtapa.id], "", 0, posicoes, sharedRegistry);

  etapas.forEach((etapa) => {
    const origem = posicoes[etapa.id];
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    const pergunta = isPergunta(etapa.atividade);
    if (pergunta) decisoes++;

    destinosSim.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa);
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
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
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa);
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
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
        loops += adicionarLoopSeNecessario(etapa.id, destinoId, etapaPorId, etapa);
        adicionarEtapasImpactadasPorRetorno(etapa.id, destinoId, etapaPorId, etapas, etapasImpactadasRetrabalho);
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

  let tempoTotal = 0;
  const atividadesTempo = [];
  const tiposTempo = {};
  const sistemasTempo = {};

  etapas.forEach((etapa) => {
    tempoTotal += etapa.tempo;
    atividadesTempo.push({ atividade: etapa.atividade, tempo: etapa.tempo });

    if (!tiposTempo[etapa.tipo]) tiposTempo[etapa.tipo] = 0;
    tiposTempo[etapa.tipo] += etapa.tempo;

    if (!sistemasTempo[etapa.sistema]) sistemasTempo[etapa.sistema] = 0;
    sistemasTempo[etapa.sistema] += etapa.tempo;
  });

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const tiposOrdenados = Object.entries(tiposTempo)
    .map(([nome, tempo]) => ({ nome, tempo }))
    .sort((a, b) => b.tempo - a.tempo);

  const sistemasOrdenados = Object.entries(sistemasTempo)
    .map(([nome, tempo]) => ({ nome, tempo }))
    .sort((a, b) => b.tempo - a.tempo);

  let tempoPotencialRetrabalho = 0;
  etapas.forEach((etapa) => {
    if (etapasImpactadasRetrabalho.has(etapa.id)) {
      tempoPotencialRetrabalho += etapa.tempo;
    }
  });

  const impactoPotencialRetrabalhoNum = tempoTotal
    ? (tempoPotencialRetrabalho / tempoTotal) * 100
    : 0;

  const taxaDecisaoNum = etapas.length
    ? (decisoes / etapas.length) * 100
    : 0;

  const infoProcessoData = {
    desenho,
    processo,
    analista,
    negocio,
    area,
    gestor
  };

  document.getElementById("infoProcesso").innerHTML =
    renderInformacoesProcessoExecutivas(infoProcessoData);

  const dadosAnalise = {
    tempoTotal,
    loops,
    conexoesExtrasCount,
    tempoPotencialRetrabalho,
    impactoPotencialRetrabalho: impactoPotencialRetrabalhoNum,
    decisoes,
    taxaDecisao: taxaDecisaoNum,
    top3Gargalos: atividadesTempo.slice(0, 3).map(item => ({
      atividade: item.atividade,
      tempo: item.tempo,
      percentual: tempoTotal ? (item.tempo / tempoTotal) * 100 : 0
    })),
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
    })()
  };

  

  document.getElementById("metricas").innerHTML =
    renderizarAnaliseExecutiva(dadosAnalise);

}

function gerarFluxoExcel() {
  const texto = document.getElementById("entrada").value;

  if (!texto.trim()) {
    alert("Cole a tabela do Excel primeiro.");
    return;
  }

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
    while (col.length < 13) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const areaEtapa = limpar(col[2]) || "Sem Área";
    const atividade = limpar(col[3]);
    const tipo = limpar(col[4]) || "Não informado";
    const sistema = limpar(col[5]) || "Sem sistema informado";
    const tempo = tempoParaSegundos(limpar(col[6]));
    const proxSim = limpar(col[7]);
    const proxNao = limpar(col[8]);
    const conexoesExtras = limpar(col[9]);
    const coluna = Number(limpar(col[10])) || 1;
    const linha = Number(limpar(col[11])) || 1;
    const cor = normalizarCor(col[12]);

    if (!id || !atividade) return;

    etapas.push({
      ordem,
      id,
      atividade,
      area: areaEtapa,
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

  const alturaPadraoNos = calcularAlturaPadraoNos(etapas);
  const maiorAlturaLosango = Math.ceil(alturaPadraoNos * CONFIG.decisionHeightFactor);
  const rowSlotHeight = Math.max(alturaPadraoNos, maiorAlturaLosango);
  const colSlotWidth = Math.max(CONFIG.boxWidth, CONFIG.decisionWidth);

  const maxColuna = Math.max(...etapas.map(e => e.coluna), 1) + 1;

  const areasOrdenadas = [];
  const areaJaExiste = new Set();
  etapas.forEach((e) => {
    const nome = e.area || "Sem Área";
    if (!areaJaExiste.has(nome)) {
      areaJaExiste.add(nome);
      areasOrdenadas.push(nome);
    }
  });

  const laneContentWidth =
    maxColuna * colSlotWidth +
    (maxColuna - 1) * EXCEL_LAYOUT.colGap;

  const laneLeft = CONFIG.marginX + EXCEL_LAYOUT.extraLeftPadding;
  const laneTop = CONFIG.marginY;

  let cursorY = laneTop;
  const lanes = [];
  let rowOffsetGlobal = 0;

  areasOrdenadas.forEach((nomeArea) => {
    const etapasArea = etapas.filter(e => (e.area || "Sem Área") === nomeArea);
    const maxLinhaArea = Math.max(...etapasArea.map(e => e.linha), 1);
    const qtdLinhas = maxLinhaArea;

    const contentHeight =
      qtdLinhas * rowSlotHeight +
      (qtdLinhas - 1) * EXCEL_LAYOUT.rowGap;

    const laneHeight =
      EXCEL_LAYOUT.lanePaddingTop +
      contentHeight +
      EXCEL_LAYOUT.lanePaddingBottom;

  lanes.push({
    area: nomeArea,
    x: laneLeft,
    y: cursorY,
    width: EXCEL_LAYOUT.laneLabelWidth + EXCEL_LAYOUT.laneEntryWidth + laneContentWidth,
    height: laneHeight,
    contentX: laneLeft + EXCEL_LAYOUT.laneLabelWidth + EXCEL_LAYOUT.laneEntryWidth,
    contentY: cursorY + EXCEL_LAYOUT.lanePaddingTop,
    contentHeight,
    rows: qtdLinhas,
    rowOffsetGlobalStart: rowOffsetGlobal
  });

    rowOffsetGlobal += qtdLinhas + 2;

    cursorY += laneHeight + EXCEL_LAYOUT.laneGap;
  });

  const larguraSvg = Math.max(
  CONFIG.marginX * 2 +
    EXCEL_LAYOUT.laneLabelWidth +
    EXCEL_LAYOUT.laneEntryWidth +
    laneContentWidth +
    100,
  1000
);

  const alturaSvg = Math.max(cursorY + CONFIG.marginY, 500);

  const svg = criarElementoSVG("svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", larguraSvg);
  svg.setAttribute("height", alturaSvg);
  svg.setAttribute("viewBox", `0 0 ${larguraSvg} ${alturaSvg}`);
  svg.setAttribute("style", "background:#ffffff");

  desenharRaiasExcel(svg, lanes);

  const laneByArea = {};
  lanes.forEach(lane => {
    laneByArea[lane.area] = lane;
  });

  const posicoes = {};

  etapas.forEach((e) => {
    const lane = laneByArea[e.area || "Sem Área"];
    const slotX = lane.contentX + (e.coluna - 1) * (colSlotWidth + EXCEL_LAYOUT.colGap);
    const slotY = lane.contentY + (e.linha - 1) * (rowSlotHeight + EXCEL_LAYOUT.rowGap);

    const pergunta = isPergunta(e.atividade);
    const w = pergunta ? CONFIG.decisionWidth : CONFIG.boxWidth;
    const h = pergunta
      ? Math.ceil(alturaPadraoNos * CONFIG.decisionHeightFactor)
      : alturaPadraoNos;

    const x = slotX + (colSlotWidth - w) / 2;
    const y = slotY + (rowSlotHeight - h) / 2;

        posicoes[e.id] = {
      id: e.id,
      x,
      y,
      w,
      h,
      cx: x + w / 2,
      cy: y + h / 2,
      top: y,
      bottom: y + h,
      left: x,
      right: x + w,
      isDecision: pergunta,
      gridCol: e.coluna,
      gridRow: e.linha,
      gridRowGlobal: lane.rowOffsetGlobalStart + e.linha,
      area: e.area || "Sem Área"
    };
  });

  const primeiraEtapa = etapas[0];
  const ultimaEtapa = etapas[etapas.length - 1];

  const lanePrimeira = laneByArea[primeiraEtapa.area || "Sem Área"];
  const laneUltima = laneByArea[ultimaEtapa.area || "Sem Área"];

  posicoes["__INICIO__"] = {
  id: "__INICIO__",
  x: posicoes[primeiraEtapa.id].x - 60 - EXCEL_LAYOUT.startGap,
  y: posicoes[primeiraEtapa.id].cy - 18,
  w: 60,
  h: 36,
  cx: posicoes[primeiraEtapa.id].x - EXCEL_LAYOUT.startGap - 30,
  cy: posicoes[primeiraEtapa.id].cy,
  top: posicoes[primeiraEtapa.id].cy - 18,
  bottom: posicoes[primeiraEtapa.id].cy + 18,
  left: posicoes[primeiraEtapa.id].x - 60 - EXCEL_LAYOUT.startGap,
  right: posicoes[primeiraEtapa.id].x - EXCEL_LAYOUT.startGap,
  isDecision: false,
  gridCol: 0,
  gridRow: primeiraEtapa.linha,
  gridRowGlobal: lanePrimeira.rowOffsetGlobalStart + primeiraEtapa.linha,
  area: primeiraEtapa.area || "Sem Área"
};

  posicoes["__FIM__"] = {
  id: "__FIM__",
  x: posicoes[ultimaEtapa.id].x + posicoes[ultimaEtapa.id].w + EXCEL_LAYOUT.endGap,
  y: posicoes[ultimaEtapa.id].cy - 18,
  w: 60,
  h: 36,
  cx: posicoes[ultimaEtapa.id].x + posicoes[ultimaEtapa.id].w + EXCEL_LAYOUT.endGap + 30,
  cy: posicoes[ultimaEtapa.id].cy,
  top: posicoes[ultimaEtapa.id].cy - 18,
  bottom: posicoes[ultimaEtapa.id].cy + 18,
  left: posicoes[ultimaEtapa.id].x + posicoes[ultimaEtapa.id].w + EXCEL_LAYOUT.endGap,
  right: posicoes[ultimaEtapa.id].x + posicoes[ultimaEtapa.id].w + EXCEL_LAYOUT.endGap + 60,
  isDecision: false,
  gridCol: ultimaEtapa.coluna + 1,
  gridRow: ultimaEtapa.linha,
  gridRowGlobal: laneUltima.rowOffsetGlobalStart + ultimaEtapa.linha,
  area: ultimaEtapa.area || "Sem Área"
};

  desenharCapsula(svg, "Início", posicoes["__INICIO__"].x, posicoes["__INICIO__"].y, 60, 36);
  desenharCapsula(svg, "Fim", posicoes["__FIM__"].x, posicoes["__FIM__"].y, 60, 36);

  etapas.forEach((etapa) => {
    desenharNoExcel(svg, etapa, posicoes[etapa.id]);
  });

  const sharedRegistry = {};
  desenharConexaoExcel(
    svg,
    posicoes["__INICIO__"],
    posicoes[primeiraEtapa.id],
    "",
    0,
    posicoes,
    sharedRegistry
  );

  etapas.forEach((etapa) => {
    const origem = posicoes[etapa.id];
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    const pergunta = isPergunta(etapa.atividade);

    destinosSim.forEach((destinoId, indice) => {
      desenharConexaoExcel(
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
      desenharConexaoExcel(
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
      desenharConexaoExcel(
        svg,
        origem,
        posicoes[destinoId],
        "",
        indice + 1,
        posicoes,
        sharedRegistry
      );
    });
  });

  etapas.forEach((etapa) => {
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    if (destinosSim.length === 0 && destinosNao.length === 0 && destinosExtras.length === 0) {
      desenharConexaoExcel(
        svg,
        posicoes[etapa.id],
        posicoes["__FIM__"],
        "",
        0,
        posicoes,
        sharedRegistry
      );
    }
  });

  return aplicarEscalaSVGExcel(svg, EXCEL_EXPORT_SCALE);
}

/* =========================
   EXPORTAÇÃO SVG E PDF
========================= */

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

  const linhasBrutas = texto
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter(l => l.trim() !== "");

  let linhas = linhasBrutas.map(l => l.split("\t"));

  if (linhas.length && ehCabecalho(linhas[0])) {
    linhas.shift();
  }

 linhas.forEach((col) => {
  while (col.length < 13) col.push("");

  const ordem = Number(limpar(col[0])) || 0;
  const id = limpar(col[1]);
  const areaEtapa = limpar(col[2]) || "Sem Área";
  const atividade = limpar(col[3]);
  const tipo = limpar(col[4]) || "Não informado";
  const sistema = limpar(col[5]) || "Sem sistema informado";
  const tempo = tempoParaSegundos(limpar(col[6]));
  const proxSim = limpar(col[7]);
  const proxNao = limpar(col[8]);
  const conexoesExtras = limpar(col[9]);

    if (!id || !atividade) return;

    etapas.push({
      ordem,
      id,
      atividade,
      area: areaEtapa,
      tipo,
      sistema,
      tempo,
      proxSim,
      proxNao,
      conexoesExtras
    });
  });

  etapas.sort((a, b) => a.ordem - b.ordem);

  const etapaPorId = {};
  etapas.forEach(e => {
    etapaPorId[e.id] = e;
  });

  let tempoTotal = 0;
  let loops = 0;
  let decisoes = 0;
  let conexoesExtrasCount = 0;
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
    if (etapasImpactadasRetrabalho.has(etapa.id)) {
      tempoPotencialRetrabalho += etapa.tempo;
    }
  });

  const impactoPotencialRetrabalho = tempoTotal
    ? (tempoPotencialRetrabalho / tempoTotal) * 100
    : 0;

  const taxaDecisao = etapas.length
    ? (decisoes / etapas.length) * 100
    : 0;

  const top3Gargalos = [...etapas]
    .sort((a, b) => b.tempo - a.tempo)
    .slice(0, 3)
    .map((e) => ({
      atividade: e.atividade,
      tempo: e.tempo,
      percentual: tempoTotal ? (e.tempo / tempoTotal) * 100 : 0
    }));

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
    .map((e) => ({
      atividade: e.atividade,
      tempo: e.tempo
    }));

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
    top3Gargalos,
    tempoPorTipo,
    tempoPorSistema,
    pareto
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

  return el.innerText
    .split("\n")
    .map(l => l.trim())
    .filter(Boolean);
}

function adicionarTextoQuebrado(doc, texto, x, y, maxWidth, lineHeight = 14, options = {}) {
  const linhas = doc.splitTextToSize(String(texto || ""), maxWidth);
  if (linhas.length === 0) return y;
  doc.text(linhas, x, y, options);
  return y + linhas.length * lineHeight;
}

function garantirEspacoPagina(doc, yAtual, alturaNecessaria, margem, pageHeight, onNewPage = null) {
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

function limparTextoPDF(txt) {
  return String(txt || "")
    .replace(/⏱/g, "")
    .replace(/\s+/g, " ")
    .trim();
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
    pageHeight
  } = config;

  const borderWidth = 0.6;
  const cellPaddingX = 6;
  const lineHeight = 11;
  const minRowHeight = 20;
  const gapAntesTitulo = 16;
  const gapTituloCabecalho = 6;
  const gapDepoisTabela = 22;

  let y = yInicial + gapAntesTitulo;

  const weights = columns.map(col => col.weight || 1);
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  const colWidths = weights.map(w => (larguraTotal * w) / totalWeight);

  const getHeaderHeight = () => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);

    let maxLines = 1;
    columns.forEach((col, i) => {
      const linhas = doc.splitTextToSize(col.header, colWidths[i] - cellPaddingX * 2);
      if (linhas.length > maxLines) maxLines = linhas.length;
    });

    return Math.max(minRowHeight, maxLines * lineHeight + 12);
  };

  const getCellBlock = (valor, width, align) => {
    if (align === "left") {
      return doc.splitTextToSize(valor, width - cellPaddingX * 2);
    }
    return [valor];
  };

  const drawTableHeader = (yHeader) => {
    const headerHeight = getHeaderHeight();

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.setLineWidth(borderWidth);

    let currentX = x;

    columns.forEach((col, i) => {
      const width = colWidths[i];
      const linhas = doc.splitTextToSize(col.header, width - cellPaddingX * 2);

      doc.rect(currentX, yHeader, width, headerHeight);

      const totalTextHeight = linhas.length * lineHeight;
      const startY = yHeader + (headerHeight - totalTextHeight) / 2 + 8;

      linhas.forEach((linha, idx) => {
        doc.text(linha, currentX + width / 2, startY + idx * lineHeight, {
          align: "center"
        });
      });

      currentX += width;
    });

    return yHeader + headerHeight;
  };

  const drawTitleAndHeader = (yStart) => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(titulo, x, yStart);

    let yLocal = yStart + gapTituloCabecalho;
    yLocal = drawTableHeader(yLocal);
    return yLocal;
  };

  y = garantirEspacoPagina(doc, y, 18 + gapTituloCabecalho + getHeaderHeight(), margem, pageHeight);
  y = drawTitleAndHeader(y);

  rows.forEach((row) => {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setLineWidth(borderWidth);

    let maxLines = 1;

    const rowLineCache = columns.map((col, i) => {
      const raw = row[col.key] !== undefined && row[col.key] !== null ? String(row[col.key]) : "";
      const valor = limparTextoPDF(raw);
      const align = col.align || "left";
      const linhas = getCellBlock(valor, colWidths[i], align);

      if (linhas.length > maxLines) maxLines = linhas.length;
      return linhas;
    });

    const rowHeight = Math.max(minRowHeight, maxLines * lineHeight + 12);

    y = garantirEspacoPagina(
      doc,
      y,
      rowHeight,
      margem,
      pageHeight,
      (novoY) => {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(12);
        doc.setLineWidth(borderWidth);
        return drawTitleAndHeader(novoY);
      }
    );

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setLineWidth(borderWidth);

    let currentX = x;

    columns.forEach((col, i) => {
      const width = colWidths[i];
      const align = col.align || "left";
      const linhas = rowLineCache[i];

      doc.rect(currentX, y, width, rowHeight);

      const totalTextHeight = linhas.length * lineHeight;
      const startY = y + (rowHeight - totalTextHeight) / 2 + 8;

      if (align === "center") {
        linhas.forEach((linha, idx) => {
          doc.text(linha, currentX + width / 2, startY + idx * lineHeight, {
            align: "center"
          });
        });
      } else if (align === "right") {
        linhas.forEach((linha, idx) => {
          doc.text(linha, currentX + width - cellPaddingX, startY + idx * lineHeight, {
            align: "right"
          });
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

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margem = 40;
  const larguraUtil = pageWidth - margem * 2;

  let y = margem;

  const processoNome = obterValorCampo("processo") || "Fluxograma do Processo";

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  y = adicionarTextoQuebrado(doc, processoNome, margem, y, larguraUtil, 18);
  y += 8;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  y = garantirEspacoPagina(doc, y, 24, margem, pageHeight);
  doc.text("Informações do Processo", margem, y);
  y += 16;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);

  infoLinhas.forEach((linha) => {
    y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
    y = adicionarTextoQuebrado(doc, linha, margem, y, larguraUtil, 13);
  });

  y += 14;

  const svgWidth = Number(svg.getAttribute("width")) || 1200;
  const svgHeight = Number(svg.getAttribute("height")) || 800;

  const escala = Math.min(larguraUtil / svgWidth, 1);
  const larguraSvgPdf = svgWidth * escala;
  const alturaSvgPdf = svgHeight * escala;

  y = garantirEspacoPagina(doc, y, alturaSvgPdf + 20, margem, pageHeight);

  await doc.svg(svg, {
    x: margem + (larguraUtil - larguraSvgPdf) / 2,
    y,
    width: larguraSvgPdf,
    height: alturaSvgPdf
  });

  y += alturaSvgPdf + 24;
  y = garantirEspacoPagina(doc, y, 28, margem, pageHeight);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("Análise do Processo", margem, y);
  y += 18;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);

  y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
  doc.text(`Tempo total do processo: ${formatarTempo(dados.tempoTotal)}`, margem, y);
  y += 18;

  y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
  doc.text(`Loops detectados: ${dados.loops}`, margem, y);
  y += 18;

  y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
  doc.text(
    `Potencial retrabalho: ${formatarTempo(dados.tempoPotencialRetrabalho)} | ${formatarPercentual(dados.impactoPotencialRetrabalho)}%`,
    margem,
    y
  );
  y += 18;

  y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
  doc.text(
    `Taxa de decisão: ${dados.decisoes} etapa(s) | ${formatarPercentual(dados.taxaDecisao)}%`,
    margem,
    y
  );
  y += 6;

  y = desenharTabelaPDF(doc, {
    titulo: "Top 3 Gargalos",
    columns: [
      { header: "Atividade", key: "atividade", weight: 5.5, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.6, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.2, align: "center" }
    ],
    rows: dados.top3Gargalos.map(item => ({
      atividade: item.atividade,
      tempoFmt: formatarTempo(item.tempo),
      percentualFmt: `${formatarPercentual(item.percentual)}%`
    })),
    x: margem,
    yInicial: y,
    larguraTotal: larguraUtil,
    margem,
    pageHeight
  });

  y = desenharTabelaPDF(doc, {
    titulo: "Tempo por Tipo",
    columns: [
      { header: "Tipo", key: "tipo", weight: 5.5, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.6, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.2, align: "center" }
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
    pageHeight
  });

  y = desenharTabelaPDF(doc, {
    titulo: "Tempo por Sistema",
    columns: [
      { header: "Sistema", key: "sistema", weight: 5.5, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.6, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.2, align: "center" }
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
    pageHeight
  });

  y = desenharTabelaPDF(doc, {
    titulo: "Pareto de Tempo",
    columns: [
      { header: "Atividade", key: "atividade", weight: 5.4, align: "left" },
      { header: "Tempo (horas)", key: "tempoFmt", weight: 1.5, align: "center" },
      { header: "%", key: "percentualFmt", weight: 1.0, align: "center" },
      { header: "Pareto", key: "paretoFmt", weight: 1.3, align: "center" }
    ],
    rows: dados.pareto.map(item => ({
      atividade: item.atividade,
      tempoFmt: formatarTempo(item.tempo),
      percentualFmt: `${formatarPercentual(item.percentual)}%`,
      paretoFmt: `${formatarPercentual(item.pareto)}%`
    })),
    x: margem,
    yInicial: y,
    larguraTotal: larguraUtil,
    margem,
    pageHeight
  });

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

function baixarSVGExcel() {
  const svg = gerarFluxoExcel();
  if (!svg) return;

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const blob = new Blob([source], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = `${gerarNomeArquivo()}_excel.svg`;
  link.click();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function baixarFluxo() {
  baixarSVG();
}