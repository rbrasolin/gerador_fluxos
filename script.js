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

function desenharRetangulo(svg, etapa, x, y, width, height, fill) {
  const group = criarElementoSVG("g");
  group.setAttribute("class", "node");

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

function desenharLosango(svg, etapa, x, y, width, height, fill) {
  const group = criarElementoSVG("g");
  group.setAttribute("class", "node");

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

function desenharLabelLinha(svg, texto, x, y) {
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

function obterCentroNo(no) {
  return {
    x: no.x + no.width / 2,
    y: no.y + no.height / 2
  };
}

function obterPortasNo(no) {
  return {
    left: { x: no.x, y: no.y + no.height / 2 },
    right: { x: no.x + no.width, y: no.y + no.height / 2 },
    top: { x: no.x + no.width / 2, y: no.y },
    bottom: { x: no.x + no.width / 2, y: no.y + no.height }
  };
}

function desenharCaminho(svg, pontos, label = "") {
  if (!pontos || pontos.length < 2) return;

  const path = criarElementoSVG("path");
  let d = `M ${pontos[0].x} ${pontos[0].y}`;

  for (let i = 1; i < pontos.length; i++) {
    d += ` L ${pontos[i].x} ${pontos[i].y}`;
  }

  path.setAttribute("d", d);
  path.setAttribute("fill", "none");
  path.setAttribute("stroke", CONFIG.stroke);
  path.setAttribute("stroke-width", CONFIG.lineWidth);
  path.setAttribute("marker-end", "url(#arrowhead)");
  path.setAttribute("stroke-linejoin", "round");
  path.setAttribute("stroke-linecap", "round");
  svg.appendChild(path);

  if (label && pontos.length >= 2) {
    let melhorSegmento = null;
    let maiorDistancia = 0;

    for (let i = 1; i < pontos.length; i++) {
      const a = pontos[i - 1];
      const b = pontos[i];
      const distancia = Math.abs(b.x - a.x) + Math.abs(b.y - a.y);

      if (distancia > maiorDistancia) {
        maiorDistancia = distancia;
        melhorSegmento = { a, b };
      }
    }

    if (melhorSegmento) {
      const x = (melhorSegmento.a.x + melhorSegmento.b.x) / 2;
      const y = (melhorSegmento.a.y + melhorSegmento.b.y) / 2 - 14;
      desenharLabelLinha(svg, label, x, y);
    }
  }
}

function calcularRotaDireta(origem, destino, label = "") {
  const portasOrigem = obterPortasNo(origem);
  const portasDestino = obterPortasNo(destino);

  const centroOrigem = obterCentroNo(origem);
  const centroDestino = obterCentroNo(destino);

  const dx = centroDestino.x - centroOrigem.x;
  const dy = centroDestino.y - centroOrigem.y;

  if (Math.abs(dx) >= Math.abs(dy)) {
    if (dx >= 0) {
      return {
        pontos: [
          portasOrigem.right,
          { x: portasDestino.left.x - CONFIG.routeGap, y: portasOrigem.right.y },
          { x: portasDestino.left.x - CONFIG.routeGap, y: portasDestino.left.y },
          portasDestino.left
        ],
        label
      };
    }

    return {
      pontos: [
        portasOrigem.left,
        { x: portasDestino.right.x + CONFIG.routeGap, y: portasOrigem.left.y },
        { x: portasDestino.right.x + CONFIG.routeGap, y: portasDestino.right.y },
        portasDestino.right
      ],
      label
    };
  }

  if (dy >= 0) {
    return {
      pontos: [
        portasOrigem.bottom,
        { x: portasOrigem.bottom.x, y: portasDestino.top.y - CONFIG.routeGap },
        { x: portasDestino.top.x, y: portasDestino.top.y - CONFIG.routeGap },
        portasDestino.top
      ],
      label
    };
  }

  return {
    pontos: [
      portasOrigem.top,
      { x: portasOrigem.top.x, y: portasDestino.bottom.y + CONFIG.routeGap },
      { x: portasDestino.bottom.x, y: portasDestino.bottom.y + CONFIG.routeGap },
      portasDestino.bottom
    ],
    label
  };
}

function calcularRotaRetorno(origem, destino, idxRetorno, label = "") {
  const portasOrigem = obterPortasNo(origem);
  const portasDestino = obterPortasNo(destino);

  const deslocamento = CONFIG.laneGap * idxRetorno;
  const xLane = Math.min(origem.x, destino.x) - CONFIG.entryExitGap - deslocamento;

  return {
    pontos: [
      portasOrigem.left,
      { x: xLane, y: portasOrigem.left.y },
      { x: xLane, y: portasDestino.bottom.y + CONFIG.routeGap },
      { x: portasDestino.bottom.x, y: portasDestino.bottom.y + CONFIG.routeGap },
      portasDestino.bottom
    ],
    label
  };
}

function calcularRotaMesmoNivel(origem, destino, label = "") {
  const portasOrigem = obterPortasNo(origem);
  const portasDestino = obterPortasNo(destino);

  if (destino.x >= origem.x) {
    return {
      pontos: [
        portasOrigem.right,
        { x: portasDestino.left.x - CONFIG.routeGap, y: portasOrigem.right.y },
        { x: portasDestino.left.x - CONFIG.routeGap, y: portasDestino.left.y },
        portasDestino.left
      ],
      label
    };
  }

  return {
    pontos: [
      portasOrigem.left,
      { x: portasDestino.right.x + CONFIG.routeGap, y: portasOrigem.left.y },
      { x: portasDestino.right.x + CONFIG.routeGap, y: portasDestino.right.y },
      portasDestino.right
    ],
    label
  };
}

function construirRotasConectores(conexoes, nosPorId, etapaPorId) {
  const rotas = [];
  let contadorRetornos = 0;

  conexoes.forEach((conexao) => {
    const origem = nosPorId[conexao.origemId];
    const destino = nosPorId[conexao.destinoId];

    if (!origem || !destino) return;

    const etapaOrigem = etapaPorId[conexao.origemId];
    const etapaDestino = etapaPorId[conexao.destinoId];

    if (!etapaOrigem || !etapaDestino) return;

    let rota;

    if (etapaDestino.ordem < etapaOrigem.ordem) {
      contadorRetornos += 1;
      rota = calcularRotaRetorno(origem, destino, contadorRetornos, conexao.label);
    } else if (Math.abs(etapaDestino.linha - etapaOrigem.linha) <= CONFIG.sameRowTolerance) {
      rota = calcularRotaMesmoNivel(origem, destino, conexao.label);
    } else {
      rota = calcularRotaDireta(origem, destino, conexao.label);
    }

    rotas.push(rota);
  });

  return rotas;
}

function ordenarEtapas(etapas) {
  return [...etapas].sort((a, b) => {
    const ordemA = Number(a.ordem) || 0;
    const ordemB = Number(b.ordem) || 0;
    if (ordemA !== ordemB) return ordemA - ordemB;

    const linhaA = Number(a.linha) || 0;
    const linhaB = Number(b.linha) || 0;
    if (linhaA !== linhaB) return linhaA - linhaB;

    const colunaA = Number(a.coluna) || 0;
    const colunaB = Number(b.coluna) || 0;
    if (colunaA !== colunaB) return colunaA - colunaB;

    return limpar(a.id).localeCompare(limpar(b.id), "pt-BR", { numeric: true });
  });
}

function montarEtapas(entrada) {
  const linhas = String(entrada || "")
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean);

  const etapas = [];

  linhas.forEach((linha) => {
    const colunas = linha.split("\t");

    if (!colunas.length) return;
    if (ehCabecalho(colunas)) return;

    const ordem = Number(limpar(colunas[0])) || 0;
    const id = limpar(colunas[1]);
    const atividade = limpar(colunas[2]);
    const tipo = limpar(colunas[3]);
    const sistema = limpar(colunas[4]);
    const tempo = tempoParaSegundos(colunas[5]);
    const proxSim = limpar(colunas[6]);
    const proxNao = limpar(colunas[7]);
    const conexoesExtras = limpar(colunas[8]);
    const coluna = Number(limpar(colunas[9])) || 1;
    const linhaPos = Number(limpar(colunas[10])) || 1;

    const categoriaOportunidade = normalizarCategoriaOportunidade(colunas[11]);
    const percentualReducao = parsePercentual(colunas[12]);
    const observacao = limpar(colunas[13]);
    const cor = normalizarCor(colunas[14]);

    if (!id || !atividade) return;

    const ganhoPotencial = tempo * (percentualReducao / 100);
    const tempoToBe = Math.max(0, tempo - ganhoPotencial);

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
      linha: linhaPos,
      categoriaOportunidade,
      percentualReducao,
      observacao,
      cor,
      ganhoPotencial,
      tempoToBe
    });
  });

  return ordenarEtapas(etapas);
}

function construirMapaEtapas(etapas) {
  const etapaPorId = {};
  etapas.forEach((etapa) => {
    etapaPorId[etapa.id] = etapa;
  });
  return etapaPorId;
}

function construirConexoes(etapas) {
  const etapaPorId = construirMapaEtapas(etapas);
  const idsValidos = new Set(etapas.map(e => e.id));
  const conexoes = [];
  let loops = 0;
  const etapasImpactadasRetorno = new Set();

  etapas.forEach((etapa) => {
    if (destinoEhValido(etapa.proxSim, idsValidos)) {
      conexoes.push({
        origemId: etapa.id,
        destinoId: etapa.proxSim,
        label: isPergunta(etapa.atividade) ? "Sim" : ""
      });

      loops += adicionarLoopSeNecessario(etapa.id, etapa.proxSim, etapaPorId, etapa);

      adicionarEtapasImpactadasPorRetorno(
        etapa.id,
        etapa.proxSim,
        etapaPorId,
        etapas,
        etapasImpactadasRetorno
      );
    }

    if (destinoEhValido(etapa.proxNao, idsValidos)) {
      conexoes.push({
        origemId: etapa.id,
        destinoId: etapa.proxNao,
        label: isPergunta(etapa.atividade) ? "Não" : ""
      });

      loops += adicionarLoopSeNecessario(etapa.id, etapa.proxNao, etapaPorId, etapa);

      adicionarEtapasImpactadasPorRetorno(
        etapa.id,
        etapa.proxNao,
        etapaPorId,
        etapas,
        etapasImpactadasRetorno
      );
    }

    quebrarListaIds(etapa.conexoesExtras).forEach((destinoExtra) => {
      if (!destinoEhValido(destinoExtra, idsValidos)) return;

      conexoes.push({
        origemId: etapa.id,
        destinoId: destinoExtra,
        label: ""
      });

      loops += adicionarLoopSeNecessario(etapa.id, destinoExtra, etapaPorId, etapa);

      adicionarEtapasImpactadasPorRetorno(
        etapa.id,
        destinoExtra,
        etapaPorId,
        etapas,
        etapasImpactadasRetorno
      );
    });
  });

  return {
    conexoes,
    loops,
    etapasImpactadasRetorno
  };
}

function calcularDimensoesLayout(etapas, alturaNo) {
  let maxCol = 1;
  let maxLin = 1;

  etapas.forEach((etapa) => {
    if (etapa.coluna > maxCol) maxCol = etapa.coluna;
    if (etapa.linha > maxLin) maxLin = etapa.linha;
  });

  const largura = CONFIG.marginX * 2 + (maxCol - 1) * (CONFIG.boxWidth + CONFIG.colGap) + CONFIG.boxWidth + 300;
  const altura = CONFIG.marginY * 2 + (maxLin - 1) * (alturaNo + CONFIG.rowGap) + alturaNo + 220;

  return { largura, altura, maxCol, maxLin };
}

function construirNos(etapas, alturaNo) {
  const nosPorId = {};

  etapas.forEach((etapa) => {
    const width = obterLarguraNo(etapa);
    const height = obterAlturaNo(etapa, alturaNo);

    const x = CONFIG.marginX + (etapa.coluna - 1) * (CONFIG.boxWidth + CONFIG.colGap);
    const y = CONFIG.marginY + (etapa.linha - 1) * (alturaNo + CONFIG.rowGap);

    nosPorId[etapa.id] = {
      id: etapa.id,
      x,
      y,
      width,
      height,
      etapa
    };
  });

  return nosPorId;
}

function renderizarFluxograma(etapas, conexoes) {
  const diagram = document.getElementById("diagram");
  if (!diagram) return null;

  diagram.innerHTML = "";

  const alturaNoBase = calcularAlturaPadraoNos(etapas);
  const dimensoes = calcularDimensoesLayout(etapas, alturaNoBase);

  const svg = criarElementoSVG("svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", dimensoes.largura);
  svg.setAttribute("height", dimensoes.altura);
  svg.setAttribute("viewBox", `0 0 ${dimensoes.largura} ${dimensoes.altura}`);
  svg.setAttribute("id", "fluxogramaSVG");

  criarMarkerArrow(svg);

  const nosPorId = construirNos(etapas, alturaNoBase);
  const etapaPorId = construirMapaEtapas(etapas);

  const rotas = construirRotasConectores(conexoes, nosPorId, etapaPorId);
  rotas.forEach((rota) => desenharCaminho(svg, rota.pontos, rota.label));

  etapas.forEach((etapa) => {
    const no = nosPorId[etapa.id];
    const fill = corHex(etapa.cor);

    if (isPergunta(etapa.atividade)) {
      desenharLosango(svg, etapa, no.x, no.y, no.width, no.height, fill);
    } else {
      desenharRetangulo(svg, etapa, no.x, no.y, no.width, no.height, fill);
    }
  });

  const primeiraEtapa = etapas[0];
  const ultimaEtapa = etapas[etapas.length - 1];

  if (primeiraEtapa && nosPorId[primeiraEtapa.id]) {
    const no = nosPorId[primeiraEtapa.id];
    desenharCapsula(svg, "Início", no.x + no.width / 2 - 40, no.y - 80, 80, 36);

    desenharCaminho(svg, [
      { x: no.x + no.width / 2, y: no.y - 44 },
      { x: no.x + no.width / 2, y: no.y }
    ]);
  }

  if (ultimaEtapa && nosPorId[ultimaEtapa.id]) {
    const no = nosPorId[ultimaEtapa.id];
    desenharCapsula(svg, "Fim", no.x + no.width / 2 - 30, no.y + no.height + 40, 60, 36);

    desenharCaminho(svg, [
      { x: no.x + no.width / 2, y: no.y + no.height },
      { x: no.x + no.width / 2, y: no.y + no.height + 40 }
    ]);
  }

  diagram.appendChild(svg);
  return svg;
}

function obterTempoTipoResumo(etapas) {
  const acumulado = {};

  etapas.forEach((etapa) => {
    const chave = etapa.tipo || "Não informado";
    acumulado[chave] = (acumulado[chave] || 0) + etapa.tempo;
  });

  return Object.entries(acumulado)
    .map(([tipo, tempo]) => ({ tipo, tempo }))
    .sort((a, b) => b.tempo - a.tempo);
}

function obterTempoCategoriaOportunidade(etapas) {
  const acumulado = {};

  etapas.forEach((etapa) => {
    const chave = etapa.categoriaOportunidade || "Sem oportunidade";
    acumulado[chave] = (acumulado[chave] || 0) + etapa.tempo;
  });

  return Object.entries(acumulado)
    .map(([categoria, tempo]) => ({ categoria, tempo }))
    .sort((a, b) => b.tempo - a.tempo);
}

function calcularSimulacaoMelhoria(etapas) {
  const rows = etapas
    .filter(etapa => etapa.percentualReducao > 0)
    .map((etapa, index) => ({
      ...etapa,
      ordemFmt: formatarOrdem(etapa.ordem, index)
    }));

  const tempoTotalAsIs = rows.reduce((acc, item) => acc + item.tempo, 0);
  const ganhoPotencialTotal = rows.reduce((acc, item) => acc + item.ganhoPotencial, 0);
  const tempoTotalToBe = rows.reduce((acc, item) => acc + item.tempoToBe, 0);
  const eficienciaPotencial = tempoTotalAsIs > 0 ? (ganhoPotencialTotal / tempoTotalAsIs) * 100 : 0;

  const ranking = [...rows]
    .sort((a, b) => b.ganhoPotencial - a.ganhoPotencial)
    .slice(0, 5);

  return {
    rows,
    ranking,
    quantidadeAtividades: rows.length,
    tempoTotalAsIs,
    ganhoPotencialTotal,
    tempoTotalToBe,
    eficienciaPotencial
  };
}

function calcularMetricasProcesso(etapas, loops, etapasImpactadasRetorno) {
  const tempoTotal = etapas.reduce((acc, etapa) => acc + etapa.tempo, 0);
  const decisoes = etapas.filter(etapa => isPergunta(etapa.atividade)).length;
  const taxaDecisao = etapas.length > 0 ? (decisoes / etapas.length) * 100 : 0;

  const tempoPotencialRetrabalho = etapas
    .filter(etapa => etapasImpactadasRetorno.has(etapa.id))
    .reduce((acc, etapa) => acc + etapa.tempo, 0);

  const impactoPotencialRetrabalho = tempoTotal > 0
    ? (tempoPotencialRetrabalho / tempoTotal) * 100
    : 0;

  const simulacao = calcularSimulacaoMelhoria(etapas);

  return {
    tempoTotal,
    loops,
    decisoes,
    taxaDecisao,
    tempoPotencialRetrabalho,
    impactoPotencialRetrabalho,
    simulacao
  };
}

function renderizarInfoProcesso(etapas) {
  const container = document.getElementById("infoProcesso");
  if (!container) return;

  const desenho = obterValorCampo("desenho");
  const processo = obterValorCampo("processo");
  const analista = obterValorCampo("analista");
  const negocio = obterValorCampo("negocio");
  const area = obterValorCampo("area");
  const gestor = obterValorCampo("gestor");

  container.innerHTML = `
    <div class="exec-card">
      <div class="exec-card-title">Informações Gerais</div>
      <div class="exec-info-grid">
        <div class="exec-info-item">
          <div class="exec-info-label">Desenho</div>
          <div class="exec-info-value">${escaparHTML(desenho || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Processo</div>
          <div class="exec-info-value">${escaparHTML(processo || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Analista</div>
          <div class="exec-info-value">${escaparHTML(analista || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Negócio</div>
          <div class="exec-info-value">${escaparHTML(negocio || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Área</div>
          <div class="exec-info-value">${escaparHTML(area || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Gestor</div>
          <div class="exec-info-value">${escaparHTML(gestor || "-")}</div>
        </div>
        <div class="exec-info-item">
          <div class="exec-info-label">Quantidade de etapas</div>
          <div class="exec-info-value">${etapas.length}</div>
        </div>
      </div>
    </div>
  `;
}

function renderizarTabelaTempoPorTipo(etapas) {
  const dados = obterTempoTipoResumo(etapas);

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">Tempo por tipo</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          <thead>
            <tr>
              <th>Tipo</th>
              <th class="th-right">Tempo</th>
            </tr>
          </thead>
          <tbody>
            ${dados.map(item => `
              <tr>
                <td>${escaparHTML(item.tipo)}</td>
                <td class="td-right">⏱ ${formatarTempo(item.tempo)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderizarTabelaOportunidadePorCategoria(etapas) {
  const dados = obterTempoCategoriaOportunidade(etapas);

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">Tempo por categoria de oportunidade</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          <thead>
            <tr>
              <th>Categoria</th>
              <th class="th-right">Tempo As Is</th>
            </tr>
          </thead>
          <tbody>
            ${dados.map(item => `
              <tr>
                <td>${escaparHTML(item.categoria)}</td>
                <td class="td-right">⏱ ${formatarTempo(item.tempo)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderizarTabelaSimulacao(simulacao) {
  if (!simulacao.rows.length) {
    return `
      <div class="exec-table-block">
        <div class="exec-table-title">Simulação de Melhoria (As Is vs To Be)</div>
        <div class="exec-card">Nenhuma atividade com potencial de redução maior que zero foi identificada.</div>
      </div>
    `;
  }

  const maiorGanho = Math.max(...simulacao.rows.map(item => item.ganhoPotencial), 0);

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">Simulação de Melhoria (As Is vs To Be)</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          <thead>
            <tr>
              <th class="th-center">Ordem</th>
              <th>Atividade</th>
              <th class="th-right">Tempo As Is</th>
              <th class="th-center">% Redução</th>
              <th class="th-right">Ganho Potencial</th>
              <th class="th-right">Tempo To Be</th>
              <th>Observação</th>
            </tr>
          </thead>
          <tbody>
            ${simulacao.rows.map(item => {
              const destaque = item.ganhoPotencial === maiorGanho && maiorGanho > 0 ? "row-highlight-opportunity" : "";
              return `
                <tr class="${destaque}">
                  <td class="td-center">${escaparHTML(item.ordemFmt)}</td>
                  <td>${escaparHTML(item.atividade)}</td>
                  <td class="td-right">${formatarTempo(item.tempo)}</td>
                  <td class="td-center">${formatarPercentual(item.percentualReducao)}%</td>
                  <td class="td-right">${formatarTempo(item.ganhoPotencial)}</td>
                  <td class="td-right">${formatarTempo(item.tempoToBe)}</td>
                  <td>${escaparHTML(item.observacao || "-")}</td>
                </tr>
              `;
            }).join("")}
            <tr class="sim-total-row">
              <td class="td-center">TOTAL</td>
              <td>${simulacao.quantidadeAtividades} atividade(s)</td>
              <td class="td-right">${formatarTempo(simulacao.tempoTotalAsIs)}</td>
              <td class="td-center">${formatarPercentual(simulacao.eficienciaPotencial)}%</td>
              <td class="td-right">${formatarTempo(simulacao.ganhoPotencialTotal)}</td>
              <td class="td-right">${formatarTempo(simulacao.tempoTotalToBe)}</td>
              <td>Eficiência total: ${formatarPercentual(simulacao.eficienciaPotencial)}%</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderizarRankingGanhos(simulacao) {
  if (!simulacao.ranking.length) return "";

  return `
    <div class="exec-table-block">
      <div class="exec-table-title">Ranking das maiores oportunidades</div>
      <div class="exec-table-wrap">
        <table class="exec-table">
          <thead>
            <tr>
              <th class="th-center">#</th>
              <th>Atividade</th>
              <th class="th-right">Ganho Potencial</th>
              <th class="th-center">% Redução</th>
              <th>Categoria</th>
            </tr>
          </thead>
          <tbody>
            ${simulacao.ranking.map((item, idx) => `
              <tr>
                <td class="td-center">${idx + 1}</td>
                <td>${escaparHTML(item.atividade)}</td>
                <td class="td-right">${formatarTempo(item.ganhoPotencial)}</td>
                <td class="td-center">${formatarPercentual(item.percentualReducao)}%</td>
                <td>${escaparHTML(item.categoriaOportunidade || "-")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderizarMetricas(etapas, metricas) {
  const container = document.getElementById("metricas");
  if (!container) return;

  container.innerHTML = `
    <div class="exec-summary-grid">
      <div class="exec-summary-item">
        <div class="exec-summary-label">Tempo total do processo</div>
        <div class="exec-summary-value">⏱ ${formatarTempo(metricas.tempoTotal)}</div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Loops detectados</div>
        <div class="exec-summary-value">${metricas.loops}</div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Potencial retrabalho</div>
        <div class="exec-summary-value">
          ⏱ ${formatarTempo(metricas.tempoPotencialRetrabalho)} | % ${formatarPercentual(metricas.impactoPotencialRetrabalho)}
        </div>
      </div>
      <div class="exec-summary-item">
        <div class="exec-summary-label">Taxa de decisão</div>
        <div class="exec-summary-value">${metricas.decisoes} etapa(s) | ${formatarPercentual(metricas.taxaDecisao)}%</div>
      </div>
    </div>

    <div class="exec-summary-grid exec-summary-grid-improvement">
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Tempo total As Is</div>
        <div class="exec-summary-value">⏱ ${formatarTempo(metricas.simulacao.tempoTotalAsIs)}</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement highlight-gain-card">
        <div class="exec-summary-label">Ganho potencial total</div>
        <div class="exec-summary-value">⏱ ${formatarTempo(metricas.simulacao.ganhoPotencialTotal)}</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Tempo total To Be</div>
        <div class="exec-summary-value">⏱ ${formatarTempo(metricas.simulacao.tempoTotalToBe)}</div>
      </div>
      <div class="exec-summary-item exec-summary-item-improvement">
        <div class="exec-summary-label">Eficiência potencial</div>
        <div class="exec-summary-value">${formatarPercentual(metricas.simulacao.eficienciaPotencial)}%</div>
      </div>
    </div>

    ${renderizarTabelaTempoPorTipo(etapas)}
    ${renderizarTabelaOportunidadePorCategoria(etapas)}
    ${renderizarTabelaSimulacao(metricas.simulacao)}
    ${renderizarRankingGanhos(metricas.simulacao)}
  `;
}

function gerarFluxo() {
  const entrada = document.getElementById("entrada")?.value || "";
  const etapas = montarEtapas(entrada);

  if (!etapas.length) {
    alert("Nenhuma etapa válida encontrada. Verifique a tabela colada do Excel.");
    return;
  }

  ultimoNomeArquivo = gerarNomeArquivo();

  const { conexoes, loops, etapasImpactadasRetorno } = construirConexoes(etapas);

  renderizarFluxograma(etapas, conexoes);
  renderizarInfoProcesso(etapas);

  const metricas = calcularMetricasProcesso(etapas, loops, etapasImpactadasRetorno);
  renderizarMetricas(etapas, metricas);
}

function obterSVGPronto() {
  return document.querySelector("#diagram svg");
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

function coletarDadosAnaliseEstruturados() {
  const entrada = document.getElementById("entrada")?.value || "";
  const etapas = montarEtapas(entrada);
  const { loops, etapasImpactadasRetorno } = construirConexoes(etapas);
  return calcularMetricasProcesso(etapas, loops, etapasImpactadasRetorno);
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
    doc.setFont(row.isTotal ? "helvetica" : "helvetica", row.isTotal ? "bold" : "normal");
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

    doc.setFont(row.isTotal ? "helvetica" : "helvetica", row.isTotal ? "bold" : "normal");
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
  y = garantirEspacoPagina(doc, y, 82, margem, pageHeight);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("Análise do Processo", margem, y);
  y += 18;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);

  const linhasResumo = [
    `Tempo total do processo: ${formatarTempo(dados.tempoTotal)}`,
    `Loops detectados: ${dados.loops}`,
    `Potencial retrabalho: ${formatarTempo(dados.tempoPotencialRetrabalho)} | ${formatarPercentual(dados.impactoPotencialRetrabalho)}%`,
    `Taxa de decisão: ${dados.decisoes} etapa(s) | ${formatarPercentual(dados.taxaDecisao)}%`,
    `Tempo total As Is (simulação): ${formatarTempo(dados.simulacao.tempoTotalAsIs)}`,
    `Ganho potencial total: ${formatarTempo(dados.simulacao.ganhoPotencialTotal)}`,
    `Tempo total To Be: ${formatarTempo(dados.simulacao.tempoTotalToBe)}`,
    `Eficiência potencial: ${formatarPercentual(dados.simulacao.eficienciaPotencial)}%`
  ];

  linhasResumo.forEach((linha) => {
    y = garantirEspacoPagina(doc, y, 18, margem, pageHeight);
    doc.text(linha, margem, y);
    y += 18;
  });

  y = desenharTabelaPDF(doc, {
    titulo: "Simulação de Melhoria (As Is vs To Be)",
    columns: [
      { header: "Ordem", key: "ordemFmt", weight: 0.9, align: "center" },
      { header: "Atividade", key: "atividade", weight: 3.1, align: "left" },
      { header: "As Is", key: "tempoAsIsFmt", weight: 1.25, align: "center" },
      { header: "% Red.", key: "reducaoFmt", weight: 1.05, align: "center" },
      { header: "Ganho", key: "ganhoFmt", weight: 1.25, align: "center" },
      { header: "To Be", key: "tempoToBeFmt", weight: 1.25, align: "center" },
      { header: "Observação", key: "observacao", weight: 2.2, align: "left" }
    ],
    rows: [
      ...dados.simulacao.rows.map(item => ({
        ordemFmt: item.ordemFmt,
        atividade: item.atividade,
        tempoAsIsFmt: formatarTempo(item.tempo),
        reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
        ganhoFmt: formatarTempo(item.ganhoPotencial),
        tempoToBeFmt: formatarTempo(item.tempoToBe),
        observacao: item.observacao || "-"
      })),
      {
        ordemFmt: "TOTAL",
        atividade: `${dados.simulacao.quantidadeAtividades} atividade(s)`,
        tempoAsIsFmt: formatarTempo(dados.simulacao.tempoTotalAsIs),
        reducaoFmt: `${formatarPercentual(dados.simulacao.eficienciaPotencial)}%`,
        ganhoFmt: formatarTempo(dados.simulacao.ganhoPotencialTotal),
        tempoToBeFmt: formatarTempo(dados.simulacao.tempoTotalToBe),
        observacao: `Eficiência total: ${formatarPercentual(dados.simulacao.eficienciaPotencial)}%`,
        isTotal: true
      }
    ],
    x: margem,
    yInicial: y,
    larguraTotal: larguraUtil,
    margem,
    pageHeight
  });

  if (dados.simulacao.ranking.length) {
    y = desenharTabelaPDF(doc, {
      titulo: "Ranking das maiores oportunidades",
      columns: [
        { header: "#", key: "posicao", weight: 0.6, align: "center" },
        { header: "Atividade", key: "atividade", weight: 3.4, align: "left" },
        { header: "Ganho", key: "ganhoFmt", weight: 1.25, align: "center" },
        { header: "% Red.", key: "reducaoFmt", weight: 1.0, align: "center" },
        { header: "Categoria", key: "categoria", weight: 1.8, align: "left" }
      ],
      rows: dados.simulacao.ranking.map((item, idx) => ({
        posicao: String(idx + 1),
        atividade: item.atividade,
        ganhoFmt: formatarTempo(item.ganhoPotencial),
        reducaoFmt: `${formatarPercentual(item.percentualReducao)}%`,
        categoria: item.categoriaOportunidade || "-"
      })),
      x: margem,
      yInicial: y,
      larguraTotal: larguraUtil,
      margem,
      pageHeight
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