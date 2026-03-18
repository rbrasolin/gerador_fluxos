let ultimoNomeArquivo = "fluxograma_processo";

const CONFIG = {
  boxWidth: 250,
  boxHeight: 110,
  colGap: 140,
  rowGap: 70,
  marginX: 80,
  marginY: 60,
  fontFamily: "Arial, sans-serif",
  fontSize: 16,
  smallFontSize: 14,
  stroke: "#111111"
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

function obterPontoSaida(origem, destino) {
  const ox = origem.x;
  const oy = origem.y;
  const dx = destino.x;
  const dy = destino.y;

  const ow = CONFIG.boxWidth;
  const oh = CONFIG.boxHeight;

  const saidaDireita = { x: ox + ow, y: oy + oh / 2 };
  const saidaEsquerda = { x: ox, y: oy + oh / 2 };
  const saidaTopo = { x: ox + ow / 2, y: oy };
  const saidaBase = { x: ox + ow / 2, y: oy + oh };

  const entradaDireita = { x: dx + ow, y: dy + oh / 2 };
  const entradaEsquerda = { x: dx, y: dy + oh / 2 };
  const entradaTopo = { x: dx + ow / 2, y: dy };
  const entradaBase = { x: dx + ow / 2, y: dy + oh };

  if (dx > ox + ow) {
    return { start: saidaDireita, end: entradaEsquerda };
  }

  if (dx + ow < ox) {
    return { start: saidaEsquerda, end: entradaDireita };
  }

  if (dy > oy) {
    return { start: saidaBase, end: entradaTopo };
  }

  return { start: saidaTopo, end: entradaBase };
}

function criarPathConector(origem, destino) {
  const { start, end } = obterPontoSaida(origem, destino);
  const meioX = start.x + (end.x - start.x) / 2;

  return `M ${start.x} ${start.y}
          C ${meioX} ${start.y},
            ${meioX} ${end.y},
            ${end.x} ${end.y}`;
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

  const maxColuna = Math.max(...etapas.map(e => e.coluna), 1) + 1;
  const maxLinha = Math.max(...etapas.map(e => e.linha), 1) + 1;

  const svgWidth =
    CONFIG.marginX * 2 +
    maxColuna * CONFIG.boxWidth +
    (maxColuna - 1) * CONFIG.colGap;

  const svgHeight =
    CONFIG.marginY * 2 +
    maxLinha * CONFIG.boxHeight +
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
    const x = CONFIG.marginX + (e.coluna - 1) * (CONFIG.boxWidth + CONFIG.colGap);
    const y = CONFIG.marginY + (e.linha - 1) * (CONFIG.boxHeight + CONFIG.rowGap);

    posicoes[e.id] = { x, y };
  });

  // Início e fim
  const inicio = { x: 20, y: CONFIG.marginY + ((etapas[0].linha - 1) * (CONFIG.boxHeight + CONFIG.rowGap)) + 20 };
  const fim = { x: svgWidth - 100, y: CONFIG.marginY + ((etapas[etapas.length - 1].linha - 1) * (CONFIG.boxHeight + CONFIG.rowGap)) + 20 };

  posicoes["__INICIO__"] = { x: inicio.x, y: inicio.y, w: 60, h: 36 };
  posicoes["__FIM__"] = { x: fim.x, y: fim.y, w: 60, h: 36 };

  function desenharCapsula(idInterno, texto, x, y) {
    const g = criarElementoSVG("g");

    const rect = criarElementoSVG("rect");
    rect.setAttribute("x", x);
    rect.setAttribute("y", y);
    rect.setAttribute("width", 60);
    rect.setAttribute("height", 36);
    rect.setAttribute("rx", 18);
    rect.setAttribute("ry", 18);
    rect.setAttribute("fill", "#ffffff");
    rect.setAttribute("stroke", CONFIG.stroke);
    rect.setAttribute("stroke-width", "2");
    g.appendChild(rect);

    const t = criarElementoSVG("text");
    t.setAttribute("x", x + 30);
    t.setAttribute("y", y + 23);
    t.setAttribute("text-anchor", "middle");
    t.setAttribute("font-family", CONFIG.fontFamily);
    t.setAttribute("font-size", "14");
    t.setAttribute("fill", "#111111");
    t.textContent = texto;
    g.appendChild(t);

    svg.appendChild(g);
  }

  desenharCapsula("__INICIO__", "Início", inicio.x, inicio.y);
  desenharCapsula("__FIM__", "Fim", fim.x, fim.y);

  // Conectores
  function desenharConexao(origemId, destinoId, rotulo = "") {
    if (!posicoes[origemId] || !posicoes[destinoId]) return;

    const origem = origemId === "__INICIO__" || origemId === "__FIM__"
      ? {
          x: posicoes[origemId].x,
          y: posicoes[origemId].y,
          w: 60,
          h: 36
        }
      : {
          x: posicoes[origemId].x,
          y: posicoes[origemId].y,
          w: CONFIG.boxWidth,
          h: CONFIG.boxHeight
        };

    const destino = destinoId === "__INICIO__" || destinoId === "__FIM__"
      ? {
          x: posicoes[destinoId].x,
          y: posicoes[destinoId].y,
          w: 60,
          h: 36
        }
      : {
          x: posicoes[destinoId].x,
          y: posicoes[destinoId].y,
          w: CONFIG.boxWidth,
          h: CONFIG.boxHeight
        };

    const adaptOrigem = {
      x: origem.x,
      y: origem.y
    };

    const adaptDestino = {
      x: destino.x,
      y: destino.y
    };

    const larguraOrigem = origem.w || CONFIG.boxWidth;
    const alturaOrigem = origem.h || CONFIG.boxHeight;
    const larguraDestino = destino.w || CONFIG.boxWidth;
    const alturaDestino = destino.h || CONFIG.boxHeight;

    const saidaDireita = { x: adaptOrigem.x + larguraOrigem, y: adaptOrigem.y + alturaOrigem / 2 };
    const saidaEsquerda = { x: adaptOrigem.x, y: adaptOrigem.y + alturaOrigem / 2 };
    const saidaTopo = { x: adaptOrigem.x + larguraOrigem / 2, y: adaptOrigem.y };
    const saidaBase = { x: adaptOrigem.x + larguraOrigem / 2, y: adaptOrigem.y + alturaOrigem };

    const entradaDireita = { x: adaptDestino.x + larguraDestino, y: adaptDestino.y + alturaDestino / 2 };
    const entradaEsquerda = { x: adaptDestino.x, y: adaptDestino.y + alturaDestino / 2 };
    const entradaTopo = { x: adaptDestino.x + larguraDestino / 2, y: adaptDestino.y };
    const entradaBase = { x: adaptDestino.x + larguraDestino / 2, y: adaptDestino.y + alturaDestino };

    let start;
    let end;

    if (adaptDestino.x > adaptOrigem.x + larguraOrigem) {
      start = saidaDireita;
      end = entradaEsquerda;
    } else if (adaptDestino.x + larguraDestino < adaptOrigem.x) {
      start = saidaEsquerda;
      end = entradaDireita;
    } else if (adaptDestino.y > adaptOrigem.y) {
      start = saidaBase;
      end = entradaTopo;
    } else {
      start = saidaTopo;
      end = entradaBase;
    }

    const meioX = start.x + (end.x - start.x) / 2;

    const path = criarElementoSVG("path");
    path.setAttribute(
      "d",
      `M ${start.x} ${start.y} C ${meioX} ${start.y}, ${meioX} ${end.y}, ${end.x} ${end.y}`
    );
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", "#111111");
    path.setAttribute("stroke-width", "2.2");
    path.setAttribute("marker-end", "url(#arrow)");
    svg.appendChild(path);

    if (rotulo) {
      const tx = criarElementoSVG("text");
      tx.setAttribute("x", meioX);
      tx.setAttribute("y", start.y - 8);
      tx.setAttribute("text-anchor", "middle");
      tx.setAttribute("font-family", CONFIG.fontFamily);
      tx.setAttribute("font-size", "12");
      tx.setAttribute("fill", "#111111");
      tx.textContent = rotulo;
      svg.appendChild(tx);
    }
  }

  // início -> primeira etapa
  desenharConexao("__INICIO__", etapas[0].id);

  // nós
  etapas.forEach((e) => {
    const x = posicoes[e.id].x;
    const y = posicoes[e.id].y;

    const g = criarElementoSVG("g");

    const rect = criarElementoSVG("rect");
    rect.setAttribute("x", x);
    rect.setAttribute("y", y);
    rect.setAttribute("width", CONFIG.boxWidth);
    rect.setAttribute("height", CONFIG.boxHeight);
    rect.setAttribute("fill", corHex(e.cor));
    rect.setAttribute("stroke", CONFIG.stroke);
    rect.setAttribute("stroke-width", "2");
    g.appendChild(rect);

    const linhas = [
      ...quebrarTextoAutomatico(e.atividade, 3, 18),
      ...quebrarTextoAutomatico(e.sistema || "Sem sistema informado", 3, 20),
      formatarTempo(e.tempo)
    ];

    const totalAlturaTexto = linhas.length * 18;
    const inicioYTexto = y + (CONFIG.boxHeight - totalAlturaTexto) / 2 + 14;

    linhas.forEach((linha, i) => {
      const text = criarElementoSVG("text");
      text.setAttribute("x", x + CONFIG.boxWidth / 2);
      text.setAttribute("y", inicioYTexto + i * 18);
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-family", CONFIG.fontFamily);
      text.setAttribute("font-size", i === linhas.length - 1 ? CONFIG.smallFontSize : CONFIG.fontSize);
      text.setAttribute("fill", "#111111");
      text.textContent = linha;
      g.appendChild(text);
    });

    svg.appendChild(g);
  });

  // setas entre etapas
  let loops = 0;
  let decisoes = 0;
  let conexoesExtrasCount = 0;
  const etapasOrigemComRetorno = new Set();

  etapas.forEach((etapa) => {
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    if (etapa.atividade.includes("?")) {
      decisoes++;
    }

    destinosSim.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        etapasOrigemComRetorno.add(etapa.id);
      }

      desenharConexao(etapa.id, destinoId, etapa.atividade.includes("?") ? (indice === 0 ? "Sim" : `Sim ${indice + 1}`) : "");
    });

    destinosNao.forEach((destinoId, indice) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        etapasOrigemComRetorno.add(etapa.id);
      }

      desenharConexao(etapa.id, destinoId, indice === 0 ? "Não" : `Não ${indice + 1}`);
    });

    destinosExtras.forEach((destinoId) => {
      const destino = etapaPorId[destinoId];
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
        etapasOrigemComRetorno.add(etapa.id);
      }

      conexoesExtrasCount++;
      desenharConexao(etapa.id, destinoId);
    });
  });

  // finais -> fim
  etapas.forEach((etapa) => {
    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosExtras = quebrarListaIds(etapa.conexoesExtras).filter(destino => destinoEhValido(destino, idsValidos));

    if (destinosSim.length === 0 && destinosNao.length === 0 && destinosExtras.length === 0) {
      desenharConexao(etapa.id, "__FIM__");
    }
  });

  document.getElementById("diagram").innerHTML = "";
  document.getElementById("diagram").appendChild(svg);

  ultimoNomeArquivo = gerarNomeArquivo();

  document.getElementById("infoProcesso").innerHTML =
    "<b>Processo:</b> " + (processo || "Não informado") + "<br>" +
    "<b>Analista:</b> " + (analista || "Não informado");

  // métricas
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

  const clone = svgOriginal.cloneNode(true);
  return clone;
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

// compatibilidade com botão antigo
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