let ultimoMermaidCode = "";
let ultimoNomeArquivo = "fluxograma_processo";

mermaid.initialize({
  startOnLoad: false,
  securityLevel: "loose",
  theme: "base",
  themeVariables: {
    fontFamily: "Arial, sans-serif",
    fontSize: "18px",
    primaryTextColor: "#111111",
    lineColor: "#222222",
    textColor: "#111111",
    nodeBorder: "#111111",
    clusterBorder: "#111111",
    edgeLabelBackground: "#ffffff",
    mainBkg: "#ffffff"
  },
  flowchart: {
    useMaxWidth: false,
    htmlLabels: false,
    curve: "basis",
    nodeSpacing: 80,
    rankSpacing: 130,
    padding: 20
  }
});

function limpar(txt) {
  if (txt === null || txt === undefined) return "";
  return String(txt)
    .replace(/"/g, "")
    .replace(/\(/g, "")
    .replace(/\)/g, "")
    .trim();
}

function normalizarCor(cor) {
  const c = limpar(cor).toLowerCase();
  const permitidas = ["blue", "yellow", "green", "red", "white"];
  return permitidas.includes(c) ? c : "white";
}

function normalizarTipoConexao(tipoConexao) {
  const valor = limpar(tipoConexao).toUpperCase();
  const permitidos = ["NORMAL", "PARALELO", "JOIN"];
  return permitidos.includes(valor) ? valor : "NORMAL";
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

  return (h * 3600) + (m * 60) + s;
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

function escapeMermaidText(texto) {
  return String(texto || "")
    .replace(/&/g, "&amp;")
    .replace(/#/g, " ")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\|/g, "/")
    .replace(/"/g, "'")
    .trim();
}

function quebrarTextoAutomatico(texto, maxPalavrasPorLinha = 3, maxCaracteresPorLinha = 18) {
  const palavras = String(texto || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (!palavras.length) return "";

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

  return linhas.join("<br/>");
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
  ultimoMermaidCode = "";
  ultimoNomeArquivo = "fluxograma_processo";
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

async function gerarFluxo() {
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
    while (col.length < 10) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]) || "Não informado";
    const sistema = limpar(col[4]);
    const tempo = tempoParaSegundos(limpar(col[5]));
    const proxSim = limpar(col[6]);
    const proxNao = limpar(col[7]);
    const cor = normalizarCor(col[8]);
    const tipoConexao = normalizarTipoConexao(col[9]);

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
      cor,
      tipoConexao
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

  let mermaidCode = "flowchart LR\n";
  mermaidCode += "%%{init: {'flowchart': {'curve': 'basis', 'htmlLabels': false}}}%%\n";

  const nodes = [];
  const links = [];
  const classLines = [];

  let tempoTotal = 0;
  const atividadesTempo = [];
  const tiposTempo = {};

  let loops = 0;
  let decisoes = 0;
  const etapasOrigemComRetorno = new Set();

  const primeiroId = etapas[0].id;
  const ultimoIds = [];

  nodes.push('INICIO(["Início"])');
  classLines.push("class INICIO white");

  etapas.forEach((etapa) => {
    const id = etapa.id;
    const atividadeOriginal = escapeMermaidText(etapa.atividade);
    const atividade = quebrarTextoAutomatico(atividadeOriginal, 3, 18);
    const tipo = etapa.tipo;

    const sistemaOriginal = escapeMermaidText(etapa.sistema || "Sem sistema informado");
    const sistema = sistemaOriginal.length > 22
      ? quebrarTextoAutomatico(sistemaOriginal, 3, 20)
      : sistemaOriginal;

    const tempo = etapa.tempo;
    const cor = etapa.cor;

    tempoTotal += tempo;
    atividadesTempo.push({ atividade: etapa.atividade, tempo });

    if (!tiposTempo[tipo]) {
      tiposTempo[tipo] = 0;
    }
    tiposTempo[tipo] += tempo;

    if (etapa.atividade.includes("?")) {
      decisoes++;
    }

    const label = `${atividade}<br/>${sistema}<br/>${formatarTempo(tempo)}`;

    if (etapa.atividade.includes("?")) {
      nodes.push(`${id}{"${label}"}`);
    } else {
      nodes.push(`${id}["${label}"]`);
    }

    classLines.push(`class ${id} ${cor}`);

    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));

    const semProxSim = destinosSim.length === 0;
    const semProxNao = destinosNao.length === 0;

    if (semProxSim && semProxNao) {
      ultimoIds.push(id);
    }
  });

  nodes.push('FIM(["Fim"])');
  classLines.push("class FIM white");

  links.push(`INICIO --> ${primeiroId}`);

  etapas.forEach((etapa) => {
    const id = etapa.id;
    const decisao = etapa.atividade.includes("?");
    const tipoConexao = etapa.tipoConexao || "NORMAL";

    const destinosSim = quebrarListaIds(etapa.proxSim).filter(destino => destinoEhValido(destino, idsValidos));
    const destinosNao = quebrarListaIds(etapa.proxNao).filter(destino => destinoEhValido(destino, idsValidos));

    if (tipoConexao === "PARALELO") {
      destinosSim.forEach((destinoId) => {
        const destino = etapaPorId[destinoId];

        if (destino && destino.ordem < etapa.ordem) {
          loops++;
          etapasOrigemComRetorno.add(id);
        }

        links.push(`${id} --> ${destinoId}`);
      });

      destinosNao.forEach((destinoId) => {
        const destino = etapaPorId[destinoId];

        if (destino && destino.ordem < etapa.ordem) {
          loops++;
          etapasOrigemComRetorno.add(id);
        }

        links.push(`${id} -->|Não| ${destinoId}`);
      });

      return;
    }

    if (destinosSim.length > 0) {
      destinosSim.forEach((destinoId, indice) => {
        const destinoSim = etapaPorId[destinoId];

        if (destinoSim && destinoSim.ordem < etapa.ordem) {
          loops++;
          etapasOrigemComRetorno.add(id);
        }

        if (decisao) {
          const rotulo = indice === 0 ? "Sim" : `Sim ${indice + 1}`;
          links.push(`${id} -->|${rotulo}| ${destinoId}`);
        } else {
          links.push(`${id} --> ${destinoId}`);
        }
      });
    }

    if (destinosNao.length > 0) {
      destinosNao.forEach((destinoId, indice) => {
        const destinoNao = etapaPorId[destinoId];

        if (destinoNao && destinoNao.ordem < etapa.ordem) {
          loops++;
          etapasOrigemComRetorno.add(id);
        }

        const rotulo = indice === 0 ? "Não" : `Não ${indice + 1}`;
        links.push(`${id} -->|${rotulo}| ${destinoId}`);
      });
    }
  });

  ultimoIds.forEach(id => {
    links.push(`${id} --> FIM`);
  });

  nodes.forEach(n => {
    mermaidCode += `${n}\n`;
  });

  links.forEach(l => {
    mermaidCode += `${l}\n`;
  });

  classLines.forEach(c => {
    mermaidCode += `${c}\n`;
  });

  mermaidCode += `
classDef white fill:#ffffff,stroke:#111111,stroke-width:2px,color:#111111;
classDef blue fill:#8ecae6,stroke:#111111,stroke-width:2px,color:#111111;
classDef yellow fill:#ffd166,stroke:#111111,stroke-width:2px,color:#111111;
classDef green fill:#95d5b2,stroke:#111111,stroke-width:2px,color:#111111;
classDef red fill:#ef476f,stroke:#111111,stroke-width:2px,color:#111111;

linkStyle default stroke:#222222,stroke-width:2.2px;
`;

  ultimoMermaidCode = mermaidCode;
  ultimoNomeArquivo = gerarNomeArquivo();

  try {
    const renderId = `mermaid_${Date.now()}`;
    const { svg } = await mermaid.render(renderId, mermaidCode);

    document.getElementById("diagram").innerHTML = svg;

    const svgElement = document.querySelector("#diagram svg");
    if (svgElement) {
      svgElement.style.maxWidth = "none";
      svgElement.style.height = "auto";
      svgElement.setAttribute("preserveAspectRatio", "xMinYMin meet");
    }
  } catch (e) {
    console.error("Erro Mermaid:", e);
    console.error("Código Mermaid gerado:\n", mermaidCode);
    alert("Erro ao gerar fluxo. Veja o console (F12).");
    return;
  }

  document.getElementById("infoProcesso").innerHTML =
    "<b>Processo:</b> " + (processo || "Não informado") + "<br>" +
    "<b>Analista:</b> " + (analista || "Não informado");

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const top3HTML = atividadesTempo
    .slice(0, 3)
    .map(a => {
      const pct = tempoTotal ? formatarPercentual((a.tempo / tempoTotal) * 100) : "0,0";

      return (
        '<div class="analytics-item">' +
          a.atividade +
          ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
          ' <span class="icon-pct">%</span>' + pct + '%' +
        '</div>'
      );
    })
    .join("");

  let paretoHTML = "";
  atividadesTempo.forEach(a => {
    const pct = tempoTotal ? formatarPercentual((a.tempo / tempoTotal) * 100) : "0,0";

    paretoHTML +=
      '<div class="analytics-item">' +
        a.atividade +
        ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
        ' <span class="icon-pct">%</span>' + pct + '%' +
      '</div>';
  });

  const tiposOrdenados = Object.entries(tiposTempo).sort((a, b) => b[1] - a[1]);

  let tiposHTML = "";
  tiposOrdenados.forEach(([tipo, tempo]) => {
    const pct = tempoTotal ? formatarPercentual((tempo / tempoTotal) * 100) : "0,0";

    tiposHTML +=
      '<div class="analytics-item">' +
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

function prepararSVGParaPDF(svgOriginal) {
  const svg = svgOriginal.cloneNode(true);

  const seletores = "svg, g, path, rect, polygon, circle, ellipse, line, polyline, text, tspan";
  const elementos = svg.querySelectorAll(seletores);
  const originalElements = svgOriginal.querySelectorAll(seletores);

  elementos.forEach((el, i) => {
    const original = originalElements[i];
    if (!original) return;

    const cs = window.getComputedStyle(original);

    const props = [
      "fill",
      "fill-opacity",
      "stroke",
      "stroke-width",
      "stroke-opacity",
      "stroke-linecap",
      "stroke-linejoin",
      "stroke-dasharray",
      "opacity",
      "color",
      "font-family",
      "font-size",
      "font-weight",
      "font-style",
      "text-anchor",
      "dominant-baseline"
    ];

    let styleInline = "";
    props.forEach((prop) => {
      const val = cs.getPropertyValue(prop);
      if (val && val !== "normal" && val !== "auto") {
        styleInline += `${prop}:${val};`;
      }
    });

    const tag = el.tagName.toLowerCase();

    if (["text", "tspan"].includes(tag)) {
      const fill = cs.getPropertyValue("fill");
      if (!fill || fill === "none") {
        styleInline += "fill:#111111;";
      }
    }

    if (["path", "rect", "polygon", "circle", "ellipse", "line", "polyline"].includes(tag)) {
      const stroke = cs.getPropertyValue("stroke");
      const fill = cs.getPropertyValue("fill");

      if (!stroke || stroke === "none") {
        styleInline += "stroke:#111111;";
      }

      if (!fill || fill === "none") {
        styleInline += "fill:none;";
      }
    }

    const existing = el.getAttribute("style") || "";
    el.setAttribute("style", `${existing};${styleInline}`);
  });

  svg.querySelectorAll("[filter]").forEach(el => el.removeAttribute("filter"));
  svg.querySelectorAll("[mask]").forEach(el => el.removeAttribute("mask"));

  return svg;
}

function obterSVGPronto() {
  const svgOriginal = document.querySelector("#diagram svg");

  if (!svgOriginal) {
    alert("Gere o fluxo primeiro.");
    return null;
  }

  const svg = prepararSVGParaPDF(svgOriginal);
  const bbox = svgOriginal.getBBox();
  const padding = 40;

  const larguraReal = Math.ceil(bbox.width + padding * 2);
  const alturaReal = Math.ceil(bbox.height + padding * 2);

  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", larguraReal);
  svg.setAttribute("height", alturaReal);
  svg.setAttribute(
    "viewBox",
    `${bbox.x - padding} ${bbox.y - padding} ${larguraReal} ${alturaReal}`
  );

  const fundo = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  fundo.setAttribute("x", bbox.x - padding);
  fundo.setAttribute("y", bbox.y - padding);
  fundo.setAttribute("width", larguraReal);
  fundo.setAttribute("height", alturaReal);
  fundo.setAttribute("fill", "white");
  svg.insertBefore(fundo, svg.firstChild);

  return {
    svg,
    larguraReal,
    alturaReal
  };
}

function baixarSVG() {
  const resultado = obterSVGPronto();
  if (!resultado) return;

  const { svg } = resultado;
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

async function baixarPDF() {
  const resultado = obterSVGPronto();
  if (!resultado) return;

  const { svg, larguraReal, alturaReal } = resultado;

  try {
    if (!window.jspdf || typeof window.jspdf.jsPDF !== "function") {
      throw new Error("jsPDF não carregado corretamente.");
    }

    const { jsPDF } = window.jspdf;
    const orientacao = larguraReal >= alturaReal ? "landscape" : "portrait";

    const pdf = new jsPDF({
      orientation: orientacao,
      unit: "pt",
      format: [larguraReal, alturaReal]
    });

    if (typeof pdf.svg === "function") {
      await pdf.svg(svg, {
        x: 0,
        y: 0,
        width: larguraReal,
        height: alturaReal
      });
    } else if (typeof window.svg2pdf === "function") {
      await window.svg2pdf(svg, pdf, {
        xOffset: 0,
        yOffset: 0,
        scale: 1
      });
    } else {
      throw new Error("svg2pdf.js não expôs integração utilizável.");
    }

    pdf.save(`${ultimoNomeArquivo}.pdf`);
  } catch (erro) {
    console.error("Erro ao gerar PDF vetorial:", erro);
    alert("Não foi possível gerar o PDF vetorial. Veja o console (F12).");
  }
}

function baixarPNG() {
  const resultado = obterSVGPronto();
  if (!resultado) return;

  const { svg, larguraReal, alturaReal } = resultado;

  const larguraExportacao = 9000;
  const escala = larguraExportacao / larguraReal;
  const alturaExportacao = Math.ceil(alturaReal * escala);

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const svg64 = btoa(unescape(encodeURIComponent(source)));
  const image64 = "data:image/svg+xml;base64," + svg64;

  const img = new Image();

  img.onload = function () {
    const canvas = document.createElement("canvas");
    canvas.width = larguraExportacao;
    canvas.height = alturaExportacao;

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