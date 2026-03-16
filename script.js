mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: "basis",
    nodeSpacing: 60,
    rankSpacing: 100
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

function ehCabecalho(colunas) {
  if (!colunas || colunas.length === 0) return false;
  return limpar(colunas[0]).toLowerCase() === "ordem";
}

function tempoParaSegundos(tempo) {
  if (!tempo) return 0;

  tempo = String(tempo).trim();

  if (!tempo.includes(":")) {
    return (Number(tempo) || 0) * 3600;
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

function gerarFluxo() {
  const texto = document.getElementById("entrada").value;

  if (!texto.trim()) {
    alert("Cole a tabela do Excel primeiro.");
    return;
  }

  const processo = document.getElementById("processo").value;
  const analista = document.getElementById("analista").value;

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
    while (col.length < 9) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]) || "Não informado";
    const sistema = limpar(col[4]);
    const tempo = tempoParaSegundos(limpar(col[5]));
    const proxSim = limpar(col[6]);
    const proxNao = limpar(col[7]);
    const cor = normalizarCor(col[8]);

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

  let mermaidCode = "flowchart LR\n";

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
    const atividade = etapa.atividade;
    const tipo = etapa.tipo;
    const sistema = etapa.sistema;
    const tempo = etapa.tempo;
    const cor = etapa.cor;

    tempoTotal += tempo;
    atividadesTempo.push({ atividade, tempo });

    if (!tiposTempo[tipo]) {
      tiposTempo[tipo] = 0;
    }
    tiposTempo[tipo] += tempo;

    if (atividade.includes("?")) {
      decisoes++;
    }

    const label = atividade + "\\n" + sistema + "\\n" + formatarTempo(tempo);

    if (atividade.includes("?")) {
      nodes.push(id + "{" + label + "}");
    } else {
      nodes.push(id + '["' + label + '"]');
    }

    classLines.push("class " + id + " " + cor);

    const semProxSim = !etapa.proxSim || !idsValidos.has(etapa.proxSim);
    const semProxNao = !etapa.proxNao || !idsValidos.has(etapa.proxNao);

    if (semProxSim && semProxNao) {
      ultimoIds.push(id);
    }
  });

  nodes.push('FIM(["Fim"])');
  classLines.push("class FIM white");

  links.push("INICIO --> " + primeiroId);

  etapas.forEach((etapa) => {
    const id = etapa.id;
    const decisao = etapa.atividade.includes("?");

    if (etapa.proxSim && idsValidos.has(etapa.proxSim)) {
      const destinoSim = etapaPorId[etapa.proxSim];

      if (destinoSim && destinoSim.ordem < etapa.ordem) {
        loops++;
        etapasOrigemComRetorno.add(id);
      }

      if (decisao) {
        links.push(id + " -->|Sim| " + etapa.proxSim);
      } else {
        links.push(id + " --> " + etapa.proxSim);
      }
    }

    if (etapa.proxNao && idsValidos.has(etapa.proxNao)) {
      const destinoNao = etapaPorId[etapa.proxNao];

      if (destinoNao && destinoNao.ordem < etapa.ordem) {
        loops++;
        etapasOrigemComRetorno.add(id);
      }

      links.push(id + " -->|Não| " + etapa.proxNao);
    }
  });

  ultimoIds.forEach(id => {
    links.push(id + " --> FIM");
  });

  nodes.forEach(n => {
    mermaidCode += n + "\n";
  });

  links.forEach(l => {
    mermaidCode += l + "\n";
  });

  classLines.forEach(c => {
    mermaidCode += c + "\n";
  });

  mermaidCode += `
classDef white fill:#ffffff,stroke:#000,stroke-width:1.5px,color:#000;
classDef blue fill:#8ecae6,stroke:#000,stroke-width:1.5px,color:#000;
classDef yellow fill:#ffd166,stroke:#000,stroke-width:1.5px,color:#000;
classDef green fill:#95d5b2,stroke:#000,stroke-width:1.5px,color:#000;
classDef red fill:#ef476f,stroke:#000,stroke-width:1.5px,color:#000;
`;

  try {
    document.getElementById("diagram").innerHTML =
      '<div class="mermaid">' + mermaidCode + "</div>";

    mermaid.init(undefined, document.querySelectorAll(".mermaid"));
  } catch (e) {
    console.error("Erro Mermaid:", e);
    console.error("Código Mermaid gerado:\n", mermaidCode);
    alert("Erro ao gerar fluxo. Veja o console (F12).");
    return;
  }

  document.getElementById("infoProcesso").innerHTML =
    "<b>Processo:</b> " + processo + "<br>" +
    "<b>Analista:</b> " + analista;

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const top3HTML = atividadesTempo
    .slice(0, 3)
    .map(a => {
      const pct = tempoTotal
        ? ((a.tempo / tempoTotal) * 100).toFixed(1).replace(".", ",")
        : "0,0";

      return '<div class="analytics-item">' +
        a.atividade +
        ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
        ' <span class="icon-pct">٪</span>' + pct + '%' +
        '</div>';
    })
    .join("");

  let paretoHTML = "";
  atividadesTempo.forEach(a => {
    const pct = tempoTotal
      ? ((a.tempo / tempoTotal) * 100).toFixed(1).replace(".", ",")
      : "0,0";

    paretoHTML += '<div class="analytics-item">' +
      a.atividade +
      ' — <span class="icon-time">⏱</span>' + formatarTempo(a.tempo) +
      ' <span class="icon-pct">٪</span>' + pct + '%' +
      '</div>';
  });

  const tiposOrdenados = Object.entries(tiposTempo).sort((a, b) => b[1] - a[1]);

  let tiposHTML = "";
  tiposOrdenados.forEach(([tipo, tempo]) => {
    const pct = tempoTotal
      ? ((tempo / tempoTotal) * 100).toFixed(1).replace(".", ",")
      : "0,0";

    tiposHTML += '<div class="analytics-item">' +
      tipo +
      ' — <span class="icon-time">⏱</span>' + formatarTempo(tempo) +
      ' <span class="icon-pct">٪</span>' + pct + '%' +
      '</div>';
  });

  let tempoPotencialRetrabalho = 0;
  etapas.forEach((etapa) => {
    if (etapasOrigemComRetorno.has(etapa.id)) {
      tempoPotencialRetrabalho += etapa.tempo;
    }
  });

  const impactoPotencialRetrabalho = tempoTotal
    ? ((tempoPotencialRetrabalho / tempoTotal) * 100).toFixed(1).replace(".", ",")
    : "0,0";

  const taxaDecisao = etapas.length
    ? ((decisoes / etapas.length) * 100).toFixed(1).replace(".", ",")
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
          '<div class="analytics-item"><b>Impacto potencial de retrabalho:</b> <span class="icon-time">⏱</span>' + formatarTempo(tempoPotencialRetrabalho) + ' <span class="icon-pct">٪</span>' + impactoPotencialRetrabalho + '%</div>' +
          '<div class="analytics-item"><b>Taxa de decisão:</b> ' + decisoes + ' etapa(s) <span class="icon-pct">٪</span>' + taxaDecisao + '%</div>' +
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

function baixarPNG() {
  const svgOriginal = document.querySelector("#diagram svg");

  if (!svgOriginal) {
    alert("Gere o fluxo primeiro.");
    return;
  }

  const svg = svgOriginal.cloneNode(true);
  const bbox = svgOriginal.getBBox();
  const padding = 40;

  const larguraReal = Math.ceil(bbox.width + padding * 2);
  const alturaReal = Math.ceil(bbox.height + padding * 2);

  const larguraExportacao = 12000; // pode reduzir para 5000 - 7000 ou 9000 se quiser arquivo menor
  const escala = larguraExportacao / larguraReal;
  const alturaExportacao = Math.ceil(alturaReal * escala);

  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", larguraReal);
  svg.setAttribute("height", alturaReal);
  svg.setAttribute(
    "viewBox",
    (bbox.x - padding) + " " + (bbox.y - padding) + " " + larguraReal + " " + alturaReal
  );

  const fundo = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  fundo.setAttribute("x", bbox.x - padding);
  fundo.setAttribute("y", bbox.y - padding);
  fundo.setAttribute("width", larguraReal);
  fundo.setAttribute("height", alturaReal);
  fundo.setAttribute("fill", "white");
  svg.insertBefore(fundo, svg.firstChild);

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
    link.download = "fluxograma.png";
    link.href = canvas.toDataURL("image/png");
    link.click();
  };

  img.src = image64;
}