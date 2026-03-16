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
    .replace(/:/g, "")
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

function formatarHoras(valor) {
  return (Number(valor) || 0).toFixed(2).replace(".", ",") + " h";
}

function gerarFluxo() {
  const texto = document.getElementById("entrada").value;
  const processo = document.getElementById("processo").value;
  const analista = document.getElementById("analista").value;

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

  // 1ª passagem: montar etapas
  linhas.forEach((col) => {
    while (col.length < 9) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]);
    const sistema = limpar(col[4]);
    const tempo = Number(String(limpar(col[5])).replace(",", ".")) || 0;
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

  let mermaidCode = "flowchart LR\n";

  const nodes = [];
  const links = [];
  const classLines = [];

  let tempoTotal = 0;
  let loops = 0;
  const atividadesTempo = [];
  const tiposTempo = {};

  const primeiroId = etapas[0].id;
  const ultimoIds = [];

  // nó de início
  nodes.push('INICIO(["Início"])');
  classLines.push("class INICIO white");

  // 2ª passagem: nós e métricas
  etapas.forEach((etapa) => {
    const id = etapa.id;
    const atividade = etapa.atividade;
    const tipo = etapa.tipo || "Não informado";
    const sistema = etapa.sistema;
    const tempo = etapa.tempo;
    const cor = etapa.cor;

    tempoTotal += tempo;
    atividadesTempo.push({ atividade, tempo });

    if (!tiposTempo[tipo]) {
      tiposTempo[tipo] = 0;
    }
    tiposTempo[tipo] += tempo;

    if (etapa.proxNao && idsValidos.has(etapa.proxNao)) {
      const destino = etapas.find(e => e.id === etapa.proxNao);
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
      }
    }

    const label = atividade + "\\n" + sistema + "\\n" + formatarHoras(tempo);

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

  // nó de fim
  nodes.push('FIM(["Fim"])');
  classLines.push("class FIM white");

  // 3ª passagem: links
  links.push("INICIO --> " + primeiroId);

  etapas.forEach((etapa) => {
    const id = etapa.id;
    const decisao = etapa.atividade.includes("?");

    if (etapa.proxSim && idsValidos.has(etapa.proxSim)) {
      if (decisao) {
        links.push(id + " -->|Sim| " + etapa.proxSim);
      } else {
        links.push(id + " --> " + etapa.proxSim);
      }
    }

    if (etapa.proxNao && idsValidos.has(etapa.proxNao)) {
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

  const indiceRetrabalho = etapas.length ? ((loops / etapas.length) * 100).toFixed(1) : "0,0";

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const top3 = atividadesTempo
    .slice(0, 3)
    .map(a => {
      const pct = tempoTotal ? ((a.tempo / tempoTotal) * 100).toFixed(1).replace(".", ",") : "0,0";
      return a.atividade + " (" + formatarHoras(a.tempo) + " | " + pct + "%)";
    })
    .join("<br>");

  let paretoHTML = "<b>Pareto de tempo</b><br><br>";
  atividadesTempo.forEach(a => {
    const pct = tempoTotal ? ((a.tempo / tempoTotal) * 100).toFixed(1).replace(".", ",") : "0,0";
    paretoHTML += a.atividade + " — " + formatarHoras(a.tempo) + " | " + pct + "%<br>";
  });

  const tiposOrdenados = Object.entries(tiposTempo).sort((a, b) => b[1] - a[1]);
  let tiposHTML = "<b>Tempo por tipo</b><br><br>";
  tiposOrdenados.forEach(([tipo, tempo]) => {
    const pct = tempoTotal ? ((tempo / tempoTotal) * 100).toFixed(1).replace(".", ",") : "0,0";
    tiposHTML += tipo + " — " + formatarHoras(tempo) + " | " + pct + "%<br>";
  });

  document.getElementById("metricas").innerHTML =
    "<b>Tempo total do processo:</b> " + formatarHoras(tempoTotal) + "<br><br>" +
    "<b>Top 3 gargalos:</b><br>" + top3 + "<br><br>" +
    "<b>Loops de retrabalho:</b> " + loops + "<br>" +
    "<b>Índice de retrabalho:</b> " + indiceRetrabalho.replace(".", ",") + "%<br><br>" +
    tiposHTML + "<br>" +
    paretoHTML;
}

function baixarPNG() {
  const svg = document.querySelector("#diagram svg");

  if (!svg) {
    alert("Gere o fluxo primeiro.");
    return;
  }

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const svg64 = btoa(unescape(encodeURIComponent(source)));
  const image64 = "data:image/svg+xml;base64," + svg64;

  const img = new Image();

  img.onload = function () {
    const canvas = document.createElement("canvas");
    canvas.width = img.width * 3;
    canvas.height = img.height * 3;

    const ctx = canvas.getContext("2d");
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