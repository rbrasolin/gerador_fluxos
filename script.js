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

  // 1ª passagem: montar estrutura correta das etapas
  linhas.forEach((col) => {
    // Precisamos de pelo menos 8 colunas do template
    // Ordem | ID | Atividade | Tipo | Sistema | Tempo | Prox_Sim | Prox_Nao | Cor
    while (col.length < 9) col.push("");

    const ordem = Number(limpar(col[0])) || 0;
    const id = limpar(col[1]);
    const atividade = limpar(col[2]);
    const tipo = limpar(col[3]).toLowerCase();
    const sistema = limpar(col[4]);
    const tempo = Number(limpar(col[5])) || 0;
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
  let manual = 0;
  let sistemaAuto = 0;
  const atividadesTempo = [];

  // 2ª passagem: criar nós e métricas
  etapas.forEach((etapa) => {
    const id = etapa.id;
    const atividade = etapa.atividade;
    const tipo = etapa.tipo;
    const sistema = etapa.sistema;
    const tempo = etapa.tempo;
    const cor = etapa.cor;

    tempoTotal += tempo;
    atividadesTempo.push({ atividade, tempo });

    if (tipo === "manual") {
      manual++;
    } else {
      sistemaAuto++;
    }

    if (etapa.proxNao && idsValidos.has(etapa.proxNao)) {
      // loop simples: volta para uma ordem anterior
      const destino = etapas.find(e => e.id === etapa.proxNao);
      if (destino && destino.ordem < etapa.ordem) {
        loops++;
      }
    }

    const label = atividade + "\\n" + sistema + "\\n" + tempo + " min";

    if (atividade.includes("?")) {
      nodes.push(id + "{" + label + "}");
    } else {
      nodes.push(id + '["' + label + '"]');
    }

    classLines.push("class " + id + " " + cor);
  });

  // 3ª passagem: criar links apenas para IDs que existem
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

  const totalAtividades = manual + sistemaAuto;
  const pctManual = totalAtividades ? ((manual / totalAtividades) * 100).toFixed(1) : "0.0";
  const pctSistema = totalAtividades ? ((sistemaAuto / totalAtividades) * 100).toFixed(1) : "0.0";
  const indiceRetrabalho = totalAtividades ? ((loops / totalAtividades) * 100).toFixed(1) : "0.0";

  atividadesTempo.sort((a, b) => b.tempo - a.tempo);

  const top3 = atividadesTempo
    .slice(0, 3)
    .map(a => a.atividade + " (" + a.tempo + " min)")
    .join("<br>");

  let paretoHTML = "<b>Pareto de tempo</b><br><br>";
  atividadesTempo.forEach(a => {
    const pct = tempoTotal ? ((a.tempo / tempoTotal) * 100).toFixed(1) : "0.0";
    paretoHTML += a.atividade + " — " + pct + "%<br>";
  });

  document.getElementById("metricas").innerHTML =
    "<b>Tempo total do processo:</b> " + tempoTotal + " min<br><br>" +
    "<b>Top 3 gargalos:</b><br>" + top3 + "<br><br>" +
    "<b>Loops de retrabalho:</b> " + loops + "<br>" +
    "<b>Índice de retrabalho:</b> " + indiceRetrabalho + "%<br><br>" +
    "<b>Manual:</b> " + pctManual + "%<br>" +
    "<b>Sistema:</b> " + pctSistema + "%<br><br>" +
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