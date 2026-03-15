mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'linear',
    nodeSpacing: 250,
    rankSpacing: 40
  }
});

function limparTexto(txt) {
  return txt
    .replace(/"/g, '')
    .replace(/\(/g, '')
    .replace(/\)/g, '')
    .replace(/:/g, '')
    .trim();
}

function gerarFluxo() {
  let texto = document.getElementById("entrada").value;
  let processo = document.getElementById("processo").value;
  let analista = document.getElementById("analista").value;
  let linhas = texto.trim().split("\n");
  
  let mermaidCode = "flowchart LR\n";
  let nodes = {};
  let links = [];
  let allIds = [];
  
  // Coleta todos IDs únicos para cadeia L→R
  linhas.forEach(function(linha) {
    if (!linha.trim()) return;
    let col = linha.trim().split(/\s{2,}|\t/);
    if (col.length < 6) return;
    allIds.push("A" + Number(col[1]));
  });
  allIds = [...new Set(allIds)].sort(function(a,b) {
    return Number(a.slice(1)) - Number(b.slice(1));
  });
  
  let tempoTotal = 0;
  let loops = 0;
  let manual = 0;
  let sistemaAuto = 0;
  let atividadesTempo = [];
  
  linhas.forEach(function(linha) {
    if (!linha.trim()) return;
    let col = linha.trim().split(/\s{2,}|\t/);
    if (col.length < 6) return;
    
    let idNumero = Number(col[1]);
    let id = "A" + idNumero;
    let atividade = limparTexto(col[2]);
    let tipo = (col[3] || "").toLowerCase();
    let sistema = limparTexto(col[4] || "");
    let tempo = Number(col[5]) || 0;
    let proxSim = col[6] ? "A" + col[6] : null;
    let proxNao = col[7] ? "A" + col[7] : null;
    let cor = (col[8] || "").toLowerCase().trim();
    
    tempoTotal += tempo;
    atividadesTempo.push({atividade: atividade, tempo: tempo});
    
    if (tipo === "manual") {
      manual++;
    } else {
      sistemaAuto++;
    }
    if (proxNao && Number(col[7]) < idNumero) {
      loops++;
    }
    
    let label = atividade + "\\n" + sistema + "\\n" + tempo + " min";
    
    if (atividade.includes("?")) {
      nodes[id] = id + "{" + label + "}";
      if (proxSim) links.push(id + " -->|Sim| " + proxSim);
      if (proxNao) links.push(id + " -->|Não| " + proxNao);
    } else {
      nodes[id] = id + '["' + label + '"]';
      if (proxSim) links.push(id + " --> " + proxSim);
      if (proxNao) links.push(id + " --> " + proxNao);
    }
    
    // COR APENAS se coluna Cor existir
    if (cor && cor !== "") {
      nodes[id] += "::: " + cor;
    }
  });
  
  // **CADEIA PRINCIPAL HORIZONTAL L→R** (resolve vertical!)
  for (let i = 0; i < allIds.length - 1; i++) {
    links.unshift(allIds[i] + " -.-> " + allIds[i+1]);
  }
  
  // Nós ordenados
  allIds.forEach(function(id) {
    mermaidCode += nodes[id] + "\n";
  });
  
  links.forEach(function(l) {
    mermaidCode += l + "\n";
  });
  
  mermaidCode += "\n";
  
  // Classes com BORDA E TEXTO PRETOS sempre
  mermaidCode += "classDef default fill:#f0f9ff,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef green fill:#95d5b2,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef red fill:#ef476f,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef white fill:#ffffff,stroke:#000,stroke-width:4px,color:#000\n";
  
  try {
    document.getElementById("diagram").innerHTML = mermaidCode;
    mermaid.init();
    
    // Atualiza métricas (IDs do seu HTML original)
    if (document.getElementById("tempoTotal")) document.getElementById("tempoTotal").textContent = tempoTotal;
    if (document.getElementById("loops")) document.getElementById("loops").textContent = loops;
    if (document.getElementById("manual")) document.getElementById("manual").textContent = manual;
    if (document.getElementById("sistemaAuto")) document.getElementById("sistemaAuto").textContent = sistemaAuto;
    
  } catch (error) {
    document.getElementById("diagram").innerHTML = "<p>Erro: " + error + "</p>";
  }
}

function exportarPNG() {
  html2canvas(document.querySelector("#diagram"), {
    scale: 2,
    useCORS: true,
    width: document.querySelector("#diagram").scrollWidth,
    height: document.querySelector("#diagram").scrollHeight
  }).then(function(canvas) {
    let link = document.createElement('a');
    link.download = 'fluxo_bpmn.png';
    link.href = canvas.toDataURL();
    link.click();
  });
}

function limparTudo() {
  document.getElementById("entrada").value = "";
  document.getElementById("processo").value = "";
  document.getElementById("analista").value = "";
  document.getElementById("diagram").innerHTML = "";
  // Limpa métricas se existirem
  ["tempoTotal", "loops", "manual", "sistemaAuto"].forEach(id => {
    let el = document.getElementById(id);
    if (el) el.textContent = "0";
  });
}
