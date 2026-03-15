mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'linear',
    nodeSpacing: 220,
    rankSpacing: 50
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
  
  // 1. Coleta IDs únicos da coluna B (ID)
  linhas.forEach(function(linha) {
    if (!linha.trim()) return;
    let col = linha.trim().split(/\s{2,}|\t/);  // Split por tabs/espaços múltiplos
    if (col.length < 6) return;
    let id = limparTexto(col[1]);  // Coluna B = ID (A, B, C...)
    if (id) allIds.push(id);
  });
  allIds = [...new Set(allIds)].sort();  // A,B,C,D...
  
  let tempoTotal = 0;
  let loops = 0;
  let manual = 0;
  let sistemaAuto = 0;
  let atividadesTempo = [];
  
  // 2. Processa cada linha
  linhas.forEach(function(linha) {
    if (!linha.trim()) return;
    let col = linha.trim().split(/\s{2,}|\t/);
    if (col.length < 6) return;
    
    let ordem = Number(col[0]);  // Coluna A (ignorada)
    let id = limparTexto(col[1]);  // Coluna B = ID
    let atividade = limparTexto(col[2]);  // Coluna C
    let tipo = (col[3] || "").toLowerCase();  // Coluna D
    let sistema = limparTexto(col[4] || "");  // Coluna E
    let tempo = Number(col[5]) || 0;  // Coluna F
    let proxSim = col[6] ? limparTexto(col[6]) : null;  // Coluna G
    let proxNao = col[7] ? limparTexto(col[7]) : null;  // Coluna H
    let cor = (col[8] || "").toLowerCase().trim();  // Coluna I
    
    tempoTotal += tempo;
    atividadesTempo.push({atividade, tempo});
    
    if (tipo === "manual") manual++;
    else sistemaAuto++;
    if (proxNao && allIds.indexOf(proxNao) < allIds.indexOf(id)) loops++;
    
    let label = atividade + "\\n" + sistema + "\\n" + tempo + "min";
    
    if (atividade.includes("?")) {
      // Losango para decisões
      nodes[id] = id + "{" + label + "}";
      if (proxSim) links.push(id + " -->|Sim| " + proxSim);
      if (proxNao) links.push(id + " -->|Não| " + proxNao);
    } else {
      // Retangular para atividades
      nodes[id] = id + '["' + label + '"]';
      if (proxSim) links.push(id + " --> " + proxSim);
      if (proxNao) links.push(id + " --> " + proxNao);
    }
    
    // COR só se coluna I existir
    if (cor && cor !== "") {
      nodes[id] += "::: " + cor;
    }
  });
  
  // 3. **CADEIA PRINCIPAL HORIZONTAL** (resolve vertical!)
  for (let i = 0; i < allIds.length - 1; i++) {
    links.unshift(allIds[i] + " -.-> " + allIds[i+1]);  // Linha tracejada sutil
  }
  
  // 4. Monta código Mermaid
  allIds.forEach(function(id) {
    mermaidCode += nodes[id] + "\n";
  });
  links.forEach(function(l) {
    mermaidCode += l + "\n";
  });
  
  mermaidCode += "\n";
  // Estilos: borda/texto PRETOS sempre
  mermaidCode += "classDef default fill:#f0f9ff,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef green fill:#95d5b2,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef red fill:#ef476f,stroke:#000,stroke-width:4px,color:#000\n";
  mermaidCode += "classDef white fill:#ffffff,stroke:#000,stroke-width:4px,color:#000\n";
  
  try {
    document.getElementById("diagram").innerHTML = mermaidCode;
    mermaid.init();
    
    // Atualiza métricas HTML
    if (document.getElementById("tempoTotal")) document.getElementById("tempoTotal").textContent = tempoTotal;
    if (document.getElementById("loops")) document.getElementById("loops").textContent = loops;
    if (document.getElementById("manual")) document.getElementById("manual").textContent = manual;
    if (document.getElementById("sistemaAuto")) document.getElementById("sistemaAuto").textContent = sistemaAuto;
    
  } catch (error) {
    document.getElementById("diagram").innerHTML = "<p>Erro: " + error.message + "</p>";
  }
}

function exportarPNG() {
  html2canvas(document.querySelector("#diagram"), {
    scale: 2,
    useCORS: true
  }).then(canvas => {
    let link = document.createElement('a');
    link.download = 'fluxo.png';
    link.href = canvas.toDataURL();
    link.click();
  }).catch(e => alert("Erro export: " + e));
}

function limparTudo() {
  document.getElementById("entrada").value = "";
  document.getElementById("processo").value = "";
  document.getElementById("analista").value = "";
  document.getElementById("diagram").innerHTML = "";
}
