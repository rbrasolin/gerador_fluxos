mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'basis',
    nodeSpacing: 120,     // Mais espaço horizontal
    rankSpacing: 100      // Menos empilhamento vertical
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
    let cor = (col[8] || "").toLowerCase();
    
    tempoTotal += tempo;
    atividadesTempo.push({ atividade, tempo });
    
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
    
    if (cor) {
      nodes[id] += "::: " + cor;
    }
  });
  
  // Ordena nós numericamente para sequência natural esquerda-direita
  Object.keys(nodes)
    .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)))
    .forEach(function(id) {
      mermaidCode += nodes[id] + "\n";
    });
  
  links.forEach(function(l) {
    mermaidCode += l + "\n";
  });
  
  // Definições de cores
  mermaidCode += "\n";
  mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,color:#000\n";
  mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,color:#000\n";
  mermaidCode += "classDef green fill:#95d5b2,stroke:#000,color:#000\n";
  mermaidCode += "classDef red fill:#ef476f,stroke:#000,color:#000\n";
  mermaidCode += "classDef white fill:#ffffff,stroke:#000,color:#000\n";
  
  // Renderização assíncrona (melhor que innerHTML direto)
  try {
    mermaid.render('diagramDiv', mermaidCode, function(svgCode) {
      document.getElementById("diagram").innerHTML = svgCode.svg;
    });
    
    // Atualiza métricas (você pode mostrar no HTML)
    document.getElementById("tempoTotal").innerText = tempoTotal;
    document.getElementById("loops").innerText = loops;
    document.getElementById("manual").innerText = manual;
    document.getElementById("sistemaAuto").innerText = sistemaAuto;
    
  } catch (error) {
    document.getElementById("diagram").innerHTML = "<p>Erro no diagrama: " + error + "</p>";
  }
}
