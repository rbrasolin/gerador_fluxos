mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'basis',
    nodeSpacing: 120,
    rankSpacing: 100
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
    
    if (tipo === "manual") manual++;
    else sistemaAuto++;
    if (proxNao && Number(col[7]) < idNumero) loops++;
    
    let label = atividade + "\\n" + sistema + "\\n" + tempo + "min";
    
    // Losango para decisões (?), retangular para processos
    if (atividade.includes("?")) {
      nodes[id] = id + "{" + label + "}";  // Losango
      if (proxSim) links.push(id + " -->|Sim| " + proxSim);
      if (proxNao) links.push(id + " -->|Não| " + proxNao);
    } else {
      nodes[id] = id + "[" + label + "]";  // Retangular (sem "")
      if (proxSim) links.push(id + " --> " + proxSim);
      if (proxNao) links.push(id + " --> " + proxNao);
    }
    
    // COR APENAS se coluna Cor existir e não estiver vazia
    if (cor && cor !== "") {
      nodes[id] += ":::cor" + cor.charAt(0).toUpperCase() + cor.slice(1);
    }
  });
  
  // Ordem sequencial para fluxo natural L→R
  Object.keys(nodes)
    .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)))
    .forEach(id => mermaidCode += nodes[id] + "\n");
  
  links.forEach(l => mermaidCode += l + "\n");
  
  // Classes com borda/texto PRETO sempre, cor só no fill
  mermaidCode += "\nclassDef corBlue fill:#8ecae6,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corYellow fill:#ffd166,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corGreen fill:#95d5b2,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corRed fill:#ef476f,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corWhite fill:#ffffff,stroke:#000,stroke-width:2px,color:#000\n";
  
  try {
    mermaid.render('diagramDiv', mermaidCode, (svg) => {
      document.getElementById("diagram").innerHTML = svg.svg;
    });
  } catch (e) {
    document.getElementById("diagram").innerHTML = "Erro: " + e.message;
  }
}
