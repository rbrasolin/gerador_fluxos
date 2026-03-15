mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'basis',
    nodeSpacing: 180,
    rankSpacing: 80
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
    
    // **PARSE EXATO**: tab OU 2+ espaços (formato Excel copiado)
    let col = linha.split(/\t| {2,}/);
    if (col.length < 6) return;
    
    let ordem = col[0].trim();      // A: 1
    let id = col[1].trim();         // B: A, B, C...
    let atividade = limparTexto(col[2]);  // C: "Receber?"
    let tipo = col[3].trim().toLowerCase(); // D: Manual
    let sistema = limparTexto(col[4]);  // E: Email
    let tempo = parseFloat(col[5]) || 0; // F: 5.0
    let proxSim = col[6] ? col[6].trim() : null; // G: B
    let proxNao = col[7] ? col[7].trim() : null; // H: 
    let cor = col[8] ? col[8].trim().toLowerCase() : ""; // I: blue
    
    tempoTotal += tempo;
    atividadesTempo.push({atividade, tempo});
    
    if (tipo === "manual") manual++;
    else sistemaAuto++;
    if (proxNao && proxNao !== "") loops++;
    
    let label = atividade + "\\n" + sistema + "\\n" + tempo + "min";
    
    if (atividade.includes("?")) {
      nodes[id] = id + "{" + label + "}";
      if (proxSim && proxSim !== "") links.push(id + " -->|Sim| " + proxSim);
      if (proxNao && proxNao !== "") links.push(id + " -->|Não| " + proxNao);
    } else {
      nodes[id] = id + '["' + label + '"]';
      if (proxSim && proxSim !== "") links.push(id + " --> " + proxSim);
      if (proxNao && proxNao !== "") links.push(id + " --> " + proxNao);
    }
    
    // Cor SÓ se coluna I tem valor
    if (cor !== "") {
      nodes[id] += "::: " + cor;
    }
  });
  
  // **LINKS NORMAIS PRIMEIRO** (Prox_Sim/Prox_Nao)
  links.forEach(l => mermaidCode += l + "\n");
  
  // Nós ordenados por ID
  Object.keys(nodes).sort().forEach(id => {
    mermaidCode += nodes[id] + "\n";
  });
  
  // Estilos - BORDA PRETA SEMPRE
  mermaidCode += "\nclassDef default fill:#e3f2fd,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef green fill:#95d5b2,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef red fill:#ef476f,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef white fill:#ffffff,stroke:#000,stroke-width:3px,color:#000\n";
  
  try {
    document.getElementById("diagram").innerHTML = mermaidCode;
    
    // **RENDER DUAS VEZES** (corrige layout Mermaid)
    mermaid.init();
    setTimeout(() => mermaid.init(), 100);
    
    // Métricas
    if (document.getElementById("tempoTotal")) document.getElementById("tempoTotal").innerHTML = tempoTotal + " min";
    if (document.getElementById("loops")) document.getElementById("loops").innerHTML = loops;
    if (document.getElementById("manual")) document.getElementById("manual").innerHTML = manual;
    if (document.getElementById("sistemaAuto")) document.getElementById("sistemaAuto").innerHTML = sistemaAuto;
    
  } catch (error) {
    console.error(error);
    document.getElementById("diagram").innerHTML = "Erro parsing: verifique formato Excel";
  }
}

function exportarPNG() {
  // Fallback simples se html2canvas não carregar
  const svg = document.querySelector("#diagram svg");
  if (svg) {
    const serializer = new XMLSerializer();
    let svgStr = serializer.serializeToString(svg);
    svgStr = '<?xml version="1.0" standalone="no"?>\r\n' + svgStr;
    const svgBlob = new Blob([svgStr], {type:"image/svg+xml;charset=utf-8"});
    const svgUrl = URL.createObjectURL(svgBlob);
    const link = document.createElement("a");
    link.href = svgUrl;
    link.download = "fluxo_bpmn.svg";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}

function limparTudo() {
  document.getElementById("entrada").value = "";
  document.getElementById("processo").value = "";
  document.getElementById("analista").value = "";
  document.getElementById("diagram").innerHTML = "";
}
