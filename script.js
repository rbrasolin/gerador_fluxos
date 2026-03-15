mermaid.initialize({
  startOnLoad: false,
  flowchart: {
    curve: 'basis',
    nodeSpacing: 150,
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
  
  console.log("Linhas processadas:", linhas.length);  // DEBUG
  
  linhas.forEach(function(linha, linhaIndex) {
    if (!linha.trim()) return;
    
    // **PARSE ROBUSTO**: tenta tab, depois múltiplos espaços
    let col = linha.split(/\t/);
    if (col.length < 2) col = linha.trim().split(/ {2,}/);
    
    console.log(`Linha ${linhaIndex}:`, col);  // DEBUG cada linha
    
    if (col.length < 6) {
      console.warn("Ignorando linha curta:", linha);
      return;
    }
    
    let idNumero = parseInt(col[1]);
    if (isNaN(idNumero)) {
      console.warn("ID inválido:", col[1]);
      return;
    }
    
    let id = "A" + idNumero;
    let atividade = limparTexto(col[2]);
    let tipo = (col[3] || "").toLowerCase();
    let sistema = limparTexto(col[4] || "");
    let tempo = parseFloat(col[5]) || 0;
    let proxSim = col[6] ? "A" + parseInt(col[6]) : null;
    let proxNao = col[7] ? "A" + parseInt(col[7]) : null;
    let cor = (col[8] || "").toLowerCase().trim();
    
    console.log(`Processado: ${id} "${atividade}" ${tempo}min cor:${cor}`);  // DEBUG
    
    tempoTotal += tempo;
    
    if (tipo === "manual") manual++;
    else sistemaAuto++;
    if (proxNao && parseInt(col[7]) < idNumero) loops++;
    
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
  
  console.log("Nodes criados:", Object.keys(nodes));  // DEBUG
  
  // Ordena por número ID
  Object.keys(nodes).sort(function(a,b) {
    return Number(a.slice(1)) - Number(b.slice(1));
  }).forEach(function(id) {
    mermaidCode += nodes[id] + "\n";
  });
  
  links.forEach(function(l) {
    mermaidCode += l + "\n";
  });
  
  mermaidCode += "\n";
  mermaidCode += "classDef default fill:#f0f9ff,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef green fill:#95d5b2,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef red fill:#ef476f,stroke:#000,stroke-width:3px,color:#000\n";
  mermaidCode += "classDef white fill:#ffffff,stroke:#000,stroke-width:3px,color:#000\n";
  
  try {
    document.getElementById("diagram").innerHTML = mermaidCode;
    mermaid.init();
    
    // Métricas
    document.getElementById("tempoTotal") && (document.getElementById("tempoTotal").textContent = tempoTotal);
    document.getElementById("loops") && (document.getElementById("loops").textContent = loops);
    document.getElementById("manual") && (document.getElementById("manual").textContent = manual);
    document.getElementById("sistemaAuto") && (document.getElementById("sistemaAuto").textContent = sistemaAuto);
    
    console.log("Mermaid code gerado:", mermaidCode);  // DEBUG final
    
  } catch (error) {
    console.error("Erro Mermaid:", error);
    document.getElementById("diagram").innerHTML = "<pre>Erro:\n" + error + "\n\nVerifique console (F12)</pre>";
  }
}

function exportarPNG() {
  const svg = document.querySelector('#diagram svg');
  if (svg) {
    const svgData = new XMLSerializer().serializeToString(svg);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    img.onload = () => {
      canvas.width = img.width;
      canvas.height = img.height;
      ctx.drawImage(img
