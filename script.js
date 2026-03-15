// Espera DOM carregar
document.addEventListener('DOMContentLoaded', function() {
  mermaid.initialize({
    startOnLoad: false,
    flowchart: {
      curve: 'basis',
      nodeSpacing: 120,
      rankSpacing: 100
    }
  });
});

function limparTexto(txt) {
  return txt.replace(/"/g, '').replace(/\(/g, '').replace(/\)/g, '').replace(/:/g, '').trim();
}

function gerarFluxo() {
  const entrada = document.getElementById("entrada").value;
  const linhas = entrada.trim().split("\n");
  let mermaidCode = "flowchart LR\n";
  let nodes = {};
  let links = [];

  linhas.forEach(linha => {
    if (!linha.trim()) return;
    const col = linha.trim().split(/\s{2,}|\t/);
    if (col.length < 6) return;
    
    const idNumero = Number(col[1]);
    const id = "A" + idNumero;
    const atividade = limparTexto(col[2]);
    const sistema = limparTexto(col[4] || "");
    const tempo = Number(col[5]) || 0;
    const proxSim = col[6] ? "A" + col[6] : null;
    const proxNao = col[7] ? "A" + col[7] : null;
    const cor = (col[8] || "").toLowerCase().trim();
    
    const label = atividade + "\\n" + sistema + "\\n" + tempo + "min";
    
    // Losango se tem "?", senão retangular
    if (atividade.includes("?")) {
      nodes[id] = id + "{" + label + "}";
      if (proxSim) links.push(id + " -->|Sim| " + proxSim);
      if (proxNao) links.push(id + " -->|Não| " + proxNao);
    } else {
      nodes[id] = id + "[" + label + "]";
      if (proxSim) links.push(id + " --> " + proxSim);
      if (proxNao) links.push(id + " --> " + proxNao);
    }
    
    // Cor só se coluna Cor tem valor
    if (cor && cor !== "") {
      nodes[id] += ":::cor" + cor.charAt(0).toUpperCase() + cor.slice(1);
    }
  });

  // Monta código ordenado
  Object.keys(nodes).sort((a,b) => Number(a.slice(1)) - Number(b.slice(1)))
    .forEach(id => mermaidCode += nodes[id] + "\n");
  links.forEach(l => mermaidCode += l + "\n");

  // Classes com borda preta
  mermaidCode += "\nclassDef corBlue fill:#8ecae6,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corYellow fill:#ffd166,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corGreen fill:#95d5b2,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corRed fill:#ef476f,stroke:#000,stroke-width:2px,color:#000\n";
  mermaidCode += "classDef corWhite fill:#ffffff,stroke:#000,stroke-width:2px,color:#000\n";

  // AUTO-RENDER no DIV (sem render manual!)
  document.getElementById("diagram").innerHTML = mermaidCode;
  mermaid.init();  // Força render
}
