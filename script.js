<script>
document.addEventListener('DOMContentLoaded', () => {
  mermaid.initialize({
    startOnLoad: false,
    flowchart: {
      curve: 'linear',
      nodeSpacing: 200,  // MAIS horizontal
      rankSpacing: 60    // MENOS vertical
    }
  });
});

function limparTexto(txt) {
  return txt.replace(/"/g, '').replace(/\(/g, '').replace(/\)/g, '').replace(/:/g, '').trim();
}

function gerarFluxo() {
  let texto = document.getElementById("entrada").value;
  let linhas = texto.trim().split("\n");
  let mermaidCode = "flowchart LR\n";  // FORÇA L→R
  let nodes = {}, links = [];
  let tempoTotal = 0, loops = 0, manual = 0, sistemaAuto = 0, atividadesTempo = [];

  linhas.forEach(linha => {
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
    let cor = col[8] ? col[8].toLowerCase().trim() : "";
    
    tempoTotal += tempo;
    atividadesTempo.push({atividade, tempo});
    if (tipo === "manual") manual++;
    else sistemaAuto++;
    if (proxNao && Number(col[7]) < idNumero) loops++;
    
    let label = `${atividade}\\n${sistema}\\n${tempo}min`;  // TEXTO VISÍVEL
    
    let nodeDef;
    if (atividade.includes("?")) {
      nodeDef = id + `{"${label}"}`;  // Losango
      if (proxSim) links.push(`${id} -->|Sim| ${proxSim}`);
      if (proxNao) links.push(`${id} -->|Não| ${proxNao}`);
    } else {
      nodeDef = id + `["${label}"]`;  // Retangular
      if (proxSim) links.push(`${id} --> ${proxSim}`);
      if (proxNao) links.push(`${id} --> ${proxNao}`);
    }
    
    // COR SÓ SE COLUNA 9 EXISTE
    if (cor && cor !== "") {
      nodeDef += `:::${cor}`;
    }
    nodes[id] = nodeDef;
  });

  // SEQUÊNCIA L→R rigorosa
  Object.keys(nodes).sort((a,b)=>Number(a.slice(1))-Number(b.slice(1)))
    .forEach(id => mermaidCode += nodes[id] + "\n");
  links.forEach(l => mermaidCode += l + "\n");

  // BORDA PRETA SEMPRE + cores específicas
  mermaidCode += "\n";
  mermaidCode += "classDef default fill:#e1f5fe,stroke:#000,stroke-width:3px,color:#000\n";  // PADRÃO
  if (Object.values(nodes).some(n=>n.includes('blue'))) 
    mermaidCode += "classDef blue fill:#8ecae6,stroke:#000,stroke-width:3px,color:#000\n";
  if (Object.values(nodes).some(n=>n.includes('yellow'))) 
    mermaidCode += "classDef yellow fill:#ffd166,stroke:#000,stroke-width:3px,color:#000\n";
  if (Object.values(nodes).some(n=>n.includes('green'))) 
    mermaidCode += "classDef green fill:#95d5b2,stroke:#000,stroke-width:3px,color:#000\n";
  if (Object.values(nodes).some(n=>n.includes('red'))) 
    mermaidCode += "classDef red fill:#ef476f,stroke:#000,stroke-width:3px,color:#000\n";

  document.getElementById("diagram").innerHTML = mermaidCode;
  mermaid.init();
  mostrarAnalises(tempoTotal, loops, manual, sistemaAuto, atividadesTempo);
}

// Resto das funções iguais...
function mostrarAnalises(tempoTotal, loops, manual, sistemaAuto, atividades) {
  atividades.sort((a,b)=>b.tempo-a.tempo);
  let pareto80 = atividades.slice(0,Math.ceil(atividades.length*0.2)).reduce((s,a)=>s+a.tempo,0);
  document.getElementById("analises").innerHTML = `
    <div class="bg-white p-6 rounded-xl shadow-lg">
      <h3 class="text-xl font-bold mb-4">📊 Análises</h3>
      <p><strong>Tempo Total:</strong> ${tempoTotal}min</p>
      <p><strong>Gargalos:</strong> ${loops}</p>
      <p><strong>Manual:</strong> ${manual} | <strong>Auto:</strong> ${sistemaAuto}</p>
    </div>
    <div class="bg-white p-6 rounded-xl shadow-lg">
      <h3 class="text-xl font-bold mb-4">📈 Pareto</h3>
      <p><strong>80%:</strong> ${pareto80}min (${((pareto80/tempoTotal)*100).toFixed(0)}%)</p>
    </div>
  `;
}

function exportarPNG() {
  const svg = document.querySelector('#diagram svg');
  if (svg) {
    const svgData = new XMLSerializer().serializeToString(svg);
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    const img = new Image();
    img.onload = function() {
      canvas.width = img.width;
      canvas.height = img.height;
      ctx.drawImage(img, 0, 0);
      const link = document.createElement("a");
      link.download = "fluxo.png";
      link.href = canvas.toDataURL();
      link.click();
    };
    img.src = "data:image/svg+xml;base64,"+btoa(unescape(encodeURIComponent(svgData)));
  }
}

function limparTudo() {
  document.querySelectorAll('input, textarea').forEach(el => el.value = '');
  document.getElementById("diagram").innerHTML = '';
  document.getElementById("analises").innerHTML = '';
}
</script>
