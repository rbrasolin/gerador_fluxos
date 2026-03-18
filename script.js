let ultimoNomeArquivo = "fluxograma_processo";

const CONFIG = {
  boxWidth: 260,
  boxHeight: 120,
  colSpacing: 140,
  rowSpacing: 100
};

function limpar(txt) {
  if (!txt) return "";
  return String(txt).trim();
}

function quebrarTexto(texto, max = 18) {
  const palavras = texto.split(" ");
  let linhas = [];
  let atual = "";

  palavras.forEach(p => {
    if ((atual + " " + p).length > max) {
      linhas.push(atual.trim());
      atual = p;
    } else {
      atual += " " + p;
    }
  });

  if (atual) linhas.push(atual.trim());

  return linhas;
}

function gerarNomeArquivo() {
  return "fluxo_svg";
}

function gerarFluxo() {
  const texto = document.getElementById("entrada").value;
  if (!texto.trim()) return alert("Cole a tabela");

  const linhas = texto.split("\n").map(l => l.split("\t"));

  const etapas = [];

  linhas.forEach(col => {
    while (col.length < 12) col.push("");

    etapas.push({
      id: limpar(col[1]),
      atividade: limpar(col[2]),
      sistema: limpar(col[4]),
      tempo: limpar(col[5]),
      proxSim: limpar(col[6]),
      proxNao: limpar(col[7]),
      extras: limpar(col[8]),
      coluna: Number(col[9]),
      linha: Number(col[10]),
      cor: limpar(col[11]) || "white"
    });
  });

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");

  const elementos = {};
  const posicoes = {};

  // 🔷 desenhar caixas
  etapas.forEach(e => {
    const x = e.coluna * (CONFIG.boxWidth + CONFIG.colSpacing);
    const y = e.linha * (CONFIG.boxHeight + CONFIG.rowSpacing);

    posicoes[e.id] = { x, y };

    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");

    const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("x", x);
    rect.setAttribute("y", y);
    rect.setAttribute("width", CONFIG.boxWidth);
    rect.setAttribute("height", CONFIG.boxHeight);
    rect.setAttribute("fill", e.cor);
    rect.setAttribute("stroke", "#111");
    rect.setAttribute("rx", 12);

    g.appendChild(rect);

    const linhasTexto = [
      ...quebrarTexto(e.atividade),
      e.sistema,
      e.tempo
    ];

    linhasTexto.forEach((t, i) => {
      const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
      text.setAttribute("x", x + CONFIG.boxWidth / 2);
      text.setAttribute("y", y + 25 + i * 18);
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-size", "14");
      text.textContent = t;

      g.appendChild(text);
    });

    svg.appendChild(g);
    elementos[e.id] = g;
  });

  // 🔶 desenhar setas
  function desenharLinha(orig, dest) {
    if (!posicoes[orig] || !posicoes[dest]) return;

    const o = posicoes[orig];
    const d = posicoes[dest];

    const linha = document.createElementNS("http://www.w3.org/2000/svg", "line");
    linha.setAttribute("x1", o.x + CONFIG.boxWidth);
    linha.setAttribute("y1", o.y + CONFIG.boxHeight / 2);
    linha.setAttribute("x2", d.x);
    linha.setAttribute("y2", d.y + CONFIG.boxHeight / 2);
    linha.setAttribute("stroke", "#000");
    linha.setAttribute("marker-end", "url(#arrow)");

    svg.appendChild(linha);
  }

  etapas.forEach(e => {
    e.proxSim.split(",").forEach(p => desenharLinha(e.id, p.trim()));
    e.proxNao.split(",").forEach(p => desenharLinha(e.id, p.trim()));
    e.extras.split(",").forEach(p => desenharLinha(e.id, p.trim()));
  });

  // seta
  const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");

  const marker = document.createElementNS("http://www.w3.org/2000/svg", "marker");
  marker.setAttribute("id", "arrow");
  marker.setAttribute("markerWidth", "10");
  marker.setAttribute("markerHeight", "10");
  marker.setAttribute("refX", "10");
  marker.setAttribute("refY", "3");
  marker.setAttribute("orient", "auto");

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M0,0 L10,3 L0,6 Z");
  path.setAttribute("fill", "#000");

  marker.appendChild(path);
  defs.appendChild(marker);
  svg.appendChild(defs);

  document.getElementById("diagram").innerHTML = "";
  document.getElementById("diagram").appendChild(svg);
}