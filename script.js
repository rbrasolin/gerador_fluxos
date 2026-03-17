let ultimoNomeArquivo = "fluxograma_processo";

mermaid.initialize({
  startOnLoad: false,
  securityLevel: "loose",
  theme: "base",
  themeVariables: {
    fontFamily: "Arial, sans-serif",
    fontSize: "18px"
  },
  flowchart: {
    htmlLabels: false, // CRÍTICO para Excel
    curve: "basis",
    nodeSpacing: 80,
    rankSpacing: 130
  }
});

function limpar(txt) {
  return (txt || "").toString().trim();
}

function gerarNomeArquivo() {
  const nome = limpar(document.getElementById("processo").value) || "fluxo";
  return nome.replace(/[^a-z0-9]/gi, "_").toLowerCase();
}

async function gerarFluxo() {
  const texto = document.getElementById("entrada").value;

  if (!texto.trim()) {
    alert("Cole os dados primeiro.");
    return;
  }

  const linhas = texto.split("\n").map(l => l.split("\t"));
  linhas.shift(); // remove cabeçalho

  let mermaidCode = "flowchart LR\n";

  linhas.forEach((col, i) => {
    const id = "A" + i;
    const atividade = limpar(col[2]);
    const sistema = limpar(col[4]);
    const tempo = limpar(col[5]);

    const label = `${atividade}\n${sistema}\n${tempo}`;

    mermaidCode += `${id}["${label}"]\n`;

    if (i > 0) {
      mermaidCode += `A${i - 1} --> ${id}\n`;
    }
  });

  const { svg } = await mermaid.render("id1", mermaidCode);
  document.getElementById("diagram").innerHTML = svg;

  ultimoNomeArquivo = gerarNomeArquivo();
}

function obterSVG() {
  const svg = document.querySelector("#diagram svg");
  if (!svg) {
    alert("Gere o fluxo primeiro.");
    return null;
  }
  return svg;
}

function baixarSVG() {
  const svg = obterSVG();
  if (!svg) return;

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);

  const blob = new Blob([source], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `${ultimoNomeArquivo}.svg`;
  a.click();
}

async function baixarPDF() {
  const svg = obterSVG();
  if (!svg) return;

  try {
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: "landscape",
      unit: "pt"
    });

    if (typeof pdf.svg !== "function") {
      throw new Error("svg2pdf não carregou corretamente");
    }

    await pdf.svg(svg, {
      x: 10,
      y: 10
    });

    pdf.save(`${ultimoNomeArquivo}.pdf`);

  } catch (e) {
    console.error(e);
    alert("Erro ao gerar PDF");
  }
}

function baixarPNG() {
  const svg = obterSVG();
  if (!svg) return;

  const serializer = new XMLSerializer();
  const source = serializer.serializeToString(svg);
  const svg64 = btoa(unescape(encodeURIComponent(source)));

  const img = new Image();
  img.src = "data:image/svg+xml;base64," + svg64;

  img.onload = function () {
    const canvas = document.createElement("canvas");
    canvas.width = 4000;
    canvas.height = 2000;

    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "white";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    const a = document.createElement("a");
    a.download = `${ultimoNomeArquivo}.png`;
    a.href = canvas.toDataURL("image/png");
    a.click();
  };
}

function limparTudo() {
  document.getElementById("entrada").value = "";
  document.getElementById("diagram").innerHTML = "";
}