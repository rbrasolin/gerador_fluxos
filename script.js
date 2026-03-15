mermaid.initialize({
startOnLoad:false,
flowchart:{
curve:'basis',
nodeSpacing:50,
rankSpacing:80
}
});

function limparTexto(txt){
return (txt || "")
.replace(/"/g,'')
.replace(/\(/g,'')
.replace(/\)/g,'')
.replace(/:/g,'')
.trim();
}

function gerarFluxo(){

let texto = document.getElementById("entrada").value;
let processo = document.getElementById("processo").value;
let analista = document.getElementById("analista").value;

let linhas = texto.trim().split("\n");

let mermaidCode = "flowchart LR\n";

let nodes = {};
let links = [];

// métricas
let tempoTotal = 0;
let loops = 0;
let manual = 0;
let sistemaAuto = 0;

let atividadesTempo = [];

linhas.forEach(linha=>{

if(!linha.trim()) return;

let col = linha.trim().split(/\s{2,}|\t/);

if(col.length < 6) return;

let id = "A"+col[1];
let idNumero = Number(col[1]);

let atividade = limparTexto(col[2]);
let tipo = (col[3] || "").toLowerCase();
let sistema = limparTexto(col[4] || "");
let tempo = Number(col[5]) || 0;

// 🔧 CORREÇÃO DO BUG DAS CORES
let proxSim = (col[6] && !isNaN(col[6])) ? "A"+col[6] : null;
let proxNao = (col[7] && !isNaN(col[7])) ? "A"+col[7] : null;

let cor = limparTexto(col[8] || "white").toLowerCase();

tempoTotal += tempo;

atividadesTempo.push({
atividade:atividade,
tempo:tempo
});

// manual vs sistema

if(tipo === "manual"){
manual++;
}else{
sistemaAuto++;
}

// detectar loop simples

if(proxNao && Number(col[7]) < idNumero){
loops++;
}

// label sem <br>

let label = `${atividade}\\n${sistema}\\n${tempo} min`;

if(atividade.includes("?")){

nodes[id] = `${id}{${label}}`;

if(proxSim){
links.push(`${id} -->|Sim| ${proxSim}`);
}

if(proxNao){
links.push(`${id} -->|Não| ${proxNao}`);
}

}else{

nodes[id] = `${id}["${label}"]`;

if(proxSim){
links.push(`${id} --> ${proxSim}`);
}

if(proxNao){
links.push(`${id} --> ${proxNao}`);
}

}

// aplicar cor

if(cor !== "white"){
nodes[id] += `:::${cor}`;
}

});


// sort correto

Object.keys(nodes)
.sort((a,b)=>Number(a.slice(1))-Number(b.slice(1)))
.forEach(id=>{
mermaidCode += nodes[id] + "\n";
});

// links

links.forEach(l=>{
mermaidCode += l + "\n";
});

// cores

mermaidCode += `
classDef blue fill:#8ecae6
classDef yellow fill:#ffd166
classDef green fill:#95d5b2
classDef red fill:#ef476f
classDef white fill:#ffffff
`;

try{

document.getElementById("diagram").innerHTML =
`<div class="mermaid">${mermaidCode}</div>`;

mermaid.init(undefined, document.querySelectorAll(".mermaid"));

}catch(e){

console.error("Erro Mermaid:", mermaidCode);

alert("Erro ao gerar fluxo. Verifique caracteres especiais.");

}

// ===== INFO PROCESSO =====

document.getElementById("infoProcesso").innerHTML = `
<b>Processo:</b> ${processo}<br>
<b>Analista:</b> ${analista}
`;


// ===== MÉTRICAS =====

let totalAtividades = manual + sistemaAuto;

let pctManual = totalAtividades ? ((manual/totalAtividades)*100).toFixed(1) : 0;
let pctSistema = totalAtividades ? ((sistemaAuto/totalAtividades)*100).toFixed(1) : 0;


// TOP 3 gargalos

atividadesTempo.sort((a,b)=>b.tempo-a.tempo);

let top3 = atividadesTempo.slice(0,3)
.map(a=>`${a.atividade} (${a.tempo} min)`)
.join("<br>");

let indiceRetrabalho = ((loops/totalAtividades)*100).toFixed(1);


// ===== PARETO =====

let paretoHTML = "<b>Pareto de tempo</b><br><br>";

atividadesTempo.forEach(a=>{
let pct = ((a.tempo/tempoTotal)*100).toFixed(1);
paretoHTML += `${a.atividade} — ${pct}%<br>`;
});

document.getElementById("metricas").innerHTML = `

<b>Tempo total do processo:</b> ${tempoTotal} min<br><br>

<b>Top 3 gargalos:</b><br>
${top3}<br><br>

<b>Loops de retrabalho:</b> ${loops}<br>
<b>Índice de retrabalho:</b> ${indiceRetrabalho}%<br><br>

<b>Manual:</b> ${pctManual}%<br>
<b>Sistema:</b> ${pctSistema}%<br><br>

${paretoHTML}

`;

}


// ===== DOWNLOAD PNG =====

function baixarPNG(){

let svg = document.querySelector("#diagram svg");

if(!svg){
alert("Gere o fluxo primeiro");
return;
}

let serializer = new XMLSerializer();
let source = serializer.serializeToString(svg);

let svg64 = btoa(unescape(encodeURIComponent(source)));
let image64 = 'data:image/svg+xml;base64,' + svg64;

let img = new Image();

img.onload = function(){

let canvas = document.createElement("canvas");

canvas.width = img.width * 3;
canvas.height = img.height * 3;

let ctx = canvas.getContext("2d");

ctx.fillStyle = "white";
ctx.fillRect(0,0,canvas.width,canvas.height);

ctx.drawImage(img,0,0,canvas.width,canvas.height);

let link = document.createElement("a");
link.download = "fluxograma.png";
link.href = canvas.toDataURL("image/png");
link.click();

};

img.src = image64;

}