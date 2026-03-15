mermaid.initialize({
startOnLoad:false,
flowchart:{
curve:'basis',
nodeSpacing:50,
rankSpacing:80
}
});

function gerarFluxo(){

let texto = document.getElementById("entrada").value;
let processo = document.getElementById("processo").value;
let analista = document.getElementById("analista").value;

let linhas = texto.trim().split("\n");

// fluxo esquerda → direita
let mermaidCode = "flowchart LR\n";

let nodes = {};
let links = [];
let rankFix = [];

let atividadesTempo = [];

// métricas
let tempoTotal = 0;
let maiorTempo = 0;
let gargalo = "";
let loops = 0;
let manual = 0;
let sistemaAuto = 0;

linhas.forEach(linha=>{

if(!linha.trim()) return;

let col = linha.trim().split(/\s{2,}|\t/);

if(col.length < 6) return;

let id = col[1];
let atividade = col[2];
let tipo = col[3] || "";
let sistema = col[4] || "";
let tempo = Number(col[5]) || 0;

let proxSim = col[6] || "";
let proxNao = col[7] || "";
let cor = col[8] ? col[8].toLowerCase() : "white";

// impedir erro da coluna cor virar próximo passo
if(isNaN(proxSim)) proxSim="";
if(isNaN(proxNao)) proxNao="";

// métricas

tempoTotal += tempo;

atividadesTempo.push({
id:id,
atividade:atividade,
tempo:tempo
});

if(tempo > maiorTempo){
maiorTempo = tempo;
gargalo = atividade;
}

if(tipo.toLowerCase() === "manual"){
manual++;
}else{
sistemaAuto++;
}

// detectar loops

if(proxNao && Number(proxNao) < Number(id)){
loops++;
}

// label da caixa

let label = `${atividade}<br>${sistema}<br>${tempo} min`;

if(atividade.includes("?")){

nodes[id] = `${id}{${label}}:::${cor}`;

rankFix.push(id);

if(proxSim){
links.push(`${id} -->|Sim| ${proxSim}`);
}

if(proxNao){
links.push(`${id} -->|Não| ${proxNao}`);
}

}else{

nodes[id] = `${id}["${label}"]:::${cor}`;

if(proxSim){
links.push(`${id} --> ${proxSim}`);
}

if(proxNao){
links.push(`${id} --> ${proxNao}`);
}

}

});


// ===== ANALISE PARETO =====

atividadesTempo.sort((a,b)=>b.tempo-a.tempo);

let top3 = atividadesTempo.slice(0,3).map(a=>a.atividade);

let paretoHTML = "<b>Pareto de tempo das atividades</b><br><br>";

atividadesTempo.forEach((a,i)=>{
paretoHTML += `${i+1}. ${a.atividade} - ${a.tempo} min<br>`;
});


// ===== DESTACAR TOP 3 GARGALOS =====

Object.keys(nodes).forEach(id=>{

let atividadeNode = atividadesTempo.find(a=>a.id==id)?.atividade;

if(top3.includes(atividadeNode)){
nodes[id] = nodes[id].replace(":::", ":::gargalo ");
}

});


// montar nodes ordenados

Object.keys(nodes)
.sort((a,b)=>Number(a)-Number(b))
.forEach(id=>{
mermaidCode += nodes[id] + "\n";
});

// montar links

links.forEach(l=>{
mermaidCode += l + "\n";
});

// corrigir layout decisões

rankFix.forEach(id=>{
mermaidCode += `{rank=same; ${id}}\n`;
});


// cores

mermaidCode += `
classDef blue fill:#8ecae6
classDef yellow fill:#ffd166
classDef green fill:#95d5b2
classDef white fill:#ffffff
classDef gargalo fill:#ff6b6b,color:#ffffff
`;


// renderizar

document.getElementById("diagram").innerHTML =
`<div class="mermaid">${mermaidCode}</div>`;

mermaid.init(undefined, document.querySelectorAll(".mermaid"));


// ===== INFORMAÇÕES PROCESSO =====

document.getElementById("infoProcesso").innerHTML = `
<b>Processo:</b> ${processo}<br>
<b>Analista:</b> ${analista}
`;


// ===== MÉTRICAS =====

let totalAtividades = manual + sistemaAuto;

let pctManual = totalAtividades ? ((manual/totalAtividades)*100).toFixed(1) : 0;
let pctSistema = totalAtividades ? ((sistemaAuto/totalAtividades)*100).toFixed(1) : 0;

let indiceRetrabalho = totalAtividades ? ((loops/totalAtividades)*100).toFixed(1) : 0;

document.getElementById("metricas").innerHTML = `

<b>Tempo total do processo:</b> ${tempoTotal} min<br><br>

<b>Gargalo principal:</b> ${gargalo} (${maiorTempo} min)<br><br>

<b>Loops de retrabalho detectados:</b> ${loops}<br>

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

let img = new Image();
let svgBlob = new Blob([source], {type:"image/svg+xml;charset=utf-8"});
let url = URL.createObjectURL(svgBlob);

img.onload = function(){

let canvas = document.createElement("canvas");

canvas.width = img.width * 3;
canvas.height = img.height * 3;

let ctx = canvas.getContext("2d");

ctx.fillStyle = "white";
ctx.fillRect(0,0,canvas.width,canvas.height);

ctx.drawImage(img,0,0,canvas.width,canvas.height);

URL.revokeObjectURL(url);

let link = document.createElement("a");
link.download = "fluxograma.png";
link.href = canvas.toDataURL("image/png");
link.click();

};

img.src = url;

}