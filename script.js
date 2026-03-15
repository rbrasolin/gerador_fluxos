mermaid.initialize({ startOnLoad:false });

function gerarFluxo(){

let texto = document.getElementById("entrada").value;
let processo = document.getElementById("processo").value;
let analista = document.getElementById("analista").value;

let linhas = texto.trim().split("\n");

// fluxo esquerda → direita
let mermaidCode = "flowchart LR\n";

let nodes = {};
let links = [];

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
let proxSim = col[6];
let proxNao = col[7];
let cor = col[8] || "white";

// métricas

tempoTotal += tempo;

if(tempo > maiorTempo){
maiorTempo = tempo;
gargalo = atividade;
}

if(tipo.toLowerCase() === "manual"){
manual++;
}else{
sistemaAuto++;
}

// detectar loops simples

if(proxNao && proxNao < id){
loops++;
}

// ===== NÓ DECISÃO OU ATIVIDADE =====

let label = `${atividade}<br>${sistema}<br>${tempo} min`;

if(atividade.includes("?")){

// losango (decisão)

nodes[id] = `${id}{${label}}:::${cor}`;

// seta SIM
if(proxSim){
links.push(`${id} -->|Sim| ${proxSim}`);
}

// seta NÃO
if(proxNao){
links.push(`${id} -->|Não| ${proxNao}`);
}

}else{

// caixa normal

nodes[id] = `${id}["${label}"]:::${cor}`;

if(proxSim){
links.push(`${id} --> ${proxSim}`);
}

if(proxNao){
links.push(`${id} --> ${proxNao}`);
}

}

});

// montar nodes

Object.values(nodes).forEach(n=>{
mermaidCode += n + "\n";
});

// montar links

links.forEach(l=>{
mermaidCode += l + "\n";
});

// cores

mermaidCode += `
classDef blue fill:#8ecae6
classDef yellow fill:#ffd166
classDef green fill:#95d5b2
classDef white fill:#ffffff
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

document.getElementById("metricas").innerHTML = `

<b>Tempo total do processo:</b> ${tempoTotal} min<br><br>

<b>Gargalo do processo:</b> ${gargalo} (${maiorTempo} min)<br><br>

<b>Loops de retrabalho detectados:</b> ${loops}<br><br>

<b>Manual:</b> ${pctManual}%<br>
<b>Sistema:</b> ${pctSistema}%

`;

}


// ===== DOWNLOAD SVG =====

function baixarSVG(){

let svg = document.querySelector("#diagram svg");

if(!svg){
alert("Gere o fluxo primeiro");
return;
}

let serializer = new XMLSerializer();
let source = serializer.serializeToString(svg);

let blob = new Blob([source], {type:"image/svg+xml;charset=utf-8"});

let url = URL.createObjectURL(blob);

let a = document.createElement("a");
a.href = url;
a.download = "fluxograma.svg";
a.click();

URL.revokeObjectURL(url);

}