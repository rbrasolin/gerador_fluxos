mermaid.initialize({ startOnLoad:false });

function gerarFluxo(){

let texto = document.getElementById("entrada").value;

let linhas = texto.trim().split("\n");

let mermaidCode = "flowchart TD\n";

let nodes = {};
let links = [];

linhas.forEach(linha=>{

if(!linha.trim()) return;

let col = linha.trim().split(/\s{2,}|\t/);

if(col.length < 6) return;

let id = col[1];
let atividade = col[2];
let sistema = col[4];
let tempo = col[5];
let proxSim = col[6];
let proxNao = col[7];
let cor = col[8] || "white";

nodes[id] = `${id}["${atividade}<br>${sistema}<br>${tempo} min"]:::${cor}`;

if(proxSim){
links.push(`${id} --> ${proxSim}`);
}

if(proxNao){
links.push(`${id} -->|Não| ${proxNao}`);
}

});

Object.values(nodes).forEach(n=>{
mermaidCode += n + "\n";
});

links.forEach(l=>{
mermaidCode += l + "\n";
});

mermaidCode += `
classDef blue fill:#8ecae6
classDef yellow fill:#ffd166
classDef green fill:#95d5b2
`;

document.getElementById("diagram").innerHTML =
`<div class="mermaid">${mermaidCode}</div>`;

mermaid.init(undefined, document.querySelectorAll(".mermaid"));

}