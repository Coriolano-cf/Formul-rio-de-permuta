// ====== Estado/UI ======
const datasPermuta = [];
const datasPagamento = [];
const TEMPLATE_URL = "templates/Permuta_de_Servico.docx";

// Escape seguro (sem regex): usa DOM
function esc(s){
  const div = document.createElement("div");
  div.innerText = String(s);
  return div.innerHTML;
}

function formatarDataBR(iso){
  if(!iso) return "";
  const [a,m,d]=iso.split("-");
  return d+"/"+m+"/"+a;
}
function renderChips(containerId, lista){
  const c = document.getElementById(containerId);
  c.innerHTML="";
  lista.forEach((iso, idx)=>{
    const s=document.createElement("span");
    s.className="chip";
    s.innerHTML = formatarDataBR(iso) + ' <button type="button" aria-label="Remover">&times;</button>';
    s.querySelector("button").addEventListener("click",()=>{
      lista.splice(idx,1);
      renderChips(containerId, lista);
    });
    c.appendChild(s);
  });
}
function adicionarData(inputId, lista, chipsId){
  const el=document.getElementById(inputId);
  const v=(el.value||"").trim();
  if(!v) return;
  if(!lista.includes(v)){
    lista.push(v);
    lista.sort();
    renderChips(chipsId, lista);
  }
  el.value="";
  el.focus();
}
document.getElementById("addPermuta").addEventListener("click", ()=> adicionarData("dataPermuta", datasPermuta, "chipsPermuta"));
document.getElementById("addPagamento").addEventListener("click", ()=> adicionarData("dataPagamento", datasPagamento, "chipsPagamento"));
["dataPermuta","dataPagamento"].forEach(id=>{
  document.getElementById(id).addEventListener("keydown", e=>{
    if(e.key==="Enter"){
      e.preventDefault();
      id==="dataPermuta"
        ? adicionarData("dataPermuta", datasPermuta, "chipsPermuta")
        : adicionarData("dataPagamento", datasPagamento, "chipsPagamento");
    }
  });
});

function validarCampos(){
  let ok=true;
  ["pmSubstituido","pmSubstituto"].forEach(id=>{
    const el=document.getElementById(id);
    el.classList.remove("error");
    if(!el.value.trim()){
      el.classList.add("error");
      ok=false;
    }
  });
  ["dataPermuta","dataPagamento"].forEach(id=> document.getElementById(id).classList.remove("error"));
  if(datasPermuta.length===0){ document.getElementById("dataPermuta").classList.add("error"); ok=false; }
  if(datasPagamento.length===0){ document.getElementById("dataPagamento").classList.add("error"); ok=false; }
  return ok;
}

// ====== Helpers DOCX ======
function desfragmentarDocx(xml){
  return xml
    .replace(/<\/w:t>\s*<w:t[^>]*>/g, "")
    .replace(/<\/w:t>\s*<\/w:r>\s*<w:r[^>]*>\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>\s*)?<w:t[^>]*>/g, "");
}
function inserirAposRotulo(xml, rotulo, valor){
  const rot = rotulo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const esp = "(?:[ \\u00A0]*)";
  let re = new RegExp(rot + esp, "g");
  let novo = xml.replace(re, rotulo + " " + valor);
  if(novo !== xml) return novo;
  re = new RegExp(rot + "[\\s\\S]*?(?=</w:p>)", "g");
  return xml.replace(re, rotulo + " " + valor);
}

async function gerarDOCX(){
  if(!validarCampos()){
    alert("Por favor, preencha os campos e adicione ao menos uma data em cada lista.");
    return;
  }
  const pmSubstituido = document.getElementById("pmSubstituido").value.trim();
  const pmSubstituto  = document.getElementById("pmSubstituto").value.trim();
  const datasPermutaStr = datasPermuta.map(formatarDataBR).join(" - ");
  const datasPagtoStr   = datasPagamento.map(formatarDataBR).join(" - ");
  const carga  = document.getElementById("carga").value;

  let arrayBuffer;
  try{
    const resp = await fetch(TEMPLATE_URL);
    if(!resp.ok) throw new Error("HTTP " + resp.status);
    arrayBuffer = await resp.arrayBuffer();
  }catch(e){
    alert("Falha ao carregar o template. Abra com Live Server e confira a pasta /templates.");
    return;
  }

  const zip = new PizZip(arrayBuffer);
  const path = "word/document.xml";
  let xml = zip.file(path).asText();
  xml = desfragmentarDocx(xml);

  // Campos
  xml = inserirAposRotulo(xml, "PM SUBSTITUÍDO:", pmSubstituido);
  xml = inserirAposRotulo(xml, "PM SUBSTITUTO:",  pmSubstituto);
  xml = inserirAposRotulo(xml, "Data do serviço permutado:",    datasPermutaStr);
  xml = inserirAposRotulo(xml, "Data do pagamento do serviço:", datasPagtoStr);

  // >>> CARGA/Total de horas — rótulo com bolinha
  xml = inserirAposRotulo(xml, "• Total de horas trabalhadas:", carga);
  // (Opcional) variações comuns — descomente se seu template alternar:
  // xml = inserirAposRotulo(xml, "• Total de horas trabalhadas :", carga);
  // xml = inserirAposRotulo(xml, "Total de horas trabalhadas:", carga);
  // xml = inserirAposRotulo(xml, "Total de horas trabalhadas :", carga);

  zip.file(path, xml);
  const out = zip.generate({
    type:"blob",
    mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
  const primeiraData = (datasPermuta[0] || "").replace(/-/g,"");
  saveAs(out, `Permuta_de_Servico_${primeiraData || "preenchida"}.docx`);
}

// ====== PDF ======
function gerarPDF(){
  if(!validarCampos()){
    alert("Por favor, preencha os campos e adicione ao menos uma data em cada lista.");
    return;
  }
  const pmSubstituido = esc(document.getElementById("pmSubstituido").value.trim());
  const pmSubstituto  = esc(document.getElementById("pmSubstituto").value.trim());
  const datasPermutaStr = esc(datasPermuta.map(formatarDataBR).join(" - "));
  const datasPagtoStr   = esc(datasPagamento.map(formatarDataBR).join(" - "));
  const carga  = esc(document.getElementById("carga").value);

  const basePath = location.origin + location.pathname.replace(/\/[^/]*$/, "");
  const logoSrc = basePath + "/img/logo-pm.png";

  const parts = [
    "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"utf-8\">",
    "<title>Permuta de Serviço (PDF)</title>",
    "<style>@page{margin:22mm 18mm}",
      "body{font-family:Calibri,Arial,sans-serif;color:#111}",
      ".cab{text-align:center;margin-bottom:8pt}.cab img{height:70px}",
      "h1{font-size:16pt;margin:8pt 0 10pt;text-transform:uppercase}",
      ".sub{font-size:10.5pt;color:#444}",
      "table{width:100%;border-collapse:collapse;margin-top:10pt}",
      "th,td{border:1px solid #bbb;padding:6pt 8pt;font-size:12pt;vertical-align:top}",
      "th{width:36%;background:#f5f6f8;text-align:left}",
      ".assinaturas{margin-top:28pt}",
      ".linhaAss{margin-top:32pt;display:flex;justify-content:space-between;gap:20pt}",
      ".linhaAss div{width:48%;text-align:center}",
      ".rod{text-align:center;margin-top:26pt;font-size:11pt}",
    "</style></head><body>",
    "<div class=\"cab\"><img src=\"", logoSrc, "\" alt=\"Logo\">",
      "<div class=\"sub\">ESTADO DO MARANHÃO — SECRETARIA DE SEGURANÇA PÚBLICA</div>",
      "<div class=\"sub\">Polícia Militar do Maranhão — 10º BPM - 2ª CIA - 4º GPPM - Turiaçu-MA</div>",
      "<h1>FORMULÁRIO DE AUTORIZAÇÃO PARA PERMUTA DE SERVIÇO</h1></div>",
    "<table>",
      "<tr><th>PM SUBSTITUÍDO</th><td>", pmSubstituido, "</td></tr>",
      "<tr><th>PM SUBSTITUTO</th><td>",  pmSubstituto,  "</td></tr>",
      "<tr><th>Data do serviço permutado</th><td>", datasPermutaStr, "</td></tr>",
      "<tr><th>Data do pagamento do serviço</th><td>", datasPagtoStr, "</td></tr>",
      "<tr><th>Total de horas trabalhadas</th><td>", carga, "</td></tr>",
    "</table>",
    "<div class=\"assinaturas\">",
      "<div class=\"linhaAss\">",
        "<div>_________________________________________<br>Assinatura do PM Substituído</div>",
        "<div>_________________________________________<br>Assinatura do PM Substituto</div>",
      "</div>",
    "</div>",
    "<div class=\"rod\">",
      "________________________________________________<br>",
      "José Ribamar Braga Junior – 1º TEM. QOPM<br>",
      "Comandante da 2ª CIA/10ºBPM",
    "</div>",
    "</body></html>"
  ];
  const docHTML = parts.join("");

  const w = window.open("", "_blank");
  w.document.open();
  w.document.write(docHTML);
  w.document.close();
  w.focus();
  w.print();
}

// ====== Limpar tudo ======
function limparFormulario(){
  document.getElementById("permutaForm").reset();
  datasPermuta.length = 0;
  datasPagamento.length = 0;
  renderChips("chipsPermuta", datasPermuta);
  renderChips("chipsPagamento", datasPagamento);
}

// Eventos
document.getElementById("btnGerarDocx").addEventListener("click", gerarDOCX);
document.getElementById("btnGerarPDF").addEventListener("click", gerarPDF);
document.getElementById("btnLimparTudo").addEventListener("click", limparFormulario);
