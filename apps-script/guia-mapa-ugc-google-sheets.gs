/**
 * OUSADIA MARKETING · Backend do Guia do Mapa do UGC (Apps Script)
 * -------------------------------------------------------------------
 * Guarda cada cadastro do guia (nome, e-mail, Instagram e WhatsApp) na aba
 * "Acessos" de uma planilha própria. O guia manda o cadastro com sendBeacon.
 *
 * COMO USAR:
 * 1. No Google Drive, crie uma planilha em branco, por exemplo
 *    "Leads Guia Mapa do UGC".
 * 2. Na planilha: Extensões > Apps Script. Apague tudo e cole este arquivo.
 *    Salve. (O script fica preso à planilha, então SHEET_ID pode ficar vazio.)
 * 3. Implantar > Nova implantação > tipo "App da Web".
 *    - Executar como: Eu.
 *    - Quem tem acesso: Qualquer pessoa.
 * 4. Autorize o acesso quando o Google pedir.
 * 5. Abra a URL gerada no navegador: deve responder "Guia do Mapa do UGC online".
 *    Se pedir login do Google, a implantação não ficou pública: repita o
 *    passo 3 com "Qualquer pessoa".
 * 6. Cole a URL (termina em `/exec`) em `ENDPOINT`, no começo do <script> do
 *    index.html, e publique de novo.
 *
 * Se mudar este código depois: Implantar > Gerenciar implantações > lápis >
 * Versão: "Nova versão". Sem isso a URL continua rodando o código antigo.
 *
 * SEGURANÇA: a URL é pública (precisa ser, pro guia gravar). Por isso o script
 * corta cada campo num tamanho máximo, tira = + - @ do começo (viraria
 * fórmula no Sheets), exige e-mail válido e consentimento, e limita os avisos
 * por e-mail a AVISOS_POR_HORA.
 */

// Vazio = a planilha onde este script foi colado.
var SHEET_ID = "";
var ABA = "Acessos";

// E-mail que recebe aviso de cada cadastro. Vazio = sem aviso.
var EMAIL_AVISO = "ousadiamkt@gmail.com";
var AVISOS_POR_HORA = 20;

var COLUNAS = ["Data", "Nome", "E-mail", "Instagram", "WhatsApp", "Origem", "Consentimento"];
var MAX = { nome: 120, email: 160, instagram: 60, whatsapp: 20, origem: 60 };

function planilha_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function aba_() {
  var ss = planilha_();
  var sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLUNAS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight("bold");
  }
  return sh;
}

function limpa_(v, max) {
  var s = String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim();
  s = s.replace(/^[=+\-@]+/, "");
  return s.slice(0, max);
}

function doGet() {
  return ContentService.createTextOutput("Guia do Mapa do UGC online");
}

function doPost(e) {
  try {
    var corpo = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var email = limpa_(corpo.email, MAX.email);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || corpo.consentimento !== true) {
      return ContentService.createTextOutput("ignorado");
    }
    var linha = [
      new Date(),
      limpa_(corpo.nome, MAX.nome),
      email,
      limpa_(corpo.instagram, MAX.instagram),
      "'" + limpa_(corpo.whatsapp, MAX.whatsapp),
      limpa_(corpo.origem, MAX.origem),
      "sim"
    ];
    aba_().appendRow(linha);
    avisa_(linha);
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("erro");
  }
}

function avisa_(linha) {
  if (!EMAIL_AVISO) return;
  var props = PropertiesService.getScriptProperties();
  var hora = Utilities.formatDate(new Date(), "America/Belem", "yyyyMMddHH");
  var chave = "avisos_" + hora;
  var n = Number(props.getProperty(chave) || 0);
  if (n >= AVISOS_POR_HORA) return;
  props.setProperty(chave, String(n + 1));
  MailApp.sendEmail(
    EMAIL_AVISO,
    "Novo acesso ao Guia do Mapa do UGC",
    "Nome: " + linha[1] + "\nE-mail: " + linha[2] + "\nInstagram: " + linha[3] + "\nWhatsApp: " + linha[4]
  );
}
