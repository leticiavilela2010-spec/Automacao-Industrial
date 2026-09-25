/**
 * Script do Google Apps Script para o "Ranking da turma" do Jogo de Digitação.
 *
 * COMO USAR:
 * 1. Crie uma planilha nova em sheets.google.com.
 * 2. Menu Extensões > Apps Script.
 * 3. Apague o conteúdo padrão (Code.gs) e cole este arquivo inteiro no lugar.
 * 4. Salve o projeto (ex: nome "Ranking Digitação").
 * 5. Clique em Implantar > Nova implantação.
 *    - Tipo: "App da Web".
 *    - Executar como: "Eu".
 *    - Quem pode acessar: "Qualquer pessoa".
 * 6. Clique em Implantar e autorize o acesso (é o seu próprio script, pode aceitar).
 * 7. Copie a URL do app da Web (termina em /exec).
 * 8. Cole essa URL na constante SHEET_WEBAPP_URL no arquivo index.html do jogo.
 *
 * Sempre que publicar uma alteração neste script, crie uma NOVA implantação
 * (ou "Gerenciar implantações > editar > Nova versão") para o link /exec atualizar.
 */

var SHEET_NAME = "Ranking";

// Opcional: defina uma palavra-chave aqui e no index.html (mesma variável) para
// dificultar que alguém de fora envie pontuações falsas para a planilha.
// Deixe como "" para não exigir chave nenhuma (mais simples de configurar).
var SHARED_SECRET = "";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (SHARED_SECRET && data.secret !== SHARED_SECRET) {
      return jsonOutput_({ ok: false, error: "unauthorized" });
    }

    var sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      String(data.name || "Anônimo").slice(0, 60),
      Number(data.score) || 0,
      Number(data.wpm) || 0,
      Number(data.accuracy) || 0,
      Number(data.mistakes) || 0
    ]);

    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var rows = readRows_();
  var json = JSON.stringify(rows);
  var callback = e.parameter && e.parameter.callback;

  if (callback) {
    // JSONP: permite o navegador ler os dados sem bloqueio de CORS.
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(rows);
}

function readRows_() {
  var values = getSheet_().getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) { // pula o cabeçalho
    if (values[i][1] === "" && values[i][2] === "") continue; // ignora linhas vazias
    rows.push({
      date: values[i][0],
      name: values[i][1],
      score: values[i][2],
      wpm: values[i][3],
      accuracy: values[i][4],
      mistakes: values[i][5]
    });
  }
  return rows;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Data", "Nome", "Pontos", "PPM", "Precisão", "Erros"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
