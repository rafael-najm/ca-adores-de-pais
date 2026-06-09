// ============================================================
// CAÇADORES DE PAIS — Apps Script API v3
// ============================================================
// IMPORTANTE: Após colar esse código:
// Implantar > Gerenciar implantações > editar > NOVA VERSÃO > salvar

const SHEET_NAME = "Planilha1";
const SPREADSHEET_ID = "1c2qtOwLuZuQhOO8Ot5OGsZXaNx4i7GcBVqCwq_DyhQk";

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "getAll";
  const callback = e.parameter && e.parameter.callback;
  let result;
  if (action === "getAll") result = getAllData();
  else result = { error: "Ação inválida" };
  return buildResponse(result, callback);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  let result;
  if      (action === "update") result = updateRow(body);
  else if (action === "add")    result = addRow(body);
  else if (action === "delete") result = deleteRow(body.id);
  else result = { error: "Ação inválida" };
  return buildResponse(result, null);
}

function buildResponse(data, callback) {
  const json = JSON.stringify(data);
  const output = callback
    ? ContentService.createTextOutput(`${callback}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function getAllData() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    result.push({
      id:       i,
      nome:     String(row[0] || ""),
      turma:    String(row[1] || ""),
      status:   String(row[2] || "Não contatado"),
      cacador:  String(row[3] || ""),
      telefone: String(row[4] || ""),
      obs:      String(row[5] || ""),
      updated:  String(row[6] || "")
    });
  }
  return result;
}

function updateRow(body) {
  const sheet = getSheet();
  const row = Number(body.id) + 1;
  sheet.getRange(row, 3).setValue(body.status   || "");
  sheet.getRange(row, 4).setValue(body.cacador  || "");
  sheet.getRange(row, 5).setValue(body.telefone || "");
  sheet.getRange(row, 6).setValue(body.obs      || "");
  sheet.getRange(row, 7).setValue(new Date().toLocaleString("pt-BR"));
  return { ok: true };
}

function addRow(body) {
  const sheet = getSheet();
  const last = sheet.getLastRow() + 1;
  sheet.getRange(last, 1).setValue(body.nome     || "");
  sheet.getRange(last, 2).setValue(body.turma    || "");
  sheet.getRange(last, 3).setValue(body.status   || "Não contatado");
  sheet.getRange(last, 4).setValue(body.cacador  || "");
  sheet.getRange(last, 5).setValue(body.telefone || "");
  sheet.getRange(last, 6).setValue(body.obs      || "");
  sheet.getRange(last, 7).setValue(new Date().toLocaleString("pt-BR"));
  return { ok: true, id: last - 1 };
}

function deleteRow(id) {
  const sheet = getSheet();
  sheet.deleteRow(Number(id) + 1);
  return { ok: true };
}
