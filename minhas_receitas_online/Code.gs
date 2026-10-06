const NOME_ABA = "Receitas";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Minhas Receitas")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOME_ABA);
  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA);
    sheet.appendRow(["ID", "Nome", "Categoria", "Ingredientes", "Modo de preparo", "Criada em"]);
  }
  return sheet;
}

function listarReceitas() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(r => ({
    id: String(r[0]),
    nome: String(r[1]),
    categoria: String(r[2]),
    ingredientes: String(r[3]),
    modo: String(r[4])
  }));
}

function adicionarReceita(r) {
  if (!r || !r.nome || !r.ingredientes || !r.modo) {
    throw new Error("Preencha nome, ingredientes e modo de preparo.");
  }
  const sheet = getSheet_();
  sheet.appendRow([
    Utilities.getUuid(),
    r.nome,
    r.categoria || "Outros",
    r.ingredientes,
    r.modo,
    new Date()
  ]);
  return true;
}

function excluirReceita(id) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function editarReceita(r) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(r.id)) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[
        r.nome, r.categoria || "Outros", r.ingredientes, r.modo
      ]]);
      return true;
    }
  }
  throw new Error("Receita não encontrada.");
}
