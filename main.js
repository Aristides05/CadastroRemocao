var planilha = SpreadsheetApp.getActiveSpreadsheet();
var sheetMain = planilha.getSheetByName("MAIN");
var sheetAmazon = planilha.getSheetByName("AMAZON");
var sheetMercadoLivre = planilha.getSheetByName("MERCADO LIVRE");
var sheetDatabase = planilha.getSheetByName("DB");
var sheetConsulta = planilha.getSheetByName("CONSULTA");
const packageID = 8;
const quantidadeCaracteresParaCodigoRastreioEmbalagemMercadoLivre = 13;
const timeDelay = 35;
const cellAmazon = "G14:J14";

function callSheetAmz() {
    callGenericSheet(sheetAmazon, "G14");
}

function callSheetML() {
    callGenericSheet(sheetMercadoLivre, "F13");
}

function callSheetConsulta() {
    callGenericSheet(sheetConsulta, "B6");
}

function callGenericSheet(sheet, cell) {
    sheet.showSheet();
    planilha.setActiveSheet(sheet);
    Utilities.sleep(timeDelay);
    sheetMain.hideSheet();
    let range = sheet.getRange(cell); 
    sheet.setCurrentCell(range);
}

function voltar() {
    let oldActiveSheet = planilha.getActiveSheet();
    callSheetMain();
    oldActiveSheet.hideSheet();
}

function callSheetMain() {
    sheetMain.showSheet();
    planilha.setActiveSheet(sheetMain);
    Utilities.sleep(timeDelay);
}

function getDateTime() {
    let objDate = new Date();
    let day = objDate.getDate();
    let month = objDate.getMonth() + 1;
    let year = objDate.getFullYear();
    let hours = objDate.getHours();
    let minutes = objDate.getMinutes();

    if (minutes < 10) { // validar
        minutes = "0" + minutes;
    }

    let fullDateTime = day + '/' + month + '/' + year + ' ' + hours + ':' + minutes;
    return fullDateTime;
}

function cadastrarAmazon() {
    let chaveConfirmacaoCell = sheetAmazon.getRange("G17:J17");
    let chaveCell = sheetAmazon.getRange("G14:J14");
    let chaveValue = chaveCell.getValue();

    if (!isAmazonKeyValid(chaveCell, chaveConfirmacaoCell)) {
        alertValorInvalido('Chave de Acesso');
        return;
    }

    var row = sheetDatabase.getLastRow() + 1;
    sheetDatabase.getRange(row, 1).setValue(getDateTime())
    sheetDatabase.getRange(row, 2).setValue(chaveValue);

    chaveCell.clearContent();
    chaveConfirmacaoCell.clearContent();

    callSheetMain();
    sheetAmazon.hideSheet();
}

function isAmazonKeyValid(chave, chaveConfirmacao) {
    let chaveValue = chave.getValue();
    let chaveConfirmacaoValue = chaveConfirmacao.getValue();
    return chaveValue == chaveConfirmacaoValue && chaveValue.length == 44;
}

function cadastraEmbalagem() {
    let cellNome = sheetMercadoLivre.getRange("F16:I16");
    let cellCodigo = sheetMercadoLivre.getRange("F19:G19");
    let cellPackageId = sheetMercadoLivre.getRange("F13:I13");

    if (!isDadosEmbalagemCorretos(cellNome, cellCodigo, cellPackageId)) {
        return;
    }

    var row = sheetDatabase.getLastRow() + 1;
    sheetDatabase.getRange(row, 1).setValue(getDateTime())
    sheetDatabase.getRange(row, 2).setValue(cellPackageId.getValue());
    sheetDatabase.getRange(row, 3).setValue(cellNome.getValue());
    sheetDatabase.getRange(row, 4).setValue(cellCodigo.getValue());

    cellNome.clearContent();
    cellCodigo.clearContent();
    cellPackageId.clearContent();

    callSheetMain();
    sheetMercadoLivre.hideSheet();
}

function isDadosEmbalagemCorretos(cellNome, cellCodigo, cellPackageId) {
    let isCorreios = sheetMercadoLivre.getRange("G24").getValue();

    if (isCorreios) {
      return isCellFilled(cellNome, 'o nome do remetente')  && isCellFilled(cellCodigo, 'o código de rastreio') && isCellValueValid(cellCodigo, 'Código de rastreio')
    }

    return isCellFilled(cellPackageId, 'Package ID') && isCellValueValid(cellPackageId, 'PackageID')

}

function isCellFilled(cell, text) {
    if (isCellValueEmpty(cell)) {
        alertInsiraValor(text)
        return false;
    }

    return true;
}

function isCellValueValid(cell, text) {
    if (isCellLengthInvalid(cell)) {
        alertValorInvalido(text);
        return false;
    }

    return true;
}

function isCellValueEmpty(cell) {
    return cell.getValue() == "";
}

function isCellLengthInvalid(cell) {
    return cell.getValue().length != quantidadeCaracteresParaCodigoRastreioEmbalagemMercadoLivre;
}

function alertInsiraValor(texto) {
    SpreadsheetApp.getUi().alert('Insira:\t' + texto);
}

function alertValorInvalido(texto) {
    SpreadsheetApp.getUi().alert(texto + '\tInválido(a).');
}

function validaUmaOpcaoCheckBox(){
  if(SpreadsheetApp.getActiveSheet().getName() != "MERCADO LIVRE"){
    return;
  }
  
  const checkPadrao = "G24";
  const checkMercadoLivre = "H24";

  let cell = SpreadsheetApp.getCurrentCell().getA1Notation();

  if(cell == checkPadrao) {setValueCheckBox(checkPadrao, checkMercadoLivre, true, false); return;}
  setValueCheckBox(checkPadrao, checkMercadoLivre, false, true) 
}

function setValueCheckBox(checkPadrao, checkMercadoLivre, value_1, value_2){
  sheetMercadoLivre.getRange(checkPadrao).setValue(value_1);
  sheetMercadoLivre.getRange(checkMercadoLivre).setValue(value_2);
}


function unlockCell(){
  let range = sheetAmazon.getRange("G14:J14");
  let protection = sheetAmazon.protect();
  protection.removeRange(range);
}

function setActiveCell(sheet, cell){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var cell = sheet.getRange('G14');
sheet.setCurrentCell(cell);
}

