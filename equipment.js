function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var sheet = ss.getSheetByName("formulario");
  var email = Session.getActiveUser().getEmail();

  var formulario = ss.getSheetByName("formulario")
  var baseDeDados = ss.getSheetByName("baseDeDados")
  
//var celula = sheet.getRange(1, 1)
//var celula = formulario.getRange(1, 1)

var NOME = formulario.getRange(2, 2).getValue()
var SERVICO = formulario.getRange(3, 2).getValue()
var PATGPON = formulario.getRange(4, 2).getValue()
var MACGPON = formulario.getRange(5, 2).getValue()
var PATROTEADOR = formulario.getRange(6, 2).getValue()
var MACROTEADOR = formulario.getRange(7, 2).getValue()
var UMFO = formulario.getRange(8, 2).getValue()
var CONAZ = formulario.getRange(9, 2).getValue()
var CONVD = formulario.getRange(10, 2).getValue()
var PTO = formulario.getRange(11, 2).getValue()
var ESTICA = formulario.getRange(12, 2).getValue()
var CORDAO = formulario.getRange(13, 2).getValue()
var PATTVBOX = formulario.getRange(14, 2).getValue()
var MACTVBOX = formulario.getRange(15, 2).getValue()
var email = Session.getActiveUser().getEmail()


var ultimaLinha = baseDeDados.getLastRow() 
var linhaSeguinte = ultimaLinha + 1

//Logger.log(Endereco)
baseDeDados.getRange(linhaSeguinte, 1).setValue(ultimaLinha)
baseDeDados.getRange(linhaSeguinte, 2).setValue(NOME)
baseDeDados.getRange(linhaSeguinte, 3).setValue(SERVICO)
baseDeDados.getRange(linhaSeguinte, 4).setValue(PATGPON)
baseDeDados.getRange(linhaSeguinte, 5).setValue(MACGPON)
baseDeDados.getRange(linhaSeguinte, 6).setValue(PATROTEADOR)
baseDeDados.getRange(linhaSeguinte, 7).setValue(MACROTEADOR)
baseDeDados.getRange(linhaSeguinte, 8).setValue(UMFO)
baseDeDados.getRange(linhaSeguinte, 9).setValue(CONAZ)
baseDeDados.getRange(linhaSeguinte, 10).setValue(CONVD)
baseDeDados.getRange(linhaSeguinte, 11).setValue(PTO)
baseDeDados.getRange(linhaSeguinte, 12).setValue(ESTICA)
baseDeDados.getRange(linhaSeguinte, 13).setValue(CORDAO)
baseDeDados.getRange(linhaSeguinte, 14).setValue(PATTVBOX)
baseDeDados.getRange(linhaSeguinte, 15).setValue(MACTVBOX)
baseDeDados.getRange(linhaSeguinte, 16).setValue(email)

formulario.getRange(1, 2).setValue('')
formulario.getRange(2, 2).setValue('')
formulario.getRange(3, 2).setValue('')
formulario.getRange(4, 2).setValue('')
formulario.getRange(5, 2).setValue('')
formulario.getRange(6, 2).setValue('')
formulario.getRange(7, 2).setValue('')
formulario.getRange(8, 2).setValue('')
formulario.getRange(9, 2).setValue('')
formulario.getRange(10, 2).setValue('')
formulario.getRange(11, 2).setValue('')
formulario.getRange(12, 2).setValue('')
formulario.getRange(13, 2).setValue('')
formulario.getRange(14, 2).setValue('')
formulario.getRange(15, 2).setValue('')
}

function leitura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var formulario = ss.getSheetByName("formulario")
  var baseDeDados = ss.getSheetByName("baseDeDados")

  var achaLinha = baseDeDados.getCurrentCell().getRow()
  
//Logger.log(achaLinha)

var id = baseDeDados.getRange(achaLinha, 1).getValue()
var NOME = baseDeDados.getRange(achaLinha, 2).getValue()
var SERVICO = baseDeDados.getRange(achaLinha, 3).getValue()
var PATGPON = baseDeDados.getRange(achaLinha, 4).getValue()
var MACGPON = baseDeDados.getRange(achaLinha, 5).getValue()
var PATROTEADOR = baseDeDados.getRange(achaLinha, 6).getValue()
var MACROTEADOR = baseDeDados.getRange(achaLinha, 7).getValue()
var UMFO = baseDeDados.getRange(achaLinha, 8).getValue()
var CONAZ = baseDeDados.getRange(achaLinha, 9).getValue()
var CONVD = baseDeDados.getRange(achaLinha, 10).getValue()
var PTO = baseDeDados.getRange(achaLinha, 11).getValue()
var ESTICA = baseDeDados.getRange(achaLinha, 12).getValue()
var CORDAO = baseDeDados.getRange(achaLinha, 13).getValue()
var PATTVBOX = baseDeDados.getRange(achaLinha, 14).getValue()
var MACTVBOX = baseDeDados.getRange(achaLinha, 15).getValue()
//var email = Session.getActiveUser().getEmail()


//var ultimalinha = baseDeDados.getLastRow() 
//var linhaseguinte = ultimalinha + 1

//Logger.log(Endereco)
formulario.getRange(1, 2).setValue(id)
formulario.getRange(2, 2).setValue(NOME)
formulario.getRange(3, 2).setValue(SERVICO)
formulario.getRange(4, 2).setValue(PATGPON)
formulario.getRange(5, 2).setValue(MACGPON)
formulario.getRange(6, 2).setValue(PATROTEADOR)
formulario.getRange(7, 2).setValue(MACROTEADOR)
formulario.getRange(8, 2).setValue(UMFO)
formulario.getRange(9, 2).setValue(CONAZ)
formulario.getRange(10, 2).setValue(CONVD)
formulario.getRange(11, 2).setValue(PTO)
formulario.getRange(12, 2).setValue(ESTICA)
formulario.getRange(13, 2).setValue(CORDAO)
formulario.getRange(14, 2).setValue(PATTVBOX)
formulario.getRange(15, 2).setValue(MACTVBOX)

formulario.activate()
}

function alterar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var formulario = ss.getSheetByName("formulario")
  var baseDeDados = ss.getSheetByName("baseDeDados")

  var id = formulario.getRange(1, 2).getValue()
  var NOME = formulario.getRange(2, 2).getValue()
  var SERVICO = formulario.getRange(3, 2).getValue()
  var PATGPON = formulario.getRange(4, 2).getValue()
  var MACGPON = formulario.getRange(5, 2).getValue()
  var PATROTEADOR = formulario.getRange(6, 2).getValue()
  var MACROTEADOR = formulario.getRange(7, 2).getValue()
  var UMFO = formulario.getRange(8, 2).getValue()
  var CONAZ = formulario.getRange(9, 2).getValue()
  var CONVD = formulario.getRange(10, 2).getValue()
  var PTO = formulario.getRange(11, 2).getValue()
  var ESTICA = formulario.getRange(12, 2).getValue()
  var CORDAO = formulario.getRange(13, 2).getValue()
  var PATTVBOX = formulario.getRange(14, 2).getValue()
  var MACTVBOX = formulario.getRange(15, 2).getValue()

  var repetidas = baseDeDados.getLastRow()

  for(var i = 1; i < repetidas; i++){
    //Logger.log('Dante')
    if(baseDeDados.getRange(i, 1).getValue() == id){
      //Logger.log(i)
      //baseDeDados.getRange(i, 1).setValue(ultimaLinha)
      baseDeDados.getRange(i, 2).setValue(NOME)
      baseDeDados.getRange(i, 3).setValue(SERVICO)
      baseDeDados.getRange(i, 4).setValue(PATGPON)
      baseDeDados.getRange(i, 5).setValue(MACGPON)
      baseDeDados.getRange(i, 6).setValue(PATROTEADOR)
      baseDeDados.getRange(i, 7).setValue(MACROTEADOR)
      baseDeDados.getRange(i, 8).setValue(UMFO)
      baseDeDados.getRange(i, 9).setValue(CONAZ)
      baseDeDados.getRange(i, 10).setValue(CONVD)
      baseDeDados.getRange(i, 11).setValue(PTO)
      baseDeDados.getRange(i, 12).setValue(ESTICA)
      baseDeDados.getRange(i, 13).setValue(CORDAO)
      baseDeDados.getRange(i, 14).setValue(PATTVBOX)
      baseDeDados.getRange(i, 15).setValue(MACTVBOX)
      //baseDeDados.getRange(i, 16).setValue(email)
    }
  }
}
/**
function treinoDeLoop() {
  for(var i = 0; i < 10; i++){
    Logger.log("Dante")
  }
}
*/
/**
  for(var i = 1; i < repetidas; i++){
    //Logger.log('Dante')
    Logger.log(baseDeDados.getRange(i + 1, 1).getValue())
  }
*/

