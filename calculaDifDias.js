function calculaDifDias(cellValor){
    var data = new Date(cellValor);

    //calcula a diferença em dias
    var hoje = new Date();
    var diffEmMilissegundos = hoje.getTime() - data.getTime();
    var diffEmDias = Math.floor(diffEmMilissegundos / (1000 * 60 * 60 * 24));
    return diffEmDias;
}

//Precisa implementar a mesma coisa para a coluna C pegando as datas da coluna B
function percorrerColuna() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colunaD = sheet.getRange("D:D"); // Coluna D

  var lastRow = sheet.getLastRow();
  var valoresColunaD = colunaD.getValues();

  for (var i = 1; i < lastRow; i++) {
    // Valor da célula na coluna D
    var celValD = valoresColunaD[i][0]; 

    //calcula diferença de dias
    var diffCalculadaE = calculaDifDias(celValD);
    //var diffCalculadaC = calculaDifDias(celValC);

    // Fazer as alterações desejadas
    var modifiedValue = diffCalculadaE + ' dias';

    // Atualizar o valor na célula da coluna E
    sheet.getRange("E" + (i+1)).setValue(modifiedValue);

  }
  
  Logger.log('Alterações concluídas.');
}