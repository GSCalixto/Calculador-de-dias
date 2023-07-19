//Precisa implementar a mesma coisa para a coluna C pegando as datas da coluna B
function percorrerColuna() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colunaB = sheet.getRange("B:B"); // Coluna B
  var colunaD = sheet.getRange("D:D"); // Coluna D

  var lastRow = sheet.getLastRow();
  var valoresColunaB = colunaB.getValues();
  var valoresColunaD = colunaD.getValues();

  for (var i = 1; i < lastRow; i++) {
    // Valor da célula na coluna B
    var celValB = valoresColunaB[i][0];   

    // Valor da célula na coluna D
    var celValD = valoresColunaD[i][0]; 

    // Verifica se a célula na coluna B está preenchida e atualiza o valor na célula da coluna C se sim
    if (celValB !== ""){
      alteraCelula(i, celValB,"C", sheet)
    }
    // Verifica se a célula na coluna D está preenchida e atualiza o valor na célula da coluna E se sim
    if (celValD !== ""){
      alteraCelula(i, celValD,"E", sheet)
    }
  }
}

//Função que insere o calculo de diferença de dias em uma célula
function alteraCelula(i, valorCelula, coluna = String, sheet) {
  var diffCalculada = calculaDifDias(valorCelula);  
  var stringDias = diffCalculada + ' dias';

  sheet.getRange(coluna + (i+1)).setValue(stringDias);
}

//Função que calcula a diferença de dias entra o valor, em formato de data, de uma célula e a data atual 
function calculaDifDias(cellValor){
  var data = new Date(cellValor);

  //calcula a diferença em dias
  var hoje = new Date();
  var diffEmMilissegundos = hoje.getTime() - data.getTime();
  var diffEmDias = Math.floor(diffEmMilissegundos / (1000 * 60 * 60 * 24));

  return diffEmDias;
}