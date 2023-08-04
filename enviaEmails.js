//Função para envia e-mail aviasando que o prazo de 15 dias foi atingido
function enviaEmail() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const colunaA = sheet.getRange("A:A"); // Coluna B
    const colunaC = sheet.getRange("C:C"); // Coluna D
  
    const ultimaLinha = sheet.getLastRow();
    const valoresColunaA = colunaA.getValues();
    const valoresColunaC = colunaC.getValues();
      
    const email = {
      to: "gcalixto.ctce@gmail.com",
      //to: "wgsilva@fazenda.rj.gov.br",
      subject: "Alerta de processos a muito tempo recebidos na corregedoria",
      htmlBody: ``,
      name: "assunto do Email"
    }
  
    // Prints 3 valuesC from the first column, starting from row 1.
    for (var i = 1; i < ultimaLinha; i++) {
      // Valor da célula na coluna A
      var celValA = valoresColunaA[i][0];   
  
      // Valor da célula na coluna C
      var celValC = valoresColunaC[i][0];
  
      var numerosEncontrados = celValC.match(/\d+/g);
      var numerosConcatenados = parseInt(numerosEncontrados ? numerosEncontrados.join('') : '');
  
      if (numerosConcatenados >= 15) {
        email.htmlBody = `
        <p>Olá Wendel! Segue a lista de processos</p>
        <br> ` + celValA;
        MailApp.sendEmail(email);
        Logger.log(email.htmlBody);
      }
    }
  }