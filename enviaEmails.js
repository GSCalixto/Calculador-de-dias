//Função para envia e-mail aviasando que o prazo de 15 dias foi atingido
function enviaEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const anexo = DriveApp.getFileById("1lL9gxTV0V3xHu7yJ5F4CmpQAn8XDVu4yXXWCm7mGvQk");

  const ultimaLinha = sheet.getLastRow();

  const colunaA = sheet.getRange(2, 1, ultimaLinha, 1); // Coluna A
  const colunaB = sheet.getRange(2, 2, ultimaLinha, 1); // Coluna B 
  const colunaC = sheet.getRange(2, 3, ultimaLinha, 1); // Coluna C

  
  const valoresColunaA = colunaA.getValues();
  const valoresColunaB = colunaB.getValues();
  const valoresColunaC = colunaC.getValues();
    
  // Prints 3 valuesC from the first column, starting from row 1.
  for (var i = 1; i < ultimaLinha; i++) {
    // Valor da célula na coluna A
    var celValA = valoresColunaA[i][0];
    
    // Valor da célula na coluna B
    var celValB = valoresColunaB[i][0].toString();

    // Valor da célula na coluna C
    var celValC = valoresColunaC[i][0];

    //Extrai o número dentro da variável celValC
    var numerosEncontrados = celValC.match(/\d+/g);
    var numerosConcatenados = parseInt(numerosEncontrados ? numerosEncontrados.join('') : '');

    if (numerosConcatenados >= 15) {
      var numProcessos = [];
      numProcessos.push(celValA);

      var dataRecebimento = [];
      dataRecebimento.push(formataData(celValB));

      for ( j in numProcessos){
        var email = {
          to: "gcalixto.ctce@gmail.com",
          //to: "wgsilva@fazenda.rj.gov.br",
          subject: "Alerta de processos",
          htmlBody: `<table border="1px" cellpadding="5px" style="border-collapse:collapse;border-color:#666">
                        <tbody>
                          <tr>
                            <th>Numero do Processo</th>
                            <th>Data de recebimento</th>
                          </tr>
                          <tr>
                            <td> `+numProcessos[j]+`</td>
                            <td> `+dataRecebimento[j]+`</td>
                          </tr>
                        </tbody>
                    </table>`,
          name: "Processos recebidos na corregedoria a mais de 15 dias",
          attachments: [anexo]
        };       
        MailApp.sendEmail(email);
      };
    };
  };
};

  //Formata a informação de data para o padrão DD/MM/AAAA usando ReGex
function formataData(data) {
  const regexNumData = [...data.matchAll(/[0-9]{1,4}/g)];
  const regexStringData = /Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/;

  var dia = regexNumData;
  var mes = data.match(regexStringData);
  var ano = regexNumData;

  var stringFormatada = dia[0][0] + "/" + mes + "/" + ano[1][0];

  return stringFormatada;
};