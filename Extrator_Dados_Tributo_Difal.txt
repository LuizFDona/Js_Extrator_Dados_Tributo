function extrairTextoPDF() {
  var folderId = "1xMtAkskm1-9OIBJQItnCQwpN5dXTf-CY"; // Substitua pelo ID da pasta
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType("application/pdf");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("BD_DIFAL");

  sheet.getRange('A:D').clearContent(); // Limpa todo o conteúdo da aba BD_DIFAL
  
  var lastRow = 1; // Começa na linha 2
  
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    
    var docFile = converterPDFparaDoc(fileId);
    
    if (docFile) {
      var docText = extrairTextoDoDoc(docFile.getId());
      
      if (docText) {
        var fileName = file.getName();
        var extractedText = extrairTextoDoNomeDoArquivo(fileName); // Extrair texto do nome do arquivo
        extractedText = extractedText.replace("NF", ""); // Remover "NF" da variável
        
        sheet.getRange(lastRow, 1).setValue(fileName);
        sheet.getRange(lastRow, 3).setValue(docText);
        sheet.getRange(lastRow, 2).setValue(extractedText); // Adicionar o texto extraído do nome do arquivo
        
        // Extrair os caracteres e realizar a conversão e divisão
        var extractedChars = docText.substring(8, 11) + docText.substring(12, 16);
        var extractedNumber = parseInt(extractedChars);
        var finalValue = extractedNumber / 100;
        
        sheet.getRange(lastRow, 4).setValue(finalValue); // Adicionar o valor calculado
        
        /*Testando nova coluna 
        var value = lastRow.toString();
        var formattedValue = Utilities.formatString("%03d", value);
        sheet.getRange(lastRow, 5).setValue(formattedValue);
        */
        lastRow++;
        
        Logger.log("Texto extraído do documento '" + fileName + "': " + docText);
        
        // Excluir o arquivo .gdoc convertido
        DriveApp.getFileById(docFile.getId()).setTrashed(true);
      } else {
        Logger.log("Falha ao extrair texto do documento");
      }
    } else {
      Logger.log("Falha ao converter PDF para documento");
    }
  }
}

function extrairTextoDoNomeDoArquivo(fileName) {
  var delimiters = /-/; // Use um delimitador que separa os elementos
  var parts = fileName.split(delimiters);

  if (parts.length >= 3) {
    return parts[1]; // O texto entre os delimitadores é o segundo elemento
  } else {
    return "";
  }
}

function converterPDFparaDoc(fileId) {
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/copy";
  
  var payload = {
    mimeType: "application/vnd.google-apps.document"
  };
  
  var headers = {
    Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    "Content-Type": "application/json"
  };
  
  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload)
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    Logger.log("Erro ao converter PDF para documento: " + result.error.message);
    return null;
  }
  
  var docId = result.id;
  var docFile = DocumentApp.openById(docId);
  
  return docFile;
}

function extrairTextoDoDoc(docId) {
  var docFile = DocumentApp.openById(docId);
  var docText = docFile.getBody().getText();
  var docText = docText.replace(/\-/g, "");
  var docText = docText.replace(/ /g, "");
  
var regex = /\d{47,48}/g;
  var matches = docText.match(regex);
  
  var result = "";
  if (matches) {
    result = matches[0];
  }
  
  return result;
}