/**
 * TransformToGoogleSheet.js
 *
 * Script para transformar a planilha GCACL_BAAPLOG_v_teste.xlsx em Google Sheets.
 * O arquivo já está no Google Drive.
 *
 * INSTRUÇÕES:
 * 1. Acesse script.google.com e crie um Novo Projeto
 * 2. Cole este código inteiro no arquivo Codigo.gs
 * 3. Selecione a função transformSpreadsheet e clique em Executar
 * 4. Autorize as permissões solicitadas
 */

// ============================================================
// CONFIGURAÇÃO
// ============================================================
const CONFIG = {
  // ID do arquivo no Google Drive (extraído da URL de compartilhamento)
  SOURCE_FILE_ID: "1Y7iffPIF-fWgst5Wa9EhAV_MioQ0j8cR",

  // Nome do arquivo de origem
  SOURCE_FILE_NAME: "GCACL_BAAPLOG_v_teste",

  // Pasta de destino no Google Drive
  DESTINATION_FOLDER_NAME: "GCACL_BAAPLOG_GoogleSheets",

  // Prefixo do nome do Google Sheet gerado
  OUTPUT_PREFIX: "GSheet_",

  // Formato de data/hora para o nome do arquivo de saída
  TIMESTAMP_FORMAT: "yyyyMMdd_HHmmss"
};

// ============================================================
// FUNÇÃO PRINCIPAL - Execute esta função
// ============================================================
function transformSpreadsheet() {
  Logger.log("=== Iniciando transformação da planilha ===");

  // 1. Abrir o arquivo diretamente pelo ID
  let sourceFile;
  try {
    sourceFile = DriveApp.getFileById(CONFIG.SOURCE_FILE_ID);
  } catch (e) {
    throw new Error(
      "ERRO: Não foi possível acessar o arquivo com ID: " + CONFIG.SOURCE_FILE_ID + "\n" +
      "Verifique se você tem permissão de acesso ao arquivo.\n" +
      "Erro: " + e.message
    );
  }

  Logger.log("Arquivo encontrado: " + sourceFile.getName());
  Logger.log("Tipo MIME: " + sourceFile.getMimeType());
  Logger.log("Tamanho: " + formatFileSize_(sourceFile.getSize()));

  // 2. Criar pasta de destino
  const destFolder = getOrCreateFolder_(CONFIG.DESTINATION_FOLDER_NAME);
  Logger.log("Pasta de destino: " + destFolder.getName());

  // 3. Converter para Google Sheets
  const googleSheet = convertToGoogleSheet_(sourceFile, destFolder);
  Logger.log("Google Sheet criado: " + googleSheet.getName());
  Logger.log("URL: " + googleSheet.getUrl());

  // 4. Formatar o Google Sheet
  formatGoogleSheet_(googleSheet);

  // 5. Resumo final
  const summary = {
    arquivoOrigem: sourceFile.getName(),
    googleSheet: googleSheet.getName(),
    url: googleSheet.getUrl(),
    pastaDestino: destFolder.getName(),
    dataConversao: new Date().toLocaleString("pt-BR")
  };

  Logger.log("\n=== CONVERSÃO CONCLUÍDA COM SUCESSO ===");
  Logger.log(JSON.stringify(summary, null, 2));

  return summary;
}

// ============================================================
// FUNÇÕES AUXILIARES
// ============================================================

/**
 * Obtém ou cria uma pasta no Google Drive.
 */
function getOrCreateFolder_(folderName) {
  if (!folderName) {
    return DriveApp.getRootFolder();
  }

  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }

  Logger.log("Criando pasta: " + folderName);
  return DriveApp.createFolder(folderName);
}

/**
 * Converte o arquivo para Google Sheets.
 */
function convertToGoogleSheet_(sourceFile, destFolder) {
  const mimeType = sourceFile.getMimeType();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    CONFIG.TIMESTAMP_FORMAT
  );
  const outputName = CONFIG.OUTPUT_PREFIX + CONFIG.SOURCE_FILE_NAME + "_" + timestamp;

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    Logger.log("Arquivo já é um Google Sheet. Criando cópia...");
    const copy = sourceFile.makeCopy(outputName, destFolder);
    return SpreadsheetApp.openById(copy.getId());
  }

  if (mimeType === MimeType.CSV || sourceFile.getName().endsWith(".csv")) {
    Logger.log("Importando arquivo CSV...");
    return importCsvToSheet_(sourceFile, outputName, destFolder);
  }

  // Excel (.xlsx, .xls) ou ODS
  Logger.log("Convertendo " + mimeType + " para Google Sheets...");
  return convertExcelToSheet_(sourceFile, outputName, destFolder);
}

/**
 * Converte arquivo Excel/ODS para Google Sheets via Drive API.
 */
function convertExcelToSheet_(sourceFile, outputName, destFolder) {
  const blob = sourceFile.getBlob();

  // Método 1: Drive API v2 (Advanced Drive Service)
  try {
    const resource = {
      title: outputName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: destFolder.getId() }]
    };
    const file = Drive.Files.insert(resource, blob, { convert: true });
    return SpreadsheetApp.openById(file.id);
  } catch (e) {
    Logger.log("Drive API v2 não disponível. Usando método REST...");
  }

  // Método 2: Drive API v3 via REST
  try {
    const accessToken = ScriptApp.getOAuthToken();
    const boundary = "-------" + Utilities.getUuid();

    const metadata = JSON.stringify({
      name: outputName,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [destFolder.getId()]
    });

    const requestBody =
      "--" + boundary + "\r\n" +
      "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
      metadata + "\r\n" +
      "--" + boundary + "\r\n" +
      "Content-Type: " + sourceFile.getMimeType() + "\r\n\r\n";

    const postData = Utilities.newBlob(requestBody).getBytes()
      .concat(blob.getBytes())
      .concat(Utilities.newBlob("\r\n--" + boundary + "--").getBytes());

    const response = UrlFetchApp.fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
      {
        method: "post",
        contentType: "multipart/related; boundary=" + boundary,
        payload: postData,
        headers: { Authorization: "Bearer " + accessToken },
        muteHttpExceptions: true
      }
    );

    const result = JSON.parse(response.getContentText());
    if (result.error) {
      throw new Error(result.error.message);
    }
    return SpreadsheetApp.openById(result.id);
  } catch (e2) {
    Logger.log("Erro no método REST: " + e2.message);
  }

  // Método 3: Cópia direta
  const tempFile = sourceFile.makeCopy(outputName + "_temp");
  try {
    const ss = SpreadsheetApp.openById(tempFile.getId());
    tempFile.moveTo(destFolder);
    tempFile.setName(outputName);
    return ss;
  } catch (e3) {
    tempFile.setTrashed(true);
    throw new Error(
      "Não foi possível converter o arquivo. Verifique se:\n" +
      "1. O arquivo é um Excel válido (.xlsx ou .xls)\n" +
      "2. O serviço 'Drive API' está habilitado no projeto\n" +
      "Erro: " + e3.message
    );
  }
}

/**
 * Importa um arquivo CSV para um novo Google Sheet.
 */
function importCsvToSheet_(csvFile, outputName, destFolder) {
  const csvContent = csvFile.getBlob().getDataAsString("UTF-8");
  const csvData = Utilities.parseCsv(csvContent);

  const spreadsheet = SpreadsheetApp.create(outputName);
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName(CONFIG.SOURCE_FILE_NAME);

  if (csvData.length > 0) {
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  }

  DriveApp.getFileById(spreadsheet.getId()).moveTo(destFolder);
  return spreadsheet;
}

/**
 * Aplica formatação ao Google Sheet convertido.
 */
function formatGoogleSheet_(spreadsheet) {
  Logger.log("Aplicando formatação...");

  const sheets = spreadsheet.getSheets();

  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) continue;

    // Cabeçalho: negrito, fundo azul, texto branco, centralizado
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285F4");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setHorizontalAlignment("center");

    // Congelar primeira linha
    sheet.setFrozenRows(1);

    // Auto-redimensionar colunas
    for (let col = 1; col <= lastCol; col++) {
      try { sheet.autoResizeColumn(col); } catch (e) { }
    }

    // Adicionar filtro
    if (lastRow > 1) {
      try { sheet.getRange(1, 1, lastRow, lastCol).createFilter(); } catch (e) { }
    }

    // Zebrado (linhas pares em cinza claro)
    if (lastRow > 1) {
      for (let row = 2; row <= lastRow; row++) {
        if (row % 2 === 0) {
          sheet.getRange(row, 1, 1, lastCol).setBackground("#F8F9FA");
        }
      }
    }

    Logger.log("Aba '" + sheet.getName() + "': " + lastRow + " linhas x " + lastCol + " colunas");
  }

  Logger.log("Formatação concluída.");
}

/**
 * Formata o tamanho do arquivo de forma legível.
 */
function formatFileSize_(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(2) + " KB";
  return (bytes / 1048576).toFixed(2) + " MB";
}
