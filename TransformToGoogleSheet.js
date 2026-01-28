/**
 * TransformToGoogleSheet.js
 *
 * Script para transformar a planilha GCACL_BAAPLOG_v_teste (Excel/CSV)
 * em Google Sheets no Google Drive.
 *
 * INSTRUÇÕES DE USO:
 * 1. Faça upload da planilha GCACL_BAAPLOG_v_teste.xlsx para o Google Drive
 * 2. Copie este script para um projeto Google Apps Script (script.google.com)
 * 3. Execute a função transformSpreadsheet()
 * 4. Autorize as permissões solicitadas
 * 5. A planilha será convertida automaticamente para Google Sheets
 */

// ============================================================
// CONFIGURAÇÃO - Altere conforme necessário
// ============================================================
const CONFIG = {
  // Nome do arquivo de origem (sem extensão)
  SOURCE_FILE_NAME: "GCACL_BAAPLOG_v_teste",

  // Pasta de destino no Google Drive (deixe vazio para a raiz)
  DESTINATION_FOLDER_NAME: "GCACL_BAAPLOG_GoogleSheets",

  // Prefixo do nome do Google Sheet gerado
  OUTPUT_PREFIX: "GSheet_",

  // Extensões suportadas para busca
  SUPPORTED_EXTENSIONS: [".xlsx", ".xls", ".csv", ".ods"],

  // Formato de data/hora para o nome do arquivo de saída
  TIMESTAMP_FORMAT: "yyyyMMdd_HHmmss"
};

// ============================================================
// FUNÇÃO PRINCIPAL
// ============================================================

/**
 * Função principal que transforma a planilha em Google Sheets.
 * Execute esta função para iniciar a conversão.
 */
function transformSpreadsheet() {
  Logger.log("=== Iniciando transformação da planilha ===");
  Logger.log("Buscando arquivo: " + CONFIG.SOURCE_FILE_NAME);

  // 1. Buscar o arquivo no Google Drive
  const sourceFile = findSourceFile_();
  if (!sourceFile) {
    const msg = "ERRO: Arquivo '" + CONFIG.SOURCE_FILE_NAME + "' não encontrado no Google Drive.\n" +
      "Faça upload do arquivo primeiro e tente novamente.\n" +
      "Extensões suportadas: " + CONFIG.SUPPORTED_EXTENSIONS.join(", ");
    Logger.log(msg);
    throw new Error(msg);
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
 * Busca o arquivo fonte no Google Drive por nome.
 * Suporta múltiplas extensões.
 * @return {GoogleAppsScript.Drive.File|null}
 */
function findSourceFile_() {
  // Buscar com cada extensão suportada
  for (const ext of CONFIG.SUPPORTED_EXTENSIONS) {
    const fileName = CONFIG.SOURCE_FILE_NAME + ext;
    const files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      return files.next();
    }
  }

  // Tentar buscar sem extensão (pode já estar no Drive como nome exato)
  const files = DriveApp.getFilesByName(CONFIG.SOURCE_FILE_NAME);
  if (files.hasNext()) {
    return files.next();
  }

  // Tentar busca parcial
  const searchResults = DriveApp.searchFiles(
    "title contains '" + CONFIG.SOURCE_FILE_NAME + "' and trashed = false"
  );
  if (searchResults.hasNext()) {
    return searchResults.next();
  }

  return null;
}

/**
 * Obtém ou cria uma pasta no Google Drive.
 * @param {string} folderName
 * @return {GoogleAppsScript.Drive.Folder}
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
 * Converte o arquivo para Google Sheets usando a API do Drive.
 * @param {GoogleAppsScript.Drive.File} sourceFile
 * @param {GoogleAppsScript.Drive.Folder} destFolder
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function convertToGoogleSheet_(sourceFile, destFolder) {
  const mimeType = sourceFile.getMimeType();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    CONFIG.TIMESTAMP_FORMAT
  );
  const outputName = CONFIG.OUTPUT_PREFIX + CONFIG.SOURCE_FILE_NAME + "_" + timestamp;

  let spreadsheet;

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    // Já é um Google Sheet - criar uma cópia
    Logger.log("Arquivo já é um Google Sheet. Criando cópia...");
    const copy = sourceFile.makeCopy(outputName, destFolder);
    spreadsheet = SpreadsheetApp.openById(copy.getId());
  } else if (mimeType === MimeType.CSV || sourceFile.getName().endsWith(".csv")) {
    // CSV - criar Google Sheet e importar dados
    Logger.log("Importando arquivo CSV...");
    spreadsheet = importCsvToSheet_(sourceFile, outputName, destFolder);
  } else {
    // Excel (.xlsx, .xls) ou ODS - usar conversão nativa do Drive
    Logger.log("Convertendo " + mimeType + " para Google Sheets...");
    spreadsheet = convertExcelToSheet_(sourceFile, outputName, destFolder);
  }

  return spreadsheet;
}

/**
 * Converte arquivo Excel/ODS para Google Sheets via Drive API.
 * @param {GoogleAppsScript.Drive.File} sourceFile
 * @param {string} outputName
 * @param {GoogleAppsScript.Drive.Folder} destFolder
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function convertExcelToSheet_(sourceFile, outputName, destFolder) {
  const blob = sourceFile.getBlob();

  // Usar Drive API avançada para converter
  const resource = {
    title: outputName,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: destFolder.getId() }]
  };

  // Método 1: Usando Drive API v2 (Advanced Drive Service)
  try {
    const file = Drive.Files.insert(resource, blob, {
      convert: true
    });
    return SpreadsheetApp.openById(file.id);
  } catch (e) {
    Logger.log("Drive API avançada não disponível. Usando método alternativo...");
  }

  // Método 2: Usando UrlFetchApp com Drive API REST
  try {
    const accessToken = ScriptApp.getOAuthToken();
    const uploadUrl = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id";

    const metadata = {
      name: outputName,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [destFolder.getId()]
    };

    const boundary = "-------" + Utilities.getUuid();
    const requestBody =
      "--" + boundary + "\r\n" +
      "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
      JSON.stringify(metadata) + "\r\n" +
      "--" + boundary + "\r\n" +
      "Content-Type: " + sourceFile.getMimeType() + "\r\n\r\n";

    const postData = Utilities.newBlob(requestBody).getBytes()
      .concat(blob.getBytes())
      .concat(Utilities.newBlob("\r\n--" + boundary + "--").getBytes());

    const options = {
      method: "post",
      contentType: "multipart/related; boundary=" + boundary,
      payload: postData,
      headers: {
        Authorization: "Bearer " + accessToken
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(uploadUrl, options);
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(result.error.message);
    }

    return SpreadsheetApp.openById(result.id);
  } catch (e2) {
    Logger.log("Erro no método REST: " + e2.message);
  }

  // Método 3: Copiar arquivo e converter
  const tempFile = sourceFile.makeCopy(outputName + "_temp");
  const tempId = tempFile.getId();

  try {
    // Tentar abrir como spreadsheet diretamente
    const ss = SpreadsheetApp.openById(tempId);
    tempFile.moveTo(destFolder);
    tempFile.setName(outputName);
    return ss;
  } catch (e3) {
    tempFile.setTrashed(true);
    throw new Error(
      "Não foi possível converter o arquivo. Verifique se:\n" +
      "1. O arquivo é um Excel válido (.xlsx ou .xls)\n" +
      "2. O serviço 'Drive API' está habilitado no projeto\n" +
      "3. Erro: " + e3.message
    );
  }
}

/**
 * Importa um arquivo CSV para um novo Google Sheet.
 * @param {GoogleAppsScript.Drive.File} csvFile
 * @param {string} outputName
 * @param {GoogleAppsScript.Drive.Folder} destFolder
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
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

  // Mover para a pasta de destino
  const file = DriveApp.getFileById(spreadsheet.getId());
  file.moveTo(destFolder);

  return spreadsheet;
}

/**
 * Aplica formatação básica ao Google Sheet convertido.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 */
function formatGoogleSheet_(spreadsheet) {
  Logger.log("Aplicando formatação...");

  const sheets = spreadsheet.getSheets();

  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) continue;

    // Formatar cabeçalho (primeira linha)
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4285F4");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setHorizontalAlignment("center");

    // Congelar primeira linha
    sheet.setFrozenRows(1);

    // Auto-redimensionar colunas
    for (let col = 1; col <= lastCol; col++) {
      try {
        sheet.autoResizeColumn(col);
      } catch (e) {
        // Ignorar erro de redimensionamento
      }
    }

    // Adicionar filtro
    if (lastRow > 1) {
      const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
      try {
        dataRange.createFilter();
      } catch (e) {
        // Filtro pode já existir
      }
    }

    // Alternar cores das linhas (zebrado)
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
 * @param {number} bytes
 * @return {string}
 */
function formatFileSize_(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(2) + " KB";
  return (bytes / 1048576).toFixed(2) + " MB";
}

// ============================================================
// FUNÇÕES UTILITÁRIAS EXTRAS
// ============================================================

/**
 * Lista todos os arquivos que contêm "GCACL_BAAPLOG" no nome.
 * Use para verificar se o arquivo está no Drive antes de converter.
 */
function listMatchingFiles() {
  Logger.log("Buscando arquivos com '" + CONFIG.SOURCE_FILE_NAME + "' no nome...\n");

  const files = DriveApp.searchFiles(
    "title contains '" + CONFIG.SOURCE_FILE_NAME + "' and trashed = false"
  );

  let count = 0;
  while (files.hasNext()) {
    const file = files.next();
    count++;
    Logger.log(count + ". " + file.getName());
    Logger.log("   Tipo: " + file.getMimeType());
    Logger.log("   Tamanho: " + formatFileSize_(file.getSize()));
    Logger.log("   Última modificação: " + file.getLastUpdated().toLocaleString("pt-BR"));
    Logger.log("   URL: " + file.getUrl());
    Logger.log("");
  }

  if (count === 0) {
    Logger.log("Nenhum arquivo encontrado. Faça upload do arquivo para o Google Drive primeiro.");
  } else {
    Logger.log("Total: " + count + " arquivo(s) encontrado(s).");
  }
}

/**
 * Faz upload direto de uma planilha Excel via URL e converte para Google Sheets.
 * Útil se o arquivo estiver hospedado em algum servidor.
 * @param {string} fileUrl - URL do arquivo Excel
 */
function transformFromUrl(fileUrl) {
  if (!fileUrl) {
    Logger.log("ERRO: Forneça a URL do arquivo.");
    return;
  }

  Logger.log("Baixando arquivo de: " + fileUrl);

  const response = UrlFetchApp.fetch(fileUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("Erro ao baixar arquivo: HTTP " + response.getResponseCode());
  }

  const blob = response.getBlob().setName(CONFIG.SOURCE_FILE_NAME + ".xlsx");
  const destFolder = getOrCreateFolder_(CONFIG.DESTINATION_FOLDER_NAME);

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    CONFIG.TIMESTAMP_FORMAT
  );
  const outputName = CONFIG.OUTPUT_PREFIX + CONFIG.SOURCE_FILE_NAME + "_" + timestamp;

  // Upload e conversão
  const resource = {
    name: outputName,
    mimeType: "application/vnd.google-apps.spreadsheet",
    parents: [destFolder.getId()]
  };

  const accessToken = ScriptApp.getOAuthToken();
  const uploadUrl = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink";

  const boundary = "-------" + Utilities.getUuid();
  const metadata = JSON.stringify(resource);
  const requestBody =
    "--" + boundary + "\r\n" +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    metadata + "\r\n" +
    "--" + boundary + "\r\n" +
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n";

  const postData = Utilities.newBlob(requestBody).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob("\r\n--" + boundary + "--").getBytes());

  const options = {
    method: "post",
    contentType: "multipart/related; boundary=" + boundary,
    payload: postData,
    headers: {
      Authorization: "Bearer " + accessToken
    },
    muteHttpExceptions: true
  };

  const result = JSON.parse(UrlFetchApp.fetch(uploadUrl, options).getContentText());

  if (result.error) {
    throw new Error("Erro na conversão: " + result.error.message);
  }

  const spreadsheet = SpreadsheetApp.openById(result.id);
  formatGoogleSheet_(spreadsheet);

  Logger.log("\n=== CONVERSÃO CONCLUÍDA ===");
  Logger.log("Nome: " + spreadsheet.getName());
  Logger.log("URL: " + spreadsheet.getUrl());

  return {
    nome: spreadsheet.getName(),
    url: spreadsheet.getUrl(),
    id: spreadsheet.getId()
  };
}
