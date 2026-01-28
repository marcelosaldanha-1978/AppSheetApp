# Transformar Planilha GCACL_BAAPLOG_v1_1 em Google Sheets

## Passo a Passo

### 1. Fazer Upload da Planilha para o Google Drive

1. Acesse [Google Drive](https://drive.google.com)
2. Navegue até a pasta desejada ou crie uma nova
3. Faça upload do arquivo `GCACL_BAAPLOG_v1_1.xlsx` (ou `.xls`, `.csv`, `.ods`) localizado em:
   ```
   D:\_BACKUP_REORGANIZACAO\20260125_205315\Documentos\GCACL_BAAPLOG\GCALC_PROJETO\DADOS_PARA_IMPORTAR
   ```
4. Aguarde o upload ser concluído

### 2. Configurar o Script de Conversão

1. Acesse [Google Apps Script](https://script.google.com)
2. Clique em **"Novo Projeto"**
3. Apague o conteúdo do arquivo `Codigo.gs` padrão
4. Copie e cole o conteúdo do arquivo `TransformToGoogleSheet.js` deste repositório
5. Renomeie o projeto para: `TransformToGoogleSheet`

### 3. Configurar o Manifesto (appsscript.json)

1. No editor do Apps Script, clique em **Configurações do Projeto** (ícone de engrenagem)
2. Marque a opção **"Mostrar arquivo de manifesto "appsscript.json" no editor"**
3. Abra o arquivo `appsscript.json` e substitua pelo conteúdo do `appsscript.json` deste repositório
4. Salve o arquivo

### 4. Habilitar a API do Drive

1. No editor do Apps Script, clique em **Serviços** (ícone de +)
2. Procure e adicione **Drive API v2**
3. Clique em **Adicionar**

### 5. Executar a Conversão

1. No editor, selecione a função `transformSpreadsheet` no dropdown
2. Clique em **Executar** (botão play)
3. Na primeira execução, autorize as permissões solicitadas
4. Acompanhe o progresso no **Log de Execução**

### 6. Resultado

Após a execução, você terá:
- Um novo Google Sheet na pasta `GCACL_BAAPLOG_GoogleSheets` no seu Drive
- Nome no formato: `GSheet_GCACL_BAAPLOG_v1_1_YYYYMMDD_HHmmss`
- Formatação aplicada automaticamente:
  - Cabeçalhos em negrito com fundo azul
  - Primeira linha congelada
  - Filtros habilitados
  - Colunas auto-redimensionadas
  - Linhas com cores alternadas (zebrado)

## Funções Disponíveis

| Função | Descrição |
|--------|-----------|
| `transformSpreadsheet()` | Converte a planilha do Drive para Google Sheets |
| `listMatchingFiles()` | Lista arquivos GCACL_BAAPLOG encontrados no Drive |
| `transformFromUrl(url)` | Converte a partir de uma URL externa |

## Solução de Problemas

| Problema | Solução |
|----------|---------|
| Arquivo não encontrado | Verifique se o upload foi concluído e o nome está correto |
| Erro de permissão | Autorize todas as permissões solicitadas pelo script |
| Drive API não disponível | Adicione o serviço Drive API nas configurações do projeto |
| Arquivo muito grande | O Google Sheets suporta até 10 milhões de células |

## Formatos Suportados

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)
- `.csv` (Valores separados por vírgula)
- `.ods` (OpenDocument Spreadsheet)
