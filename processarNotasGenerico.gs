/**
 * Script genérico para processar notas fiscais no Google Drive.
 * - Lê arquivos (JPG, PNG, PDF) de uma pasta do Drive
 * - Extrai texto via OCR nativo
 * - Localiza data, valor, CNPJ ou CPF
 * - Salva em planilha Google
 * - Move arquivo processado para subpasta "Processados"
 *
 * Personalize os IDs abaixo antes de usar.
 */

function processarNotasSimples() {
  var pastaId = 'COLE_AQUI_O_ID_DA_PASTA';
  var planilhaId = 'COLE_AQUI_O_ID_DA_PLANILHA';
  var sheet = SpreadsheetApp.openById(planilhaId).getSheetByName('Sheet1');
  var pasta = DriveApp.getFolderById(pastaId);

  // Criar/pegar subpasta "Processados"
  var subPastas = pasta.getFoldersByName('Processados');
  var pastaProcessados = subPastas.hasNext() ? subPastas.next() : pasta.createFolder('Processados');

  var arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    var tipo = arquivo.getMimeType();

    if (tipo != 'image/jpeg' && tipo != 'image/png' && tipo != 'application/pdf') {
      Logger.log("Pulando arquivo não suportado: " + arquivo.getName());
      continue;
    }

    // Criar Google Doc temporário para OCR
    var doc = DocumentApp.create('OCR Temp');
    var body = doc.getBody();
    var blob = arquivo.getBlob();
    body.appendImage(blob);

    var docId = doc.getId();
    var docFile = DriveApp.getFileById(docId);
    var texto = docFile.getAs(MimeType.PLAIN_TEXT).getDataAsString();

    // Extrair dados com regex
    var data = texto.match(/\d{2}\/\d{2}\/\d{4}/);
    var valor = texto.match(/R\$?\s?\d+([.,]\d{2})?/);
    var cnpj = texto.match(/\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}/);
    var cpf = texto.match(/\d{3}\.\d{3}\.\d{3}-\d{2}/);

    sheet.appendRow([
      data ? data[0] : '',
      valor ? valor[0] : '',
      cnpj ? cnpj[0] : (cpf ? cpf[0] : ''),
      arquivo.getName()
    ]);

    // Apagar doc temporário
    docFile.setTrashed(true);

    // Mover arquivo processado para "Processados"
    pastaProcessados.addFile(arquivo);
    pasta.removeFile(arquivo);
  }
}
