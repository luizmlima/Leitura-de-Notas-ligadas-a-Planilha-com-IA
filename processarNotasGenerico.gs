// IDs — ajuste se necessário
var PASTA_ID = "ID";
var PLANILHA_ID = "ID";
var ABA_PLANILHA = "ABA";

/**
 * Processa todos os arquivos na pasta configurada:
 * - Faz OCR (via Drive API)
 * - Extrai campos (data, total, CNPJ, fornecedor, nº NF, itens)
 * - Registra resultado na planilha
 * - Move arquivo para subpasta Processados ou Erros
 */
function processarNotasPorPaginas() {
  var pasta = DriveApp.getFolderById(PASTA_ID);
  var sheet = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_PLANILHA);
  garantirCabecalho(sheet);

  var pastaProcessados = obterOuCriarSubpasta(pasta, 'Processados');
  var pastaErros = obterOuCriarSubpasta(pasta, 'Erros');

  var arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    try {
      Logger.log('Iniciando: %s (%s)', arquivo.getName(), arquivo.getMimeType());

      if (arquivo.isTrashed()) {
        Logger.log('Arquivo na lixeira — pulando.');
        continue;
      }

      var mime = arquivo.getMimeType();
      if (!/^image\/|application\/pdf/.test(mime)) {
        Logger.log('Tipo não suportado: %s', mime);
        continue;
      }

      var textoOCR = processarBlobOCR(arquivo.getBlob(), arquivo.getName(), mime);
      Logger.log('OCR (amostra): %s', textoOCR ? textoOCR.substring(0, 500) : '[vazio]');

      var texto = normalizarTextoOCR(textoOCR);

      // Extração de campos
      var data = encontrarData(texto);
      var total = encontrarValorTotal(texto);
      var cnpj = encontrarCNPJ(texto);
      var fornecedor = encontrarFornecedor(texto);
      var numeroNF = encontrarNumeroNota(texto);
      var itens = encontrarItens(texto);

      // Grava linha na planilha
      sheet.appendRow([
        data || 'Não encontrado',
        total != null ? total : 'Não encontrado',
        '', // Link (inserido abaixo como RichText)
        fornecedor || '',
        cnpj || '',
        numeroNF || '',
        itens.length ? itens.join('; ') : '',
        arquivo.getName()
      ]);

      // Insere link na coluna 3
      var ultima = sheet.getLastRow();
      sheet.getRange(ultima, 3).setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(arquivo.getName())
          .setLinkUrl(arquivo.getUrl())
          .build()
      );

      // Renomeia arquivo genérico e move para Processados
      if (isNomeGenerico(arquivo.getName())) {
        try {
          var novoNome = sugerirNomeArquivo(arquivo.getName(), data, total);
          arquivo.setName(novoNome);
          Logger.log('Renomeado para: %s', novoNome);
        } catch (e) {
          Logger.log('Falha ao renomear: %s', e.message);
        }
      }

      pastaProcessados.addFile(arquivo);
      pasta.removeFile(arquivo);

    } catch (e) {
      Logger.log('Erro processando %s: %s', arquivo.getName(), e.message);
      try {
        pastaErros.addFile(arquivo);
        pasta.removeFile(arquivo);
      } catch (ee) {
        Logger.log('Falha ao mover para Erros: %s', ee.message);
      }
    }
  }
}

// ----------------------- Utilitários e Extração -----------------------

function garantirCabecalho(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Data', 'Total (R$)', 'Link', 'Fornecedor', 'CNPJ', 'Nº NF', 'Itens (CSV)', 'Nome Arquivo']);
  }
}

function obterOuCriarSubpasta(pastaPai, nome) {
  var it = pastaPai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pastaPai.createFolder(nome);
}

function sugerirNomeArquivo(nomeAntigo, data, total) {
  var safeData = data ? data.replace(/\//g, '-') : 'data_desconhecida';
  var safeTotal = (typeof total === 'number') ? ('R$' + total.toFixed(2).replace('.', ',')) : 'valor_desconhecido';
  var ts = new Date().getTime();
  var ext = nomeAntigo.lastIndexOf('.') > -1 ? nomeAntigo.slice(nomeAntigo.lastIndexOf('.')) : '';
  return 'nota_' + safeData + '_' + safeTotal + '_' + ts + ext;
}

function isNomeGenerico(nome) {
  var nl = nome.toLowerCase();
  return (/^scan\d|^image\d|^doc\d|^file\d|^page\d/.test(nl) || nl.length <= 10);
}

function normalizarTextoOCR(texto) {
  if (!texto) return '';
  var t = texto.replace(/\r/g, '\n').replace(/\t/g, ' ').replace(/[ ]{2,}/g, ' ');
  // Correções típicas de OCR
  t = t.replace(/[lI]/g, '1').replace(/[Oo]/g, '0');
  return t;
}

// ----------------------- Data -----------------------

function encontrarData(texto) {
  if (!texto) return null;
  var linhas = texto.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);

  var padroes = [
    /(\d{1,2}[\/\-\. ]\d{1,2}[\/\-\. ]\d{2,4})\s+(\d{1,2}:\d{2}:\d{2}|\d{1,2}:\d{2})/,
    /(\d{1,2}[\/\-\. ]\d{1,2}[\/\-\. ]\d{2,4})/,
    /(20\d{2}[\-\/]\d{1,2}[\-\/]\d{1,2})/
  ];

  for (var i = linhas.length - 1; i >= 0; i--) {
    var l = linhas[i];
    for (var p = 0; p < padroes.length; p++) {
      var m = l.match(padroes[p]);
      if (m) return padronizarData(m[1]);
    }
    var mMes = l.match(/(\d{1,2})\s*(de\s*)?(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez|janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s*(de\s*)?(\d{2,4})/i);
    if (mMes) {
      var dia = mMes[1].padStart(2, '0');
      var mes = mesNomeParaNumero(mMes[3]).padStart(2, '0');
      var ano = mMes[5].length === 2 ? '20' + mMes[5] : mMes[5];
      return dia + '/' + mes + '/' + ano;
    }
  }
  return null;
}

function mesNomeParaNumero(nome) {
  var m = nome.toLowerCase();
  var mapa = { jan: '01', janeiro: '01', fev: '02', fevereiro: '02', mar: '03', março: '03', abr: '04', abril: '04', mai: '05', maio: '05', jun: '06', junho: '06', jul: '07', julho: '07', ago: '08', agosto: '08', set: '09', setembro: '09', out: '10', outubro: '10', nov: '11', novembro: '11', dez: '12', dezembro: '12' };
  return mapa[m] || '01';
}

function padronizarData(dataStr) {
  dataStr = dataStr.split(/\s+/)[0];
  var partes = dataStr.split(/[\/\-\. ]/).filter(Boolean);
  if (partes.length === 3) {
    if (partes[0].length === 4) {
      return partes[2].padStart(2, '0') + '/' + partes[1].padStart(2, '0') + '/' + partes[0];
    }
    var dia = partes[0].padStart(2, '0');
    var mes = partes[1].padStart(2, '0');
    var ano = partes[2].length === 2 ? '20' + partes[2] : partes[2];
    return dia + '/' + mes + '/' + ano;
  }
  return dataStr;
}

// ----------------------- Valores -----------------------

function encontrarValorTotal(texto) {
  if (!texto) return null;
  var linhas = texto.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);

  // 1) Linhas finais com referência a pagamento/total
  for (var i = linhas.length - 1; i >= 0; i--) {
    var l = linhas[i];
    if (/valor\s*a\s*pagar|valor\s*pago|total\s*pago|pagamento|troco/i.test(l)) {
      var vals = extrairValoresDaLinha(l);
      if (vals.length) {
        if (/troco/i.test(l) && vals.length > 1) return vals[vals.length - 2];
        return vals[vals.length - 1];
      }
    }
  }

  // 2) Linhas com 'TOTAL' ou similares
  for (var j = linhas.length - 1; j >= 0; j--) {
    var lj = linhas[j];
    if (/t[o0]t[a@lI]/i.test(lj) || /valor\s*total/i.test(lj) || /total\s*geral/i.test(lj)) {
      var v = extrairValoresDaLinha(lj);
      if (v.length) return v[v.length - 1];
      var nearby = procurarValorProximo(linhas, j);
      if (nearby != null) return nearby;
    }
  }

  // 3) Fallback: último R$ no texto
  var all = [];
  var re = /R\$?\s*([\d\.,]{2,})/g;
  var m;
  while ((m = re.exec(texto)) !== null) {
    var n = normalizarValor(m[1]);
    if (n != null) all.push(n);
  }
  if (all.length) return all[all.length - 1];

  // 4) Último número plausível
  var re2 = /([\d]{1,3}(?:[\.\,]\d{3})*[\,\.]\d{2})/g;
  while ((m = re2.exec(texto)) !== null) {
    var n2 = normalizarValor(m[1]);
    if (n2 != null) all.push(n2);
  }
  return all.length ? all[all.length - 1] : null;
}

function procurarValorProximo(linhas, idx) {
  for (var d = 1; d <= 3; d++) {
    var up = linhas[idx - d];
    if (up) {
      var v = extrairValoresDaLinha(up);
      if (v.length) return v[v.length - 1];
    }
    var down = linhas[idx + d];
    if (down) {
      var v2 = extrairValoresDaLinha(down);
      if (v2.length) return v2[v2.length - 1];
    }
  }
  return null;
}

function extrairValoresDaLinha(linha) {
  if (!linha) return [];
  var res = [];
  var re = /R\$?\s*([\d\.,]+)/g;
  var m;
  while ((m = re.exec(linha)) !== null) {
    var n = normalizarValor(m[1]); if (n != null) res.push(n);
  }
  if (res.length) return res;
  var re2 = /([\d\.,]{2,})/g;
  while ((m = re2.exec(linha)) !== null) {
    var n2 = normalizarValor(m[1]); if (n2 != null) res.push(n2);
  }
  return res;
}

function normalizarValor(valorStr) {
  if (!valorStr) return null;
  var s = String(valorStr).trim();
  s = s.replace(/[ \u00A0]/g, '');
  s = s.replace(/[Oo]/g, '0').replace(/[lI]/g, '1');
  s = s.replace(/[^0-9\.,-]/g, '').replace(/\(.*?\)/g, '');

  var hasComma = s.indexOf(',') !== -1;
  var hasDot = s.indexOf('.') !== -1;

  if (hasDot && hasComma) {
    if (s.lastIndexOf('.') > s.lastIndexOf(',')) {
      s = s.replace(/\./g, '').replace(/,/g, '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (hasComma && !hasDot) {
    s = s.replace(/,/g, '.');
  } else if (hasDot && !hasComma) {
    var parts = s.split('.');
    if (!(parts.length === 2 && parts[1].length === 2)) s = s.replace(/\./g, '');
  }

  if (s.indexOf('.') === -1) {
    var only = s.replace(/\D/g, '');
    if (only.length > 2) s = only.slice(0, only.length - 2) + '.' + only.slice(-2);
    else if (only.length === 2) s = '0.' + only;
  }

  var num = parseFloat(s);
  return isNaN(num) ? null : num;
}

// ----------------------- CNPJ / Fornecedor / Nº NF -----------------------

function encontrarCNPJ(texto) {
  if (!texto) return null;
  var m = texto.match(/(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})/);
  if (m) return m[1];
  var m2 = texto.match(/(\d{14})/g);
  if (m2) {
    for (var i = 0; i < m2.length; i++) {
      var c = m2[i];
      if (validarCNPJ(c)) return c.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    }
  }
  return null;
}

function validarCNPJ(cnpj) {
  if (!cnpj) return false;
  var s = cnpj.replace(/\D/g, '');
  if (s.length !== 14) return false;
  if (/^(\d)\1+$/.test(s)) return false;
  return true;
}

function encontrarFornecedor(texto) {
  if (!texto) return null;
  var linhas = texto.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
  for (var i = 0; i < Math.min(6, linhas.length); i++) {
    var l = linhas[i];
    if (/[a-zA-Z]/.test(l) && !/cpf|cnpj|inscr|data|nota|nf\b|consumidor|cupom/i.test(l) && l.length > 3) return l;
  }
  return null;
}

function encontrarNumeroNota(texto) {
  if (!texto) return null;
  var m = texto.match(/(?:n\.?º|n\.?\s|nf\s*[:#\-]?|nota\s*fiscal\s*[:#\-]?|numero\s*[:#\-]?|nº|extrato\s*nº)\s*(\d{2,10})/i);
  if (m) return m[1];
  var m2 = texto.match(/\bNF\b\s*(\d{3,})/i);
  if (m2) return m2[1];
  return null;
}

// ----------------------- Itens -----------------------

function encontrarItens(texto) {
  if (!texto) return [];
  var inicio = texto.search(/c(upom|urom)\s*fiscal/i);
  if (inicio === -1) inicio = 0;
  else inicio = texto.indexOf('\n', inicio + 10) + 1;

  var fim = texto.search(/(qtd\.?|valor|total|subtotal|forma\s+de\s+pagamento|pagamento\s+total|ie\s*:)/i);
  var bloco = (fim !== -1 && fim > inicio) ? texto.substring(inicio, fim) : texto.substring(inicio);

  var linhas = bloco.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
  var itens = [];

  var regexItemDetalhe = /(\d+)\s+([\d\.\,]+)\s*x\s*([\d\.\,]+)\s+([\d\.\,]+)$/;
  var regexItemSimples = /^(\d{1,4})\s+([a-zA-Z].*?)\s+([\d\.\,]+)$/;

  for (var i = 0; i < linhas.length; i++) {
    var l = linhas[i];

    var m = l.match(regexItemDetalhe);
    if (m) {
      var desc = i > 0 ? linhas[i - 1] : 'Item Desconhecido';
      itens.push(desc.substring(0, 50).trim() + ' [Qtd: ' + m[2] + ', Preço: ' + m[3] + ', Total: ' + m[4] + ']');
      i++;
      continue;
    }

    var ms = l.match(regexItemSimples);
    if (ms) {
      itens.push(ms[2].substring(0, 50).trim() + ' [Total: ' + ms[3] + ']');
      continue;
    }

    if (/desconto\s*no\s*item|desconto\s*de|acréscimo/i.test(l)) {
      itens.push(l);
    }
  }

  return itens;
}

// ----------------------- OCR (Drive API) -----------------------

/**
 * Cria um Google Doc temporário via Drive API para ativar OCR e retorna o texto.
 * Requer Drive Advanced Service ativado.
 */
function processarBlobOCR(blob, nomeArquivo, mime) {
  try {
    var resource = { title: 'temp-ocr-' + new Date().getTime() + '-' + nomeArquivo };
    var options = { ocr: true };
    var res = Drive.Files.insert(resource, blob, options);
    var docId = res.id;

    for (var t = 0; t < 12; t++) {
      try {
        var doc = DocumentApp.openById(docId);
        var txt = doc.getBody().getText();
        DriveApp.getFileById(docId).setTrashed(true);
        return txt;
      } catch (e) {
        Utilities.sleep(800);
      }
    }
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
    throw new Error('Timeout OCR');
  } catch (e) {
    Logger.log('processarBlobOCR falhou: %s', e.message);
    throw e;
  }
}
