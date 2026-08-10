// ═══════════════════════════════════════════════════════════════
// RUMO – Leads do Google Ads
// Google Apps Script – Recebe os dados do formulário de captura
// (landing page do anúncio) e salva na planilha "Leads do Google
// Ads Multiplica".
//
// INSTRUÇÕES DE USO:
//  1. Abra a planilha "Leads do Google Ads Multiplica" no Google Sheets
//  2. Menu: Extensões > Apps Script
//  3. Cole todo este código no editor (apague o que estiver lá)
//  4. Salve (Ctrl+S)
//  5. Clique em "Implantar" > "Nova implantação"
//  6. Tipo: App da Web | Acesso: Qualquer pessoa
//  7. Copie a URL gerada (termina em /exec) e envie SÓ ela para
//     quem for construir o formulário — nunca o link da planilha.
// ═══════════════════════════════════════════════════════════════

const NOME_ABA = 'Respostas'; // nome da aba onde os dados serão salvos

const COLUNAS = [
  'Data/Hora',
  'Nome',
  'WhatsApp',
  'Categoria',
];

function doGet(e) {
  return processarDados(e);
}

function doPost(e) {
  return processarDados(e);
}

function processarDados(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(NOME_ABA);

    // Cria a aba se não existir
    if (!sheet) {
      sheet = ss.insertSheet(NOME_ABA);
    }

    // Cria cabeçalho se a planilha estiver vazia
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUNAS);
      sheet.getRange(1, 1, 1, COLUNAS.length)
        .setBackground('#3E658E')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(10);
      sheet.setFrozenRows(1);
    }

    const raw = (e.parameter && e.parameter.payload) || (e.postData && e.postData.contents);
    const d = JSON.parse(raw);

    sheet.appendRow([
      d.dataEnvio  || new Date().toLocaleString('pt-BR'),
      d.nome       || '',
      d.whatsapp   || '',
      d.categoria  || '', // "Adolescente" ou "Adulto"
    ]);

    // Ajusta largura das colunas automaticamente
    sheet.autoResizeColumns(1, COLUNAS.length);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Função de teste – rode manualmente para verificar se está funcionando
function testeLocal() {
  const dadosFicticios = {
    dataEnvio: '07/07/2026 10:00:00',
    nome:      'Maria da Silva',
    whatsapp:  '(21) 99999-9999',
    categoria: 'Adolescente',
  };

  const e = { postData: { contents: JSON.stringify(dadosFicticios) } };
  const resultado = doPost(e);
  Logger.log(resultado.getContent());
}
