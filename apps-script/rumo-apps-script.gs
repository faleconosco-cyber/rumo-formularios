// ═══════════════════════════════════════════════════════════════
// RUMO – Orientação Profissional
// Google Apps Script – Recebe respostas do formulário e salva
// na planilha do Google Sheets.
//
// INSTRUÇÕES DE USO:
//  1. Abra sua planilha no Google Sheets
//  2. Menu: Extensões > Apps Script
//  3. Cole todo este código no editor (apague o que estiver lá)
//  4. Salve (Ctrl+S)
//  5. Clique em "Implantar" > "Nova implantação"
//  6. Tipo: App da Web | Acesso: Qualquer pessoa
//  7. Copie a URL gerada e cole em rumo-formulario.html
//     na variável APPS_SCRIPT_URL
// ═══════════════════════════════════════════════════════════════

const NOME_ABA = 'Respostas'; // nome da aba onde os dados serão salvos

const COLUNAS = [
  'Data/Hora',
  'Nome do Jovem',
  'Data de Nascimento',
  'Colégio',
  'Série',
  'Nome do Responsável',
  'WhatsApp',
  'Como chegou',
  'Orientação anterior',
  'Detalhe orientação',
  'Motivação',
  'Quando surgiu',
  'Reação ao tema',
  'Expectativas',
  'Psicoterapia',
  'Diagnóstico',
  'Detalhe diagnóstico',
  'Psiquiatra',
  'Medicação',
  'Detalhe medicação',
  'Observações',
  'Dias disponíveis',
  'Períodos disponíveis',
  'Preferência de atendimento',
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
      d.dataEnvio             || new Date().toLocaleString('pt-BR'),
      d.nomeJovem             || '',
      d.dataNasc              || '',
      d.colegio               || '',
      d.serie                 || '',
      d.nomeResp              || '',
      d.whatsapp              || '',
      d.comoChegou            || '',
      d.orientacaoAnterior    || '',
      d.detailOrientacao      || '',
      d.motivacao             || '',
      d.quandoSurgiu          || '',
      d.reacaoTema            || '',
      d.expectativas          || '',
      d.psicoterapia          || '',
      d.diagnostico           || '',
      d.detailDiagnostico     || '',
      d.psiquiatra            || '',
      d.medicacao             || '',
      d.detailMedicacao       || '',
      d.observacoes               || '',
      d.diasDisponiveis           || '',
      d.periodosDisponiveis       || '',
      d.preferenciaAtendimento    || '',
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
    dataEnvio:          '06/05/2026 10:00:00',
    nomeJovem:          'João da Silva',
    dataNasc:           '10/03/2008',
    colegio:            'Colégio Exemplo',
    serie:              '3º ano EM',
    nomeResp:           'Maria da Silva',
    whatsapp:           '(11) 99999-9999',
    comoChegou:         'Instagram',
    orientacaoAnterior: 'Não',
    detailOrientacao:   '',
    motivacao:          'Meu filho não sabe o que quer fazer e está ansioso com o vestibular.',
    quandoSurgiu:       'Se intensificou recentemente',
    reacaoTema:         'Fica ansioso/a, Evita o assunto',
    expectativas:       'Autoconhecimento, Clareza na escolha profissional',
    psicoterapia:       'Nunca teve',
    diagnostico:        'Não',
    detailDiagnostico:  '',
    psiquiatra:         'Não',
    medicacao:          'Não',
    detailMedicacao:    '',
    observacoes:        '',
  };

  const e = { postData: { contents: JSON.stringify(dadosFicticios) } };
  const resultado = doPost(e);
  Logger.log(resultado.getContent());
}
