// =====================================================
// Apps Script – Rumo Cadastro e Autorização
// Conta: faleconosco@institutorumo.com
// Planilha: Respostas Cadastro Rumo
// =====================================================

function doGet(e)  { return processarDados(e); }
function doPost(e) { return processarDados(e); }

function processarDados(e) {
  try {
    // Lê o payload independente de GET ou POST
    var raw = (e.parameter && e.parameter.payload)
           || (e.postData  && e.postData.contents);
    if (!raw) throw new Error('Sem payload');
    var d = JSON.parse(raw);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Respostas');

    // Cria a aba 'Respostas' e cabeçalho na primeira execução
    if (!sheet) {
      sheet = ss.insertSheet('Respostas');

      var cabecalho = [
        // Controle
        'Data e Hora',
        // Responsável 1
        'Resp1 – Nome completo',
        'Resp1 – Relação com adolescente',
        'Resp1 – Formação',
        'Resp1 – Atua na área?',
        'Resp1 – Área de atuação',
        'Resp1 – WhatsApp',
        'Resp1 – E-mail',
        'Resp1 – CPF',
        'Resp1 – CEP',
        'Resp1 – Endereço',
        'Resp1 – Autoriza pesquisa acadêmica',
        // Responsável 2
        'Resp2 – Nome completo',
        'Resp2 – Relação com adolescente',
        'Resp2 – Formação',
        'Resp2 – Atua na área?',
        'Resp2 – Área de atuação',
        'Resp2 – WhatsApp',
        'Resp2 – E-mail',
        'Resp2 – CPF',
        'Resp2 – CEP',
        'Resp2 – Endereço',
        'Resp2 – Autoriza pesquisa acadêmica',
        // Situação familiar
        'Situação familiar',
        // Adolescente
        'Adolescente – Nome completo',
        'Adolescente – Data de nascimento',
        'Adolescente – Série/ano',
        'Adolescente – Escola',
        // Expectativas e percepções
        'Expectativas com o processo',
        'Habilidades percebidas',
        'Carreiras que combinam',
        'Carreiras que não combinam',
        // Saúde mental
        'Psicoterapia',
        'Acompanhamento psiquiátrico',
        'Diagnóstico',
        'Qual diagnóstico',
        'Medicação',
        'Qual medicação',
        // LGPD e autorizações
        'Autoriza participação',
        'Autoriza uso de imagem',
        'Autoriza contato para pesquisa',
        'Observações sobre separação',
        // Documentos
        'Confirmação envio de documentos'
      ];

      sheet.appendRow(cabecalho);

      // Estilo do cabeçalho
      var headerRange = sheet.getRange(1, 1, 1, cabecalho.length);
      headerRange.setBackground('#3E658E');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('Arial');
      headerRange.setFontSize(10);
      sheet.setFrozenRows(1);

      // Largura das colunas
      sheet.setColumnWidth(1, 160);  // Data e Hora
      for (var c = 2; c <= cabecalho.length; c++) {
        sheet.setColumnWidth(c, 200);
      }
    }

    // Monta a linha de dados
    var linha = [
      d.dataEnvio              || '',
      // Responsável 1
      d.resp1Nome              || '',
      d.resp1Relacao           || '',
      d.resp1Formacao          || '',
      d.resp1AtuaNaArea        || '',
      d.resp1Atuacao           || '',
      d.resp1Whatsapp          || '',
      d.resp1Email             || '',
      d.resp1Cpf               || '',
      d.resp1Cep               || '',
      d.resp1Endereco          || '',
      d.resp1AutorizaPesquisa  || '',
      // Responsável 2
      d.resp2Nome              || '',
      d.resp2Relacao           || '',
      d.resp2Formacao          || '',
      d.resp2AtuaNaArea        || '',
      d.resp2Atuacao           || '',
      d.resp2Whatsapp          || '',
      d.resp2Email             || '',
      d.resp2Cpf               || '',
      d.resp2Cep               || '',
      d.resp2Endereco          || '',
      d.resp2AutorizaPesquisa  || '',
      // Situação familiar
      d.situacaoFamiliar       || '',
      // Adolescente
      d.nomeAdolescente        || '',
      d.dataNascAdolescente    || '',
      d.serieAdolescente       || '',
      d.escolaAdolescente      || '',
      // Expectativas e percepções
      d.expectativas           || '',
      d.habilidades            || '',
      d.carreirasCombinam      || '',
      d.carreirasNaoCombinam   || '',
      // Saúde mental
      d.psicoterapia           || '',
      d.psiquiatra             || '',
      d.diagnostico            || '',
      d.detalheDiagnostico     || '',
      d.medicacao              || '',
      d.detalheMedicacao       || '',
      // LGPD e autorizações
      d.autorizaParticipacao   || '',
      d.autorizaImagem         || '',
      d.autorizaContato        || '',
      d.observacoesSeparacao   || '',
      // Documentos
      d.confirmacaoDoc         || ''
    ];

    sheet.appendRow(linha);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
