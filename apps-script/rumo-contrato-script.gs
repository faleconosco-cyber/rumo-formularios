// =====================================================
// Apps Script – Rumo Contrato de Prestação de Serviço
// Conta: faleconosco@institutorumo.com
// Planilha: Contratos Assinados Rumo
// =====================================================

function doGet(e)  { return processarDados(e); }
function doPost(e) { return processarDados(e); }

function processarDados(e) {
  try {
    var raw = (e.parameter && e.parameter.payload)
           || (e.postData  && e.postData.contents);
    if (!raw) throw new Error('Sem payload');
    var d = JSON.parse(raw);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contratos');

    if (!sheet) {
      sheet = ss.insertSheet('Contratos');

      var cabecalho = [
        'Data e Hora do Aceite',
        'Nome do Responsável',
        'CPF do Responsável',
        'E-mail do Responsável',
        'Psicóloga Escolhida',
        'Investimento Total',
        'Forma de Pagamento'
      ];

      sheet.appendRow(cabecalho);

      var headerRange = sheet.getRange(1, 1, 1, cabecalho.length);
      headerRange.setBackground('#3E658E');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('Arial');
      headerRange.setFontSize(10);
      sheet.setFrozenRows(1);

      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 220);
      sheet.setColumnWidth(3, 150);
      sheet.setColumnWidth(4, 260);
      sheet.setColumnWidth(5, 260);
      sheet.setColumnWidth(6, 140);
      sheet.setColumnWidth(7, 220);
    }

    var linha = [
      d.dataAceite         || '',
      d.nomeAceite         || '',
      d.cpfAceite          || '',
      d.emailResponsavel   || '',
      d.psicologa          || '',
      d.valorTotal         || '',
      d.formaPagamento     || ''
    ];

    sheet.appendRow(linha);

    // ── E-mail de confirmação ─────────────────────
    if (d.emailResponsavel) {
      var nome  = d.nomeAceite || 'Responsável';
      var data  = d.dataAceite || '';
      var psi   = d.psicologa  || '';
      var valor = d.valorTotal || '';
      var pgto  = d.formaPagamento || '';

      MailApp.sendEmail({
        to:      d.emailResponsavel,
        subject: 'Seu Acordo de Serviço com o Rumo está registrado ✅',
        body:
          'Olá, ' + nome + '!\n\n' +
          'O Acordo de Serviço do processo de Orientação Profissional foi registrado com sucesso em ' + data + '.\n\n' +
          'Resumo da contratação:\n\n' +
          '• Psicóloga: ' + psi + '\n' +
          '• Investimento total: ' + valor + '\n' +
          '• Forma de pagamento: ' + pgto + '\n\n' +
          'O que foi acordado:\n\n' +
          '• O processo é composto por aproximadamente 8 sessões online de 50 minutos cada.\n' +
          '• Em caso de cancelamento ou reagendamento, o aviso deve ser feito com no mínimo 24 horas de antecedência útil. Cancelamentos fora desse prazo implicam cobrança integral da sessão.\n' +
          '• O limite máximo é de 2 cancelamentos ou reagendamentos ao longo de todo o processo.\n' +
          '• Todo contato dos responsáveis deve ser feito exclusivamente com a Diretora do Instituto, Cláudia Botelho.\n' +
          '• O prazo máximo para conclusão do processo é de 12 semanas a partir da data do primeiro encontro.\n\n' +
          'Em breve entraremos em contato para os próximos passos. 🌱\n\n' +
          'Instituto Rumo — Orientação Profissional\n' +
          'faleconosco@institutorumo.com'
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'erro', mensagem: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
