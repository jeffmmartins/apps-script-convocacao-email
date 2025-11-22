/**
 * Função principal para servir a interface web.
 * Deve referenciar o nome do arquivo HTML SEM a extensão ".html".
 * * ATENÇÃO: O arquivo HTML está definido como 'index.html'.
 */
function doGet() {
  // O nome do arquivo HTML agora é 'index'
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('Gerador de Convocação e Minuta')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Função auxiliar para extrair o ID do arquivo do Google Drive a partir de uma URL ou de um ID bruto.
 * @param {string} urlOrId - A URL ou ID do arquivo do Drive.
 * @returns {string} O ID do arquivo.
 */
function extractFileId(urlOrId) {
  // Expressão regular para encontrar o ID em URLs do Drive
  const regex = /[-\w]{25,}/; 
  const match = urlOrId.match(regex);
  
  if (match) {
    return match[0];
  }
  
  // Se não for uma URL, assume que é o ID bruto
  return urlOrId;
}

/**
 * Função do lado do servidor que envia o e-mail via GmailApp com anexo do Drive e negrito em HTML.
 * @param {Object} emailData - Objeto contendo assunto, corpo (simples), htmlCorpo, para, cc, cco e documentoDriveID.
 * @returns {string} O endereço de e-mail do destinatário.
 */
function enviarEmailConvocacao(emailData) {
  
  // 1. Extrai o ID do arquivo
  const fileId = extractFileId(emailData.documentoDriveID);
  
  let anexo = null;
  
  try {
    // 2. Busca o arquivo no Google Drive
    const file = DriveApp.getFileById(fileId);
    // Cria o anexo como PDF
    anexo = file.getAs(MimeType.PDF); 
    
  } catch (e) {
    Logger.log('Erro ao buscar arquivo no Drive com ID %s: %s', fileId, e.toString());
    // Se a falha for no Drive, lança um erro específico.
    throw new Error('Falha na busca do PDF no Drive. Verifique se o ID/URL está correto e se as permissões de acesso ao Drive foram concedidas.');
  }

  try {
    // 3. Envia o e-mail com o anexo e o corpo em HTML (com negrito)
    GmailApp.sendEmail(
      emailData.para, // Destinatário principal (aceita múltiplos separados por vírgula)
      emailData.assunto,
      emailData.corpo, // Corpo de texto simples (fallback)
      {
        // NOVO: FORÇA O REMETENTE PARA O ALIAS CONFIGURADO
        from: 'contratacao@sefaz.ce.gov.br', 
        cc: emailData.cc, // E-mails em cópia (aceita múltiplos separados por vírgula)
        bcc: emailData.cco, // E-mails em cópia oculta (aceita múltiplos separados por vírgula)
        name: 'Secretaria da Fazenda do Estado do Ceará (Sefaz/CE) - Cecoc', // Nome de exibição
        htmlBody: emailData.htmlCorpo, // CORPO PRINCIPAL COM NEGRITO/HTML
        attachments: [anexo] // Anexa o arquivo do Drive
      }
    );
    
    // Retorna o endereço principal para feedback de sucesso
    return emailData.para; 
    
  } catch (e) {
    Logger.log('Erro ao enviar e-mail: %s', e.toString());
    // Se a falha for no Gmail (permissão ou formato de e-mail inválido), lança um erro.
    throw new Error('Falha no envio do e-mail. Verifique o formato dos endereços, se o alias está configurado e se as permissões de Gmail foram concedidas.');
  }
}
