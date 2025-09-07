/**
 * Abre uma caixa de diálogo HTML personalizada para configurar propriedades:
 * - Nomes das duas secretárias
 */
function openPropertiesModal() { 
    // Valida e configura propriedades automaticamente se necessário (mesmo padrão da sidebar)
    if (!validateScriptProperties()) {
        setupScriptProperties();
        
        if (!validateScriptProperties()) {
            DocumentApp.getUi().alert('❌ Erro: Não foi possível configurar as Propriedades do Script automaticamente. Execute a função setupScriptProperties() manualmente.');
            return;
        }
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    
    // 1. Lê os valores atuais das secretárias
    const currentSecretary13 = scriptProperties.getProperty('secretary13') || '';
    const currentSecretary12 = scriptProperties.getProperty('secretary12') || '';
    
    // 2. Cria o diálogo a partir de um arquivo de template HTML
    const template = HtmlService.createTemplateFromFile('SP_client');
    
    // 3. Passa os valores atuais para o template HTML
    template.currentSecretary13 = currentSecretary13;
    template.currentSecretary12 = currentSecretary12;
  
    // 4. Constrói e exibe o diálogo
    const html = template.evaluate()
        .setWidth(600)
        .setHeight(300);
    DocumentApp.getUi().showModalDialog(html, 'Configurar Propriedades');
}

/**
 * Recebe os dados do diálogo HTML e os salva nas ScriptProperties.
 * Esta função é chamada pelo JavaScript do lado do cliente.
 * @param {object} formData - Objeto contendo os dados do formulário.
 */
function savePropertiesData(formData) {
  const ui = DocumentApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  try {
    // Usa setProperties para salvar múltiplos valores de uma vez (mais eficiente)
    scriptProperties.setProperties({
      'secretary13': formData.secretary13 || '', // Salva o valor ou uma string vazia
      'secretary12': formData.secretary12 || ''
    });

  } catch (e) {
    // Lança um erro que será capturado pelo withFailureHandler no cliente
    throw new Error('Não foi possível salvar as configurações. Erro: ' + e.message);
  }
}