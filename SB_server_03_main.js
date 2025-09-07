function openSidebarServerSide() {
  const htmlFileName = 'SB_client_INDEX';
  const title = 'Sidebar';

  // ========================================================================
  // 1. VALIDAÇÃO DE PROPRIEDADES (destacado em amarelo)
  // ========================================================================
  console.log('📋 Etapa 1: Validação de Propriedades');
  if (!validateScriptProperties()) {
    setupScriptProperties();
    
    if (!validateScriptProperties()) {
      DocumentApp.getUi().alert('❌ Erro: Não foi possível configurar as Propriedades do Script automaticamente. Execute a função setupScriptProperties() manualmente.');
      return;
    }
  }
  console.log('✅ Propriedades validadas');

  // ========================================================================
  // 2. BUSCA DE DADOS (destacado em verde) - VERSÃO OTIMIZADA
  // ========================================================================
  console.log('📊 Etapa 2: Busca de Dados');
  let sidebarData;
  try {
    console.time('⏱️ Busca de dados');
    
    const jsonResponse = getSidebarData();
    sidebarData = JSON.parse(jsonResponse);
    
    console.timeEnd('⏱️ Busca de dados');
    
    // Verifica se houve erro na busca de dados
    if (sidebarData.error) {
      DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + sidebarData.error);
      return;
    }
    
    console.log(`✅ Dados obtidos: ${sidebarData.items?.length || 0} itens, ${sidebarData.promptTitles?.length || 0} prompts`);
    
  } catch (error) {
    console.error('❌ Erro ao buscar dados da sidebar:', error);
    DocumentApp.getUi().alert('❌ Erro ao buscar dados: ' + error.message);
    return;
  }

  // ========================================================================
  // 3. CRIAÇÃO DA SIDEBAR (destacado em roxo)
  // ========================================================================
  console.log('🎨 Etapa 3: Criação da Sidebar');
  const html = HtmlService.createTemplateFromFile(htmlFileName);
  html.initialSearchInput = "Anamnese";
  html.preloadedData = JSON.stringify(sidebarData); // Passa os dados já obtidos

  const sidebar = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
    .setTitle(title)
    .setWidth(300);

  console.log('✅ Sidebar criada');

  // ========================================================================
  // 4. EXIBIÇÃO DA SIDEBAR
  // ========================================================================
  console.log('📱 Exibindo Sidebar');
  DocumentApp.getUi().showSidebar(sidebar);
  console.log('✅ Sidebar ativa (etapas 4-5: cliente silencioso)');
}


/**
 * Creates optimized sidebar with injected data
 * @param {Object} data - Processed data from stage 2
 * @returns {Object} Creation result
 */
function createOptimizedSidebar(data) {
  console.log('🎨 Etapa 3: Criação da Sidebar');
  
  try {
    // Create HTML template with injected data
    const html = HtmlService.createTemplateFromFile('SB_client_05_main');
    html.initialSearchInput = "Anamnese";
    html.preloadedData = JSON.stringify(data);
    
    // Evaluate and configure sidebar
    const sidebar = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setTitle('Sidebar')
      .setWidth(300);
    
    // Display sidebar
    DocumentApp.getUi().showSidebar(sidebar);
    
    console.log('✅ Sidebar ativa');
    
    return { success: true };
    
  } catch (error) {
    console.error('❌ Erro na criação da sidebar:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Include server-side function for HTML templates
 * @param {string} filename - File to include
 * @returns {string} File content
 */
function includeServerSide(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Legacy validation function for backward compatibility
 * @returns {Object} Validation result
 */
function validateSingleHeadline1() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Basic validation logic
    const h1Pattern = /^#\s+/gm;
    const h1Matches = bodyText.match(h1Pattern);
    
    if (!h1Matches || h1Matches.length === 0) {
      return {
        isValid: false,
        message: 'Nenhum Headline1 encontrado no documento.'
      };
    }
    
    if (h1Matches.length === 1) {
      return {
        isValid: true,
        message: 'Documento válido com um único Headline1.'
      };
    }
    
    return {
      isValid: false,
      showDialog: true,
      uniqueHeadline1s: h1Matches.map((match, index) => `H1 ${index + 1}`),
      message: `Foram encontrados ${h1Matches.length} Headline1 diferentes no documento.`
    };
    
  } catch (error) {
    return {
      isValid: false,
      message: 'Erro ao analisar a estrutura do documento: ' + error.message
    };
  }
}

/**
 * Builds document structure with H1 handling
 * @param {string} text - Text to process
 * @param {Array} headline1Rules - H1 rules
 * @param {Array} headline2Rules - H2 rules
 * @returns {Array} Document structure
 */
function buildDocumentStructureWithH1Only(text, headline1Rules, headline2Rules) {
  const lines = text.split('\n');
  let currentHeadline1 = null;
  let documentStructure = [];
  let foundFirstStructure = false;
  let awaitingHeadline2 = false;

  lines.forEach((line, idx) => {
    if (line.trim() === "") return;

    if (headline1Rules.some(rule => rule.condition(line))) {
      if (awaitingHeadline2 && currentHeadline1) {
        documentStructure.push({
          headline1: currentHeadline1,
          headline2: "Documento Indefinido",
          text: ''
        });
      }
      
      currentHeadline1 = line;
      awaitingHeadline2 = true;
    }
    else if (headline2Rules.some(rule => rule.condition(line))) {
      documentStructure.push({
        headline1: currentHeadline1,
        headline2: line,
        text: ''
      });
      foundFirstStructure = true;
      awaitingHeadline2 = false;
    }
    else if (awaitingHeadline2) {
      documentStructure.push({
        headline1: currentHeadline1,
        headline2: "Título",
        text: line + '\n'
      });
      foundFirstStructure = true;
      awaitingHeadline2 = false;
    }
    else if (documentStructure.length > 0) {
      documentStructure[documentStructure.length - 1].text += line + '\n';
    }
  });

  if (awaitingHeadline2 && currentHeadline1) {
    documentStructure.push({
      headline1: currentHeadline1,
      headline2: "Documento Indefinido",
      text: ''
    });
  }

  return documentStructure;
}

/**
 * Opens append choice dialog
 * @param {Array} uniqueHeadline1s - Unique H1 options
 * @param {string} textToAppend - Text to append
 * @param {boolean} shouldFormat - Whether to format
 */
function openAppendChoiceDialog(uniqueHeadline1s, textToAppend, shouldFormat) {
  const html = HtmlService.createTemplateFromFile('SB_modelessDialog');
  html.uniqueHeadline1s = uniqueHeadline1s;
  html.textToAppend = textToAppend;
  html.shouldFormat = shouldFormat;
  
  const dialog = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.NATIVE)
    .setWidth(400)
    .setHeight(300);
  
  DocumentApp.getUi().showModalDialog(dialog, 'Escolha o Headline1');
}

/**
 * Validates H1s for brush operation
 * @returns {Object} Validation result
 */
function validateHeadline1sForBrush() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Use validation logic
    const validation = validateSingleHeadline1();
    
    if (!validation.isValid && validation.showDialog) {
      return {
        isValid: false,
        showDialog: true,
        uniqueHeadline1s: validation.uniqueHeadline1s,
        message: validation.message
      };
    }
    
    return validation;
    
  } catch (error) {
    return {
      isValid: false,
      message: 'Erro ao validar estrutura para brush: ' + error.message
    };
  }
}

/**
 * Gets brush to headline2s mapping
 * @returns {Array} Mapping array
 */
function getBrushToHeadline2s() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stored = properties.getProperty('brushToHeadline2s');
    
    if (stored) {
      return JSON.parse(stored);
    }
    
    return [];
    
  } catch (error) {
    console.error('❌ Erro ao obter brush mapping:', error);
    return [];
  }
}

/**
 * Saves brush to headline2s mapping
 * @param {Array} toHeadline2sArray - Mapping array
 */
function saveBrushToHeadline2s(toHeadline2sArray) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('brushToHeadline2s', JSON.stringify(toHeadline2sArray));
    
    console.log('✅ Brush mapping salvo');
    
  } catch (error) {
    console.error('❌ Erro ao salvar brush mapping:', error);
  }
} 