/**
 * Versão customizada do buildDocumentStructure que trata H1 sozinho
 * @param {string} text - Texto para processar
 * @param {Array} headline1Rules - Regras para H1
 * @param {Array} headline2Rules - Regras para H2
 * @return {Array} Estrutura do documento
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
      // Se há um H1 pendente, criar entrada para ele antes de processar o novo
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

  // Se terminou o processamento e ainda há um H1 pendente, criar entrada
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
 * Valida se o documento possui apenas um Headline1
 * @return {Object} Objeto com isValid (boolean) e message (string)
 */
function validateSingleHeadline1() {
  try {
    // Obter o documento e seu texto
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Usar TextManipulator para processar o texto primeiro
    const formattingRules = getFormattingRules();
    const placeholders = getRulesForPlaceholders();
    const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
    const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
    
    const textManipulator = new TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
    const processedText = textManipulator
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks()
      .getResult();
    
    // Criar uma instância customizada do HeadlineManager
    const headlineManager = new DocumentHeadlineManager();
    
    // Substituir o texto do corpo pela versão processada
    headlineManager.bodyText = processedText;
    
    // Usar uma versão customizada do buildDocumentStructure que trata H1 sozinho
    headlineManager.bodyStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    
    const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
    
    if (uniqueHeadline1s.length === 0) {
      return {
        isValid: false,
        message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um Headline1 para usar esta função.'
      };
    }
    
    if (uniqueHeadline1s.length === 1) {
      return {
        isValid: true,
        message: 'Documento válido com um único Headline1.'
      };
    }
    
    // Mais de um Headline1 encontrado - abrir diálogo de seleção
    return {
      isValid: false,
      showDialog: true,
      uniqueHeadline1s: uniqueHeadline1s,
      message: `Foram encontrados ${uniqueHeadline1s.length} Headline1 diferentes no documento.`
    };
    
  } catch (error) {
    return {
      isValid: false,
      message: 'Erro ao analisar a estrutura do documento: ' + error.message
    };
  }
}

/**
 * Abre o diálogo de seleção de H1 quando há múltiplos H1 disponíveis
 * @param {Array} uniqueHeadline1s - Array com os H1 únicos encontrados
 * @param {string} textToAppend - Texto a ser adicionado
 * @param {boolean} shouldFormat - Se deve formatar após adicionar
 */
function openAppendChoiceDialog(uniqueHeadline1s, textToAppend, shouldFormat) {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('SB_client_APPEND_multiH1s');
    htmlTemplate.uniqueHeadline1s = uniqueHeadline1s;
    htmlTemplate.textToAppend = textToAppend;
    htmlTemplate.shouldFormat = shouldFormat;
    
    const ui = DocumentApp.getUi();
    const dialog = htmlTemplate.evaluate()
      .setWidth(400)
      .setHeight(500);
    
    ui.showModalDialog(dialog, 'Escolher Paciente');
    
  } catch (error) {
    DocumentApp.getUi().alert('Erro ao abrir diálogo de seleção: ' + error.message);
  }
}


/**
 * Executa todo o fluxo do botão brush em uma única operação
 * Combina: obtenção de toHeadline2s, validação de H1s e processamento com IA
 * @param {Object} brushData - Dados do brush incluindo prompts do cliente
 * @return {Object} Resultado da operação
 */
function executeCompleteBrushFlow(brushData) {
  const { fixedFromHeadline2 } = brushData;
  
  try {
    // ETAPA 1: Obter toHeadline2s do PropertiesService
    const properties = PropertiesService.getScriptProperties();
    const toHeadline2sString = properties.getProperty('BRUSH_TO_HEADLINE2S');
    
    if (!toHeadline2sString) {
      return {
        success: false,
        action: 'no_config',
        message: 'Nenhum toHeadline2 configurado. Configure os toHeadline2s nas propriedades do script primeiro.'
      };
    }
    
    // Converte a string CSV em array
    const toHeadline2sArray = toHeadline2sString
      .split(',')
      .map(item => item.trim())
      .filter(item => item.length > 0);
    
    if (toHeadline2sArray.length === 0) {
      return {
        success: false,
        action: 'no_config',
        message: 'Configuração de toHeadline2s está vazia.'
      };
    }
    
    // NOVO: Obter prompts diretamente do servidor
    const documentId = properties.getProperty('docData');
    const sheetId = properties.getProperty('sheetData');
    
    if (!documentId || !sheetId) {
      return {
        success: false,
        action: 'no_config',
        message: 'Configuração de documentId ou sheetId não encontrada.'
      };
    }
    
    const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
    const data = dataFetcher.extractAllDataFromSources();
    const allPrompts = data.allPrompts;
    
    if (!allPrompts || Object.keys(allPrompts).length === 0) {
      return {
        success: false,
        action: 'no_prompts',
        message: 'Nenhum prompt encontrado no sistema.'
      };
    }
    
    // ETAPA 2: Validar H1s do documento
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Usar TextManipulator para processar o texto
    const formattingRules = getFormattingRules();
    const placeholders = getRulesForPlaceholders();
    const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
    const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
    
    const textManipulator = new TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
    const processedText = textManipulator
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks()
      .getResult();
    
    // Criar estrutura do documento
    const headlineManager = new DocumentHeadlineManager();
    headlineManager.bodyText = processedText;
    headlineManager.bodyStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    
    const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
    
    // Verificar cenários de H1s
    if (uniqueHeadline1s.length === 0) {
      return {
        success: false,
        action: 'no_h1',
        message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um paciente para usar esta função.'
      };
    }
    
    if (uniqueHeadline1s.length > 1) {
      return {
        success: false,
        action: 'multiple_h1',
        message: `Encontrados ${uniqueHeadline1s.length} Headline1s no documento. O processamento de múltiplos pacientes ainda não está implementado para o botão Brush.`
      };
    }
    
    // Verificar se existe "Prontuário Médico"
    const headline1 = uniqueHeadline1s[0];
    const hasProtuarioPair = headlineManager.bodyStructure.some(item => 
      item.headline1 === headline1 && 
      item.headline2 === "Prontuário Médico"
    );
    
    if (!hasProtuarioPair) {
      return {
        success: false,
        action: 'no_prontuario',
        message: 'Não foi encontrado o Headline2 "Prontuário Médico" no documento. Esta função requer a seção "Prontuário Médico" para funcionar.'
      };
    }
    
    // ETAPA 3: Montar prompts válidos
    const validPrompts = [];
    const pairs = [];
    
    toHeadline2sArray.forEach(toHeadline2 => {
      // Procura o prompt correspondente (com normalização Unicode)
        const promptKey = Object.keys(allPrompts).find(key => {
          const prompt = allPrompts[key];
          
          // Normalizar strings para resolver problemas de acentos (NFC vs NFD)
          const normalizedFromPrompt = (prompt.fromHeadline2 || '').normalize('NFC');
          const normalizedToPrompt = (prompt.toHeadline2 || '').normalize('NFC');
          const normalizedFromFixed = fixedFromHeadline2.normalize('NFC');
          const normalizedToSearch = toHeadline2.normalize('NFC');
          
                    const fromMatch = normalizedFromPrompt === normalizedFromFixed;
          const toMatch = normalizedToPrompt === normalizedToSearch;
          
          return fromMatch && toMatch;
      });
      
      if (promptKey) {
        const promptContent = allPrompts[promptKey].content;
        
        validPrompts.push({
          fromHeadline2: fixedFromHeadline2,
          toHeadline2: toHeadline2,
          promptContent: promptContent,
          optionText: `Melhorar ${toHeadline2}`
        });
        
        pairs.push({
          headline1: headline1,
          headline2: toHeadline2
        });
      }
    });
    
    if (validPrompts.length === 0) {
      return {
        success: false,
        action: 'no_prompts',
        message: 'Nenhum prompt válido encontrado para os toHeadline2s configurados. Verifique a configuração dos prompts.'
      };
    }
    
    // ETAPA 4: Executar processamento com IA
    const headline1s = [{ headline1: headline1 }];
    
    const processingData = {
      headline1s: headline1s,
      prompts: validPrompts,
      requireAugmentedContext: true
    };
    
    // Chamar a função de processamento que está em AI_server_main.js
    let result;
    try {
      result = processSequentialPromptsForPairs(processingData);
    } catch (error) {
      return {
        success: false,
        action: 'processing_error',
        message: 'Erro ao chamar função de processamento: ' + error.message
      };
    }
    
    if (result.success) {
      return {
        success: true,
        action: 'processed',
        processedCount: validPrompts.length,
        message: `Processamento concluído com sucesso! ${validPrompts.length} seção(ões) melhorada(s).`
      };
    } else {
      return {
        success: false,
        action: 'processing_error',
        message: `Erro no processamento: ${result.message || 'Erro desconhecido'}`
      };
    }
    
  } catch (error) {
    return {
      success: false,
      action: 'error',
      message: 'Erro durante o processamento brush: ' + error.message
    };
  }
}


/**
 * Valida o número de Headline1s no documento para o botão brush
 * @return {Object} Objeto com count, uniqueHeadline1s, hasProntuario e message
 */
function validateHeadline1sForBrush() {
  try {
    // Obter o documento e seu texto
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Usar TextManipulator para processar o texto primeiro
    const formattingRules = getFormattingRules();
    const placeholders = getRulesForPlaceholders();
    const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
    const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
    
    const textManipulator = new TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
    const processedText = textManipulator
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks()
      .getResult();
    
    // Criar uma instância customizada do HeadlineManager
    const headlineManager = new DocumentHeadlineManager();
    
    // Substituir o texto do corpo pela versão processada
    headlineManager.bodyText = processedText;
    
    // Usar uma versão customizada do buildDocumentStructure que trata H1 sozinho
    headlineManager.bodyStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    
    const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
    
    // Verificar se existe "Prontuário Médico" quando há apenas um H1
    let hasProntuario = false;
    if (uniqueHeadline1s.length === 1) {
      // Verificar se existe o par H1 + "Prontuário Médico"
      const hasProtuarioPair = headlineManager.bodyStructure.some(item => 
        item.headline1 === uniqueHeadline1s[0] && 
        item.headline2 === "Prontuário Médico"
      );
      hasProntuario = hasProtuarioPair;
    }
    
    return {
      count: uniqueHeadline1s.length,
      uniqueHeadline1s: uniqueHeadline1s,
      hasProntuario: hasProntuario,
      message: `${uniqueHeadline1s.length} Headline1(s) encontrado(s) no documento.`
    };
    
  } catch (error) {
    console.error('❌ Erro ao validar H1s para brush:', error);
    throw new Error('Erro ao validar Headline1s do documento: ' + error.message);
  }
}

/**
 * Obtém os toHeadline2s configurados no PropertiesService para o botão brush
 * @return {Array} Array de toHeadline2s
 */
function getBrushToHeadline2s() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const toHeadline2sString = properties.getProperty('BRUSH_TO_HEADLINE2S');
    
    if (!toHeadline2sString) {
      console.log('📝 Nenhum toHeadline2 configurado no PropertiesService');
      return [];
    }
    
    // Converte a string CSV em array e remove espaços em branco
    const toHeadline2sArray = toHeadline2sString
      .split(',')
      .map(item => item.trim())
      .filter(item => item.length > 0);
    
    console.log('📝 toHeadline2s obtidos do PropertiesService:', toHeadline2sArray);
    return toHeadline2sArray;
    
  } catch (error) {
    console.error('❌ Erro ao obter toHeadline2s do PropertiesService:', error);
    throw new Error('Erro ao obter configuração de toHeadline2s: ' + error.message);
  }
}

/**
 * Salva os toHeadline2s no PropertiesService
 * @param {Array} toHeadline2sArray - Array de toHeadline2s
 */
function saveBrushToHeadline2s(toHeadline2sArray) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const toHeadline2sString = toHeadline2sArray.join(', ');
    
    properties.setProperty('BRUSH_TO_HEADLINE2S', toHeadline2sString);
    console.log('✅ toHeadline2s salvos no PropertiesService:', toHeadline2sString);
    
    return true;
  } catch (error) {
    console.error('❌ Erro ao salvar toHeadline2s no PropertiesService:', error);
    throw new Error('Erro ao salvar configuração de toHeadline2s: ' + error.message);
  }
}

/**
 * Executa o fluxo completo do brush usando pares fromHeadline2,toHeadline2
 * das script properties brushButtonPairs
 * @return {Object} Resultado da operação
 */
function executeCompleteBrushFlowWithPairs() {
  try {
    // ETAPA 1: Obter pares fromHeadline2,toHeadline2 do PropertiesService
    const properties = PropertiesService.getScriptProperties();
    let brushButtonPairsString = properties.getProperty('brushButtonPairs');
    
    if (!brushButtonPairsString) {
      // Configurar pares padrão automaticamente
      const setupResult = setupBrushButtonPairs();
      if (setupResult.success) {
        brushButtonPairsString = properties.getProperty('brushButtonPairs');
      } else {
        return {
          success: false,
          action: 'no_config',
          message: 'Não foi possível configurar os pares automaticamente: ' + setupResult.message
        };
      }
    }
    
    // Converte a string em array de pares
    // Formato: "Prontuário Médico, Prontuário Médico; Prontuário Médico, Prescrição de Óculos; Laudo de Mapeamento de Retina, Laudo de Mapeamento de Retina"
    const pairsArray = brushButtonPairsString
      .split(';')
      .map(pair => pair.trim())
      .filter(pair => pair.length > 0)
      .map(pair => {
        const [fromHeadline2, toHeadline2] = pair.split(',').map(item => item.trim());
        return { fromHeadline2, toHeadline2 };
      });
    
    if (pairsArray.length === 0) {
      return {
        success: false,
        action: 'no_config',
        message: 'Configuração de pares está vazia ou mal formatada.'
      };
    }
    
    // ETAPA 2: Obter prompts diretamente do servidor
    const documentId = properties.getProperty('docData');
    const sheetId = properties.getProperty('sheetData');
    
    if (!documentId || !sheetId) {
      return {
        success: false,
        action: 'no_config',
        message: 'Configuração de documentId ou sheetId não encontrada.'
      };
    }
    
    const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
    const data = dataFetcher.extractAllDataFromSources();
    const allPrompts = data.allPrompts;
    
    if (!allPrompts || Object.keys(allPrompts).length === 0) {
      return {
        success: false,
        action: 'no_prompts',
        message: 'Nenhum prompt encontrado no sistema.'
      };
    }
    
    // ETAPA 3: Validar estrutura do documento
    const doc = DocumentApp.getActiveDocument();
    const bodyText = doc.getBody().getText();
    
    // Usar TextManipulator para processar o texto
    const formattingRules = getFormattingRules();
    const placeholders = getRulesForPlaceholders();
    const heading1Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
    const heading2Rules = formattingRules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
    
    const textManipulator = new TextManipulator(bodyText, placeholders, heading1Rules, heading2Rules);
    const processedText = textManipulator
      .splitInLines()
      .removeEmptyLines()
      .trimLeadingSpaces()
      .checkAndModifyFirstLine()
      .processLines()
      .structureH2Blocks()
      .getResult();
    
    // Criar estrutura do documento
    const headlineManager = new DocumentHeadlineManager();
    headlineManager.bodyText = processedText;
    headlineManager.bodyStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    
    const uniqueHeadline1s = headlineManager.getUniqueHeadline1s();
    
    // Verificar cenários de H1s
    if (uniqueHeadline1s.length === 0) {
      return {
        success: false,
        action: 'no_h1',
        message: 'Nenhum Headline1 encontrado no documento. É necessário ter pelo menos um paciente para usar esta função.'
      };
    }
    
    if (uniqueHeadline1s.length > 1) {
      return {
        success: false,
        action: 'multiple_h1',
        message: `Encontrados ${uniqueHeadline1s.length} Headline1s no documento. O processamento de múltiplos pacientes ainda não está implementado para o botão Brush.`
      };
    }
    
    // ETAPA 4: Verificar compatibilidade dos pares com o documento
    const headline1 = uniqueHeadline1s[0];
    const validPrompts = [];
    
    pairsArray.forEach(pair => {
      const { fromHeadline2, toHeadline2 } = pair;
      
      // Verificar se existe o fromHeadline2 no documento
      const hasFromHeadline2 = headlineManager.bodyStructure.some(item => 
        item.headline1 === headline1 && 
        item.headline2 === fromHeadline2
      );
      
      if (!hasFromHeadline2) {
        console.log(`⚠️ Par ignorado: ${fromHeadline2} -> ${toHeadline2}. Não encontrado "${fromHeadline2}" no documento.`);
        return; // Pula este par
      }
      
      // Procurar o prompt correspondente
      const promptKey = Object.keys(allPrompts).find(key => {
        const prompt = allPrompts[key];
        
        // Normalizar strings para resolver problemas de acentos (NFC vs NFD)
        const normalizedFromPrompt = (prompt.fromHeadline2 || '').normalize('NFC');
        const normalizedToPrompt = (prompt.toHeadline2 || '').normalize('NFC');
        const normalizedFromSearch = fromHeadline2.normalize('NFC');
        const normalizedToSearch = toHeadline2.normalize('NFC');
        
        const fromMatch = normalizedFromPrompt === normalizedFromSearch;
        const toMatch = normalizedToPrompt === normalizedToSearch;
        
        return fromMatch && toMatch;
      });
      
      if (promptKey) {
        const promptContent = allPrompts[promptKey].content;
        
        validPrompts.push({
          fromHeadline2: fromHeadline2,
          toHeadline2: toHeadline2,
          promptContent: promptContent,
          optionText: `Melhorar ${toHeadline2}`
        });
        
        console.log(`✅ Par válido encontrado: ${fromHeadline2} -> ${toHeadline2}`);
      } else {
        console.log(`⚠️ Par ignorado: ${fromHeadline2} -> ${toHeadline2}. Prompt não encontrado.`);
      }
    });
    
    if (validPrompts.length === 0) {
      return {
        success: false,
        action: 'no_valid_pairs',
        message: 'Nenhum par válido encontrado. Verifique se existem H2s compatíveis no documento e se os prompts estão configurados corretamente.'
      };
    }
    
    // ETAPA 5: Executar processamento com IA
    const headline1s = [{ headline1: headline1 }];
    
    const processingData = {
      headline1s: headline1s,
      prompts: validPrompts,
      requireAugmentedContext: true
    };
    
    // Chamar a função de processamento que está em AI_server_main.js
    let result;
    try {
      result = processSequentialPromptsForPairs(processingData);
    } catch (error) {
      return {
        success: false,
        action: 'processing_error',
        message: 'Erro ao chamar função de processamento: ' + error.message
      };
    }
    
    if (result.success) {
      return {
        success: true,
        action: 'processed',
        processedCount: validPrompts.length,
        message: `Processamento concluído com sucesso! ${validPrompts.length} seção(ões) melhorada(s).`
      };
    } else {
      return {
        success: false,
        action: 'processing_error',
        message: `Erro no processamento: ${result.message || 'Erro desconhecido'}`
      };
    }
    
  } catch (error) {
    return {
      success: false,
      action: 'error',
      message: 'Erro durante o processamento brush com pares: ' + error.message
    };
  }
}

/**
 * Configura os pares padrão do botão brush nas script properties se não existirem
 */
function setupBrushButtonPairs() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const existingPairs = properties.getProperty('brushButtonPairs');
    
    if (!existingPairs) {
      const defaultPairs = "Prontuário Médico, Prontuário Médico; Prontuário Médico, Prescrição de Óculos; Laudo de Mapeamento de Retina, Laudo de Mapeamento de Retina";
      
      properties.setProperty('brushButtonPairs', defaultPairs);
      console.log('✅ Pares padrão do botão brush configurados:', defaultPairs);
      
      return {
        success: true,
        message: 'Pares padrão configurados com sucesso.'
      };
    } else {
      console.log('📝 Pares do botão brush já existem:', existingPairs);
      return {
        success: true,
        message: 'Pares já configurados.'
      };
    }
  } catch (error) {
    console.error('❌ Erro ao configurar pares do botão brush:', error);
    return {
      success: false,
      message: 'Erro ao configurar pares: ' + error.message
    };
  }
}

/**
 * Recebe a solicitação da UI e chama o modal de seleção mágica.
 * @param {Array} promptsData - Array com os dados dos prompts do cliente
 */
function callMagicModal(promptsData) {
  console.log('[SB_server_buttonMagic] Iniciando modal mágico com', promptsData ? promptsData.length : 0, 'prompts');
  
  try {
    // Se os prompts não foram fornecidos, carrega do servidor
    const prompts = promptsData && promptsData.length > 0 ? promptsData : getAllPromptsForMagic();
    
    // NOVA LÓGICA: Buscar todos os H1s únicos do documento
    const uniqueHeadline1s = getAllUniqueDocumentHeadline1s();
    
    showMagicModal(prompts, uniqueHeadline1s);
  } catch (error) {
    console.error('[SB_server_buttonMagic] Erro ao abrir modal mágico:', error);
    throw error;
  }
}


/**
 * Obtém todos os H1s únicos do documento
 * @return {Array} Array de objetos no formato [{ headline1: "..." }]
 */
function getAllUniqueDocumentHeadline1s() {
  try {
    const documentManager = new DocumentHeadlineManager();
    const uniqueHeadline1s = documentManager.getUniqueHeadline1s();
    
    // Formatar para o formato esperado pelo cliente
    const formattedHeadline1s = uniqueHeadline1s.map(h1 => ({ headline1: h1 }));
    
    console.log('[SB_server_buttonMagic] Total de H1s únicos encontrados:', formattedHeadline1s.length);
    
    return formattedHeadline1s;
  } catch (error) {
    console.error('[SB_server_buttonMagic] Erro ao obter H1s únicos:', error);
    throw error;
  }
}

/**
 * Obtém todos os prompts disponíveis para o modal mágico
 * @return {Array} Array de objetos de prompt
 */
function getAllPromptsForMagic() {
  try {
    const sheetId = getSheetId();
    const documentId = getDocumentId();

    if (!sheetId || !documentId) {
      throw new Error('IDs de planilha ou documento não encontrados');
    }

    const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
    const data = dataFetcher.extractAllDataFromSources();
    
    // Converte o objeto allPrompts em um array de objetos
    const promptsArray = [];
    
    if (data.allPrompts) {
      Object.keys(data.allPrompts).forEach(key => {
        const prompt = data.allPrompts[key];
        promptsArray.push({
          optionText: key,
          fromHeadline2: prompt.fromHeadline2,
          toHeadline2: prompt.toHeadline2,
          promptContent: prompt.content
        });
      });
    }
    
    console.log('[SB_server_buttonMagic] Prompts carregados:', promptsArray.length);
    return promptsArray;
  } catch (error) {
    console.error('[SB_server_buttonMagic] Erro ao obter prompts:', error);
    throw error;
  }
}

/**
 * Exibe o modal mágico com duas colunas
 * @param {Array} prompts - Array de prompts
 * @param {Array} uniqueHeadline1s - Array de H1 únicos
 */
function showMagicModal(prompts, uniqueHeadline1s) {
  try {
    console.log('[SB_server_buttonMagic] Preparando modal com', prompts.length, 'prompts e', uniqueHeadline1s.length, 'H1 únicos');
    
    // Cria o template a partir do arquivo
    const template = HtmlService.createTemplateFromFile('AI_client_magic');
    
    // Anexa os dados como propriedades no template
    template.promptData = JSON.stringify(prompts);
    template.uniqueHeadline1sData = JSON.stringify(uniqueHeadline1s);
    
    // Avalia o template
    const html = template.evaluate()
      .setWidth(900)
      .setHeight(700);
    
    // Mostra o modal
    DocumentApp.getUi().showModalDialog(html, 'Processamento de IA');
    
    console.log('[SB_server_buttonMagic] Modal exibido com sucesso');
  } catch (error) {
    console.error('[SB_server_buttonMagic] Erro ao exibir modal:', error);
    throw error;
  }
}