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
 * @deprecated A lógica foi substituída por executeBrushFlow.
 */
function executeCompleteBrushFlowWithPairs() {
  throw new Error("Esta função está obsoleta e foi substituída por executeBrushFlow.");
}

/**
 * @deprecated A lógica foi substituída por executeBrushFlow.
 */
function openBrushMagicModal() {
  throw new Error("Esta função está obsoleta e foi substituída por executeBrushFlow.");
}

/**
 * Executa o fluxo de melhoria (Brush) de forma automática e sem interface.
 * @returns {Object} Objeto com o resultado da operação para o cliente.
 */
function executeBrushFlow() {
  console.log('[Brush Flow] Iniciando execução do Brush.');
  try {
    // Etapa 1: Validar se há apenas um H1 no documento
    const docManager = new DocumentHeadlineManager();
    const uniqueHeadline1s = docManager.getUniqueHeadline1s();

    if (uniqueHeadline1s.length !== 1) {
      const message = uniqueHeadline1s.length === 0
        ? 'Nenhum paciente (Headline1) encontrado no documento.'
        : `O modo Brush funciona apenas com um paciente (Headline1) por vez. Encontrados: ${uniqueHeadline1s.length}`;
      console.warn(`[Brush Flow] Validação de H1 falhou: ${message}`);
      return { success: false, message: message };
    }
    const headline1 = uniqueHeadline1s[0];
    console.log(`[Brush Flow] Paciente validado: ${headline1}`);

    // Etapa 2: Carregar e mapear todos os prompts por prefixo
    const promptsMap = _getPromptsMap();
    if (promptsMap.size === 0) {
      return { success: false, message: 'Nenhum arquivo de prompt encontrado ou configurado.' };
    }
    console.log(`[Brush Flow] ${promptsMap.size} prompts carregados e mapeados.`);

    // Etapa 3: Ler e processar os pares da propriedade de script
    const properties = PropertiesService.getScriptProperties();
    const brushPairsString = properties.getProperty('brushButtonPairs');
    if (!brushPairsString) {
      return { success: false, message: 'A propriedade de script "brushButtonPairs" não está configurada.' };
    }
    // CORREÇÃO: Mudar a ordem de parsing para "prefixo, h2"
    const brushPairs = brushPairsString.split(';').map(p => {
      const [prefix, h2] = p.split(',').map(s => s.trim());
      return { promptPrefix: prefix, fromH2: h2, toH2: h2 }; // fromH2 e toH2 são os mesmos
    });
    console.log(`[Brush Flow] ${brushPairs.length} pares de Brush configurados.`);

    // Etapa 4: Verificar quais seções (H2) existem para o paciente
    const docStructure = docManager.getDocumentHeadlinePairsAndContent();
    const existingH2s = new Set(docStructure
      .filter(item => item.headline1 === headline1)
      .map(item => item.headline2.normalize('NFC')));
    console.log(`[Brush Flow] Seções existentes para o paciente: ${[...existingH2s].join(', ')}`);

    // Etapa 5: Montar a lista de prompts que serão executados
    const promptsToExecute = [];
    for (const pair of brushPairs) {
      const normalizedFromH2 = pair.fromH2.normalize('NFC');
      if (existingH2s.has(normalizedFromH2)) {
        const normalizedPrefix = pair.promptPrefix.normalize('NFC');
        if (promptsMap.has(normalizedPrefix)) {
          const promptData = promptsMap.get(normalizedPrefix);
          promptsToExecute.push({
            fromHeadline2: pair.fromH2, // O H2 que deve existir
            toHeadline2: promptData.toHeadline2, // O H2 que será criado/atualizado (do nome do arquivo)
            promptContent: promptData.content,
            optionText: `Brush: ${promptData.toHeadline2}` // Nome para referência interna
          });
          console.log(`[Brush Flow] Prompt agendado: ${normalizedPrefix} para a seção ${pair.fromH2}`);
        } else {
          console.warn(`[Brush Flow] Seção "${pair.fromH2}" encontrada, mas o prompt com prefixo "${pair.promptPrefix}" (normalizado: "${normalizedPrefix}") não foi localizado.`);
        }
      }
    }

    if (promptsToExecute.length === 0) {
      return { success: false, message: 'Nenhuma seção elegível para melhoria foi encontrada no documento para este paciente.' };
    }

    // Etapa 6: Executar o processamento em lote
    console.log(`[Brush Flow] Enviando ${promptsToExecute.length} prompts para processamento.`);
    const processingData = {
      headline1s: [{ headline1: headline1 }],
      prompts: promptsToExecute,
      requireAugmentedContext: true // O Brush sempre usa o contexto completo do H1
    };

    const result = processSequentialPromptsForPairs(processingData);
    if (result.success) {
      const successMessage = `Processamento Brush concluído com sucesso! ${promptsToExecute.length} seção(ões) atualizada(s).`;
      console.log(`[Brush Flow] ${successMessage}`);
      return { success: true, message: successMessage };
    } else {
      throw new Error(result.message || 'Erro desconhecido durante o processamento sequencial.');
    }

  } catch (error) {
    console.error('[Brush Flow] Erro fatal durante a execução:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Helper para buscar todos os prompts e retorná-los como um Map por prefixo.
 * @returns {Map<string, Object>} Mapa onde a chave é o prefixo do prompt e o valor são os dados do prompt.
 * @private
 */
function _getPromptsMap() {
  const allPrompts = getAllPromptsForMagic(); // Reutiliza a função existente
  const promptsMap = new Map();

  for (const prompt of allPrompts) {
    // CORREÇÃO: a chave do mapa é o optionText, que já é o prefixo do prompt.
    const prefix = prompt.optionText;
    promptsMap.set(prefix, {
      toHeadline2: prompt.toHeadline2,
      content: prompt.promptContent
    });
  }
  return promptsMap;
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