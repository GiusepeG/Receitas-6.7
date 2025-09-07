// =========================================================================
// FUNÇÕES DE INSERÇÃO E FORMATAÇÃO
// =========================================================================

/**
 * Ponto de entrada para inserção. Verifica se há texto selecionado.
 * Se houver, abre um diálogo modeless. Se não, insere o texto diretamente.
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {string} status da operação
 */
function insertTextAndFormat(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
  
    // Se o texto estiver vazio, apenas formata se solicitado.
    if (!textToInsert) {
      if (shouldFormat) {
        executeFormat();
      }
      return "FORMATTED_ONLY";
    }

    const selection = DocumentApp.getActiveDocument().getSelection();
  
    // Verifica se há texto selecionado
    if (selection && selection.getRangeElements().length > 0) {
      
      // Prepara as propriedades para a comunicação entre o diálogo e a sidebar
      initializeDialogState();
      
      const htmlTemplate = HtmlService.createTemplateFromFile('SB_modelessDialog');
      htmlTemplate.textToInsert = textToInsert;
      htmlTemplate.shouldFormat = shouldFormat;
      
      const ui = DocumentApp.getUi();
      ui.showModelessDialog(htmlTemplate.evaluate().setWidth(300).setHeight(200), 'Posicione o Cursor');
      
      return "DIALOG_OPENED";
      
    } else {
      // Se não há seleção, insere diretamente
      const success = insertTextAtCursor(jsonObj);
      return success ? "INSERTED_DIRECTLY" : "INSERTION_FAILED";
    }
}

/**
 * Ponto de entrada para append. Sempre adiciona o texto ao final do documento.
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {string} status da operação
 */
function appendTextAndFormat(jsonObj) {
    const obj = JSON.parse(jsonObj);
    const textToInsert = obj.text;
    const shouldFormat = obj.shouldFormat;
  
    // Se o texto estiver vazio, apenas formata se solicitado.
    if (!textToInsert) {
      if (shouldFormat) {
        executeFormat();
      }
      return "FORMATTED_ONLY";
    }

    // Sempre adiciona ao final do documento, independente de seleção
    const success = appendTextAtEnd(jsonObj);
    return success ? "INSERTED_DIRECTLY" : "INSERTION_FAILED";
}

/**
 * Adiciona o texto ao final do documento.
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {boolean} true se bem-sucedido
 */
function appendTextAtEnd(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Sempre adiciona ao final do documento
  try {
    body.appendParagraph(textToInsert);
  } catch (e) {
    DocumentApp.getUi().alert('Erro ao adicionar texto ao final do documento: ' + e.message);
    return false;
  }

  // Formata o documento se solicitado
  if (shouldFormat) {
    executeFormat();
  }

  return true;
}

/**
 * Insere o texto na posição atual do cursor no documento.
 * Esta função é chamada diretamente ou pelo diálogo modeless.
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {boolean} true se bem-sucedido
 */
function insertTextAtCursor(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  // Garante que não há texto selecionado no momento da inserção
  if (selection && selection.getRangeElements().length > 0) {
    throw new Error('Ainda há texto selecionado. Por favor, posicione o cursor sem selecionar nenhum texto e tente novamente.');
  }

  const body = doc.getBody();
  const cursor = doc.getCursor();

  // Tenta inserir no cursor
  if (cursor) {
    try {
      const element = cursor.insertText(textToInsert + '\n');
      if (!element) {
        body.appendParagraph(textToInsert);
        DocumentApp.getUi().alert('O cursor estava em uma posição inválida. O texto foi inserido ao final do documento.');
      }
    } catch (e) {
      try {
        body.appendParagraph(textToInsert);
        DocumentApp.getUi().alert('Ocorreu um erro ao inserir no cursor. O texto foi adicionado ao final do documento.');
      } catch (e2) {
        DocumentApp.getUi().alert('Erro crítico ao inserir texto: ' + e2.message);
        return false;
      }
    }
  } else {
    // Insere ao final se o cursor não estiver disponível
    try {
      body.appendParagraph(textToInsert);
    } catch (e) {
      DocumentApp.getUi().alert('Erro ao inserir texto: ' + e.message);
      return false;
    }
  }

  // Formata o documento se solicitado
  if (shouldFormat) {
    executeFormat();
  }

  return true;
}

/**
 * Insere texto na posição do cursor sem abrir diálogo (versão simplificada)
 * Usado pelo botão cursor-text que já verificou se há seleção
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {boolean} true se bem-sucedido
 */
function insertTextAtCursorOnly(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  // Verifica novamente se há texto selecionado (segurança)
  if (selection && selection.getRangeElements().length > 0) {
    throw new Error('Ainda há texto selecionado. Por favor, posicione o cursor sem selecionar nenhum texto e tente novamente.');
  }

  const cursor = doc.getCursor();

  // Se não há cursor disponível, cancela a operação
  if (!cursor) {
    throw new Error('Cursor não disponível. Posicione o cursor no documento e tente novamente.');
  }

  // Tenta inserir no cursor
  try {
    const element = cursor.insertText(textToInsert + '\n');
    if (!element) {
      throw new Error('Não foi possível inserir o texto na posição do cursor. Tente posicionar o cursor em outro local.');
    }
  } catch (e) {
    throw new Error('Erro ao inserir texto na posição do cursor: ' + e.message);
  }

  // Formata o documento se solicitado (embora cursor-text não formate)
  if (shouldFormat) {
    executeFormat();
  }

  return true;
}

/**
 * Adiciona texto diretamente ao final do documento (versão simplificada)
 * Usado pelo botão append quando não há H1s no documento
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {boolean} true se bem-sucedido
 */
function appendTextDirectlyToEnd(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Adiciona texto ao final do documento
  try {
    body.appendParagraph(textToInsert);
  } catch (e) {
    throw new Error('Erro ao adicionar texto ao final do documento: ' + e.message);
  }

  // Formata o documento se solicitado
  if (shouldFormat) {
    executeFormat();
  }

  return true;
}

/**
 * Executa o append do texto no H1 selecionado pelo usuário
 * @param {string} selectedH1 - O H1 selecionado pelo usuário
 * @param {string} textToAppend - O texto a ser adicionado
 * @param {boolean} shouldFormat - Se deve formatar após adicionar
 * @return {Object} Resultado da operação
 */
function executeAppendToSelectedH1(selectedH1, textToAppend, shouldFormat) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const bodyText = body.getText();
    
    // Usar o mesmo processamento que a validação usa
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
    
    // Usar buildDocumentStructureWithH1Only para obter os H1s da mesma forma que a validação
    const documentStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    const uniqueH1s = [...new Set(documentStructure.map(item => item.headline1).filter(Boolean))];
    
    // Verificar se o H1 selecionado existe no documento processado
    if (!uniqueH1s.includes(selectedH1)) {
      return {
        success: false,
        message: `O Headline1 "${selectedH1}" não foi encontrado no documento.`
      };
    }
    
    // Modificar o texto processado para incluir o novo conteúdo no H1 específico
    const modifiedText = insertTextIntoSpecificH1(processedText, selectedH1, textToAppend, heading1Rules, heading2Rules);
    
    // Atualizar o documento com o texto modificado
    body.setText(modifiedText);
    
    // Formatar o documento se solicitado
    if (shouldFormat) {
      executeFormat();
    }
    
    return {
      success: true,
      message: `Texto adicionado com sucesso ao paciente "${selectedH1}".`
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Erro ao adicionar texto: ' + error.message
    };
  }
}

/**
 * Insere texto especificamente em um H1 no texto processado
 * @param {string} processedText - Texto já processado
 * @param {string} targetH1 - H1 onde inserir o texto
 * @param {string} textToInsert - Texto a ser inserido
 * @param {Array} heading1Rules - Regras para H1
 * @param {Array} heading2Rules - Regras para H2
 * @return {string} Texto modificado
 */
function insertTextIntoSpecificH1(processedText, targetH1, textToInsert, heading1Rules, heading2Rules) {
  const lines = processedText.split('\n');
  const modifiedLines = [];
  let foundTargetH1 = false;
  let targetH1EndIndex = -1;
  
  // Primeira passada: encontrar onde termina o conteúdo do H1 target
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Verificar se é um H1
    if (heading1Rules.some(rule => rule.condition(line))) {
      if (foundTargetH1) {
        // Encontramos o próximo H1, então o H1 target termina na linha anterior
        targetH1EndIndex = i - 1;
        break;
      }
      
      if (line === targetH1) {
        foundTargetH1 = true;
      }
    }
  }
  
  // Se não encontramos outro H1 após o target, o target vai até o final
  if (foundTargetH1 && targetH1EndIndex === -1) {
    targetH1EndIndex = lines.length - 1;
  }
  
  // Segunda passada: construir o texto modificado
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Se chegamos ao final do H1 target, inserir o texto
    if (i === targetH1EndIndex) {
      modifiedLines.push(line);
      
      // Verificar se a última linha não está vazia para adicionar uma linha em branco
      if (line.trim() !== '') {
        modifiedLines.push('');
      }
      
      // Inserir o texto
      modifiedLines.push(textToInsert);
      continue;
    }
    
    modifiedLines.push(line);
  }
  
  // Se não conseguimos inserir o texto (H1 não foi encontrado), adicionar ao final
  if (foundTargetH1 && targetH1EndIndex === -1) {
    modifiedLines.push('');
    modifiedLines.push(textToInsert);
  }
  
  return modifiedLines.join('\n');
}

/**
 * Verifica se há seleção e insere texto na posição do cursor em uma única operação
 * Usado pelo botão cursor-text para otimizar e reduzir chamadas ao servidor
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {Object} Objeto com status da operação e mensagem
 */
function checkSelectionAndInsertAtCursor(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  // Verifica se há texto selecionado
  if (selection && selection.getRangeElements().length > 0) {
    return {
      success: false,
      hasSelection: true,
      message: 'Há texto selecionado no documento. Posicione o cursor sem selecionar texto e tente novamente.'
    };
  }

  const cursor = doc.getCursor();

  // Se não há cursor disponível, cancela a operação
  if (!cursor) {
    return {
      success: false,
      hasSelection: false,
      message: 'Cursor não disponível. Posicione o cursor no documento e tente novamente.'
    };
  }

  // Tenta inserir no cursor
  try {
    const element = cursor.insertText(textToInsert + '\n');
    if (!element) {
      return {
        success: false,
        hasSelection: false,
        message: 'Não foi possível inserir o texto na posição do cursor. Tente posicionar o cursor em outro local.'
      };
    }
  } catch (e) {
    return {
      success: false,
      hasSelection: false,
      message: 'Erro ao inserir texto na posição do cursor: ' + e.message
    };
  }

  // Formata o documento se solicitado (embora cursor-text não formate)
  if (shouldFormat) {
    try {
      executeFormat();
    } catch (e) {
      console.warn('Aviso: Erro ao formatar documento:', e.message);
    }
  }

  return {
    success: true,
    hasSelection: false,
    message: 'Texto inserido com sucesso na posição do cursor.'
  };
}

/**
 * Valida a estrutura do documento e executa o append apropriado em uma única operação
 * Usado pelo botão append para otimizar e reduzir chamadas ao servidor
 * @param {string} jsonObj - JSON com texto e flag de formatação
 * @return {Object} Objeto com status da operação e ações necessárias
 */
function validateAndExecuteAppend(jsonObj) {
  const obj = JSON.parse(jsonObj);
  const textToInsert = obj.text;
  const shouldFormat = obj.shouldFormat;

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
    
    // Usar buildDocumentStructureWithH1Only para obter os H1s únicos
    const documentStructure = buildDocumentStructureWithH1Only(processedText, heading1Rules, heading2Rules);
    const uniqueHeadline1s = [...new Set(documentStructure.map(item => item.headline1).filter(Boolean))];
    
    // Cenário 1: Nenhum H1 encontrado - executa append direto
    if (uniqueHeadline1s.length === 0) {
      try {
        const body = doc.getBody();
        body.appendParagraph(textToInsert);
        
        if (shouldFormat) {
          executeFormat();
        }
        
        return {
          success: true,
          action: 'append_direct',
          message: 'Texto adicionado ao final do documento (sem H1s).'
        };
      } catch (e) {
        return {
          success: false,
          action: 'append_direct',
          message: 'Erro ao adicionar texto ao final do documento: ' + e.message
        };
      }
    }
    
    // Cenário 2: Um H1 encontrado - executa append normal
    if (uniqueHeadline1s.length === 1) {
      try {
        const success = appendTextAtEnd(JSON.stringify(obj));
        return {
          success: success,
          action: 'append_normal',
          message: success ? 'Texto adicionado com sucesso.' : 'Falha ao adicionar texto.'
        };
      } catch (e) {
        return {
          success: false,
          action: 'append_normal',
          message: 'Erro ao adicionar texto: ' + e.message
        };
      }
    }
    
    // Cenário 3: Múltiplos H1s encontrados - precisa abrir diálogo
    return {
      success: false,
      action: 'show_dialog',
      uniqueHeadline1s: uniqueHeadline1s,
      textToInsert: textToInsert,
      shouldFormat: shouldFormat,
      message: `Foram encontrados ${uniqueHeadline1s.length} Headline1 diferentes no documento.`
    };
    
  } catch (error) {
    return {
      success: false,
      action: 'error',
      message: 'Erro ao analisar a estrutura do documento: ' + error.message
    };
  }
}
