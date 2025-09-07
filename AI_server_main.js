function showModalDocumentStructure(prompts) {
    const allPairs = getDocumentHeadlinePairs();
    // 1. Crie o template a partir do arquivo
    const template = HtmlService.createTemplateFromFile('AI_client_index');
  
    // 2. Anexe o array 'prompts' e 'allPairs' como propriedades no template.
    //    JSON.stringify() vai transformar os dados em uma string JSON.
    template.promptData = JSON.stringify(prompts);
    template.allPairsData = JSON.stringify(allPairs);
  
    // 3. Avalie o template.
    // O título do modal será uma concatenação dos 'optionText' de cada prompt.
    const title = prompts.map(p => p.optionText).join(' & ');
    const html = template.evaluate()
      .setWidth(500)
      .setHeight(600);
    
    // 4. Mostre o modal com os dados já injetados.
    DocumentApp.getUi().showModalDialog(html, title);
  }
  
  /**
   * Obtém os pares de Headline1 e Headline2 do documento
   * @return {Array} Array de objetos com headline1 e headline2
   */
  function getDocumentHeadlinePairs() {
    try {
      // Cria estrutura do documento (construtor agora obtém dados automaticamente)
      const documentManager = new DocumentHeadlineManager();
      
      // Usa o método da classe para obter os pares
      const pairs = documentManager.getDocumentHeadlinePairsAndContent();
      
      // Log adicional para verificar se a estrutura está vazia
      if (!pairs || pairs.length === 0) {
        console.warn('[AI_server_main] Atenção: Nenhum par H1/H2 foi retornado pelo DocumentHeadlineManager.');
        documentManager.logDocumentStructure('[AI_server_main] Estrutura do Documento (vazia)');
      }
      
      return pairs;
    } catch (error) {
      console.error('[AI_server_main] Erro fatal em getDocumentHeadlinePairs:', error);
      // Lança o erro novamente para que o cliente o receba na `onFailure`
      throw error;
    }
  }
  
  /**
   * Processa múltiplos pacientes selecionados com o prompt
   */
  function processSequentialPromptsForPairs(processingData) {
    const { headline1s, prompts, requireAugmentedContext } = processingData;

    if (!headline1s || headline1s.length === 0 || !prompts || prompts.length === 0) {
      console.error("Dados de processamento inválidos. Headline1s ou prompts ausentes.");
      return { success: false, message: "Dados de processamento inválidos." };
    }

    const doc = DocumentApp.getActiveDocument();
    Logger.log(`bodyText with empty one space lines like " ": ${doc.getBody().getText()}`);
    const documentManager = new DocumentHeadlineManager(); // Inicia com o conteúdo atual
    const oldBodyText = documentManager.getGroupedText();
    Logger.log(`bodyText without empty one space lines like " ": ${oldBodyText}`);
    let iterationCounter = 0; // Contador para o log

    try {
      // Itera sobre cada PROMPT (tarefa de IA, ex: "Corrigir Prontuário")
      prompts.forEach((prompt, promptIndex) => {
        const { fromHeadline2, toHeadline2, promptContent, optionText } = prompt;
        
        // Itera sobre cada PACIENTE (H1 único selecionado)
        headline1s.forEach(headline1Item => {
          const { headline1 } = headline1Item;
          let clinicalContent;

          if (requireAugmentedContext) {
            // --- NOVO CAMINHO: CONTEXTO AMPLIADO ---
            clinicalContent = documentManager.getAugmentedContentForH1(headline1);
            
          } else {
            // --- CAMINHO ANTIGO: CONTEÚDO ESPECÍFICO ---
            // Busca o conteúdo MAIS ATUALIZADO para o H2 de origem (fromHeadline2)
            clinicalContent = documentManager.getContentForPair(headline1, fromHeadline2);
          }

          // Pula este paciente se o conteúdo não for encontrado em nenhum dos caminhos
          if (clinicalContent === null || clinicalContent.trim() === '') {
            console.warn(`Atenção: Pulando paciente "${headline1}" para o prompt "${optionText}" porque nenhum conteúdo relevante foi encontrado.`);
            return; // 'continue' para o forEach do 'headline1s'
          }
          
          // Monta o clinicalContent final com o H1
          const clinicalContentWithHeadline1 = `${headline1}\n\n${clinicalContent}`;

          // Monta o prompt final
          const finalPrompt = `${promptContent}\n\n${clinicalContentWithHeadline1}`;
          
          // Chama a IA
          const newContent = callAIModel(finalPrompt);

          // Registra a interação no documento de log (opcional)
          iterationCounter++;
          // logInteractionToDoc(promptContent, headline1, fromHeadline2,toHeadline2, clinicalContentWithHeadline1, newContent, iterationCounter);
          
          // Atualiza a estrutura do documento EM MEMÓRIA com o novo conteúdo
          documentManager.createOrUpdateBodyHeadline2(newContent, headline1, toHeadline2, fromHeadline2);
        });
      });

      // Após todos os prompts serem executados, atualiza o documento e formata
      const newBodyText = documentManager.getGroupedText();

      //Logger.log(`newBodyText: ${newBodyText}`);

      doc.getBody().setText(newBodyText);
      
      executeFormat();
      
      return { success: true };

    } catch (error) {
      console.error("❌ Erro fatal durante o processamento sequencial:", error);
      return { success: false, message: error.message };
    }
  }
  
  /**
   * Atualiza o progresso no cliente (função auxiliar)
   * @param {number} progress - Porcentagem de progresso
   * @param {string} status - Mensagem de status
   */
  function updateClientProgress(progress, status) {
    // Esta função pode ser expandida para enviar atualizações em tempo real
    // Por enquanto, apenas loga o progresso
  }  

/**
 * Registra a interação da IA (prompt e resposta) em um documento de log do Google Docs.
 * @param {string} finalPrompt - O prompt completo enviado para o modelo de IA.
 * @param {string} newContent - O conteúdo gerado pela IA.
 * @param {number} iteration - O número da iteração para identificação no log.
 */
function logInteractionToDoc(promptContent, headline1, fromHeadline2, toHeadline2, clinicalContentWithHeadline1, newContent, iterationCounter) {
  try {
    // 1. Acessa a script property com o ID do documento de log
    let logDocFileId = PropertiesService.getScriptProperties().getProperty('logDocFileId');
    
    // 2. Se não estiver configurado, tenta configurar automaticamente
    if (!logDocFileId) {
      // Tenta usar o mesmo documento principal como fallback
      const docData = PropertiesService.getScriptProperties().getProperty('docData');
      if (docData) {
        PropertiesService.getScriptProperties().setProperty('logDocFileId', docData);
        logDocFileId = docData;
      } else {
        console.warn("A property 'logDocFileId' não está definida. Log de interação será pulado.");
        return;
      }
    }

    // 3. Acessa o Google Doc de log pelo ID
    const logDoc = DocumentApp.openById(logDocFileId);
    const body = logDoc.getBody();
    const iterationStr = String(iterationCounter).padStart(2, '0');

    // 4. Adiciona o header Datetime
    body.appendParagraph(`DateTime: ${new Date().toLocaleString()}, ${headline1}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // 5. Adiciona o header da iteração
    body.appendParagraph(`iterationStr: ${iterationStr}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(`fromHeadline2: ${fromHeadline2}`);
    body.appendParagraph(`toHeadline2: ${toHeadline2}`);
    
    // 6. Adiciona o promptContent
    body.appendParagraph('promptContent').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    //body.appendParagraph(promptContent);

    // 7. Adiciona o clinicalContent
    body.appendParagraph('clinicalContent').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(clinicalContentWithHeadline1);

    // 8. Adiciona o newContent
    body.appendParagraph('newContent').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(newContent);

  } catch (e) {
    console.error(`Falha ao registrar a interação no documento de log. Erro: ${e.toString()}`);
    // Não lança o erro para não interromper o processo principal.
  }
}  
