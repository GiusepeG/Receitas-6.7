// Função principal para processar pares selecionados com IA
function processAISelectedPairsWithPrompt(selectedPairs) {

  Logger.log(`[AI] ✅ Processando pares selecionados com IA: ${selectedPairs.length}`);

  const doc = DocumentApp.getActiveDocument();
  const bodyText = doc.getBody().getText();

  // Obtém regras de formatação
  const rules = getFormattingRules();
  const headline1Rules = rules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
  const headline2Rules = rules.filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);

  // Cria estrutura do documento
  const documentManager = new DocumentHeadlineManager(bodyText, null, headline1Rules, headline2Rules);

  let processedCount = 0;
  const totalPairs = selectedPairs.length;

  // Processa cada par selecionado
  selectedPairs.forEach((pair, index) => {
    try {
      // Calcula o progresso
      const progress = Math.round(((index) / totalPairs) * 100);
      const status = `[AI] Processando par ${index + 1} de ${totalPairs}: ${pair.headline1}`;

      // Monta o prompt final para IA
      const finalPrompt = `${pair.promptContent}\n\n${pair.headline1}\n${pair.toHeadline2}\n${pair.clinicalContent}`;

      // Chamada à IA (placeholder)
      // TODO: Integrar com modelo de IA real
      const aiResponse = callAIModelPlaceholder(pair.headline1, pair.fromHeadline2, finalPrompt);

      // Atualiza o documento usando o DocumentHeadlineManager
      documentManager.createOrUpdateBodyHeadline2(aiResponse, pair.headline1, pair.toHeadline2, pair.fromHeadline2);
      processedCount++;
    } catch (error) {
      console.error(`[AI] ❌ Erro ao processar par ${index + 1}:`, error);
    }
  });

  // Aplica formatação se pelo menos um par foi processado
  if (processedCount > 0) {
    try {
      // Atualiza o documento com as mudanças
      const newBodyText = documentManager.getText();

      Logger.log(`[AI] ✅ Documento atualizado com sucesso: ${newBodyText}`);

      doc.getBody().setText(newBodyText);
      // Executa a formatação padrão
      executeFormat();
    } catch (error) {
      console.error('[AI] ❌ Erro ao atualizar documento ou formatar:', error);
    }
  }
  return {
    success: true
  };
}

// Placeholder para chamada à IA (substitua por integração real)
function callAIModelPlaceholder(headline1, fromHeadline2, prompt) {
  // Apenas retorna o prompt para simulação
  return `# [IA SIMULADA]\n${prompt}`;
}

