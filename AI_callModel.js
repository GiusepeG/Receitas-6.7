/**
 * Chama o modelo Gemini 2.0 Flash-Lite (Google) via REST.
 * @param {string} headline1 – título H1 usado no início do prompt
 * @param {string} headline2 – título H2 usado no início do prompt
 * @param {string} finalPrompt – prompt completo montado
 * @param {boolean} isTest – modo stub (devolve texto fixo) - opcional
 * @return {string} – resposta do modelo
 */
function callAIModel(finalPrompt, isTest = false) {

  if (isTest) return "<texto da IA (stub)>";


  // ----- parâmetros do Gemini ----------------------------------------------
  // Tenta obter a chave da API das Script Properties primeiro, depois do arquivo
  let apiKey;
  try {
    // Primeiro tenta das Script Properties
    apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      // Se não encontrar, tenta do arquivo (função que pode não existir)
      try {
        apiKey = getContentFromFile('GEMINI_KEY.txt').trim();
      } catch (e) {
        Logger.log("❌ Erro ao ler arquivo GEMINI_KEY.txt:", e.message);
        throw new Error("Chave da API Gemini não encontrada. Configure a propriedade GEMINI_API_KEY nas Script Properties.");
      }
    }
  } catch (e) {
    Logger.log("❌ Erro ao obter chave da API:", e.message);
    throw new Error("Chave da API Gemini não configurada. Configure a propriedade GEMINI_API_KEY nas Script Properties.");
  }
  
  const modelId  = "gemini-2.5-flash-lite";                     // custo ≈ $0,075 / Mi tokens de entrada
  const url      = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;

  const requestBody = {
    contents: [
      {
        parts: [{ text: finalPrompt }]   // formato exigido pelo Gemini
      }
    ],
    generationConfig: {
      temperature: 0,
      maxOutputTokens: 3500             // similar ao max_tokens do OpenAI
    }
  };

  // ----- chamada HTTPS ------------------------------------------------------
  const response = UrlFetchApp.fetch(url, {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(requestBody)
  });

  const json = JSON.parse(response.getContentText());

  // Estrutura da resposta:
  // { candidates: [ { content: { parts: [ { text: "..." } ] } } ] }
  const responseText = json.candidates[0].content.parts[0].text;

  return responseText;
}
