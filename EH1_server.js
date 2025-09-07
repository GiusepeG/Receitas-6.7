/**
 * @OnlyCurrentDoc
 */

// Busca a chave de API das Propriedades do Script.
// Veja as instruções na documentação sobre como configurar esta propriedade.
// https://developers.google.com/apps-script/guides/properties
const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

/**
 * Função principal que cria e exibe a caixa de diálogo.
 */
function extractHeadlineOneFromText() {
  if (!apiKey) {
    DocumentApp.getUi().alert(
      'Chave de API não configurada',
      'Por favor, configure a GEMINI_API_KEY nas Propriedades do Projeto antes de usar esta função.',
      DocumentApp.getUi().ButtonSet.OK);
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('EH1_client.html')
      .setWidth(600)
      .setHeight(480);
  DocumentApp.getUi().showModalDialog(html, 'Filtrar e Listar Pacientes');
}

/**
 * Processa o texto, chama a API Gemini, ordena, filtra o resultado e o retorna.
 * @param {string} pastedText O texto bruto colado pelo usuário.
 * @param {string} filterType O critério de filtragem ('toda', 'matutino', 'vespertino').
 * @return {string} Uma string JSON com os dados prontos para serem escritos.
 */
function processPastedText(pastedText, filterType) {
  const modelName = 'gemini-1.5-flash-latest'; // Usando um modelo público estável.
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  const initialPrompt = `Você é um assistente de IA especializado em extrair dados de textos não estruturados.
Sua tarefa é extrair o horário do agendamento e o nome completo do paciente a partir do texto que será fornecido.

Analise cada agendamento no texto e siga estas regras para a extração:

1.  **Extraia o Horário:** O horário do agendamento está no formato HH:MM e aparece após a data e um código de setor (ex: JRE, US, YL).

2.  **Extraia o Nome do Paciente:** O nome do paciente está geralmente em letras maiúsculas.
    * Na maioria dos casos, o nome aparece diretamente após a informação do convênio (ex: 'IOP-UNIMED', 'PARTICULAR', 'MEDPREV').
    * **REGRA ESPECIAL:** Se o nome do paciente for "- NAO CADASTRADO -", o nome real estará na coluna de observações (Obs.). Neste caso, você deve procurar por um nome completo em letras maiúsculas na mesma linha. Frequentemente, este nome é seguido por um hífen e o nome de um funcionário (ex: "JOANE CRISTINE MARTINS GONCALO - NAIANE"). Você deve extrair apenas o nome completo do paciente.

3.  **Formato da Saída:** Apresente os resultados em uma lista, com cada item contendo o horário e o nome do paciente extraído.

**Exemplo 1:**
**Texto:** \`16/06/2025 JRE 14:10 CN PARTICULAR FRANCISCO PEREIRA GAMA 85 CONSULTA C2\`
**Saída Esperada:**
* Horário: 14:10
* Paciente: FRANCISCO PEREIRA GAMA

**Exemplo 2:**
**Texto:** \`16/06/2025 JRE 15:30 IOP-UNIMED - NAO CADASTRADO - 50 CONSULTA C1 JOANE CRISTINE MARTINS GONCALO - NAIANE\`
**Saída Esperada:**
* Horário: 15:30
* Paciente: JOANE CRISTINE MARTINS GONCALO

Agora, processe o texto a seguir e extraia o horário e o nome de cada paciente.`;

  const fullPrompt = initialPrompt + "\n\n### Texto para Processar:\n" + pastedText;
  const schema = { "type": "object", "properties": { "agendamento": { "type": "array", "items": { "type": "object", "properties": { "horario": { "type": "string" }, "paciente": { "type": "string" } }, "required": ["horario", "paciente"] } } }, "required": ["agendamento"] };
  const payload = { "contents": [{ "parts": [{ "text": fullPrompt }] }], "generationConfig": { "responseMimeType": "application/json", "responseSchema": schema } };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const parsedResponse = JSON.parse(responseBody);
      const extractedText = parsedResponse.candidates[0].content.parts[0].text;
      let finalJsonObject = JSON.parse(extractedText);
      let agendamentos = finalJsonObject.agendamento || [];

      // 1. ORDENAR: Sempre ordenar por horário.
      agendamentos.sort((a, b) => a.horario.localeCompare(b.horario));

      // 2. FILTRAR: Filtrar conforme a escolha do usuário.
      if (filterType === 'matutino') {
        agendamentos = agendamentos.filter(item => item.horario <= '12:59');
      } else if (filterType === 'vespertino') {
        agendamentos = agendamentos.filter(item => item.horario > '12:59');
      }
      // Para 'toda', nenhum filtro é aplicado.

      finalJsonObject.agendamento = agendamentos;
      
      return JSON.stringify(finalJsonObject);
    } else {
      throw new Error(`Erro na API: ${responseCode} - ${responseBody}`);
    }
  } catch (e) {
    Logger.log(`Exceção capturada: ${e.toString()}`);
    throw new Error(`Ocorreu uma exceção: ${e.message}`);
  }
}

/**
 * Limpa o corpo do documento ativo e escreve os nomes dos pacientes filtrados.
 * @param {string} processedJsonString A string JSON com os dados já ordenados e filtrados.
 * @return {string} Uma mensagem de sucesso para ser exibida ao usuário.
 */
function writeToDocument(processedJsonString) {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const data = JSON.parse(processedJsonString);
    const agendamentos = data.agendamento;

    body.clear();
    
    if (agendamentos && agendamentos.length > 0) {
      agendamentos.forEach(item => {
        body.appendParagraph("Nome: " + item.paciente);
        body.appendParagraph("");
      });
      const successMessage = `Documento atualizado com ${agendamentos.length} nomes.`;
      return successMessage;
    } else {
      body.appendParagraph("Nenhum paciente encontrado para o filtro selecionado.");
      return "Nenhum dado encontrado para escrever no documento.";
    }
  } catch (e) {
    Logger.log("Erro ao escrever no documento: " + e.toString());
    throw new Error("Falha ao escrever dados no documento. Verifique os logs.");
  }
}