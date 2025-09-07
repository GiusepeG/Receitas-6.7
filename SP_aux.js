
/**
 * Obtém o ID do documento das propriedades
 * @return {string|null} ID do documento ou null
 */
function getDocumentId() {
    const SP = PropertiesService.getScriptProperties();
    const documentId = SP.getProperty('docData');
    
    if (!documentId) {
      console.error('❌ ID do documento não encontrado');
      return null;
    }
    
    return documentId;
  }
  
  /**
   * Obtém o ID da planilha das propriedades
   * @return {string|null} ID da planilha ou null
   */
  function getSheetId() {
    const SP = PropertiesService.getScriptProperties();
    const sheetId = SP.getProperty('sheetData');
    
    if (!sheetId) {
      console.error('❌ ID da planilha não encontrado');
      return null;
    }
    
    return sheetId;
  }
  
  /**
   * Valida se as propriedades necessárias estão configuradas
   * @return {boolean} true se todas configuradas
   */
  function validateScriptProperties() {
      const SP = PropertiesService.getScriptProperties();
      const sheetDataId = SP.getProperty('sheetData');
      const docDataId = SP.getProperty('docData');
      const logDocFileId = SP.getProperty('logDocFileId');
      
      if (!sheetDataId || !docDataId) {
        console.error('❌ Propriedades principais não configuradas (sheetData ou docData)');
        return false;
      }
      
      if (!logDocFileId) {
        console.warn('⚠️ Propriedade logDocFileId não configurada - logs de interação não funcionarão');
        // Não retorna false pois o log não é crítico para o funcionamento principal
      }
      
      return true;
  }
  
  /**
   * Configura propriedades com valores padrão
   */
  function setupScriptProperties() {
      const SP = PropertiesService.getScriptProperties();
      
      // IDs padrão - você pode alterar estes valores
      const defaultSheetData = "1uaStYZhSDK7tXtBsJlLJrCm59RQwBjKV7keLVFME2nM";
      const defaultDocData = "1wA4vP5qoUnNpvR7YxrV3nCGEwwKGjJ6gzgZfgLQJnrM";
      // ID padrão para documento de log de interações da IA - você pode alterar este valor
      const defaultLogDocFileId = "1wA4vP5qoUnNpvR7YxrV3nCGEwwKGjJ6gzgZfgLQJnrM"; // Mesmo que docData por padrão
      
      if (!SP.getProperty('sheetData')) {
        SP.setProperty('sheetData', defaultSheetData);
      }
      
      if (!SP.getProperty('docData')) {
        SP.setProperty('docData', defaultDocData);
      }
      
      // Configura o ID do documento de log se não existir
      if (!SP.getProperty('logDocFileId')) {
        SP.setProperty('logDocFileId', defaultLogDocFileId);
        console.log("✅ Propriedade logDocFileId configurada com valor padrão:", defaultLogDocFileId);
      }
      
      validateScriptProperties();
  }
  
  /**
   * Configura especificamente o ID do documento de log
   * @param {string} logDocFileId - ID do documento para log (opcional)
   * @return {boolean} true se configurado com sucesso
   */
  function setupLogDocumentId(logDocFileId = null) {
      const SP = PropertiesService.getScriptProperties();
      
      if (!logDocFileId) {
          // Se não fornecido, usa o mesmo ID do documento principal por padrão
          logDocFileId = SP.getProperty('docData');
          
          if (!logDocFileId) {
              console.error('❌ Não foi possível configurar logDocFileId: docData não encontrado');
              return false;
          }
      }
      
      try {
          // Testa se o documento existe e se temos acesso
          const testDoc = DocumentApp.openById(logDocFileId);
          testDoc.getName(); // Testa acesso
          
          SP.setProperty('logDocFileId', logDocFileId);
          console.log('✅ logDocFileId configurado com sucesso:', logDocFileId);
          return true;
          
      } catch (error) {
          console.error('❌ Erro ao configurar logDocFileId:', error.message);
          console.error('Verifique se o ID do documento está correto e se você tem acesso a ele');
          return false;
      }
  }
  
  /**
   * Cria um novo documento especificamente para logs de interação da IA
   * @return {string|null} ID do documento criado ou null se falhou
   */
  function createLogDocument() {
      try {
          const newDoc = DocumentApp.create('Log de Interações IA - ' + new Date().toLocaleDateString());
          const newDocId = newDoc.getId();
          
          // Adiciona conteúdo inicial ao documento
          const body = newDoc.getBody();
          body.appendParagraph('Log de Interações da IA').setHeading(DocumentApp.ParagraphHeading.TITLE);
          body.appendParagraph('Este documento registra todas as interações entre os prompts e a IA.')
              .setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
          body.appendParagraph('Criado automaticamente em: ' + new Date().toLocaleString());
          body.appendParagraph(''); // Linha em branco
          
          console.log('✅ Novo documento de log criado:', newDocId);
          
          // Configura automaticamente como logDocFileId
          if (setupLogDocumentId(newDocId)) {
              return newDocId;
          } else {
              return null;
          }
          
      } catch (error) {
          console.error('❌ Erro ao criar documento de log:', error.message);
          return null;
      }
  }
  
  /**
   * Obtém prompts filtrados por fromHeadline2 (função de servidor para fallback)
   * @param {string} fromHeadline2 - Valor do fromHeadline2 para filtrar
   * @return {Array} Array com prompts filtrados
   */
  function getFilteredPromptsFromServer(fromHeadline2) {
    try {
      // Obtém os IDs necessários das propriedades (mesmo padrão de getSidebarData)
      const sheetId = getSheetId();
      const documentId = getDocumentId();
  
      if (!sheetId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID da planilha (sheetData) não encontrado nas propriedades do script.';
        console.error(errorMsg);
        return [];
      }
  
      if (!documentId) {
        const errorMsg = '🛑 FALHA CRÍTICA: ID do Documento (docData) não encontrado nas propriedades do script.';
        console.error(errorMsg);
        return [];
      }
  
      // Cria DataFetcher com os IDs corretos (mesmo padrão de getSidebarData)
      const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
      const sidebarData = dataFetcher.extractAllDataFromSources();
      const allPrompts = sidebarData.allPrompts || {};
      
      const filteredPrompts = [];
      
      Object.keys(allPrompts).forEach(promptKey => {
        const prompt = allPrompts[promptKey];
        if (prompt.fromHeadline2 === fromHeadline2) {
          filteredPrompts.push({
            key: promptKey,
            title: promptKey,
            fromHeadline2: prompt.fromHeadline2,
            toHeadline2: prompt.toHeadline2,
            content: prompt.content
          });
        }
      });
      
      console.log('✅ Prompts filtrados obtidos do servidor:', filteredPrompts.length);
      return filteredPrompts;
      
    } catch (error) {
      console.error('❌ Erro ao obter prompts filtrados do servidor:', error);
      return [];
    }
  }
  