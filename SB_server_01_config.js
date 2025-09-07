/**
 * SERVER STAGE 1: CONFIGURATION & VALIDATION
 * Handles all configuration, validation, and property setup
 */

/**
 * Configures and validates all script properties
 * @returns {Object} Configuration result with success status
 */
function initializeServerConfig() {
  console.log('📋 Etapa 1: Configuração do Servidor');
  
  try {
    // Validate existing properties
    if (!validateScriptProperties()) {
      console.log('⚙️ Configurando propriedades...');
      setupScriptProperties();
      
      if (!validateScriptProperties()) {
        return {
          success: false,
          error: 'Não foi possível configurar as Propriedades do Script automaticamente'
        };
      }
    }
    
    // Get and validate all required IDs
    const config = {
      sheetId: getSheetId(),
      documentId: getDocumentId(),
      promptsFolderId: getPromptsFolderId()
    };
    
    // Validate all IDs are present
    if (!config.sheetId || !config.documentId || !config.promptsFolderId) {
      return {
        success: false,
        error: 'IDs de configuração não encontrados'
      };
    }
    
    console.log('✅ Configuração validada');
    return {
      success: true,
      config: config
    };
    
  } catch (error) {
    console.error('❌ Erro na configuração:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets sheet ID from script properties
 * @returns {string} Sheet ID
 */
function getSheetId() {
  return PropertiesService.getScriptProperties().getProperty('sheetId');
}

/**
 * Gets document ID from script properties
 * @returns {string} Document ID
 */
function getDocumentId() {
  return PropertiesService.getScriptProperties().getProperty('documentId');
}

/**
 * Gets prompts folder ID from script properties
 * @returns {string} Prompts folder ID
 */
function getPromptsFolderId() {
  return PropertiesService.getScriptProperties().getProperty('promptsFolderId');
}

/**
 * Validates script properties
 * @returns {boolean} True if valid
 */
function validateScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  const requiredProps = ['sheetId', 'documentId', 'promptsFolderId'];
  
  return requiredProps.every(prop => {
    const value = properties.getProperty(prop);
    return value && value.trim() !== '';
  });
}

/**
 * Sets up script properties
 */
function setupScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  
  // Set default values if not exist
  if (!properties.getProperty('sheetId')) {
    properties.setProperty('sheetId', 'YOUR_SHEET_ID_HERE');
  }
  
  if (!properties.getProperty('documentId')) {
    properties.setProperty('documentId', 'YOUR_DOCUMENT_ID_HERE');
  }
  
  if (!properties.getProperty('promptsFolderId')) {
    properties.setProperty('promptsFolderId', 'YOUR_PROMPTS_FOLDER_ID_HERE');
  }
} 