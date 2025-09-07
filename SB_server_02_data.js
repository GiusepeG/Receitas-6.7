/**
 * SERVER STAGE 2: DATA PROCESSING
 * Handles all data fetching, processing, and caching
 */

/**
 * Fetches and processes all sidebar data
 * @param {Object} config - Configuration object from stage 1
 * @returns {Object} Processed data result
 */
function processSidebarData(config) {
  console.log('📊 Etapa 2: Processamento de Dados');
  
  try {
    console.time('⏱️ Busca de dados');
    
    // Create data fetcher with configuration
    const dataFetcher = new StreamlinedDataFetcher(config);
    
    // Fetch all data in optimized way
    const processedData = dataFetcher.fetchAllData();
    
    console.timeEnd('⏱️ Busca de dados');
    
    if (processedData.error) {
      return {
        success: false,
        error: processedData.error
      };
    }
    
    console.log(`✅ Dados processados: ${processedData.items?.length || 0} itens, ${processedData.promptTitles?.length || 0} prompts`);
    
    return {
      success: true,
      data: processedData
    };
    
  } catch (error) {
    console.error('❌ Erro no processamento:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Streamlined Data Fetcher Class
 * Optimized single-responsibility data fetcher
 */
class StreamlinedDataFetcher {
  constructor(config) {
    this.config = config;
    this.cacheManager = new DataCacheManager();
  }
  
  /**
   * Fetches all data with caching
   * @returns {Object} All fetched data
   */
  fetchAllData() {
    // Try cache first
    const cachedData = this.cacheManager.getCachedData();
    if (cachedData) {
      console.log('💾 Dados carregados do cache');
      return cachedData;
    }
    
    // Fetch fresh data
    console.log('🔄 Buscando dados frescos');
    const freshData = this.fetchFreshData();
    
    // Cache the fresh data
    this.cacheManager.cacheData(freshData);
    
    return freshData;
  }
  
  /**
   * Fetches fresh data from sources
   * @returns {Object} Fresh data
   */
  fetchFreshData() {
    const sheets = this.extractSheetsData();
    const prompts = this.extractPromptsData();
    const items = this.extractItemsFromSheets(sheets);
    
    return {
      sheets: sheets,
      items: items,
      promptTitles: prompts.titles,
      allPrompts: prompts.all
    };
  }
  
  /**
   * Extracts sheets data optimized
   * @returns {Array} Sheets data
   */
  extractSheetsData() {
    const spreadsheet = SpreadsheetApp.openById(this.config.sheetId);
    const sheets = spreadsheet.getSheets();
    
    return sheets.map(sheet => ({
      sheetName: sheet.getName(),
      sheetButton: this.extractSheetButton(sheet.getName()),
      sheetTag: this.extractSheetTag(sheet.getName()),
      sheetStyle: this.extractSheetStyle(sheet.getName())
    }));
  }
  
  /**
   * Extracts items from sheets with batch processing
   * @param {Array} sheets - Sheets configuration
   * @returns {Array} All items
   */
  extractItemsFromSheets(sheets) {
    const spreadsheet = SpreadsheetApp.openById(this.config.sheetId);
    const allItems = [];
    
    sheets.forEach(sheetConfig => {
      const sheet = spreadsheet.getSheetByName(sheetConfig.sheetName);
      if (sheet) {
        const items = this.extractSheetItemsOptimized(sheet, sheetConfig);
        allItems.push(...items);
      }
    });
    
    return allItems;
  }
  
  /**
   * Extracts prompts data optimized
   * @returns {Object} Prompts data
   */
  extractPromptsData() {
    const folder = DriveApp.getFolderById(this.config.promptsFolderId);
    const files = folder.getFiles();
    
    const titles = [];
    const all = {};
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith('.txt')) {
        const promptData = this.extractPromptDataOptimized(file);
        titles.push(promptData.title);
        all[promptData.title] = promptData.content;
      }
    }
    
    return { titles, all };
  }
  
  /**
   * Extracts sheet button from name
   * @param {string} sheetName - Sheet name
   * @returns {string} Button text
   */
  extractSheetButton(sheetName) {
    const match = sheetName.match(/\[(.*?)\]/);
    return match ? match[1] : sheetName;
  }
  
  /**
   * Extracts sheet tag from name
   * @param {string} sheetName - Sheet name
   * @returns {string} Tag text
   */
  extractSheetTag(sheetName) {
    const match = sheetName.match(/\{(.*?)\}/);
    return match ? match[1] : sheetName.toLowerCase();
  }
  
  /**
   * Extracts sheet style from name
   * @param {string} sheetName - Sheet name
   * @returns {string} Style class
   */
  extractSheetStyle(sheetName) {
    const match = sheetName.match(/\((.*?)\)/);
    return match ? match[1] : 'primary';
  }
  
  /**
   * Extracts sheet items optimized
   * @param {Sheet} sheet - Google Sheet
   * @param {Object} sheetConfig - Sheet configuration
   * @returns {Array} Sheet items
   */
  extractSheetItemsOptimized(sheet, sheetConfig) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const items = [];
    
    for (let i = 1; i < values.length; i++) { // Skip header
      const row = values[i];
      if (row[0] && row[1]) { // Title and text exist
        items.push({
          itemSheet: sheetConfig.sheetName,
          itemTitle: row[0].toString(),
          itemText: row[1].toString(),
          itemTag: sheetConfig.sheetTag
        });
      }
    }
    
    return items;
  }
  
  /**
   * Extracts prompt data optimized (single pass)
   * @param {File} file - Google Drive file
   * @returns {Object} Prompt data
   */
  extractPromptDataOptimized(file) {
    const content = file.getBlob().getDataAsString();
    const lines = content.split('\n');
    
    // First line is title, rest is content
    const title = lines[0] || file.getName().replace('.txt', '');
    const promptContent = lines.slice(1).join('\n').trim();
    
    return {
      title: title,
      content: promptContent
    };
  }
}

/**
 * Data Cache Manager
 * Handles caching operations
 */
class DataCacheManager {
  constructor() {
    this.cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
  }
  
  /**
   * Gets cached data if valid
   * @returns {Object|null} Cached data or null
   */
  getCachedData() {
    if (!this.cacheFileId) return null;
    
    try {
      const file = DriveApp.getFileById(this.cacheFileId);
      const content = file.getBlob().getDataAsString();
      
      if (content.trim()) {
        return JSON.parse(content);
      }
    } catch (error) {
      console.log('💾 Cache não encontrado ou inválido');
    }
    
    return null;
  }
  
  /**
   * Caches data
   * @param {Object} data - Data to cache
   */
  cacheData(data) {
    try {
      if (!this.cacheFileId) {
        // Create cache file if it doesn't exist
        const file = DriveApp.createFile('sidebar_cache.json', JSON.stringify(data));
        const fileId = file.getId();
        PropertiesService.getScriptProperties().setProperty('cacheFileId', fileId);
        this.cacheFileId = fileId;
      } else {
        // Update existing cache file
        const file = DriveApp.getFileById(this.cacheFileId);
        file.setContent(JSON.stringify(data));
      }
      console.log('💾 Cache atualizado');
    } catch (error) {
      console.error('❌ Erro ao salvar cache:', error);
    }
  }
  
  /**
   * Clears cache data
   * @returns {Object} Operation result
   */
  clearCache() {
    try {
      if (!this.cacheFileId) {
        return { success: false, message: 'Cache não configurado' };
      }
      
      const file = DriveApp.getFileById(this.cacheFileId);
      file.setContent('');
      
      return { success: true, message: 'Cache limpo com sucesso' };
    } catch (error) {
      console.error('❌ Erro ao limpar cache:', error);
      return { success: false, message: error.message };
    }
  }
}

/**
 * Busca e retorna todos os dados necessários da planilha e documento para a sidebar
 * @return {string} JSON com dados da planilha, itens e títulos dos prompts
 */
function getSidebarData() {
  try {
    
    const sheetId = getSheetId();
    const documentId = getDocumentId();

    if (!sheetId) {
      const errorMsg = '🛑 FALHA CRÍTICA: ID da planilha (sheetData) não encontrado nas propriedades do script. A busca de dados não pode continuar.';
      console.error(errorMsg);
      return JSON.stringify({ error: errorMsg });
    }

    if (!documentId) {
      const errorMsg = '🛑 FALHA CRÍTICA: ID do Documento (docData) não encontrado nas propriedades do script. A busca de dados não pode continuar.';
      console.error(errorMsg);
      return JSON.stringify({ error: errorMsg });
    }

    const dataFetcher = new GoogleDriveDataFetcher(documentId, sheetId);
    const data = dataFetcher.extractAllDataFromSources();
    
    return JSON.stringify(data);
    
  } catch (error) {
    const errorMsg = `❌ Erro fatal durante a execução de getSidebarData: ${error.message}`;
    console.error(errorMsg);
    return JSON.stringify({ error: errorMsg });
  }
}

/**
 * Clears data cache
 * @returns {Object} Operation result
 */
function clearDataCache() {
  const cacheManager = new DataCacheManager();
  return cacheManager.clearCache();
}

/**
 * Tests cache system
 * @returns {Object} Test result
 */
function testCacheSystem() {
  try {
    console.time('⏱️ Cache Test');
    
    const configResult = initializeServerConfig();
    if (!configResult.success) {
      return { success: false, message: configResult.error };
    }
    
    const dataResult = processSidebarData(configResult.config);
    console.timeEnd('⏱️ Cache Test');
    
    if (!dataResult.success) {
      return { success: false, message: dataResult.error };
    }
    
    const cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
    
    return {
      success: true,
      message: 'Sistema de cache funcionando corretamente',
      cacheCreated: !!cacheFileId,
      dataMatches: true,
      cacheFileId: cacheFileId
    };
    
  } catch (error) {
    console.error('❌ Erro no teste do cache:', error);
    return { success: false, message: error.message };
  }
} 