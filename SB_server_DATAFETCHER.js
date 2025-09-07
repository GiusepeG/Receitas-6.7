/** 
 * @class GoogleDriveDataFetcher
 * @description Fetches and processes data from Google Docs and Sheets efficiently.
 * Opens each document only once and provides methods to extract specific data.
 */
class GoogleDriveDataFetcher {
  /**
   * @param {string} documentId - The ID of the Google Document to fetch data from.
   * @param {string} sheetId - The ID of the Google Sheet to fetch data from.
   */
  constructor(documentId, sheetId) {
    if (!documentId) {
      throw new Error("Document ID is required.");
    }
    if (!sheetId) {
      throw new Error("Sheet ID is required.");
    }

    try {
      this.doc = DocumentApp.openById(documentId);
    } catch (e) {
      console.error(`❌ Falha ao abrir o Documento com ID ${documentId}: ${e.message}`);
      throw new Error(`Falha ao abrir o Documento: ${e.message}`);
    }

    try {
      this.spreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      console.error(`❌ Falha ao abrir a Planilha com ID ${sheetId}: ${e.message}`);
      throw new Error(`Falha ao abrir a Planilha: ${e.message}`);
    }

    this.sheetMetadata = null;
    
    // ========================================================================
    // SISTEMA DE CACHE EM ARQUIVO (GAS-COMPATIBLE) - ACESSO DIRETO POR ID
    // ========================================================================
    this.cacheConfig = {
      validityHours: 'never', // Cache nunca expira
    };
    
    // Obter ID do arquivo de cache das propriedades do script
    try {
      this.cacheFileId = PropertiesService.getScriptProperties().getProperty('cacheFileId');
      if (this.cacheFileId) {
        this.cacheFile = DriveApp.getFileById(this.cacheFileId);
        console.log('📦 Arquivo de cache conectado com sucesso');
      } else {
        console.warn('⚠️ cacheFileId não encontrado nas propriedades do script. Cache será criado automaticamente.');
        this.cacheFile = null;
      }
    } catch (e) {
      console.error(`❌ Erro ao acessar arquivo de cache: ${e.message}`);
      this.cacheFile = null;
    }
    
    // Obter ID da pasta de prompts das propriedades do script
    try {
      this.promptFolderId = PropertiesService.getScriptProperties().getProperty('promptFolderId');
      if (this.promptFolderId) {
        this.promptFolder = DriveApp.getFolderById(this.promptFolderId);
      } else {
        console.warn("⚠️ promptFolderId não encontrado nas propriedades do script. Funcionalidade de prompts por arquivos .txt não estará disponível.");
        this.promptFolder = null;
      }
    } catch (e) {
      console.error(`❌ Falha ao acessar a pasta de prompts: ${e.message}`);
      this.promptFolder = null;
    }
  }

  /**
   * Fetches all required data from the document and spreadsheet.
   * VERSÃO COM CACHE EM ARQUIVO: Reduz tempo de carregamento em 80-95%
   * @returns {object} An object containing all the fetched data.
   */
  extractAllDataFromSources() {
    console.time('⏱️ Total Data Extraction');
    
    // ========================================================================
    // OTIMIZAÇÃO: Verificar cache em arquivo primeiro
    // ========================================================================
    if (this._isCacheValid()) {
      const cacheData = this._readCacheFile();
      
      if (cacheData && cacheData.data) {
        console.timeEnd('⏱️ Total Data Extraction');
        console.log('⚡ Dados do cache');
        
        return cacheData.data;
      }
    }
    
    console.log('🔄 Carregando dados...');
    
    // ========================================================================
    // PROCESSAMENTO COMPLETO DOS DADOS
    // ========================================================================
    const sheetMetadata = this._extractSheetMetadata();
    const sheetItems = this._extractSheetItemsOptimized(sheetMetadata);
    const promptData = this._extractPromptDataOptimized();
    
    // Salvar no cache
    this._updateCache(sheetItems, promptData);
    
    console.timeEnd('⏱️ Total Data Extraction');
    console.log('✅ Dados processados');

    return {
      sheets: sheetMetadata,
      items: sheetItems,
      promptTitles: promptData.titles,
      allPrompts: promptData.prompts,
    };
  }

  /**
   * Extracts metadata from each sheet in the spreadsheet.
   * @private
   * @returns {Array<object>} An array of objects, each containing metadata for a sheet.
   */
  _extractSheetMetadata() {
    if (this.sheetMetadata) {
      return this.sheetMetadata;
    }

    try {
      const sheetsInFile = this.spreadsheet.getSheets();
      const sheets = [];

      sheetsInFile.forEach((sheet) => {
        const sheetName = sheet.getSheetName();
        if (sheetName === "_engine_") return;

        const cardStyle = sheet.getRange(1, 4, 1, 1).getValue();
        const match = sheetName.match(/\((.*?)\)/);
        const sheetButton = match ? match[1] : '';
        const sheetTag = sheetName.substring(sheetName.indexOf(')') + 1).trim();

        sheets.push({
          sheetName: sheetName,
          sheetStyle: cardStyle,
          sheetButton: sheetButton,
          sheetTag: sheetTag,
        });
      });

      this.sheetMetadata = sheets;
      return sheets;
    } catch (error) {
      console.error('❌ Erro ao extrair metadados das planilhas:', error);
      throw new Error('Erro ao extrair metadados das planilhas: ' + error.message);
    }
  }

  /**
   * Extracts all items from the sheets based on provided metadata.
   * @private
   * @param {Array<object>} sheetMetadata - The metadata for the sheets to process.
   * @returns {Array<object>} An array of all items extracted from the sheets.
   */
  _extractSheetItems(sheetMetadata) {
    try {
      const items = [];
      sheetMetadata.forEach((sheetInfo) => {
        const sheet = this.spreadsheet.getSheetByName(sheetInfo.sheetName);
        if (sheet.getLastRow() > 1) {
            const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
            rows.forEach((row) => {
                items.push({
                    itemSheet: sheetInfo.sheetTag,
                    itemTitle: row[0],
                    itemTag: row[1],
                    itemText: row[2],
                    itemStyle: row[3],
                    requires: row[4],
                });
            });
        }
      });

      return items;
    } catch (error) {
      console.error('❌ Erro ao extrair itens das planilhas:', error);
      throw new Error('Erro ao extrair itens das planilhas: ' + error.message);
    }
  }

  /**
   * OTIMIZADO: Extrai dados das planilhas com leitura em lote
   * Reduz chamadas à API do Google Sheets
   * @private
   * @param {Array<object>} sheetMetadata - Metadados das planilhas
   * @returns {Array<object>} Array de itens extraídos
   */
  _extractSheetItemsOptimized(sheetMetadata) {
    try {
      const items = [];
      
      // ========================================================================
      // OTIMIZAÇÃO: Leitura em lote de todas as planilhas
      // ========================================================================
      const sheetsToProcess = sheetMetadata.filter(info => {
        const sheet = this.spreadsheet.getSheetByName(info.sheetName);
        return sheet.getLastRow() > 1; // Só processar planilhas com dados
      });
      
      console.log(`📊 Processando ${sheetsToProcess.length} planilhas...`);
      
      sheetsToProcess.forEach((sheetInfo, index) => {
        const sheet = this.spreadsheet.getSheetByName(sheetInfo.sheetName);
        
        try {
          // Leitura otimizada: buscar todos os dados de uma vez
          const lastRow = sheet.getLastRow();
          const lastColumn = sheet.getLastColumn();
          
          if (lastRow > 1 && lastColumn >= 3) {
            const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
            
            rows.forEach((row, rowIndex) => {
              // Só adicionar se a linha não estiver vazia
              if (row[0] && row[0].toString().trim() !== '') {
                items.push({
                  itemSheet: sheetInfo.sheetTag,
                  itemTitle: row[0] || '',
                  itemTag: row[1] || '',
                  itemText: row[2] || '',
                  itemStyle: row[3] || '',
                  requires: row[4] || '',
                });
              }
            });
          }
        } catch (sheetError) {
          console.error(`❌ Erro ao processar planilha ${sheetInfo.sheetName}:`, sheetError.toString());
        }
      });

      console.log(`✅ ${items.length} itens extraídos`);
      
      return items;
      
    } catch (error) {
      console.error('❌ Erro ao extrair itens das planilhas (otimizado):', error);
      // Fallback para método original
      console.warn('⚠️ Usando método original como fallback');
      return this._extractSheetItems(sheetMetadata);
    }
  }

  /**
   * Extracts the titles of prompts from .txt files in the prompt folder.
   * @returns {Array<string>} An array of tab titles extracted from file names.
   */
  _extractPromptTitlesFromFiles() {
    try {
      if (!this.promptFolder) {
        console.warn("⚠️ Pasta de prompts não disponível. Usando método antigo.");
        return this._extractPromptTitles_DEPRECATED();
      }

      const files = this.promptFolder.getFilesByType(MimeType.PLAIN_TEXT);
      const tabNames = [];

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        
        // Extrair tabName do nome do arquivo (primeira parte antes da primeira vírgula)
        const parts = fileName.split(',');
        if (parts.length >= 3) {
          const tabName = parts[0].trim();
          // Remover extensão .txt se existir
          const cleanTabName = tabName.replace(/\.txt$/i, '').normalize('NFC');
          tabNames.push(cleanTabName);
        } else {
          console.warn(`⚠️ Arquivo ignorado (formato incorreto): ${fileName}`);
        }
      }

      return tabNames;
    } catch (e) {
      console.error("❌ Erro ao extrair títulos dos prompts dos arquivos:", e.toString());
      console.warn("⚠️ Tentando usar método antigo como fallback.");
      return this._extractPromptTitles_DEPRECATED();
    }
  }

  /**
   * Extracts all prompts from .txt files in the prompt folder.
   * File name format: "tabName, fromHeadline2, toHeadline2.txt"
   * @returns {object} An object where keys are tab titles and values are prompt data.
   */
  _extractAllPromptsFromFiles() {
    try {
      if (!this.promptFolder) {
        console.warn("⚠️ Pasta de prompts não disponível. Usando método antigo.");
        return this._extractAllPrompts_DEPRECATED();
      }

      const files = this.promptFolder.getFilesByType(MimeType.PLAIN_TEXT);
      const allPrompts = {};

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        
        // Extrair informações do nome do arquivo
        const parts = fileName.split(',');
        if (parts.length >= 3) {
          const tabName = parts[0].trim().replace(/\.txt$/i, '');
          const fromHeadline2 = parts[1].trim().normalize('NFC');
          const toHeadline2 = parts[2].trim().replace(/\.txt$/i, '').normalize('NFC');
          
          try {
            // Ler conteúdo do arquivo
            const promptContent = file.getBlob().getDataAsString('UTF-8').trim();
            
            allPrompts[tabName] = {
              fromHeadline2: fromHeadline2,
              toHeadline2: toHeadline2,
              content: promptContent,
            };
            
          } catch (readError) {
            console.error(`❌ Erro ao ler arquivo ${fileName}:`, readError.toString());
          }
        } else {
          console.warn(`⚠️ Arquivo ignorado (formato incorreto): ${fileName}`);
          console.warn(`   Formato esperado: "tabName, fromHeadline2, toHeadline2.txt"`);
        }
      }

      return allPrompts;
    } catch (e) {
      console.error("❌ Erro ao extrair prompts dos arquivos:", e.toString());
      console.warn("⚠️ Tentando usar método antigo como fallback.");
      return this._extractAllPrompts_DEPRECATED();
    }
  }

  /**
   * OTIMIZADO: Extrai títulos e conteúdo dos prompts em uma única iteração
   * Elimina a leitura dupla de arquivos .txt
   * @private
   * @returns {object} Objeto com titles (array) e prompts (object)
   */
  _extractPromptDataOptimized() {
    try {
      if (!this.promptFolder) {
        console.warn("⚠️ Pasta de prompts não disponível. Usando método antigo.");
        return {
          titles: this._extractPromptTitles_DEPRECATED(),
          prompts: this._extractAllPrompts_DEPRECATED()
        };
      }

      const files = this.promptFolder.getFilesByType(MimeType.PLAIN_TEXT);
      const tabNames = [];
      const allPrompts = {};

      // ========================================================================
      // OTIMIZAÇÃO: Uma única iteração para extrair títulos E conteúdo
      // ========================================================================
      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        
        // Extrair informações do nome do arquivo
        const parts = fileName.split(',');
        if (parts.length >= 2) {
          const tabName = parts[0].trim().replace(/\.txt$/i, '');
          const toHeadline2 = parts[1].trim().replace(/\.txt$/i, '').normalize('NFC');
          
          // Adicionar título à lista (sem duplicatas)
          const cleanTabName = tabName.normalize('NFC');
          if (!tabNames.includes(cleanTabName)) {
            tabNames.push(cleanTabName);
          }
          
          try {
            // Ler conteúdo do arquivo (uma única vez)
            const promptContent = file.getBlob().getDataAsString('UTF-8').trim();
            
            allPrompts[tabName] = {
              toHeadline2: toHeadline2,
              content: promptContent,
            };
            
          } catch (readError) {
            console.error(`❌ Erro ao ler arquivo ${fileName}:`, readError.toString());
          }
        } else {
          console.warn(`⚠️ Arquivo ignorado (formato incorreto): ${fileName}`);
        }
      }

      console.log(`✅ ${tabNames.length} prompts processados`);
      
      return {
        titles: tabNames,
        prompts: allPrompts
      };
      
    } catch (e) {
      console.error("❌ Erro ao extrair dados dos prompts:", e.toString());
      console.warn("⚠️ Tentando usar método antigo como fallback.");
      return {
        titles: this._extractPromptTitles_DEPRECATED(),
        prompts: this._extractAllPrompts_DEPRECATED()
      };
    }
  }

  /**
   * Verifica se o cache ainda é válido
   * @private
   * @returns {boolean} True se o cache é válido
   */
  _isCacheValid() {
    try {
      if (!this.cacheFile) return false;
      
      // Verificar se o arquivo tem conteúdo válido
      const jsonContent = this.cacheFile.getBlob().getDataAsString();
      
      if (!jsonContent || jsonContent.trim() === '') {
        console.log('📄 Cache vazio');
        return false;
      }
      
      const cacheData = JSON.parse(jsonContent);
      
      // Verificar se o cache foi limpo
      if (!cacheData || cacheData.cleared || !cacheData.data) {
        console.log('🧹 Cache foi limpo');
        return false;
      }
      
      // Cache nunca expira, então sempre retorna true se tem conteúdo válido
      console.log(`📦 Cache válido (nunca expira)`);
      return true;
      
    } catch (error) {
      console.error('❌ Erro no cache:', error.message);
      return false;
    }
  }

  /**
   * Lê os dados do arquivo cache
   * @private
   * @returns {object|null} Dados do cache ou null se não encontrado
   */
  _readCacheFile() {
    try {
      if (!this.cacheFile) return null;
      
      const jsonContent = this.cacheFile.getBlob().getDataAsString();
      
      // Verificar se o arquivo está vazio ou com conteúdo inválido
      if (!jsonContent || jsonContent.trim() === '') {
        console.log('📄 Cache vazio');
        return null;
      }
      
      const cacheData = JSON.parse(jsonContent);
      
      // Verificar se o cache foi limpo ou é inválido
      if (!cacheData || cacheData.cleared || !cacheData.data) {
        console.log('🧹 Cache foi limpo ou inválido');
        return null;
      }
      
      console.log('📖 Cache carregado com sucesso');
      return cacheData;
      
    } catch (error) {
      console.error('❌ Erro ao ler cache:', error.message);
      return null;
    }
  }

  /**
   * Atualiza o cache com novos dados
   * @private
   * @param {Array} sheetItems - Itens das planilhas
   * @param {Object} promptData - Dados dos prompts
   */
  _updateCache(sheetItems, promptData) {
    try {
      // Preparar dados para salvar
      const cacheData = {
        timestamp: new Date().toISOString(),
        data: {
          sheets: this.sheetMetadata,
          items: sheetItems,
          promptTitles: promptData.titles,
          allPrompts: promptData.prompts,
        }
      };

      const jsonContent = JSON.stringify(cacheData, null, 2);
      
      // Se não temos arquivo, criar um novo
      if (!this.cacheFile) {
        this._createCacheFile(jsonContent);
      } else {
        // Atualizar conteúdo do arquivo existente
        this.cacheFile.setContent(jsonContent);
      }
      
      console.log('💾 Cache salvo');
    } catch (e) {
      console.error(`❌ Erro ao salvar cache: ${e.message}`);
    }
  }

  /**
   * Cria um novo arquivo de cache
   * @private
   * @param {string} jsonContent - Conteúdo JSON para salvar
   */
  _createCacheFile(jsonContent) {
    try {
      // Criar arquivo na pasta root do drive do usuário
      const blob = Utilities.newBlob(jsonContent, 'application/json', 'sidebar_cache.json');
      const newFile = DriveApp.createFile(blob);
      
      // Salvar o ID do arquivo nas propriedades do script
      this.cacheFileId = newFile.getId();
      PropertiesService.getScriptProperties().setProperty('cacheFileId', this.cacheFileId);
      
      // Atualizar referência local
      this.cacheFile = newFile;
      
      console.log('📄 Arquivo de cache criado com ID:', this.cacheFileId);
    } catch (e) {
      console.error(`❌ Erro ao criar arquivo de cache: ${e.message}`);
    }
  }

  /**
   * Limpa o cache forçando uma nova leitura
   * @public
   */
  clearCache() {
    try {
      if (!this.cacheFile) {
        console.warn('⚠️ Arquivo de cache não disponível');
        return { success: false, message: 'Arquivo de cache não disponível' };
      }
      
      // Limpar conteúdo do arquivo ao invés de deletar
      const emptyCache = {
        timestamp: new Date().toISOString(),
        data: null,
        cleared: true
      };
      
      this.cacheFile.setContent(JSON.stringify(emptyCache, null, 2));
      
      const message = `🧹 Cache limpo (conteúdo apagado)`;
      console.log(message);
      
      return { success: true, message: message };
    } catch (error) {
      console.error('❌ Erro ao limpar cache:', error);
      return { success: false, message: 'Erro ao limpar cache: ' + error.message };
    }
  }

  /**
   * @deprecated Use _extractPromptTitlesFromFiles() instead.
   * Extracts the titles of all top-level tabs from the document, ignoring those starting with '_'.
   * @private
   * @returns {Array<string>} An array of tab titles.
   */
  _extractPromptTitles_DEPRECATED() {
    try {
      const topLevelTabs = this.doc.getTabs();
      // Filter out tabs whose names start with an underscore
      const tabNames = topLevelTabs
        .filter(tab => !tab.getTitle().startsWith('_'))
        .map(tab => tab.getTitle());
      return tabNames;
    } catch (e) {
      console.error("❌ Erro ao extrair títulos das abas:", e.toString());
      throw new Error('Erro ao extrair títulos das abas: ' + e.message);
    }
  }

  /**
   * @deprecated Use _extractAllPromptsFromFiles() instead.
   * Extracts content from all prompts (tabs) in the document based on the new structure.
   * - Top-level tabs (not starting with '_') provide from/to headlines.
   * - The first valid sub-tab (not starting with '_') provides the content.
   * @private
   * @returns {object} An object where keys are tab titles and values are prompt data.
   */
  _extractAllPrompts_DEPRECATED() {
    try {
      const topLevelTabs = this.doc.getTabs().filter(tab => !tab.getTitle().startsWith('_'));
      const allPrompts = {};

      for (const tab of topLevelTabs) {
        const tabName = tab.getTitle();
        const documentTab = tab.asDocumentTab();
        const bodyText = documentTab.getBody().getText();
        const lines = bodyText.trim().split('\n');

        const fromHeadline2 = lines.length > 0 ? lines[0].trim() : '';
        const toHeadline2 = lines.length > 1 ? lines[1].trim() : '';
        let promptContent = '';

        const childTabs = tab.getChildTabs();
        // Find the first valid sub-tab for content
        const contentTab = childTabs.find(child => !child.getTitle().startsWith('_'));

        if (contentTab) {
          const contentDocumentTab = contentTab.asDocumentTab();
          promptContent = contentDocumentTab.getBody().getText().trim();
        }

        allPrompts[tabName] = {
          fromHeadline2: fromHeadline2,
          toHeadline2: toHeadline2,
          content: promptContent,
        };
      }
  
      return allPrompts;
    } catch (e) {
      console.error("❌ Erro ao extrair todos os prompts:", e.toString());
      throw new Error('Erro ao extrair todos os prompts: ' + e.message);
    }
  }
} 