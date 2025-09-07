/**
 * Salva todas as propriedades do script atual em um arquivo de texto no Google Drive.
 * @param {string} fileId O ID do arquivo de texto para salvar as propriedades.
 */
function savePropertiesToFile() {

  const fileId = '1It53sYDoYcYNE0P2PmHekB-FZwDCXxFK'
  try {
    // 1. Obtém as propriedades do script atual.
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    // 2. Converte as propriedades para uma string no formato JSON.
    const propertiesString = JSON.stringify(properties, null, 2);
    
    // 3. Obtém o arquivo de texto pelo ID.
    const file = DriveApp.getFileById(fileId);
    
    // 4. Edita o conteúdo do arquivo com a string JSON.
    file.setContent(propertiesString);
    
    console.log(`Propriedades salvas com sucesso no arquivo: ${file.getName()} (ID: ${fileId})`);
    
  } catch (error) {
    console.error(`Erro ao salvar as propriedades: ${error.message}`);
  }
}

/**
 * Lê um arquivo de texto e restaura as propriedades do script.
 * O arquivo de texto deve conter um objeto JSON válido.
 * @param {string} fileId O ID do arquivo de texto para ler as propriedades.
 */
function loadPropertiesFromFile() {

  const fileId = '1It53sYDoYcYNE0P2PmHekB-FZwDCXxFK'

  try {
    // 1. Obtém o arquivo de texto pelo ID.
    const file = DriveApp.getFileById(fileId);
    
    // 2. Obtém o conteúdo do arquivo como uma string.
    const content = file.getBlob().getDataAsString();
    
    // 3. Converte a string JSON de volta para um objeto.
    const properties = JSON.parse(content);
    
    // 4. Limpa as propriedades existentes para evitar conflitos.
    PropertiesService.getScriptProperties().deleteAllProperties();
    
    // 5. Define as novas propriedades.
    PropertiesService.getScriptProperties().setProperties(properties);
    
    console.log(`Propriedades carregadas com sucesso do arquivo: ${file.getName()} (ID: ${fileId})`);
    
  } catch (error) {
    console.error(`Erro ao carregar as propriedades: ${error.message}`);
  }
}