function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Abrir Sidebar', 'openSidebarServerSide')
    .addItem('Cabeçalhos e Rodapés', 'chooseHeaderFooter')
    .addSeparator()
    .addItem('Extrair `Nome:` da agenda', 'extractHeadlineOneFromText')
    .addSeparator()
    .addItem('Configurar Propriedades', 'openPropertiesModal')
    .addItem('Limpar Cache', 'clearDataCacheFromMenu')
    .addToUi();
}

/**
 * Função para limpar cache via menu
 */
function clearDataCacheFromMenu() {
  try {
    const result = clearDataCache();
    
    if (result.success) {
      DocumentApp.getUi().alert('✅ Cache limpo com sucesso!\n\n' + result.message);
    } else {
      DocumentApp.getUi().alert('❌ Erro ao limpar cache:\n\n' + result.message);
    }
  } catch (error) {
    DocumentApp.getUi().alert('❌ Erro inesperado:\n\n' + error.message);
  }
}