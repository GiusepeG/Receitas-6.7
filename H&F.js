
/**
 * Etapa 1: Busca e classificação das imagens no Google Drive
 */
function getAvailableImages() {
  FOLDER_ID = PropertiesService.getScriptProperties().getProperty('imageFolderId');
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  
  const images = {
    headers: [],
    footers: []
  };
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileId = file.getId();
    
    // Extrai o tipo e nome baseado na convenção: "tipo nome.extensão"
    const spaceIndex = fileName.indexOf(' ');
    if (spaceIndex !== -1) {
      const type = fileName.substring(0, spaceIndex).toLowerCase();
      const name = fileName.substring(spaceIndex + 1).replace(/\.[^/.]+$/, ''); // Remove extensão
      
      const imageObject = {
        name: name,
        id: fileId
      };
      
      if (type === 'header') {
        images.headers.push(imageObject);
      } else if (type === 'footer') {
        images.footers.push(imageObject);
      }
    }
  }
  
  return images;
}

/**
 * Função auxiliar para criar os radio buttons
 */
function createRadios(items, name, preselectedId) {
  let html = `<div class="form-check">
    <input class="form-check-input" type="radio" name="${name}" id="${name}-none" value="" ${!preselectedId ? 'checked' : ''}>
    <label class="form-check-label" for="${name}-none">
      Nenhum
    </label>
  </div>`;
  
  for (let i = 0; i < items.length; i++) {
    const isChecked = items[i].id === preselectedId ? 'checked' : '';
    html += `<div class="form-check">
      <input class="form-check-input" type="radio" name="${name}" id="${name}-${i}" value="${items[i].id}" ${isChecked}>
      <label class="form-check-label" for="${name}-${i}">
        ${items[i].name}
      </label>
    </div>`;
  }
  
  return html;
}

/**
 * Etapa 2: Geração da interface de usuário dinâmica
 */
function chooseHeaderFooter() {
  const images = getAvailableImages();

  // Encontra o rodapé "Carimbo" para pré-seleção.
  const carimboFooter = images.footers.find(footer => footer.name.toLowerCase() === 'carimbo');
  const preselectedFooterId = carimboFooter ? carimboFooter.id : '';
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css" 
        integrity="sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N" crossorigin="anonymous">
      <style>
        body {
          padding: 1rem;
        }
        .btn-block {
          margin-top: 0.75rem;
        }
        h5 {
          margin-top: 1rem;
          margin-bottom: 0.5rem;
        }
      </style>
    </head>
    <body>
      <div class="container">
        
        <h5>Escolha o Cabeçalho</h5>
        ${createRadios(images.headers, 'headerRadio', '')}
        
        <h5 class="mt-3">Escolha o Rodapé</h5>
        ${createRadios(images.footers, 'footerRadio', preselectedFooterId)}

        <button class="btn btn-primary btn-block" id="applyBtn" onclick="submitSelection();">
          Aplicar Seleção
        </button>
        <button class="btn btn-secondary btn-block" onclick="google.script.host.close();">
          Cancelar
        </button>
        
      </div>

      <script>
        /**
         * Etapa 3: Captura a escolha do usuário e envia para o servidor.
         */
        function submitSelection() {
          // Desabilita o botão para prevenir cliques múltiplos
          document.getElementById('applyBtn').disabled = true;
          document.getElementById('applyBtn').textContent = 'Aplicando...';

          // Coleta os IDs selecionados
          const selectedHeaderId = document.querySelector('input[name="headerRadio"]:checked').value;
          const selectedFooterId = document.querySelector('input[name="footerRadio"]:checked').value;

          // Chama diretamente a função insertHeaderAndFooter no servidor
          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('Falha ao aplicar: ' + error.message);
              google.script.host.close();
            })
            .insertHeaderAndFooter(selectedHeaderId, selectedFooterId);
        }
      </script>
    </body>
    </html>`;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(500);
  
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Escolha o cabeçalho e rodapé');
}



function insertHeaderAndFooter(headerId, footerId) {
  clearHeaderAndFooter();

  if (headerId != '') {
    insertHeaderLogo(headerId);
  }

  if (footerId != '') {
    insertFooterLogo(footerId);
  }

  // Inserir a data no cabeçalho
  var hasHeaderImage = headerId != '';
  insertHeaderDate(hasHeaderImage);
}



function clearHeaderAndFooter() { 
  var header = DocumentApp.getActiveDocument().getHeader();
  var footer = DocumentApp.getActiveDocument().getFooter();

  if (header) {
    header.clear();
  }
  if (footer) {
    footer.clear();
  }
}


function insertLogo(id, isHeader) {
  if (!id) {
    return;
  }

  try {
    var sectionType = isHeader ? 'cabeçalho' : 'rodapé';
    var section = isHeader ?
      DocumentApp.getActiveDocument().getHeader() :
      DocumentApp.getActiveDocument().getFooter();

    if (!section) {
      section = isHeader ?
        DocumentApp.getActiveDocument().addHeader() :
        DocumentApp.getActiveDocument().addFooter();
    }

    // Se for o rodapé, remove todas as imagens posicionadas existentes antes de inserir a nova.
    // Isso garante que o rodapé não acumule imagens antigas.
    if (!isHeader && section) {
      const paragraphs = section.getParagraphs();
      for (const p of paragraphs) {
        const images = p.getPositionedImages();
        for (const img of images) {
          p.removePositionedImage(img.getId());
        }
      }
    }

    // Obtém uma imagem no Drive pelo ID
    var image = DriveApp.getFileById(id).getBlob();

    const style = {};
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
    style[DocumentApp.Attribute.FONT_FAMILY] = 'IBM Plex Sans';
    style[DocumentApp.Attribute.FONT_SIZE] = 6;
    style[DocumentApp.Attribute.BOLD] = true;
    style[DocumentApp.Attribute.SPACING_AFTER] = 0;
    style[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    style[DocumentApp.Attribute.LINE_SPACING] = 0;
    style[DocumentApp.Attribute.INDENT_START] = 0;
    style[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
    style[DocumentApp.Attribute.INDENT_END] = 0;

    // Aplica offset específico baseado no tipo de seção
    if (isHeader) {
      var paragraph = section.appendParagraph('');
      paragraph.setAttributes(style);

      paragraph.addPositionedImage(image)
        .setWidth(1880 / 3)
        .setHeight(350 / 3)
        .setLeftOffset(0)
        .setTopOffset(-40);

    } else {
      var paragraph = section.appendParagraph('');
      paragraph.setAttributes(style);

      paragraph.addPositionedImage(image)
        .setWidth(1880 / 3)
        .setHeight(350 / 3)
        .setLeftOffset(0)
        .setTopOffset(0);
    }
  } catch (e) {
    Logger.log(`[H&F - insertLogo] ERRO: ${e.toString()}`);
    // Re-throw the error so it can be caught by the client-side failure handler
    throw new Error(`Erro em insertLogo: ${e.message}`);
  }

}

function insertHeaderLogo(id_header) {
  insertLogo(id_header, true);
}

function insertFooterLogo(id_footer) {
  insertLogo(id_footer, false);
}


function insertHeaderDate(hasHeaderImage) {
  var header = DocumentApp.getActiveDocument().getHeader();

  if (!header) {
    header = DocumentApp.getActiveDocument().addHeader();
  }

  var numberOfParagraphs = 3;
  if (hasHeaderImage) {
    numberOfParagraphs = 1;
  }

  // Inserir os parágrafos em branco conforme necessário
  for (var i = 0; i < numberOfParagraphs; i++) {
    header.appendParagraph('');
  }

  // Inserir um parágrafo com a data atual
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var dateParagraph = header.appendParagraph(formattedDate);

  // Estilizar o parágrafo com a data
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'IBM Plex Sans';
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
  style[DocumentApp.Attribute.INDENT_END] = 0;
  style[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
  style[DocumentApp.Attribute.INDENT_START] = 0;
  style[DocumentApp.Attribute.ITALIC] = null;
  style[DocumentApp.Attribute.LEFT_TO_RIGHT] = true;
  style[DocumentApp.Attribute.LINE_SPACING] = 1;
  style[DocumentApp.Attribute.SPACING_AFTER] = 0;
  style[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  style[DocumentApp.Attribute.STRIKETHROUGH] = null;
  style[DocumentApp.Attribute.UNDERLINE] = null;

  dateParagraph.setAttributes(style);

  // O código que removia a imagem do cabeçalho foi retirado daqui.

}
