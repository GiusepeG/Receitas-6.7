function getRulesForWholeParagraphs() {
    return [
        {
        keywords: ['OD:', 'OE:', 'AO:', 'ADIÇÃO:'], // e também "Papila normocorada;" -> linhas mais comuns antes!
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase().trim(); 
          return this.keywords.some(keyword => text.startsWith(keyword)) || text.endsWith(";"); // Verifica se a linha começa com um dos keywords ou termina com ";"
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 10,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['NOME:', 'REF:', 'PACIENTE:'],
        heading: DocumentApp.ParagraphHeading.HEADING1,
        requiresPageBreak: true,
        condition: function(line) {
          var text = line.toUpperCase();
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 16,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [
          'PRONTUÁRIO', 'LAUDO', 'RECEITUÁRIO', 'PRESCRIÇÃO', 'RECEITA',
          'RELATÓRIO', 'ATESTADO', 'SOLICITAÇÃO', 'ORIENTAÇÃO', 'ENCAMINHAMENTO',
          'RECOMENDAÇÕES', 'DESCRIÇÃO', 'GUIA','JUSTIFICATIVA', 'TRANSCRIÇÃO', 'DOCUMENTO INDEFINIDO'
        ],
        heading: DocumentApp.ParagraphHeading.HEADING2,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.toUpperCase(); // Não sensível a maiúsculas e minúsculas
          return this.keywords.some(keyword => text.startsWith(keyword.toUpperCase()));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 18,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 20,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['História', 'Exame Físico', 'Exames Complementares', 'Hipótese Diagnóstica', 'Conclusão', 'Conduta', 'Prezad', 'Uso', 'Orientações' ],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          // Verifica se a linha começa com uma palavra-chave que inicia com maiúscula
          return this.keywords.some(keyword => line.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Mono',
          [DocumentApp.Attribute.FONT_SIZE]: 14,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-Z][0-9]{2}[.][0-9]{1}\s*\(.*\)$/; // H10.0 (Conjuntivite mucopurulenta)
          return regex.test(text);
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) { // Simbrinza (colírio de uso contínuo)
          var text = line.trim();
          var startsWithLetter = /^[a-zA-ZçÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]/.test(text.charAt(0));
          var endsWithParenthesis = text.endsWith(")");
          return startsWithLetter && endsWithParenthesis;
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 10,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['OLHO DIREITO', 'OLHO ESQUERDO'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          var regex = /^[A-ZÀ-Ÿ0-9\s:/çÇãÃõÕáÁéÉíÍóÓúÚâÂêÊîÎôÔûÛàÀèÈìÌòÒùÙ]+:$/;  // "ACUIDADE VISUAL SEM CORREÇÃO:" e "REFRAÇÃO:"
          return !text.startsWith('CID') && (regex.test(text) || this.keywords.some(keyword => text.includes(keyword)));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 5,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [], // composição: carmelose 5mg/mL
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return line.startsWith("composição:");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Sans Condensed',
          [DocumentApp.Attribute.FONT_SIZE]: 8,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ['CID'],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword)) && text.endsWith(":");
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: ["À disposição", "Auxiliar"],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          var text = line.trim();
          return this.keywords.some(keyword => text.startsWith(keyword));
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.RIGHT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1.5,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 15,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      },
      {
        keywords: [],
        heading: DocumentApp.ParagraphHeading.NORMAL,
        requiresPageBreak: false,
        condition: function(line) {
          return true; // Aplica-se a todos os parágrafos que não correspondem a outras regras
        },
        style: {
          [DocumentApp.Attribute.BACKGROUND_COLOR]: null,
          [DocumentApp.Attribute.BOLD]: null,
          [DocumentApp.Attribute.FONT_FAMILY]: 'IBM Plex Serif',
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.FOREGROUND_COLOR]: null,
          [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
          [DocumentApp.Attribute.INDENT_END]: 0,
          [DocumentApp.Attribute.INDENT_FIRST_LINE]: 50,
          [DocumentApp.Attribute.INDENT_START]: 0,
          [DocumentApp.Attribute.ITALIC]: null,
          [DocumentApp.Attribute.LEFT_TO_RIGHT]: true,
          [DocumentApp.Attribute.LINE_SPACING]: 1,
          [DocumentApp.Attribute.SPACING_AFTER]: 0,
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.STRIKETHROUGH]: null,
          [DocumentApp.Attribute.UNDERLINE]: null
        }
      }
    ];
  }

/**
 * Retorna as regras de formatação para os títulos.
 * Esta função atua como um alias para getRulesForWholeParagraphs,
 * garantindo compatibilidade com o DocumentHeadlineManager.
 * @returns {Array} As regras de formatação.
 */
function getFormattingRules() {
  return getRulesForWholeParagraphs();
}