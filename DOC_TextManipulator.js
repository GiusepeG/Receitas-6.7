/**
 * @class TextManipulator
 * @classdesc A class to process and format text using a fluent, chainable interface.
 */
class TextManipulator {
  /**
   * @param {string} text The initial text to be manipulated.
   * @param {Array<Object>} placeholders
   * @param {Array<Object>} heading1Rules Rules for HEADING1 paragraphs.
   * @param {Array<Object>} heading2Rules Rules for HEADING2 paragraphs.
   */
  constructor(text, placeholders, heading1Rules, heading2Rules) {
    this.text = text;
    this.placeHolders = placeholders || [];
    this.heading1Rules = heading1Rules || [];
    this.heading2Rules = heading2Rules || [];

    this.lines = null;
    this.lastHeading1 = null;
    this.currentPatientName = null;
    this.processedLines = []; // Array of dictionaries with {text, heading}
    this.shouldLog = false;
  }

  /**
   * Enables or disables logging for debugging purposes.
   * @param {boolean} [shouldLog=true]
   * @returns {TextManipulator}
   */
  withLogging(shouldLog = true) {
    this.shouldLog = shouldLog;
    return this;
  }

  /**
   * Splits the initial text into an array of lines.
   * @returns {TextManipulator} The instance for chaining.
   */
  splitInLines() {
    this.lines = this.text.split(/\r\n|\r|\n/g);
    this._log('Text split into ' + this.lines.length + ' lines.');
    return this;
  }

  /**
   * Removes empty lines from the current array of lines.
   * @returns {TextManipulator} The instance for chaining.
   */
  removeEmptyLines() {
    if (this.lines === null) {
      this.splitInLines();
    }
    const originalCount = this.lines.length;
    this.lines = this.lines.filter(line => line !== "");
    this._log(`${originalCount - this.lines.length} empty lines removed.`);
    return this;
  }

  /**
   * Removes unnecessary leading spaces from lines while preserving intentional indentation.
   * @returns {TextManipulator} The instance for chaining.
   */
  trimLeadingSpaces() {
    if (this.lines === null) {
      this.splitInLines();
    }
    
    const originalCount = this.lines.length;
    this.lines = this.lines.map(line => {
      // Check if line starts with spaces followed by non-space characters
      const regexLeadingSpaces = /^\s+\S+/;
      if (regexLeadingSpaces.test(line)) {
        return line.trim();
      }
      return line;
    });
    
    this._log('Leading spaces trimmed from lines.');
    return this;
  }

  /**
   * Removes escape characters from all lines.
   * Replaces "\" with "" and "*" with " ".
   * @returns {TextManipulator} The instance for chaining.
   */
  removeEscapeCharacters() {
    if (this.lines === null) {
      this.splitInLines();
    }
    
    this.lines = this.lines.map(line => {
      return line.replace(/\\/g, '').replace(/\*/g, ' ');
    });
    
    this._log('Escape characters removed from lines.');
    return this;
  }

  /**
   * Checks the first line and modifies it if it appears to be a person's name.
   * Prefixes with "Nome: " if it matches the criteria.
   * @returns {TextManipulator} The instance for chaining.
   */
  checkAndModifyFirstLine() {
    if (this.lines === null) {
      this.splitInLines();
    }
    if (this.lines.length === 0) return this;

    const firstLine = this.lines[0];
    if (typeof firstLine !== 'string') {
      this.lines[0] = "";
      return this;
    }

    const trimmedText = firstLine.trim();
    if (trimmedText === "") {
        this.lines[0] = "";
        return this;
    }

    const regexUppercaseWithAccents = /^[A-ZÇÃÕÁÉÍÓÚÂÊÎÔÛÀÈÌÒÙÄËÏÖÜŸŠŽČĐÑ ]+$/;
    const words = trimmedText.split(" ").filter(word => word.length > 0);

    if (regexUppercaseWithAccents.test(trimmedText) && words.length >= 2) {
      this.lines[0] = "Nome: " + trimmedText;
      this._log('First line modified to: ' + this.lines[0]);
    } else {
      this.lines[0] = trimmedText;
    }
    return this;
  }

  /**
   * Processes all lines according to the configured formatting rules and placeholders.
   * This is the main processing method that orchestrates the manipulation.
   * @returns {TextManipulator} The instance for chaining.
   */
  processLines() {
    if (this.lines === null) {
      this._log('Lines not available. Run splitInLines() first.');
      return this;
    }

    this.lines.forEach(line => {
      let currentLine = this._trimSpacesAndTest(line);
      
      // Check if the line is a name declaration (e.g., "Nome: John Doe")
      if (currentLine.trim().toUpperCase().startsWith("NOME:")) {
          const extractedName = this._extractName(currentLine);
          if (extractedName) {
            this.currentPatientName = capitalizeFirstLetter(extractedName);
          }
      }
      
      currentLine = this._replacePlaceholders(currentLine);

      // Check HEADING1 rules first
      const heading1Rule = this.heading1Rules.find(rule => rule.condition(currentLine));
      if (heading1Rule) {
        this._handleAppliedRule(heading1Rule, currentLine);
        return;
      }

      // Check HEADING2 rules
      const heading2Rule = this.heading2Rules.find(rule => rule.condition(currentLine));
      if (heading2Rule) {
        this._handleAppliedRule(heading2Rule, currentLine);
        return;
      }

      // If no rules match, add as regular line
      this.processedLines.push({
        text: currentLine,
        heading: null
      });
    });

    this._log('Finished processing all lines.');
    return this;
  }

  /**
   * Structures the document by grouping H2 sections under their parent H1,
   * optionally prioritizing specific H2s, and ensuring the H1 is repeated for each section.
   * @param {Array<string>|null} [priorityHeadlines=null] - A list of H2 texts to place first, in order.
   * @returns {TextManipulator} The instance for chaining.
   */
  structureH2Blocks(priorityHeadlines = null) {
    if (this.processedLines.length === 0) {
      this._log('No lines to structure.');
      return this;
    }

    const h1Groups = new Map();
    let currentH1 = null;
    let currentH2Block = null;

    // Etapa 1: Agrupar linhas, criando H1s e H2s padrão conforme necessário.
    this.processedLines.forEach(item => {
      if (item.heading === DocumentApp.ParagraphHeading.HEADING1) {
        currentH1 = item;
        currentH2Block = null; // Redefinir para um novo H1
        if (!h1Groups.has(currentH1.text)) {
          h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
        }
      } else if (item.heading === DocumentApp.ParagraphHeading.HEADING2) {
        if (!currentH1) {
          currentH1 = {
            text: "Nome: PACIENTE INDEFINIDO",
            heading: DocumentApp.ParagraphHeading.HEADING1
          };
          h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
        }
        currentH2Block = [item];
        h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
      } else { // Parágrafo de texto regular
        if (!currentH1) {
          currentH1 = {
            text: "Nome: PACIENTE INDEFINIDO",
            heading: DocumentApp.ParagraphHeading.HEADING1
          };
          h1Groups.set(currentH1.text, { h1: currentH1, h2Blocks: [] });
        }
        if (!currentH2Block) {
          const defaultH2 = {
            text: "Documento Indefinido",
            heading: DocumentApp.ParagraphHeading.HEADING2
          };
          currentH2Block = [defaultH2];
          h1Groups.get(currentH1.text).h2Blocks.push(currentH2Block);
        }
        currentH2Block.push(item);
      }
    });
    this._log(`Finished grouping. H1 groups count: ${h1Groups.size}`);

    // Etapa 3: Reordenar blocos H2 dentro de cada grupo H1 por prioridade.
    if (priorityHeadlines && priorityHeadlines.length > 0) {
      const priorityTexts = priorityHeadlines.map(h => h.trim().toUpperCase());

      h1Groups.forEach(group => {
        group.h2Blocks.sort((a, b) => {
          const aText = a[0].text.trim().toUpperCase();
          const bText = b[0].text.trim().toUpperCase();

          const aIndex = priorityTexts.indexOf(aText);
          const bIndex = priorityTexts.indexOf(bText);

          if (aIndex !== -1 && bIndex !== -1) {
            return aIndex - bIndex; // Both are priority, sort by index
          }
          if (aIndex !== -1) {
            return -1; // a is priority, b is not
          }
          if (bIndex !== -1) {
            return 1; // b is priority, a is not
          }
          return 0; // neither is priority
        });
      });
      this._log(`H2 blocks reordered based on priorities: "${priorityHeadlines.join(', ')}"`);
    }

    // Etapa 4: Planificar a estrutura agrupada de volta para uma lista de linhas.
    const finalStructuredLines = [];
    h1Groups.forEach(group => {
      if (group.h2Blocks.length === 0) {
        finalStructuredLines.push(group.h1);
      } else {
        group.h2Blocks.forEach(block => {
          finalStructuredLines.push(group.h1);
          finalStructuredLines.push(...block);
        });
      }
    });

    this.processedLines = finalStructuredLines;
    this._log(`Final line structure rebuilt. Total lines: ${this.processedLines.length}`);
    return this;
  }

  /**
   * Returns the final processed text.
   * @param {string} [joiner='\n'] - The character to join lines with.
   * @returns {string} The final text.
   */
  getResult(joiner = '\n') {
    return this.processedLines.map(item => item.text).join(joiner);
  }

  /**
   * Logs the array of dictionaries for debugging purposes.
   * @returns {TextManipulator} The instance for chaining.
   */
  logDictionaries() {
    if (this.shouldLog) {
      this._log('=== Processed Lines Dictionaries ===');
      this.processedLines.forEach((item, index) => {
        this._log(`[${index}] text: "${item.text}", heading: ${item.heading}`);
      });
      this._log('=== End of Dictionaries ===');
    }
    return this;
  }

  /**
   * Internal log helper.
   * @private
   */
  _log(message) {
    if (this.shouldLog) {
      // In a Google Apps Script environment, Logger.log is preferred.
      if (typeof Logger !== 'undefined') {
        Logger.log(message);
      } else {
        console.log(message);
      }
    }
  }

  /**
   * Trims leading spaces if the line contains a mix of spaces and characters.
   * @private
   */
  _trimSpacesAndTest(text) {
    // If the line is only spaces, preserve it as is
    if (text.trim() === '') {
      return text;
    }
    
    // If the line starts with spaces followed by non-space characters, trim leading spaces
    const regex = /^\s+\S+/;
    return regex.test(text) ? text.trim() : text;
  }

  /**
   * Replaces placeholders in the text.
   * @param {string} text The text to process.
   * @returns {string} The text with placeholders replaced.
   */
  _replacePlaceholders(text) {
    if (!this.placeHolders || this.placeHolders.length === 0) {
      return text;
    }
    let modifiedText = text;
    this.placeHolders.forEach(rule => {
      const placeholderRegex = new RegExp(this._escapeRegExp(rule.text), 'gi'); // Case-insensitive
      if (placeholderRegex.test(modifiedText)) {
        const replacementValue = typeof rule.replacement === 'function'
          ? rule.replacement(this.currentPatientName)
          : rule.replacement;
        
        modifiedText = modifiedText.replace(placeholderRegex, replacementValue);
      }
    });
    return modifiedText;
  }

  /**
   * Handles a line where a formatting rule was applied.
   * @private
   */
  _handleAppliedRule(appliedRule, text) {
    if (this._isHeading1(appliedRule)) {
      let heading1Text = text;
      const regexNomeVariations = /^(NOME:|Nome:|nome:|nOME:)/i;
      if (regexNomeVariations.test(heading1Text)) {
        heading1Text = this._formatNomeLine(heading1Text);
      }
      this._storeHeading1(heading1Text);
    } else if (this._isHeading2(appliedRule)) {
      this.processedLines.push({
        text: text,
        heading: DocumentApp.ParagraphHeading.HEADING2
      });
    } else {
      this.processedLines.push({
        text: text,
        heading: null
      });
    }
  }

  /**
   * Formats a "Nome:" line to a consistent case.
   * @private
   */
  _formatNomeLine(text) {
    const regexNomeVariations = /^(NOME:|Nome:|nome:|nOME:)\s*/i;
    const processedText = text.replace(regexNomeVariations, "Nome: ");
    return processedText.replace(/Nome: (.+)/, (_, namePart) => "Nome: " + namePart.toUpperCase());
  }

  /**
   * Stores a HEADING1 line and updates the lastHeading1 reference.
   * @private
   */
  _storeHeading1(text) {
    this.processedLines.push({
      text: text,
      heading: DocumentApp.ParagraphHeading.HEADING1
    });
    this.lastHeading1 = text;
  }

  /**
   * Extracts the first and full name from a "Nome:" heading.
   * @private
   */
  _extractName(heading1Text) {
    const parts = heading1Text.split(":");
    if (parts.length > 1) {
      return parts[1].trim();
    }
    return null;
  }

  /**
   * Checks if a rule is for HEADING1.
   * @private
   */
  _isHeading1(appliedRule) {
    return appliedRule.heading === DocumentApp.ParagraphHeading.HEADING1;
  }

  /**
   * Checks if a rule is for HEADING2.
   * @private
   */
  _isHeading2(appliedRule) {
    return appliedRule.heading === DocumentApp.ParagraphHeading.HEADING2;
  }

  /**
   * Escapes a string for use in a regular expression.
   * @private
   */
  _escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
  }
}

/**
 * Capitalizes the first letter of a string.
 * @param {string} str - The string to capitalize.
 * @returns {string} The string with the first letter capitalized.
 */
function capitalizeFirstLetter(str) {
  if (!str || typeof str !== 'string') return str;
  // Capitalizes the first letter of each word
  return str.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
} 