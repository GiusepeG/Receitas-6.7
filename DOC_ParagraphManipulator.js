/**
 * @class ParagraphManipulator
 * @classdesc A class to format and manipulate paragraphs in a Google Document using a fluent, chainable interface.
 */
class ParagraphManipulator {
  /**
   * @param {Array<Object>} processedLines Array of processed lines with text and heading properties.
   * @param {Array<Object>} formattingRules Array of formatting rules for paragraphs.
   * @param {Array<Object>} insideParagraphRules Array of formatting rules for text inside paragraphs.
   */
  constructor(processedLines, formattingRules, insideParagraphRules) {
    this.processedLines = processedLines || [];
    this.formattingRules = formattingRules || [];
    this.insideParagraphRules = insideParagraphRules || [];
    this.shouldLog = false;
    this.document = DocumentApp.getActiveDocument();
    this.body = this.document.getBody();
  }

  /**
   * Enables or disables logging for debugging purposes.
   * @param {boolean} [shouldLog=true]
   * @returns {ParagraphManipulator}
   */
  withLogging(shouldLog = true) {
    this.shouldLog = shouldLog;
    return this;
  }

  /**
   * Clears the document body and prepares it for new content.
   * @returns {ParagraphManipulator} The instance for chaining.
   */
  clearDocument() {
    this._log('Clearing document body...');
    this.body.clear();
    this._log(`Document cleared. Number of children: ${this.body.getNumChildren()}`);
    return this;
  }

  /**
   * Formats all processed lines into the document with proper styling and page breaks.
   * @returns {ParagraphManipulator} The instance for chaining.
   */
  formatParagraphs() {
    if (this.processedLines.length === 0) {
      this._log('No processed lines available for formatting.');
      return this;
    }

    this._log('Starting paragraph formatting...');
    
    const processedH1Texts = new Set();
    const h1Rule = this._getH1Rule();
    const defaultRule = this._getDefaultRule();
    
    let firstMeaningfulParagraphAdded = false;

    this.processedLines.forEach((item, index) => {
      const line = item.text;
      const heading = item.heading;
      
      this._log(`Processing line #${index}: "${line}" with heading: ${heading}`);
      
      // Check if line is completely empty (no characters at all)
      const isLineCompletelyEmpty = line === '';
      
      if (!isLineCompletelyEmpty) {
        this._log(`Line #${index}: Adding paragraph with text: "${line}"`);
        const paragraph = this.body.appendParagraph(line);
        const paragraphNativeIndex = this.body.getChildIndex(paragraph);
        
        firstMeaningfulParagraphAdded = true;
        
        // Apply formatting based on heading or rules
        this._applyFormatting(paragraph, line, heading, h1Rule, defaultRule, processedH1Texts);
        
        // Handle page breaks
        this._handlePageBreak(paragraph, paragraphNativeIndex, line, h1Rule);
        
      } else {
        this._log(`Line #${index}: Completely empty line detected`);
        if (firstMeaningfulParagraphAdded) {
          const emptyPara = this.body.appendParagraph(line);
          this._log(`Line #${index}: Empty paragraph added at index: ${this.body.getChildIndex(emptyPara)}`);
        } else {
          this._log(`Line #${index}: Initial empty line ignored`);
        }
      }
    });

    this._log('Paragraph formatting completed.');
    return this;
  }

  /**
   * Performs final cleanup of the document.
   * @returns {ParagraphManipulator} The instance for chaining.
   */
  cleanupDocument() {
    this._log('Starting document cleanup...');
    
    const paragraphs = this.body.getParagraphs();
    this._log(`Total paragraphs before cleanup: ${paragraphs.length}`);
    
    if (paragraphs.length > 0) {
      const firstPara = paragraphs[0];
      const firstParaText = firstPara.getText();
      
      this._log(`First paragraph text: "${firstParaText}"`);
      
      if (firstParaText.trim() === '') {
        this._log('First paragraph is empty, removing...');
        try {
          firstPara.removeFromParent();
          this._log('First empty paragraph removed successfully.');
        } catch (e) {
          this._log(`Error removing first paragraph: ${e.message}`);
        }
      } else {
        this._log('First paragraph is not empty, no cleanup needed.');
      }
    }
    
    this._log('Document cleanup completed.');
    return this;
  }

  /**
   * Internal log helper.
   * @private
   */
  _log(message) {
    if (this.shouldLog) {
      if (typeof Logger !== 'undefined') {
        Logger.log(message);
      } else {
        console.log(message);
      }
    }
  }

  /**
   * Gets the H1 rule from formatting rules.
   * @private
   */
  _getH1Rule() {
    return this.formattingRules.find(rule => 
      rule && rule.heading === DocumentApp.ParagraphHeading.HEADING1 && rule.matchType !== 'default'
    );
  }

  /**
   * Gets the default rule from formatting rules.
   * @private
   */
  _getDefaultRule() {
    return this.formattingRules.find(rule => rule.matchType === 'default');
  }

  /**
   * Applies formatting to a paragraph based on its heading and rules.
   * @private
   */
  _applyFormatting(paragraph, line, heading, h1Rule, defaultRule, processedH1Texts) {
    if (heading === DocumentApp.ParagraphHeading.HEADING1) {
      if (processedH1Texts.has(line)) {
        this._log(`H1 repeated: "${line}". Applying NORMAL heading with H1 style.`);
        paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
        if (h1Rule && h1Rule.style) {
          paragraph.setAttributes(h1Rule.style);
        }
      } else {
        this._log(`New H1: "${line}". Applying HEADING1.`);
        paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        if (h1Rule && h1Rule.style) {
          paragraph.setAttributes(h1Rule.style);
        }
        processedH1Texts.add(line);
      }
    } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      this._log(`H2: "${line}". Applying HEADING2.`);
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const h2Rule = this.formattingRules.find(rule => 
        rule && rule.heading === DocumentApp.ParagraphHeading.HEADING2
      );
      if (h2Rule && h2Rule.style) {
        paragraph.setAttributes(h2Rule.style);
      }
    } else {
      // Apply formatting rules for regular text
      this._applyFormattingRules(paragraph, line, defaultRule);
    }
  }

  /**
   * Applies formatting rules to a paragraph.
   * @private
   */
  _applyFormattingRules(paragraph, line, defaultRule) {
    let ruleApplied = null;
    
    for (const rule of this.formattingRules) {
      if (!rule || rule.matchType === 'default') continue;
      
      if (rule.condition && typeof rule.condition === 'function' && rule.condition(line)) {
        this._log(`Rule applied: "${line}"`);
        if (rule.heading) paragraph.setHeading(rule.heading);
        if (rule.style) paragraph.setAttributes(rule.style);
        ruleApplied = rule;
        break;
      }
    }
    
    if (!ruleApplied && defaultRule) {
      this._log(`Default rule applied: "${line}"`);
      if (defaultRule.heading) paragraph.setHeading(defaultRule.heading);
      if (defaultRule.style) paragraph.setAttributes(defaultRule.style);
    }
  }

  /**
   * Handles page breaks for paragraphs.
   * @private
   */
  _handlePageBreak(paragraph, paragraphNativeIndex, line, h1Rule) {
    // Check if page break is required
    let requiresPageBreak = false;
    
    if (h1Rule && h1Rule.condition && h1Rule.condition(line)) {
      requiresPageBreak = h1Rule.requiresPageBreak || false;
    }
    
    if (requiresPageBreak) {
      this._log(`Checking page break for paragraph at index ${paragraphNativeIndex}`);
      
      let suppressPageBreak = false;
      
      if (paragraphNativeIndex === 0) {
        suppressPageBreak = true;
        this._log('Page break suppressed (first paragraph)');
      } else if (paragraphNativeIndex === 1) {
        const firstChild = this.body.getChild(0);
        if (firstChild.getType() === DocumentApp.ElementType.PARAGRAPH && 
            firstChild.asParagraph().getText().trim() === '') {
          suppressPageBreak = true;
          this._log('Page break suppressed (second paragraph with ghost first)');
        }
      }
      
      if (!suppressPageBreak) {
        this._log(`Inserting page break before paragraph at index ${paragraphNativeIndex}`);
        this.body.insertPageBreak(paragraphNativeIndex);
      } else {
        this._log('Page break suppressed');
      }
    }
  }

  /**
   * Formats specific texts inside paragraphs based on formatting rules.
   * @returns {ParagraphManipulator} The instance for chaining.
   */
  formatTextInsideParagraphs() {
    if (!this.insideParagraphRules || this.insideParagraphRules.length === 0) {
      this._log('No inside paragraph formatting rules available.');
      return this;
    }

    this._log('Starting formatting inside paragraphs...');

    this.insideParagraphRules.forEach((rule, ruleIndex) => {
      this._log(`Processing rule ${ruleIndex + 1}: ${rule.styleName}`);
      
      if (!rule.targets || !Array.isArray(rule.targets)) {
        this._log(`Warning: Rule ${ruleIndex + 1} has no valid targets.`);
        return;
      }

      rule.targets.forEach((target, targetIndex) => {
        this._log(`Processing target ${targetIndex + 1}: "${target}"`);
        
        let searchResult = this.body.findText(target);
        let matchCount = 0;
        
        while (searchResult !== null) {
          const element = searchResult.getElement();
          
          // Check if the element is of text type
          if (element.getType() === DocumentApp.ElementType.TEXT) {
            const textElement = element.asText();
            const startOffset = searchResult.getStartOffset();
            const endOffset = searchResult.getEndOffsetInclusive();
            
            this._log(`Found match ${matchCount + 1} for "${target}" at positions ${startOffset}-${endOffset}`);

            // Apply the formatting based on the styleName
            switch (rule.styleName) {
              case 'bold':
                textElement.setBold(startOffset, endOffset, true);
                this._log(`Applied bold formatting to "${target}"`);
                break;
              case 'italic':
                textElement.setItalic(startOffset, endOffset, true);
                this._log(`Applied italic formatting to "${target}"`);
                break;
              case 'underline':
                textElement.setUnderline(startOffset, endOffset, true);
                this._log(`Applied underline formatting to "${target}"`);
                break;
              default:
                this._log(`Warning: Unknown style name "${rule.styleName}" for target "${target}"`);
                break;
            }
            
            matchCount++;
          } else {
            this._log(`Element is not of text type: ${element.getType()}`);
          }

          // Find the next match
          searchResult = this.body.findText(target, searchResult);
        }
        
        this._log(`Total matches found for "${target}": ${matchCount}`);
      });
    });

    this._log('Inside paragraph formatting completed.');
    return this;
  }
} 