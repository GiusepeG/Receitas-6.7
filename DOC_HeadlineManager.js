class DocumentHeadlineManager {
    constructor() {
      // Obtém automaticamente os dados necessários do documento ativo
      const doc = DocumentApp.getActiveDocument();
      const bodyText = doc.getBody().getText();
      
      // Obtém o texto selecionado (pode ser vazio se nada estiver selecionado)
      let selectionText = '';
      try {
        const selection = doc.getSelection();
        if (selection) {
          selectionText = selection.getText();
        }
      } catch (e) {
        // Se não houver seleção, selectionText permanece vazio
      }
      
      // Obtém as regras de formatação
      const headline1Conditions = getFormattingRules().filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING1);
      const headline2Conditions = getFormattingRules().filter(rule => rule.heading === DocumentApp.ParagraphHeading.HEADING2);
      
      this.bodyText = bodyText;
      this.selectionText = selectionText;

      this.bodyStructure = this.buildDocumentStructure(this.bodyText, headline1Conditions, headline2Conditions);
      this.selectionStructure = selectionText 
        ? this.buildDocumentStructure(this.selectionText, headline1Conditions, headline2Conditions, true) 
        : null;

    }
  
    buildDocumentStructure(text, headline1Conditions, headline2Conditions, limitToOneHeadline2 = false) {
      const lines = text.split('\n');
      let currentHeadline1 = null;
      let documentStructure = [];
      let foundFirstStructure = false;
      let awaitingHeadline2 = false;

      lines.forEach((line, idx) => {
  
        if (headline1Conditions.some(rule => rule.condition(line))) {
          currentHeadline1 = line;
          awaitingHeadline2 = true;
        }
        else if (headline2Conditions.some(rule => rule.condition(line))) {
          if (limitToOneHeadline2 && foundFirstStructure) return;
          documentStructure.push({
            headline1: currentHeadline1,
            headline2: line,
            text: ''
          });
          foundFirstStructure = true;
          awaitingHeadline2 = false;
        }
        else if (awaitingHeadline2) {
          documentStructure.push({
            headline1: currentHeadline1,
            headline2: "Título",
            text: line + '\n'
          });
          foundFirstStructure = true;
          awaitingHeadline2 = false;
        }
        else if (documentStructure.length > 0) {
          documentStructure[documentStructure.length - 1].text += line + '\n';
        }
      });
  
      return documentStructure;
    }
  
    GroupedHeadlines() {
      const groupedStructure = [];
      const headline1Map = new Map();
  
      this.bodyStructure.forEach(item => {
        if (!headline1Map.has(item.headline1)) {
          headline1Map.set(item.headline1, []);
        }
        headline1Map.get(item.headline1).push(item);
      });
  
      headline1Map.forEach((items, headline1) => {
        // Ordena os itens para colocar "Prontuário Médico" no topo de cada grupo
        const sortedItems = items.sort((a, b) => {
          if (a.headline2 === "Prontuário Médico" && b.headline2 !== "Prontuário Médico") {
            return -1; // a vem antes
          }
          if (a.headline2 !== "Prontuário Médico" && b.headline2 === "Prontuário Médico") {
            return 1; // b vem antes
          }
          return 0; // mantém ordem original para outros casos
        });
        
        groupedStructure.push(...sortedItems);
      });
  
      this.bodyStructure = groupedStructure;
      
      return this;
    }
  
    hasHeadline1AndHeadline2(structure, headline2) {
      return structure && structure.some(doc => doc.headline2 === headline2);
    }
  
    hasHeadline1AndHeadline2InSelection(headline2) {
      return this.hasHeadline1AndHeadline2(this.selectionStructure, headline2);
    }
  
    hasHeadline1AndHeadline2InBody(headline2) {
      return this.hasHeadline1AndHeadline2(this.bodyStructure, headline2);
    }
  
    getHeadline1fromSelection() {
      if (this.selectionStructure && this.selectionStructure.length > 0) {
        return this.selectionStructure[0].headline1;
      } else {
        throw new Error("Nenhum headline1 encontrado na seleção.");
      }
    }
  
    getHeadline1fromBody(headline2) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline2 === headline2) {
            return item.headline1;
          }
        }
        throw new Error(`Nenhum headline1 encontrado para o headline2: ${headline2}`);
      } else {
        throw new Error("Nenhuma estrutura encontrada no corpo do documento.");
      }
    }
  
    getFirstHeadline2fromBody(headline1) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline1 === headline1) {
            return item.headline2;
          }
        }
        throw new Error(`Nenhum headline2 encontrado para o headline1: ${headline1}`);
      } else {
        throw new Error("Nenhum headline1 encontrado no corpo do documento.");
      }
    }
  
    getFirstTextfromBody(headline2) {
      if (this.bodyStructure && this.bodyStructure.length > 0) {
        for (const item of this.bodyStructure) {
          if (item.headline2 === headline2) {
            return item.text;
          }
        }
        throw new Error(`Nenhum texto encontrado para o headline2: ${headline2}`);
      } else {
        throw new Error("Nenhuma estrutura encontrada no corpo do documento.");
      }
    }
  
    /**
     * Obtém apenas os Headline1 únicos do documento
     * @return {Array<string>} Array com os Headline1 únicos
     */
    getUniqueHeadline1s() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }
      
      // Extrai apenas os Headline1 únicos e não vazios
      const uniqueHeadlines = [...new Set(
        this.bodyStructure
          .map(item => item.headline1)
          .filter(Boolean) // Remove valores null/undefined/vazios
      )];
      
      return uniqueHeadlines;
    }
  
    /**
     * Obtém todos os Headline2 associados a um Headline1 específico
     * @param {string} headline1 - O Headline1 para buscar
     * @return {Array<string>} Array com os Headline2 associados
     */
    getHeadline2sForHeadline1(headline1) {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }
      
      const headline2s = this.bodyStructure
        .filter(item => item.headline1 === headline1)
        .map(item => item.headline2)
        .filter(Boolean);
      
      return [...new Set(headline2s)]; // Remove duplicatas
    }
  
    /**
     * Reúne todo o conteúdo de todos os H2 sob um H1 específico.
     * @param {string} headline1 - O Headline1 (paciente) para o qual o conteúdo será reunido.
     * @return {string|null} Uma string única com todo o conteúdo ou null se o H1 não for encontrado.
     */
    getAugmentedContentForH1(headline1) {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        console.warn(`[getAugmentedContentForH1] A estrutura do corpo está vazia. Não é possível buscar por "${headline1}".`);
        return null;
      }
  
      // Filtra todos os itens que pertencem ao headline1 especificado
      const patientSections = this.bodyStructure.filter(item => item.headline1 === headline1);
  
      if (patientSections.length === 0) {
        console.warn(`[getAugmentedContentForH1] Nenhum conteúdo encontrado para o paciente: "${headline1}".`);
        return null; // Retorna nulo se nenhum H2 for encontrado para este H1
      }
      // Concatena o texto de todas as seções, cada uma com seu H2
      const combinedContent = patientSections
        .map(section => {
          // Garante que o texto não seja nulo ou indefinido antes de juntar
          const textContent = section.text || '';
          // Remove a quebra de linha final para evitar espaçamento duplo
          const cleanedContent = textContent.endsWith('\n') ? textContent.slice(0, -1) : textContent;
          return `${section.headline2}\n${cleanedContent}`;
        })
        .join('\n\n'); // Separa as seções com uma linha em branco
  
      
      return combinedContent;
    }
  
         /**
      * Cria ou atualiza uma estrutura H1/H2 no documento
      * @param {string} newText - Novo texto a ser inserido
      * @param {string} headline1 - Headline1 (paciente)
      * @param {string} newHeadline2 - Headline2 de destino
      * @param {string} fromHeadline2 - Headline2 de origem (opcional)
      * 
      * REGRAS DE COMPORTAMENTO:
      * - Se fromHeadline2 === newHeadline2: ATUALIZAÇÃO - Remove todas as estruturas antigas {H1, H2} e substitui pela nova
      * - Se fromHeadline2 = "Prontuário Médico" e newHeadline2 = "Prescrição de Óculos": ATUALIZAÇÃO - Mantém "Prontuário Médico" e atualiza "Prescrição de Óculos"
      * - Se fromHeadline2 !== newHeadline2: ADIÇÃO - Adiciona nova estrutura, mantendo todas as existentes
      * - Se fromHeadline2 for null: ADIÇÃO - Comportamento padrão de adicionar nova estrutura
      * - Evita duplicatas exatas (mesmo H1, H2 e texto)
      */
    createOrUpdateBodyHeadline2(newText, headline1, newHeadline2, fromHeadline2 = null) {
      let cleanedText = newText;
      // Remove o headline2 do início do texto, permitindo espaços em branco antes
      if (cleanedText.trim().startsWith(newHeadline2)) {
        cleanedText = cleanedText.substring(cleanedText.indexOf(newHeadline2) + newHeadline2.length);
        // Se o restante for apenas espaços/quebras de linha, remove-os
        if (cleanedText.trim().length === 0) {
          cleanedText = '';
        } else {
          // Remove apenas a quebra de linha inicial, se houver
          cleanedText = cleanedText.startsWith('\n') ? cleanedText.substring(1) : cleanedText;
        }
      }

      // ✅ REGRA ESPECÍFICA BRUSH&MAGIC: Se fromHeadline2 === toHeadline2, é uma ATUALIZAÇÃO
      if (fromHeadline2 && fromHeadline2 === newHeadline2) {
        // Remove todas as estruturas antigas com o mesmo H1 e H2
        this.bodyStructure = this.bodyStructure.filter(doc => 
          !(doc.headline1 === headline1 && doc.headline2 === fromHeadline2)
        );
        
        // Adiciona a nova estrutura atualizada
        this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
        return;
      }
      
      // ✅ REGRA ESPECÍFICA: Prontuário Médico → Prescrição de Óculos (mantém prontuário, atualiza prescrição)
      if (fromHeadline2 === "Prontuário Médico" && newHeadline2 === "Prescrição de Óculos") {
        // Remove todas as estruturas antigas de "Prescrição de Óculos" para este H1
        this.bodyStructure = this.bodyStructure.filter(doc => 
          !(doc.headline1 === headline1 && doc.headline2 === "Prescrição de Óculos")
        );
        
        // Adiciona a nova estrutura de "Prescrição de Óculos" (mantém o Prontuário Médico)
        this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
        return;
      }
      
      // ✅ Verifica se já existe um item EXATAMENTE igual (mesmo H1, H2 e texto)
      const exactMatch = this.bodyStructure.find(doc => 
        doc.headline1 === headline1 && 
        doc.headline2 === newHeadline2 && 
        doc.text === cleanedText
      );
      
      if (exactMatch) {
        // Se o item é exatamente igual, não faz nada para evitar duplicatas totais
        return;
      }
      
      // ✅ Permite múltiplos H2s iguais para o mesmo H1 - sempre adiciona nova entrada
      this.bodyStructure.push({ headline1, headline2: newHeadline2, text: cleanedText });
    }
  
    /**
     * Obtém os pares de Headline1 e Headline2 do documento
     * @return {Array} Array de objetos com headline1 e headline2
     */
    getDocumentHeadlinePairs() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }
      
      return this.bodyStructure.map(item => ({
        headline1: item.headline1,
        headline2: item.headline2
      }));
    }
  
    /**
     * Obtém os pares de Headline1 e Headline2 do documento junto com seu conteúdo
     * @return {Array} Array de objetos com headline1, headline2 e content
     */
    getDocumentHeadlinePairsAndContent() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return [];
      }
      
      return this.bodyStructure.map(item => ({
        headline1: item.headline1,
        headline2: item.headline2,
        content: item.text ? item.text.trim() : ''
      }));
    }
  
    /**
     * Obtém o conteúdo de texto para um par H1/H2 específico.
     * @param {string} headline1 - O Headline1 a ser buscado.
     * @param {string} headline2 - O Headline2 a ser buscado.
     * @return {string|null} O conteúdo de texto, ou null se não encontrado.
     */
    getContentForPair(headline1, headline2) {
      const item = this.bodyStructure.find(doc => doc.headline1 === headline1 && doc.headline2 === headline2);
      if (item && typeof item.text !== 'undefined' && item.text !== null) {
        return item.text;
      }
      // Retorna string vazia se o par existe mas não tem texto,
      // e null se o par não foi encontrado de todo.
      return item ? '' : null;
    }
  
    /**
     * Agrupa os Headline2 conforme seu Headline1 e retorna o texto organizado
     * @return {string} Texto do documento com Headline2 agrupados por Headline1
     */
    getGroupedText() {
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        return '';
      }
  
      // Agrupa os itens por Headline1
      const groupedByHeadline1 = new Map();
      
      this.bodyStructure.forEach(item => {
        if (!groupedByHeadline1.has(item.headline1)) {
          groupedByHeadline1.set(item.headline1, []);
        }
        groupedByHeadline1.get(item.headline1).push(item);
      });
  
      // Constrói o texto agrupado
      let groupedText = '';
      
      groupedByHeadline1.forEach((items, headline1) => {
        // Adiciona o Headline1 uma vez
        groupedText += `${headline1}\n`;
        
        // Adiciona todos os Headline2 e seus textos associados
        items.forEach(item => {
          groupedText += `${item.headline2}\n`;
          if (item.text) { // Garante que a propriedade text exista
            // Remove o último \n para evitar linhas duplas, mas mantém o resto
            const textContent = item.text.endsWith('\n') ? item.text.slice(0, -1) : item.text;
            groupedText += `${textContent}\n`;
          }
        });
        
        // Adiciona uma linha em branco entre grupos
        groupedText += '\n';
      });
  
      return groupedText.trim();
    }
  
    getText() {
      return this.bodyStructure
        .map(item => {
          const textContent = item.text || '';
          const cleanedContent = textContent.endsWith('\n') ? textContent.slice(0, -1) : textContent;
          return `${item.headline1}\n${item.headline2}\n${cleanedContent}`;
        })
        .join("\n\n");
    }
  
    logDocumentStructure(title = "Document Structure") {
      Logger.log(`=== ${title} ===`);
      if (!this.bodyStructure || this.bodyStructure.length === 0) {
        Logger.log("Estrutura do documento vazia.");
        return;
      }
      this.bodyStructure.forEach((item, idx) => {
        Logger.log(`[${idx}] Headline1: "${item.headline1}" | Headline2: "${item.headline2}" | Texto: "${(item.text || '')}"`);
      });
      Logger.log(`=== Fim de ${title} ===`);
    }
  
  
    // ✅ Novo método embutido para processar e atualizar o documento
    processAndUpdateDocument(newText, headline1, toHeadline2, doc, fromHeadline2 = null) {
  
      this.createOrUpdateBodyHeadline2(newText, headline1, toHeadline2, fromHeadline2);
  
      this.GroupedHeadlines();
      const newBodyText = this.getText();
      doc.getBody().setText(newBodyText);
    }
  }
  