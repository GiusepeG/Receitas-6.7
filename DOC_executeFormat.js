function executeFormat() {

  const headlineManager = new DocumentHeadlineManager();
  let patientName = null;

  // Tenta extrair o nome do paciente do primeiro H1 encontrado no documento
  if (headlineManager.bodyStructure && headlineManager.bodyStructure.length > 0) {
    const firstHeadline1 = headlineManager.bodyStructure[0].headline1;
    if (firstHeadline1) {
      const parts = firstHeadline1.split(':');
      if (parts.length > 1) {
        const extractedName = parts[1].trim();
        patientName = capitalizeFirstLetter(extractedName);
      }
    }
  }

  const bodyText = DocumentApp.getActiveDocument().getBody().getText();

  const rulesForPlaceholders = getRulesForPlaceholders();

  const rulesForWholeParagraphs = getRulesForWholeParagraphs();
  const rulesForHeading1 = filterRulesByHeading(rulesForWholeParagraphs, DocumentApp.ParagraphHeading.HEADING1);
  const rulesForHeading2 = filterRulesByHeading(rulesForWholeParagraphs, DocumentApp.ParagraphHeading.HEADING2);

  const rulesForInsideParagraphs = getRulesForInsideParagraphs();
  const priorityHeadlines = ["Prontuário Médico", "Laudo de Mapeamento de Retina"];

  const textManipulator = new TextManipulator(bodyText, rulesForPlaceholders, rulesForHeading1, rulesForHeading2);
  textManipulator.currentPatientName = patientName;

  const processedLines = textManipulator
  .withLogging(false)
  .splitInLines()
  .removeEmptyLines()
  .trimLeadingSpaces()
  .removeEscapeCharacters()
  .checkAndModifyFirstLine()
  .processLines()
  .structureH2Blocks(priorityHeadlines)
  .logDictionaries()
  .processedLines;

  // Use ParagraphManipulator to format the document
  const paragraphManipulator = new ParagraphManipulator(processedLines, rulesForWholeParagraphs, rulesForInsideParagraphs);
  paragraphManipulator
    .withLogging(false)
    .clearDocument()
    .formatParagraphs()
    .formatTextInsideParagraphs()
    .cleanupDocument();

  // TODO: Implement the document manipulator
}

/// UTILITIES ///
/**
 * Filter rules by heading
 * @param {Array} rules - The rules to filter
 * @param {DocumentApp.ParagraphHeading} heading - The heading to filter by
 * @returns {Array} The filtered rules
 */
function filterRulesByHeading(rules, heading) {
  return rules.filter(function(rule) {
    return rule.heading === heading;
  });
}