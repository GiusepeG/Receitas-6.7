function getRulesForInsideParagraphs() {
    return [
      {
        styleName: 'italic',
        targets: [
          "screening",
          "floaters",
          "flashes",
          "lattice",
          "leak point"
        ],
      },
      {
        styleName: 'bold',
        targets: [
          "No olho direito",
          "no olho direito",
          "No olho esquerdo",
          "no olho esquerdo",
          "Em ambos os olhos",
          "em ambos os olhos",
          "com urgência",
          "3219 9700"
        ],
      },
      {
        styleName: 'underline',
        targets: [
          "OD:",
          "OE:",
          "AO:",
          "ORIENTAÇÕES",
          "CID 10:",
          "composição:"
        ],
      },
    ];
  }
  