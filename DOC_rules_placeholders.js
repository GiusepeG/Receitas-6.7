function getRulesForPlaceholders() {
    return [
      {
        text: "{Nome}",
        replacement: (patientName) => {
          if (!patientName || typeof patientName !== 'string') return "Nome";
          const firstName = patientName.split(' ')[0];
          return firstName;
        }
      },
      {
        text: "{Nome Completo}",
        replacement: (patientName) => patientName || "Nome Completo"
      },
      {
        text: "{Secretária do consultório 13}",
        replacement: PropertiesService.getScriptProperties().getProperty('secretary13')
      },
      {
        text: "{Secretária do consultório 12}",
        replacement: PropertiesService.getScriptProperties().getProperty('secretary12')
      },
      {
        text: "{HH:MM}",
        replacement: (patientName) => {
          const now = new Date();
          let hours = now.getHours();
          let minutes = now.getMinutes();

          // Arredonda para o próximo múltiplo de 15 minutos
          let roundedMinutes = Math.ceil(minutes / 15) * 15;
          if (roundedMinutes === 60) {
            roundedMinutes = 0;
            hours = (hours + 1) % 24;
          }

          // Formata com zero à esquerda se necessário
          const hoursStr = hours.toString().padStart(2, '0');
          const minutesStr = roundedMinutes.toString().padStart(2, '0');
          return `${hoursStr}:${minutesStr}`;
        }
      },
      {
        text: "{À}",
        replacement: "À disposição para quaisquer outros esclarecimentos,"
      },
      {
        text: "{prontuário médico}",
        replacement: "Prontuário Médico"
      },
      {
        text: "{encaminhamento}",
        replacement: "Encaminhamento"
      },
      {
        text: "{laudo de mapeamento de retina}",
        replacement: "Laudo de Mapeamento de Retina"
      },
      {
        text: "{relatório médico}",
        replacement: "Relatório Médico"
      },
      {
        text: "{laudo de tomografia de coerência óptica}",
        replacement: "Laudo de Tomografia de Coerência Óptica"
      },
      {
        text: "{laudo de angiografia fluoresceínica digital}",
        replacement: "Laudo de Angiografia Fluoresceínica Digital"
      },
      {
        text: "{laudo de retinografia colorida digital}",
        replacement: "Laudo de Retinografia Colorida Digital"
      }
      // Adicione mais placeholders conforme necessário
    ];
  }