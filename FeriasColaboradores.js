function f_calculaDiasDeFerias() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rangeInicio = planilha.getRange("L8:L").getValues();
    var rangeFim = planilha.getRange("M8:M").getValues();
    var colunaResultados = planilha.getRange("O8:O");
  
    // Limpar a coluna O antes de preencher os novos valores
    colunaResultados.clearContent();
  
    var hoje = new Date();
  
    // Ajustar para pegar a semana passada (segunda a sexta-feira)
    var ultimaSegunda = new Date(hoje);
    ultimaSegunda.setDate(hoje.getDate() - hoje.getDay() - 6); // Segunda-feira da semana passada
  
    var ultimaSexta = new Date(ultimaSegunda);
    ultimaSexta.setDate(ultimaSegunda.getDate() + 4); // Sexta-feira da semana passada
  
    // Obter feriados da coluna B8:B
    var rangeFeriados = planilha.getRange("B8:B").getValues();
    var feriados = rangeFeriados.flat().filter(String).map(data => {
      var d = new Date(data);
      return !isNaN(d.getTime()) ? d.toISOString().split('T')[0] : null;
    }).filter(Boolean);
  
    function formatarData(data) {
      return data.toISOString().split('T')[0];
    }
  
    function contarDiasUteis(dataInicio, dataFim) {
      var count = 0;
      var dataAtual = new Date(dataInicio);
      while (dataAtual <= dataFim) {
        var diaSemana = dataAtual.getDay();
        var dataFormatada = formatarData(dataAtual);
        if (diaSemana !== 0 && diaSemana !== 6 && !feriados.includes(dataFormatada)) count++;
        dataAtual.setDate(dataAtual.getDate() + 1);
      }
      return count;
    }
  
    var resultados = [];
    for (var i = 0; i < rangeInicio.length; i++) {
      var dataInicio = new Date(rangeInicio[i][0]);
      var dataFimOriginal = new Date(rangeFim[i][0]);
  
      if (!isNaN(dataInicio.getTime()) && !isNaN(dataFimOriginal.getTime())) {
        var dataFim = new Date(dataFimOriginal);
        dataFim.setDate(dataFim.getDate() + 1); // adiciona 1 dia ao fim das férias
  
        var inicioContagem = dataInicio < ultimaSegunda ? ultimaSegunda : dataInicio;
        var fimContagem = dataFim > ultimaSexta ? ultimaSexta : dataFim;
  
        if (fimContagem >= inicioContagem) {
          resultados.push([contarDiasUteis(inicioContagem, fimContagem)]);
        } else {
          resultados.push([0]);
        }
      } else {
        resultados.push([""]);
      }
    }
  
    colunaResultados.setValues(resultados);
  }
  