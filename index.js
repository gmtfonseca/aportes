const linhasCategoria = {
  rendaFixa: 8,
  acoes: 9,
  fiis: 10,
  stocks: 13,
  reits: 14,
  cripto: 17,
};

const numLinhas = 50;
const aporteMinDolar = 15;

const macroAlocacaoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Macro Alocação");

function estimarAportes() {
  const valorAporte = macroAlocacaoSheet.getRange("B1").getValue();
  const cotacaoDolar = macroAlocacaoSheet.getRange("E1").getValue();

  const macroDiferencas = calcularMacroDiferencas(valorAporte);

  const alocacoes = calcularAlocacoes(macroDiferencas, valorAporte);

  const diferencas = calcularDiferencas(alocacoes, cotacaoDolar);

  Logger.log(JSON.stringify({ macroDiferencas, alocacoes, diferencas }, null, 2));

  const aportes = calcularAportes(alocacoes, diferencas, cotacaoDolar);

  Logger.log(JSON.stringify(aportes, null, 2));

  atualizarPlanilha(alocacoes, aportes);
}

function calcularMacroDiferencas(valorAporte) {
  const totalAtual = macroAlocacaoSheet.getRange("B19").getValue();
  const totalFuturo = totalAtual + valorAporte;

  const macroDiferencas = Object.fromEntries(
    Object.entries(linhasCategoria).map(([categoria, linha]) => {
      const percentualObjetivo = macroAlocacaoSheet.getRange(`E${linha}`).getValue();
      const valorMercadoAtual = macroAlocacaoSheet.getRange(`B${linha}`).getValue();

      const diferenca = percentualObjetivo - valorMercadoAtual / totalFuturo;

      return [categoria, diferenca];
    })
  );

  return macroDiferencas;
}

function calcularAlocacoes(diferencas, valorAporte) {
  const totalNormalizado = Object.values(diferencas).reduce((soma, valor) => soma + Math.max(valor, 0), 0);

  const percentuais = Object.fromEntries(
    Object.entries(diferencas).map(([key, valor]) => [key, Math.max(valor, 0) / totalNormalizado])
  );

  return Object.fromEntries(Object.entries(percentuais).map(([key, percentual]) => [key, percentual * valorAporte]));
}

function calcularDiferencas(alocacoes, cotacaoDolar) {
  return {
    rendaFixa: calcularDiferencasRendaFixa(alocacoes.rendaFixa),
    acoes: calcularDiferencasAcoes(alocacoes.acoes),
    fiis: calcularDiferencasFiis(alocacoes.fiis),
    stocks: calcularDiferencasStocks(alocacoes.stocks, cotacaoDolar),
    reits: calcularDiferencasReits(alocacoes.reits, cotacaoDolar),
    cripto: calcularDiferencasCripto(alocacoes.cripto, cotacaoDolar),
  };
}

function calcularDiferencasRendaFixa(valorAporte) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Renda Fixa");
  const totalAtual = sheet.getRange("B2").getValue();
  const totalFuturo = totalAtual + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    valorAtual: 1,
    percentualObjetivo: 3,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual];
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          valorAtual,
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularDiferencasAcoes(valorAporte) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ações");
  const totalAtual = sheet.getRange("H2").getValue();
  const totalFuturo = totalAtual + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    cotacao: 5,
    quantidade: 6,
    valorAtual: 7,
    percentualObjetivo: 9,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual];
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          quantidade: range[i][colunas.quantidade],
          cotacao: range[i][colunas.cotacao],
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularDiferencasFiis(valorAporte) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FIIs");
  const totalAtual = sheet.getRange("G2").getValue();
  const totalFuturo = totalAtual + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    cotacao: 4,
    quantidade: 5,
    valorAtual: 6,
    percentualObjetivo: 8,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual];
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          quantidade: range[i][colunas.quantidade],
          cotacao: range[i][colunas.cotacao],
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularDiferencasStocks(valorAporte, cotacaoDolar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stocks");
  const totalAtual = sheet.getRange("F2").getValue();
  const totalFuturo = totalAtual * cotacaoDolar + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    cotacao: 3,
    quantidade: 4,
    valorAtual: 5,
    percentualObjetivo: 7,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual] * cotacaoDolar;
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          quantidade: range[i][colunas.quantidade],
          cotacao: range[i][colunas.cotacao],
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularDiferencasReits(valorAporte, cotacaoDolar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REITs");
  const totalAtual = sheet.getRange("F2").getValue();
  const totalFuturo = totalAtual * cotacaoDolar + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    cotacao: 3,
    quantidade: 4,
    valorAtual: 5,
    percentualObjetivo: 7,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual] * cotacaoDolar;
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          quantidade: range[i][colunas.quantidade],
          cotacao: range[i][colunas.cotacao],
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularDiferencasCripto(valorAporte, cotacaoDolar) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cripto");
  const totalAtual = sheet.getRange("D2").getValue();
  const totalFuturo = totalAtual * cotacaoDolar + valorAporte;

  const linhaInicial = 2;

  const colunas = {
    ticker: 0,
    cotacao: 1,
    quantidade: 2,
    valorAtual: 3,
    percentualObjetivo: 5,
  };

  const ultimaColuna = Math.max(...Object.values(colunas)) + 1;

  const range = sheet.getRange(1, 1, numLinhas, ultimaColuna).getValues();

  const ativos = [];
  for (let i = linhaInicial; i < numLinhas; i++) {
    const percentualObjetivo = range[i][colunas.percentualObjetivo];

    if (percentualObjetivo > 0) {
      const valorAtual = range[i][colunas.valorAtual] * cotacaoDolar;
      const valorIdeal = totalFuturo * percentualObjetivo;
      const valorDiferenca = valorIdeal - valorAtual;

      if (valorDiferenca > 0) {
        ativos.push({
          ticker: range[i][colunas.ticker],
          quantidade: range[i][colunas.quantidade],
          cotacao: range[i][colunas.cotacao],
          valorDiferenca,
          linha: i + 1,
        });
      }
    }
  }

  return ativos;
}

function calcularAportes(alocacoes, diferencas, cotacaoDolar) {
  return {
    rendaFixa: calcularAportesRendaFixa(alocacoes.rendaFixa, diferencas.rendaFixa),
    acoes: calcularAportesAcoes(alocacoes.acoes, diferencas.acoes),
    fiis: calcularAportesFiis(alocacoes.fiis, diferencas.fiis),
    stocks: calcularAportesStocks(alocacoes.stocks, diferencas.stocks, cotacaoDolar),
    reits: calcularAportesReits(alocacoes.reits, diferencas.reits, cotacaoDolar),
    cripto: calcularAportesCripto(alocacoes.cripto, diferencas.cripto, cotacaoDolar),
  };
}

function calcularAportesRendaFixa(valorAlocado, ativos) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = proporcao * valorAlocado;

    if (valorAporte > 0) {
      result.push({
        ticker: ativo.ticker,
        valorAporte,
        valorFinalPlanilha: formatar(ativo.valorAtual + valorAporte, "real"),
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function calcularAportesAcoes(valorAlocado, ativos) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = proporcao * valorAlocado;
    const quantidade = Math.trunc(valorAporte / ativo.cotacao);

    if (quantidade > 0) {
      result.push({
        ticker: ativo.ticker,
        quantidade,
        quantidadeFinal: ativo.quantidade + quantidade,
        valorAporte: quantidade * ativo.cotacao,
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function calcularAportesFiis(valorAlocado, ativos) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = proporcao * valorAlocado;
    const quantidade = Math.trunc(valorAporte / ativo.cotacao);

    if (quantidade > 0) {
      result.push({
        ticker: ativo.ticker,
        quantidade,
        quantidadeFinal: ativo.quantidade + quantidade,
        valorAporte: quantidade * ativo.cotacao,
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function calcularAportesStocks(valorAlocado, ativos, cotacaoDolar) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = (proporcao * valorAlocado) / cotacaoDolar;
    const quantidade = valorAporte / ativo.cotacao;

    if (valorAporte > aporteMinDolar) {
      result.push({
        ticker: ativo.ticker,
        valorAporte,
        quantidade,
        quantidadeFinal: ativo.quantidade + quantidade,
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function calcularAportesReits(valorAlocado, ativos, cotacaoDolar) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = (proporcao * valorAlocado) / cotacaoDolar;
    const quantidade = valorAporte / ativo.cotacao;

    if (valorAporte > aporteMinDolar) {
      result.push({
        ticker: ativo.ticker,
        valorAporte,
        quantidade,
        quantidadeFinal: ativo.quantidade + quantidade,
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function calcularAportesCripto(valorAlocado, ativos, cotacaoDolar) {
  const totalDiferenca = ativos.reduce((prev, curr) => prev + curr.valorDiferenca, 0);

  return ativos.reduce((result, ativo) => {
    const proporcao = ativo.valorDiferenca / totalDiferenca;
    const valorAporte = (proporcao * valorAlocado) / cotacaoDolar;
    const quantidade = valorAporte / ativo.cotacao;

    if (valorAporte > aporteMinDolar) {
      result.push({
        ticker: ativo.ticker,
        valorAporte,
        quantidade,
        quantidadeFinal: ativo.quantidade + quantidade,
        linha: ativo.linha,
      });
    }

    return result;
  }, []);
}

function formatar(num, formato) {
  const formatador = new Intl.NumberFormat(formato === "real" ? "pt-BR" : "en-US", {
    style: "currency",
    currency: formato === "real" ? "BRL" : "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  return formatador.format(num);
}

function atualizarPlanilha(alocacoes, aportes) {
  for (const [categoria, valorAporte] of Object.entries(alocacoes)) {
    const linha = linhasCategoria[categoria];
    macroAlocacaoSheet.getRange(`I${linha}`).setValue(valorAporte);
  }

  const rendaFixaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Renda Fixa");
  for (const aporte of aportes.rendaFixa) {
    rendaFixaSheet.getRange(`H${aporte.linha}`).setValue(aporte.valorAporte);
  }

  const acoesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ações");
  for (const aporte of aportes.acoes) {
    acoesSheet.getRange(`O${aporte.linha}`).setValue(aporte.valorAporte);
    acoesSheet.getRange(`P${aporte.linha}`).setValue(aporte.quantidade);
  }

  const fiisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FIIs");
  for (const aporte of aportes.fiis) {
    fiisSheet.getRange(`N${aporte.linha}`).setValue(aporte.valorAporte);
    fiisSheet.getRange(`O${aporte.linha}`).setValue(aporte.quantidade);
  }

  const stocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stocks");
  for (const aporte of aportes.stocks) {
    stocksSheet.getRange(`M${aporte.linha}`).setValue(aporte.valorAporte);
    stocksSheet.getRange(`N${aporte.linha}`).setValue(aporte.quantidade);
  }

  const reitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REITs");
  for (const aporte of aportes.reits) {
    reitsSheet.getRange(`M${aporte.linha}`).setValue(aporte.valorAporte);
    reitsSheet.getRange(`N${aporte.linha}`).setValue(aporte.quantidade);
  }

  const criptoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cripto");
  for (const aporte of aportes.cripto) {
    criptoSheet.getRange(`K${aporte.linha}`).setValue(aporte.valorAporte);
    criptoSheet.getRange(`L${aporte.linha}`).setValue(aporte.quantidade);
  }
}

function limparAportes() {
  for (const linha of Object.values(linhasCategoria)) {
    macroAlocacaoSheet.getRange(`I${linha}`).clearContent();
  }

  const rendaFixaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Renda Fixa");
  for (let i = 2; i < numLinhas; i++) {
    rendaFixaSheet.getRange(`H${i}`).clearContent();
  }

  const acoesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ações");
  for (let i = 2; i < numLinhas; i++) {
    acoesSheet.getRange(`O${i}`).clearContent();
    acoesSheet.getRange(`P${i}`).clearContent();
  }

  const fiisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FIIs");
  for (let i = 2; i < numLinhas; i++) {
    fiisSheet.getRange(`N${i}`).clearContent();
    fiisSheet.getRange(`O${i}`).clearContent();
  }

  const stocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stocks");
  for (let i = 2; i < numLinhas; i++) {
    stocksSheet.getRange(`M${i}`).clearContent();
    stocksSheet.getRange(`N${i}`).clearContent();
  }

  const reitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REITs");
  for (let i = 2; i < numLinhas; i++) {
    reitsSheet.getRange(`M${i}`).clearContent();
    reitsSheet.getRange(`N${i}`).clearContent();
  }

  const criptoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cripto");
  for (let i = 2; i < numLinhas; i++) {
    criptoSheet.getRange(`K${i}`).clearContent();
    criptoSheet.getRange(`L${i}`).clearContent();
  }
}

// eslint-disable-next-line no-unused-vars
function efetivarAportes() {
  const rendaFixaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Renda Fixa");
  for (let i = 2; i < numLinhas; i++) {
    const valorAporteCell = rendaFixaSheet.getRange(`H${i}`);

    if (valorAporteCell.getValue() > 0) {
      const valorAtualCell = rendaFixaSheet.getRange(`B${i}`);

      valorAtualCell.setValue(valorAtualCell.getValue() + valorAporteCell.getValue());
    }
  }

  const acoesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ações");
  for (let i = 2; i < numLinhas; i++) {
    const quantidadeAporteCell = acoesSheet.getRange(`P${i}`);

    if (quantidadeAporteCell.getValue() > 0) {
      const quantidadeAtualCell = acoesSheet.getRange(`G${i}`);

      quantidadeAtualCell.setValue(quantidadeAtualCell.getValue() + quantidadeAporteCell.getValue());
    }
  }

  const fiisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FIIs");
  for (let i = 2; i < numLinhas; i++) {
    const quantidadeAporteCell = fiisSheet.getRange(`O${i}`);

    if (quantidadeAporteCell.getValue() > 0) {
      const quantidadeAtualCell = fiisSheet.getRange(`F${i}`);

      quantidadeAtualCell.setValue(quantidadeAtualCell.getValue() + quantidadeAporteCell.getValue());
    }
  }

  const stocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stocks");
  for (let i = 2; i < numLinhas; i++) {
    const quantidadeAporteCell = stocksSheet.getRange(`N${i}`);

    if (quantidadeAporteCell.getValue() > 0) {
      const quantidadeAtualCell = stocksSheet.getRange(`E${i}`);

      quantidadeAtualCell.setValue(quantidadeAtualCell.getValue() + quantidadeAporteCell.getValue());
    }
  }

  const reitsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REITs");
  for (let i = 2; i < numLinhas; i++) {
    const quantidadeAporteCell = reitsSheet.getRange(`N${i}`);

    if (quantidadeAporteCell.getValue() > 0) {
      const quantidadeAtualCell = reitsSheet.getRange(`E${i}`);

      quantidadeAtualCell.setValue(quantidadeAtualCell.getValue() + quantidadeAporteCell.getValue());
    }
  }

  const criptoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cripto");
  for (let i = 2; i < numLinhas; i++) {
    const quantidadeAporteCell = criptoSheet.getRange(`L${i}`);

    if (quantidadeAporteCell.getValue() > 0) {
      const quantidadeAtualCell = criptoSheet.getRange(`C${i}`);

      quantidadeAtualCell.setValue(quantidadeAtualCell.getValue() + quantidadeAporteCell.getValue());
    }
  }

  limparAportes();

  macroAlocacaoSheet.getRange("B1").clearContent();
}

// eslint-disable-next-line no-unused-vars
function onEdit(e) {
  if (e.range.getA1Notation() === "B1") {
    const valor = e.range.getValue();

    if (valor > 0) {
      estimarAportes();
    } else {
      limparAportes();
    }
  }
}
