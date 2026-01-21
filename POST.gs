//FUNÇÃO QUE IDENTICA QUAL A FUNÇÃO FOI CHAMADA PELO FRONTEND

function roteadorSolicitacoes(data) {
  if (!data || !data.qualFuncao) {
    return { erro: "Parâmetros inválidos" };
  }

  console.log("Executando: " + data.qualFuncao + " | Matrícula: " + data.matricula);

  switch (data.qualFuncao) {
    case "registrarRetirada":
      return registrarRetiradaEstoque(data.itens);

    case "buscarColaborador": 
      return buscarColaboradorPorMatricula(data);

    case "salvaritem":
      return salvaritem(data);

    case "salvarColaborador":
      return salvarColaborador(data);

    case "buscarItem":
      return buscarItemPorCodigo(data);
    
    case "buscarDadosEstoque":
      return buscarDadosEstoque();
    
    case "salvarMovimentacaoEPI":
      return salvarMovimentacaoEPI(data);

    case "buscarCompletoDevolucao":
      return buscarCompletoDevolucao(data.matricula);

    case "atualizarDevolucaoEPI":
      return atualizarDevolucaoEPI(data);

    case "buscarPendentesDevolucao":
      return buscarPendentesDevolucao();

    case "buscarItensPendentesPorMatricula":
      return buscarItensPendentesPorMatricula(data);
    
    case "salvarDatasDevolucaoFardamento":
      return salvarDatasDevolucaoFardamento(data.lista);

    case "buscarMovimentacoesUniformes":
      return buscarMovimentacoesUniformes(data);

    case "atualizarCardsMetricas":
      return buscarMetricasEstoque();

    case "buscarEntregasHoje":
      return buscarEntregasHoje();

    case "buscarParaDevolucao":
      return buscarParaDevolucao(data);

    case "registrarDevolucao":
      return registrarDevolucao(data);

    case "buscarItemPorCodigoEpi":
      return buscarItemPorCodigoEpi(data);

      case "buscarTodosEpisPendentes":
    return buscarTodosEpisPendentes(data);

      case "processarDevolucaoLote":
    return processarDevolucaoLote(data);

      case "buscarDadosFichaCompleta":
    return buscarDadosFichaCompleta(data);

      case "buscarDadosRelatorioEpi":
    return buscarDadosRelatorioEpi(data);

      /*case "processarLoteDevolucoes":
    return processarLoteDevolucoes(data.listaDevolucoes);*/
    
      case "buscarEPIsPendentes":
    return buscarEPIsPendentes(data);

    default:
      return { erro: "Função '" + data.qualFuncao + "' não reconhecida no roteador." };
  }
}


//FUNÇÃO QUE ALIMENTA O DASHBOARD //

function buscarMetricasEstoque() {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaEstoque = ss.getSheetByName("Estoque");
    const abaMov = ss.getSheetByName("movimentacoes");
    
    const dadosEstoque = abaEstoque.getDataRange().getValues();
    const dadosMov = abaMov.getDataRange().getValues();
    
    const hojeStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

    //CÁLCULOS DO ESTOQUE Card 1 e Card 2 
    let totalModelosCadastrados = dadosEstoque.length - 1; 
    let itensComSaldoPositivo = 0; 
    let itensEstoqueBaixo = 0;     
    let nomeItemMaisCritico = "Nenhum";
    let menorValorEncontrado = Infinity;

    for (let i = 1; i < dadosEstoque.length; i++) {
      let saldoAtual = Number(dadosEstoque[i][3]) || 0; // Coluna D (Índice 3)
      
      //  i 1 Coluna B para i 2 Coluna C
      let nomeItem = dadosEstoque[i][2]; 

      //  Conta itens com valor diferente de 0
      if (saldoAtual > 0) {
        itensComSaldoPositivo++;
      }

      // Contar itens com valor menor que 7
      if (saldoAtual < 7) {
        itensEstoqueBaixo++;
        // identifica qual deles é o menor de todos para exibir o nome
        if (saldoAtual < menorValorEncontrado) {
          menorValorEncontrado = saldoAtual;
          nomeItemMaisCritico = nomeItem;
        }
      }
    }

    // Porcentagens
    let porcentagemDisponibilidade = totalModelosCadastrados > 0 
      ? Math.round((itensComSaldoPositivo / totalModelosCadastrados) * 100) : 0;
    
    let porcentagemCritica = totalModelosCadastrados > 0 
      ? Math.round((itensEstoqueBaixo / totalModelosCadastrados) * 100) : 0;

    // CÁLCULOS DE MOVIMENTAÇÃO  
    let entregasHoje = 0;
    let resumoHoje = {};
    for (let i = 1; i < dadosMov.length; i++) {
      let dataMovStr = (dadosMov[i][1] instanceof Date) 
        ? Utilities.formatDate(dadosMov[i][1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") 
        : String(dadosMov[i][1]);
      
      if (dataMovStr === hojeStr) {
        entregasHoje += Number(dadosMov[i][4]) || 0;
        let itemNome = dadosMov[i][3];
        resumoHoje[itemNome] = (resumoHoje[itemNome] || 0) + 1;
      }
    }
    let itemMaisEntregue = Object.keys(resumoHoje).reduce((a, b) => resumoHoje[a] > resumoHoje[b] ? a : b, "Nenhum");

    return {
      // Card 1
      totalDisponivel: itensComSaldoPositivo, 
      percDisponivel: porcentagemDisponibilidade,
      // Card 2
      totalCritico: itensEstoqueBaixo,
      percCritico: porcentagemCritica,
      nomeCritico: nomeItemMaisCritico, 
      // Card 3
      entregasQtd: entregasHoje,
      entregasNome: itemMaisEntregue,
      dataHoje: hojeStr
    };

  } catch (e) {
    return { erro: e.toString() };
  }
}

// ATUALIZA TABELA DE MOVIMENTAÇÕES //


function buscarEntregasHoje() {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("movimentacoes");
    const dados = abaMov.getDataRange().getValues();
    
    // Pega a data formatada para comparação
    const hojeStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    
    let entregasDeHoje = [];

    // Ignora o cabeçalho i=1
    for (let i = 1; i < dados.length; i++) {
      let dataLinha = dados[i][1]; 
      let dataFormatada = (dataLinha instanceof Date) 
        ? Utilities.formatDate(dataLinha, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") 
        : String(dataLinha);

      if (dataFormatada === hojeStr) {
        entregasDeHoje.push({
          codigo: dados[i][2],      
          descricao: dados[i][3],   
          quantidade: dados[i][4],  
          tamanho: dados[i][5],     
          colaborador: dados[i][7]  
        });
      }
    }

    return entregasDeHoje;

  } catch (e) {
    return { erro: e.toString() };
  }
}

//FUNÇÃO DE CADASTRO DE COLABORADORES //

function salvarColaborador(data) {
  const ss = SpreadsheetApp.openById("");
  const main = ss.getSheetByName("colaboradores");

  // Garante que a aba existe e tem cabeçalho
  if (main.getLastRow() === 0) {
    main.appendRow(["ID", "Matricula", "Nome", "Função", "idcolaborador"]);
  }

  const idExistente = data.idUsuario; 
  const matriculaNova = data.matricula;

  //verificação de duplicidade Apenas para novos cadastros
  if (!idExistente || idExistente === "") {
    const ultimaLinhaAtual = main.getLastRow();
    if (ultimaLinhaAtual > 1) {
      const matriculas = main.getRange(2, 2, ultimaLinhaAtual - 1).getValues().flat();
      if (matriculas.some(m => m.toString() === matriculaNova.toString())) {
        return "ERRO: Esta matrícula já está cadastrada!";
      }
    }
  }

  // LÓGICA DE SALVAMENTO 

  if (!idExistente || idExistente === "") {
    //AÇÃO NOVO CADASTRO 
    const ultimaLinha = main.getLastRow();
    let maiorId = 0;

    if (ultimaLinha > 1) {
      const idsExistentes = main.getRange(2, 1, ultimaLinha - 1).getValues().flat();
      maiorId = idsExistentes.length > 0 ? Math.max(...idsExistentes.map(Number)) : 0;
    }

    const novoId = maiorId + 1;

    main.appendRow([
      novoId,
      data.matricula,
      data.nome,
      data.funcao,
      data.idcolaborador,
    ]);

    return "Usuário cadastrado com sucesso!";

  } else {
    // AÇÃO EDIÇÃO DE USUÁRIO EXISTENTE 
    const ultimaLinha = main.getLastRow();
    const ids = main.getRange(2, 1, ultimaLinha - 1).getValues().flat();
    
    // localiza a linha correta pelo ID Coluna A
    const indexLinha = ids.indexOf(Number(idExistente));

    if (indexLinha !== -1) {
      const linhaParaEditar = indexLinha + 2; 
      
      // Atualiza as colunas B, C, D e E Matricula, Nome, Função
      main.getRange(linhaParaEditar, 2, 1, 4).setValues([[
        data.matricula,
        data.nome,
        data.funcao,
        data.idcolaborador
      ]]);

      return "Dados atualizados com sucesso!";
    } else {
      return "ERRO: ID não encontrado para edição.";
    }
  }
}

//FUNÇÃO BUSCA COLABORADOR  

function buscarColaboradorPorMatricula(data) { 
  const ss = SpreadsheetApp.openById("");
  const sheet = ss.getSheetByName("colaboradores");
  const dados = sheet.getDataRange().getValues();
  const linhas = dados.slice(1);
  
  const matriculaProcurada = data.matricula.toString().trim(); 

  const colaborador = linhas.find(linha => linha[1].toString().trim() === matriculaProcurada);
  
  if (colaborador) {
    return {
      id: colaborador[0],
      matricula: colaborador[1],
      nome: colaborador[2],
      funcao: colaborador[3],
      idcolaborador: colaborador[4]
    };
  } else {
    return null; // Retorna null para o frontend tratar
  }
}


//FUNÇÃO SALVAR ITEM 

function salvaritem(data) {
  const ss = SpreadsheetApp.openById("");
  const sheet = ss.getSheetByName("estoque"); 

  if (!sheet) return "ERRO: Aba 'estoque' não encontrada!";

  // Garante cabeçalho 
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Código", "Descrição", "Quantidade", "Tamanho", "Extra"]);
  }

  const idExistente = data.idItem;
  const codigoNovo = data.codigo;

  if (!idExistente || idExistente === "") {
    const ultimaLinhaAtual = sheet.getLastRow();
    if (ultimaLinhaAtual > 1) {
      // Pega todos os valores da coluna B 
      const codigosExistentes = sheet.getRange(2, 2, ultimaLinhaAtual - 1).getValues().flat();
      
      // Verifica se o código já existe (converte para string para comparar corretamente)
      if (codigosExistentes.some(c => c.toString() === codigoNovo.toString())) {
        return "ERRO: Este código de item já está cadastrado!";
      }
    }
  }

  if (!idExistente || idExistente === "") {
    // AÇÃO NOVO ITEM 
    const ultimaLinha = sheet.getLastRow();
    let novoId = 1;
    
    if (ultimaLinha > 1) {
       const ids = sheet.getRange(2, 1, ultimaLinha - 1).getValues().flat();
       const idsNumericos = ids.map(Number).filter(n => !isNaN(n));
       novoId = idsNumericos.length > 0 ? Math.max(...idsNumericos) + 1 : 1;
    }

    sheet.appendRow([
      novoId,
      data.codigo,
      data.descricao,
      data.quantidade,
      data.tamanho,
      data.extra
    ]);
    return "Item cadastrado com sucesso!";

  } else {
    // AÇÃO EDITAR ITEM 
    const ultimaLinha = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, ultimaLinha - 1).getValues().flat();
    const index = ids.indexOf(Number(idExistente));

    if (index !== -1) {
      const linha = index + 2;
      sheet.getRange(linha, 2, 1, 5).setValues([[
        data.codigo,
        data.descricao,
        data.quantidade,
        data.tamanho,
        data.extra
      ]]);
      return "Item atualizado com sucesso!";
    }
    return "ERRO: Item não encontrado.";
  }
}

function buscarItemPorCodigo(data) {
  const ss = SpreadsheetApp.openById("");
  const sheet = ss.getSheetByName("estoque");
  const dados = sheet.getDataRange().getValues();
  const linhas = dados.slice(1); // Pula o cabeçalho
  
  const codigoProcurado = data.codigo;

  // Busca na coluna B 
  const item = linhas.find(linha => linha[1].toString() === codigoProcurado.toString());
  
  if (item) {
    return {
      id: item[0],
      codigo: item[1],
      descricao: item[2],
      quantidade: item[3],
      tamanho: item[4],
      extra: item[5]
    };
  } else {
    return "ERRO: Item não encontrado no estoque!";
  }
}

//FUNÇÃO QUE SUBTRAI ITENS DO ESTOQUE E SALVA NA ABA MOVIMENTACOES

function registrarRetiradaEstoque(itens) {
  const ss = SpreadsheetApp.openById("");
  const sheetEstoque = ss.getSheetByName("Estoque"); 
  const dataEstoque = sheetEstoque.getDataRange().getValues();
  const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  
  try {
    for (let j = 0; j < itens.length; j++) {
      let item = itens[j];
      let sheetMov = ss.getSheetByName(item.destino);
      
      if (!sheetMov) throw new Error("A aba '" + item.destino + "' não foi encontrada.");

      // Baixa no Estoque
      let itemEncontrado = false;
      let qtdTotal = Number(item.quantidade);

      for (let i = 1; i < dataEstoque.length; i++) {
        if (dataEstoque[i][1].toString() === item.codigo.toString()) {
          itemEncontrado = true;
          let estoqueAtual = Number(dataEstoque[i][3]);
          if (estoqueAtual < qtdTotal) throw new Error("Estoque insuficiente para " + item.descricao);
          sheetEstoque.getRange(i + 1, 4).setValue(estoqueAtual - qtdTotal);
          break;
        }
      }
      if (!itemEncontrado) throw new Error("Código " + item.codigo + " não encontrado.");

      // Gravação Linha por Linha
      for (let k = 0; k < qtdTotal; k++) {
        const lastRow = sheetMov.getLastRow();
        let nextId = 1;
        if (lastRow > 1) {
          const ids = sheetMov.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Number.isFinite);
          if (ids.length > 0) nextId = Math.max(...ids) + 1;
        }

        const novaLinha = [
          nextId,    // A ID
          dataHoje,      // B Data
          item.codigo,      // C Código
          item.descricao,   // D Descrição
          1,         // E Quantidade
          item.tamanho,   // F Tamanho
          item.colaboradorID, // G Matrícula
          item.colaborador, // H Nome
          item.ca,   // I CA
          "",          // J Data Devolução (Vazia)
          item.motivo   // K Motivo (Admissão, Desgaste, etc.)
        ];

        sheetMov.appendRow(novaLinha);
        sheetMov.getRange(sheetMov.getLastRow(), 1).setNumberFormat("0");
      }
    }
    return "Sucesso! Itens registrados com Motivo e CA.";
  } catch (e) {
    return "ERRO: " + e.message;
  }
}

//BUSCA ESTOQUE PARA ALIMENTAR TABELA 

function buscarDadosEstoque() {
  const ss = SpreadsheetApp.openById("");
  const sheet = ss.getSheetByName("estoque");
  
  // Pega todos os dados (pulando o cabeçalho)
  const valores = sheet.getDataRange().getValues();
  const dados = [];
  
  // Começa em 1 para pular a linha de cabeçalho
  for (let i = 1; i < valores.length; i++) {
    dados.push({
      codigo: valores[i][1],      
      descricao: valores[i][2],   
      quantidade: valores[i][3],  
      tamanho: valores[i][4]      
    });
  }
  console.log("Dados processados: ", dados);
  return dados;
}


// FUNÇÃO PARA GRAVAR NA ABA EPIMOVIMENTACOES

function salvarMovimentacaoEPI(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("epimovimentacoes");
    
    // LÓGICA PARA ID SEQUENCIAL 
    const ultimaLinha = abaMov.getLastRow();
    let novoId = 1; // caso a planilha esteja vazia só cabeçalho

    if (ultimaLinha > 1) {
      // Pega o valor da última célula da coluna A
      const ultimoIdValor = abaMov.getRange(ultimaLinha, 1).getValue();
       
      if (!isNaN(ultimoIdValor)) {
        novoId = Number(ultimoIdValor) + 1;
      }
    }

    // MAPEAMENTO DAS COLUNAS (A até J)
    const novaLinha = [
      novoId,      // A 
      data.matricula,  // B
      data.nome,     // C
      data.funcao,    // D
      data.epiSerial,   // E
      data.dataRetirada,  // F
      "",            // G
      data.extra,     // H
      data.descricao,  // I
      data.codigo    // J
    ];

    abaMov.appendRow(novaLinha);

    return "Sucesso: Retirada registrada";
    
  } catch (e) {
    return "ERRO no servidor: " + e.toString();
  }
}

// FUNÇÃO BUSCAR EPI PARA DEVOLUÇÃO 

function buscarCompletoDevolucao(matriculaSolicitada) {
  try {
    const ss = SpreadsheetApp.openById(""); 
    const abaMov = ss.getSheetByName("epimovimentacoes");

    if (!abaMov) return { erro: "Aba 'epimovimentacoes' não encontrada." };

    const dados = abaMov.getDataRange().getValues();
    const matriculaTexto = matriculaSolicitada.toString().trim();

    // Filtra as linhas buscando pela matrícula Coluna B 
    const registros = dados.filter(r => r[1] && r[1].toString().trim() === matriculaTexto);

    if (registros.length === 0) {
      return { erro: "Nenhum registro de retirada encontrado para esta matrícula." };
    }

    // Pega o registro mais recente 
    const ultimaMov = registros[registros.length - 1];
 
    return {
      id: ultimaMov[0],
      matricula: ultimaMov[1],
      nome: ultimaMov[2],
      funcao: ultimaMov[3],
      epi: ultimaMov[4],
      dataRetirada: ultimaMov[5] instanceof Date ? Utilities.formatDate(ultimaMov[5], "GMT-3", "dd/MM/yyyy") : ultimaMov[5]
    };

  } catch (e) {
    return { erro: "Erro no servidor: " + e.message };
  }
}

//SALVA DEVOLUÇÃO DE EPI HIGIENIZAÇÃO

function atualizarDevolucaoEPI(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("epimovimentacoes");
    
    const valores = abaMov.getDataRange().getValues();
    const matriculaBusca = String(data.matricula).trim();
    let linhaEncontrada = -1;

    for (let i = valores.length - 1; i >= 0; i--) {
      if (String(valores[i][1]).trim() === matriculaBusca && !valores[i][6]) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada !== -1) {
      abaMov.getRange(linhaEncontrada, 7).setValue(data.dataDevolucao);
      return "Sucesso: Devolução registrada!"; 
    } else {
      return "Erro: Nenhuma retirada pendente encontrada.";
    }

  } catch (e) {

    // texto com a descrição do erro
    return "Erro no servidor: " + e.toString();
  }
}

//PREENCHE TABELA DE EPIs PEDENTES DE DEVOLUÇÃO 

function buscarPendentesDevolucao() {
  try {
    const ss = SpreadsheetApp.openById("");
    const aba = ss.getSheetByName("epimovimentacoes");
    const dados = aba.getDataRange().getValues();
    
    // Remove o cabeçalho
    dados.shift();

    // Filtra apenas onde a coluna G está vazia
    const pendentes = dados.filter(r => !r[6] || r[6].toString().trim() === "").map(r => {
      return {
        matricula: r[1],
        nome: r[2],
        funcao: r[3],
        epi: r[4],
        dataRetirada: r[5] instanceof Date ? Utilities.formatDate(r[5], "GMT-3", "dd/MM/yyyy") : r[5],
        dataDevolucao: "" 
      };
    });

    return pendentes;
  } catch (e) {
    return { erro: e.message };
  }
}

//BUSCA NOME PARA DEVOLUÇÃO DE FARDAMENTOS

function buscarItensPendentesPorMatricula(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("movimentacoes");
    const valores = abaMov.getDataRange().getValues();
    
    const matriculaBusca = String(data.matricula).trim();
    let pendentes = [];
    let nomeColaborador = "";

    // Percorre a planilha pula cabeçalho i=1
    for (let i = 1; i < valores.length; i++) {
      const matriculaLinha = String(valores[i][6]).trim(); 
      const dataDevolucao = valores[i][8]; 
      
      if (matriculaLinha === matriculaBusca && (!dataDevolucao || dataDevolucao.toString().trim() === "")) {
        nomeColaborador = valores[i][7]; 
        pendentes.push({
          linha: i + 1, 
          itemNome: valores[i][3] 
        });
      }
    }

    if (pendentes.length === 0) {
      return { erro: "Nenhum item pendente encontrado para esta matrícula." };
    }

    return {
      nome: nomeColaborador,
      itens: pendentes
    };
    
  } catch (e) {
    return { erro: e.toString() };
  }
}

// FUNÇÃO PARA SALVAR AS DATAS

function salvarDatasDevolucaoFardamento(listaDevolucao) {
  try {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("movimentacoes");

    listaDevolucao.forEach(dev => {
      if (dev.data && dev.linha) {
        const celula = abaMov.getRange(dev.linha, 9); 
        
        // Define o formato da célula para data antes de inserir o valor
        celula.setNumberFormat("dd/mm/yyyy");
        
        // Insere o valor
        celula.setValue(dev.data);
      }
    });

    return "Sucesso: Devoluções registradas ";
  } catch (e) {
    return "ERRO: " + e.toString();
  }
}


//ACOMPANHAMENTO DE FARDAMENTOS //

function buscarMovimentacoesUniformes(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const aba = ss.getSheetByName("movimentacoes");
    const valores = aba.getDataRange().getValues();
    const matriculaBusca = String(data.matricula).trim();
    
    let resultado = { nome: "", itens: [] };

    for (let i = 1; i < valores.length; i++) {
      if (String(valores[i][6]).trim() === matriculaBusca) {
        if (!resultado.nome) resultado.nome = valores[i][7];

        // CONVERSÃO DE TIPOS 
        let dataRetiradaRaw = valores[i][1];
        let dataDevolucaoRaw = valores[i][8];

        let dataRetiradaStr = (dataRetiradaRaw instanceof Date) 
          ? Utilities.formatDate(dataRetiradaRaw, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy")
          : String(dataRetiradaRaw);

        let dataDevolucaoStr = (dataDevolucaoRaw instanceof Date)
          ? Utilities.formatDate(dataDevolucaoRaw, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy")
          : String(dataDevolucaoRaw);

        resultado.itens.push({
          dataRetirada: dataRetiradaStr, 
          quantidade: valores[i][4],    
          descricao: valores[i][3],     
          tamanho: valores[i][5],       
          dataDevolucao: dataDevolucaoStr
        });
      }
    }

    if (resultado.itens.length === 0) return { erro: "Nenhum registro encontrado." };
    return resultado;

  } catch (e) {
    return { erro: "Erro no processamento: " + e.toString() };
  }
}

//DEVOLUÇÃO DE EPI

// Busca apenas registros da matrícula informada que ainda não foram devolvidos
function buscarParaDevolucao(data) {
  try {
    const ss = SpreadsheetApp.openById("1wGKWaGj3YKd63Kl6IMyFwajSs5Ce0UqvZkPUlduKMd8");
    const sheet = ss.getSheetByName("registroepi");
    if (!sheet) return { erro: "Aba 'registroepi' não encontrada" };

    const dados = sheet.getDataRange().getValues();
    const matriculaProcurada = data.matricula.toString().trim();
    const resultados = [];
    const fusoHorario = ss.getSpreadsheetTimeZone();

    for (let i = 1; i < dados.length; i++) {

      let matriculaLinha = dados[i][6] ? dados[i][6].toString().trim() : "";
      let dataDevolucao = dados[i][9];

      if (matriculaLinha === matriculaProcurada && (!dataDevolucao || dataDevolucao === "")) {
        
        let dataRetiradaStr = dados[i][1] instanceof Date 
          ? Utilities.formatDate(dados[i][1], fusoHorario, "dd/MM/yyyy") 
          : (dados[i][1] ? dados[i][1].toString() : "");

        resultados.push({
          id: dados[i][0].toString(),  // Coluna A
          dataRetirada: dataRetiradaStr,  // Coluna B
          codigo: dados[i][2] || "",    // Coluna C
          descricao: dados[i][3] || "",  // Coluna D
          ca: dados[i][8] || "N/A",  // Coluna I
          colaborador: dados[i][7] || "",  // Coluna H
          motivoExistente: dados[i][10] || "Não informado" // Coluna K
        });
      }
    }
    return resultados; 
  } catch (e) {
    return { erro: e.message };
  }
}

// Localiza o ID na aba e preenche as colunas J e K
function registrarDevolucao(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const sheet = ss.getSheetByName("registroepi");
    
    const lista = data.listaDevolucoes;
    
    if (!lista || lista.length === 0) return { sucesso: false, erro: "Nenhum dado recebido." };

    lista.forEach(item => {
      const linhaPlanilha = parseInt(item.id);
      
      const partesData = item.data.split("/"); // Assume formato dd/mm/yyyy do Flatpickr
      const dataObjeto = new Date(partesData[2], partesData[1] - 1, partesData[0]);

      //  Grava na coluna J 
      if (!isNaN(linhaPlanilha)) {
        sheet.getRange(linhaPlanilha, 10).setValue(dataObjeto);
        sheet.getRange(linhaPlanilha, 10).setNumberFormat("dd/MM/yyyy");
      }
    });

    return { sucesso: true };
    
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor: " + e.message };
  }
}

// FUNÇÃO QUE BUSCA DOS ITENS PARA ALIMENTAR ABA DE HIGIENIZAÇÃO 

function buscarItemPorCodigoEpi(data) {
  const ss = SpreadsheetApp.openById("");
  const sheet = ss.getSheetByName("estoque");
  const dados = sheet.getDataRange().getValues();
  
  const codigoProcurado = data.codigo.toString().trim();
  
  // Procura na Coluna B (índice 1) - Código
  const item = dados.find(linha => linha[1].toString().trim() === codigoProcurado);
  
  if (item) {
    return {
      descricao: item[2], // Coluna C
      extra: item[5]  // Coluna F
    };
  } else {
    return "ERRO: Item não encontrado!";
  }
}

//BUSCA TODOS OS EPIS PEDENTES DE DEVOLUÇÃO 

function buscarTodosEpisPendentes(data) {
  const ss = SpreadsheetApp.openById("");
  const abaMov = ss.getSheetByName("epimovimentacoes");
  const valores = abaMov.getDataRange().getValues();
  
  const matricula = data.matricula.toString().trim();
  const itensEncontrados = [];
  let nomeColaborador = "";

  // Percorre da linha 2 em diante
  for (let i = 1; i < valores.length; i++) {
    const row = valores[i];
    if (row[1].toString().trim() === matricula && row[6] === "") {
      if (!nomeColaborador) nomeColaborador = row[2]; 
      
      itensEncontrados.push({
        idLinha: i + 1,        
        descricao: row[8],     
        epiOriginal: row[4]    
      });
    }
  }

  return {
    nome: nomeColaborador,
    itens: itensEncontrados
  };
}

function processarDevolucaoLote(data) {
  const ss = SpreadsheetApp.openById("");
  const abaMov = ss.getSheetByName("epimovimentacoes");
  
  data.itens.forEach(item => {
    abaMov.getRange(item.idLinha, 7).setValue(item.dataDevolucao);
  });
  
  return "Sucesso: " + data.itens.length + " item(ns) devolvido(s) com êxito!";
}


//GERA RELATORIO DE HIGIENIZAÇÃO DE EPI


function buscarDadosFichaCompleta(data) {
    const ss = SpreadsheetApp.openById("");
    const abaMov = ss.getSheetByName("epimovimentacoes");
    const valores = abaMov.getDataRange().getValues();
    
    const matricula = data.matricula.toString().trim();
    
    // Filtra todos os registros da matrícula
    const registros = valores.filter(r => r[1].toString().trim() === matricula);

    if (registros.length === 0) return { erro: "Matrícula não encontrada." };

    const temCampoEmBranco = registros.some(r => {
        const valorG = r[6];
        return valorG === "" || valorG === null || valorG === undefined;
    });

    if (temCampoEmBranco) {
        return { 
            erro: "Atenção: Existem pendências faça devolução antes de gerar o relatorio." 
        };
    }

    // Retorno normal caso esteja tudo preenchido
    return {
        nome: registros[0][2],
        funcao: registros[0][3],
        itens: registros.map(r => ({
            data: r[5] instanceof Date ? Utilities.formatDate(r[5], "GMT-3", "dd/MM/yyyy") : r[5],
            ca: r[7],
            descricao: r[8]
        }))
    };
}

//FUNÇÃO QUE ALIMENTA MODAL RELATORIO DE EPI

function buscarDadosRelatorioEpi(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const sheetEpi = ss.getSheetByName("registroepi");
    const sheetColab = ss.getSheetByName("colaboradores");
    
    const matriculaProcurada = data.matricula.toString().trim();
    const dadosEpi = sheetEpi.getDataRange().getValues();
    
    let nomeEncontrado = "";
    let funcaoEncontrada = "Não cadastrada";
    const itensParaTabela = [];
    const fusoHorario = ss.getSpreadsheetTimeZone();

    //Loop na aba registroepi para buscar Itens, Nome e Matrícula
    for (let i = 1; i < dadosEpi.length; i++) {
      if (dadosEpi[i][6].toString().trim() === matriculaProcurada) {
        
        if (!nomeEncontrado) nomeEncontrado = dadosEpi[i][7];

        itensParaTabela.push({
          motivo: dadosEpi[i][10] || "---", 
          qtd: dadosEpi[i][4] || "1",       
          descricao: dadosEpi[i][3] || "",  
          ca: dadosEpi[i][8] || "",         
          dataEntrega: dadosEpi[i][1] instanceof Date ? Utilities.formatDate(dadosEpi[i][1], fusoHorario, "dd/MM/yyyy") : dadosEpi[i][1],
          dataDevolucao: dadosEpi[i][9] instanceof Date ? Utilities.formatDate(dadosEpi[i][9], fusoHorario, "dd/MM/yyyy") : (dadosEpi[i][9] || "---")
        });
      }
    }

    if (sheetColab) {
      const dadosColab = sheetColab.getDataRange().getValues();
      
      const registroColab = dadosColab.find(r => r[1].toString().trim() === matriculaProcurada);
      
      if (registroColab) {
        funcaoEncontrada = registroColab[3] || "Função não preenchida";
      }
    }

    // Validação caso não encontre nenhum item de EPI para a matrícula
    if (itensParaTabela.length === 0) {
      return { erro: "Nenhum registro de EPI encontrado para a matrícula: " + matriculaProcurada };
    }

    return {
      nome: nomeEncontrado || "Nome não encontrado",
      matricula: matriculaProcurada,
      funcao: funcaoEncontrada,
      itens: itensParaTabela
    };

  } catch (e) {
    return { erro: "Erro no servidor: " + e.message };
  }
}

//FUNÇÃO QUE BUSCA EPIS PEDENTES DE DEVOLUÇÃO 

function buscarEPIsPendentes(data) {
  try {
    const ss = SpreadsheetApp.openById("");
    const sheetEpi = ss.getSheetByName("registroepi");
    
    const matriculaProcurada = data.matricula.toString().trim();
    const dadosEpi = sheetEpi.getDataRange().getValues();
    
    let nomeEncontrado = "";
    const itensPendentes = [];

    for (let i = 1; i < dadosEpi.length; i++) {
      const matriculaLinha = dadosEpi[i][6].toString().trim(); 
      const dataDevolucao = dadosEpi[i][9]; 

      if (matriculaLinha === matriculaProcurada && (!dataDevolucao || dataDevolucao.toString().trim() === "")) {
        
        if (!nomeEncontrado) nomeEncontrado = dadosEpi[i][7]; 

        itensPendentes.push({
          id: i + 1,            // Número da linha para o update
          itemNome: dadosEpi[i][3], // Coluna D
          ca: dadosEpi[i][8]  // Coluna I
        });
      }
    }

    if (itensPendentes.length === 0) {
      return { erro: "Não há EPIs pendentes de devolução para esta matrícula." };
    }

    return {
      nome: nomeEncontrado || "Colaborador",
      itens: itensPendentes
    };

  } catch (e) {
    return { erro: "Erro ao buscar pendências: " + e.message };
  }
}