/**
 * ==================================================
 * ETAPA 1 — AUTENTICAÇÃO COM O BLING
 * ==================================================
 */

const PROP_CLIENT_ID = 'BLING_CLIENT_ID';
const PROP_CLIENT_SECRET = 'BLING_CLIENT_SECRET';

/**
 * Configura e retorna o serviço OAuth2 apontando para o Bling.
 */
function getBlingService() {
  const props = PropertiesService.getScriptProperties();

  return OAuth2.createService('bling')
    .setAuthorizationBaseUrl('https://www.bling.com.br/Api/v3/oauth/authorize')
    .setTokenUrl('https://api.bling.com.br/Api/v3/oauth/token')
    .setClientId(props.getProperty(PROP_CLIENT_ID))
    .setClientSecret(props.getProperty(PROP_CLIENT_SECRET))
    .setCallbackFunction('authCallback')
    .setPropertyStore(props)
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(
        props.getProperty(PROP_CLIENT_ID) + ':' + props.getProperty(PROP_CLIENT_SECRET)
      )
    });
}

/**
 * Rode manualmente para gerar o link de autorização (aparece no Registro de execução).
 */
function iniciarAutorizacao() {
  const service = getBlingService();

  if (service.hasAccess()) {
    Logger.log('Já estamos autenticados com o Bling. Nada a fazer.');
  } else {
    Logger.log('Abra este link no navegador para autorizar o acesso:');
    Logger.log(service.getAuthorizationUrl());
  }
}

/**
 * Chamada automaticamente pelo Bling após a autorização. Não rode manualmente.
 */
function authCallback(request) {
  const service = getBlingService();
  const isAuthorized = service.handleCallback(request);

  return HtmlService.createHtmlOutput(
    isAuthorized
      ? 'Autorização concluída com sucesso! Pode fechar esta aba.'
      : 'Falha na autorização. Volte ao Apps Script e tente novamente.'
  );
}

/**
 * Confirma que a conexão com o Bling está funcionando.
 */
function testarConexao() {
  const service = getBlingService();

  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }

  Logger.log('Conectado com sucesso ao Bling!');
  Logger.log('Token (início): ' + service.getAccessToken().substring(0, 10) + '...');
}

/**
 * Desconecta o script do Bling, caso precise reautorizar do zero.
 */
function resetarAutorizacao() {
  getBlingService().reset();
  Logger.log('Autorização removida. Rode "iniciarAutorizacao" para reconectar.');
}

/**
 * Serve o link de autorização OU o painel web, dependendo do estado da autorização.
 * Para usar: Implantar > Nova implantação > Tipo "Aplicativo da Web" >
 * Executar como "Eu" > Quem pode acessar (sua escolha) > Implantar.
 */
function doGet() {
  const service = getBlingService();

  if (!service.hasAccess()) {
    const authorizationUrl = service.getAuthorizationUrl();
    return HtmlService.createHtmlOutput(
      '<p>Clique no botão abaixo para autorizar o acesso ao Bling:</p>' +
      '<a href="' + authorizationUrl + '" target="_blank" ' +
      'style="display:inline-block;padding:12px 20px;background:#0b93f6;' +
      'color:white;text-decoration:none;border-radius:6px;font-family:sans-serif;">' +
      'Autorizar acesso ao Bling</a>'
    );
  }

  return HtmlService.createHtmlOutput(gerarPaginaDashboard())
    .setTitle('Painel de Sincronizações')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ==================================================
 * ETAPA 3 — GERAÇÃO DO RELATÓRIO DIÁRIO
 * ==================================================
 */

const SPREADSHEET_ID = 'COLE_AQUI_O_ID_DA_SUA_PLANILHA';
const ABA_REGISTRO = 'Registro';
const ABA_RELATORIO = 'Relatorio';
const ABA_GRAFICOS = 'Graficos';
const DIAS_SEMANA = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'];

const WORK_START_HOUR = 8;
const WORK_END_HOUR = 17;
const HORAS_TRABALHO = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

const META_DIARIA = 25;
const META_SEMANAL = 125;
const META_MENSAL = 500;

/**
 * Garante que qualquer horário fora do expediente (ex: sincronização feita
 * antes das 8h ou depois das 17h) caia no balde mais próximo do gráfico de
 * hora, em vez de ser descartado silenciosamente.
 */
function clampHoraTrabalho(hora) {
  if (hora < WORK_START_HOUR) return WORK_START_HOUR;
  if (hora > WORK_END_HOUR) return WORK_END_HOUR;
  return hora;
}

/**
 * Gatilho simples: toda vez que você edita a aba "Registro", grava o
 * horário exato na coluna B, ao lado do SKU colado na coluna A.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== ABA_REGISTRO) return;
  if (e.range.getColumn() !== 1) return; // só reage à coluna A (SKU)

  const linhaInicial = e.range.getRow();
  const numLinhas = e.range.getNumRows();

  for (let i = 0; i < numLinhas; i++) {
    const linha = linhaInicial + i;
    if (linha === 1) continue; // pula cabeçalho
    const sku = sheet.getRange(linha, 1).getValue();
    if (sku) {
      sheet.getRange(linha, 2).setValue(new Date());
    }
  }
}

/**
 * Função principal. Lê os SKUs colados na aba "Registro" (com o horário
 * gravado automaticamente pelo onEdit), descobre o produto-pai de cada
 * um, remove duplicados, escreve o resultado na aba "Relatorio" (com uma
 * linha de resumo no final do dia, contendo a data), atualiza o painel
 * de gráficos e limpa a aba "Registro" para o dia seguinte.
 */
function gerarRelatorioDiario() {
  const lock = LockService.getScriptLock();
  const conseguiuLock = lock.tryLock(10000);
  if (!conseguiuLock) {
    Logger.log('Já existe uma execução em andamento. Abortando para evitar conflito.');
    return;
  }

  try {
    const service = getBlingService();
    if (!service.hasAccess()) {
      Logger.log('Não autorizado. Rode "iniciarAutorizacao" primeiro.');
      return;
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaRegistro = ss.getSheetByName(ABA_REGISTRO);
    const abaRelatorio = ss.getSheetByName(ABA_RELATORIO);

    const ultimaLinha = abaRegistro.getLastRow();
    if (ultimaLinha < 2) {
      Logger.log('Nenhum SKU registrado para processar hoje.');
      return;
    }

    const dados = abaRegistro.getRange(2, 1, ultimaLinha - 1, 2).getValues(); // A: SKU, B: horário
    const entradas = dados
      .map(function (linha) { return { sku: String(linha[0]).trim(), timestamp: linha[1] }; })
      .filter(function (item) { return item.sku.length > 0; });

    if (entradas.length === 0) {
      Logger.log('Nenhum SKU válido encontrado na aba Registro.');
      return;
    }

    const paisEncontrados = {}; // { codigoPai: nomePai }
    const horaPorProduto = {}; // { codigoPai: hora }
    const erros = [];

    entradas.forEach(function (item) {
      try {
        const pai = buscarProdutoPai(service, item.sku);
        if (pai) {
          if (!paisEncontrados[pai.codigo]) {
            paisEncontrados[pai.codigo] = pai.nome;
            const hora = item.timestamp instanceof Date ? item.timestamp.getHours() : new Date().getHours();
            horaPorProduto[pai.codigo] = clampHoraTrabalho(hora);
          }
        } else {
          erros.push(item.sku + ' -> não encontrado no Bling');
        }
      } catch (e) {
        erros.push(item.sku + ' -> erro: ' + e.message);
      }
      Utilities.sleep(400); // respeita o limite de requisições por segundo do Bling
    });

    const codigosPais = Object.keys(paisEncontrados);
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // Linhas de produto: SEM data (a data fica só na linha de resumo, abaixo)
    codigosPais.forEach(function (codigo) {
      abaRelatorio.appendRow(['', codigo, paisEncontrados[codigo]]);
    });

    abaRelatorio.appendRow([
      'RESUMO ' + hoje, '', '',
      'Total: ' + codigosPais.length + ' produto(s) sincronizado(s)'
    ]);

    const contagemPorHora = {};
    codigosPais.forEach(function (codigo) {
      const hora = horaPorProduto[codigo];
      contagemPorHora[hora] = (contagemPorHora[hora] || 0) + 1;
    });

    atualizarDadosGraficoHoje(codigosPais.length, contagemPorHora);

    if (erros.length > 0) {
      Logger.log('Atenção — alguns SKUs não foram processados corretamente:');
      Logger.log(erros.join('\n'));
    }

    // Limpa a aba Registro (SKU e horário), mantém o cabeçalho da linha 1
    abaRegistro.getRange(2, 1, ultimaLinha - 1, 2).clearContent();

    Logger.log('Relatório gerado com sucesso! ' + codigosPais.length + ' produto(s) único(s) registrados.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dado um SKU (pode ser do produto-pai OU de uma variação), retorna
 * { codigo, nome } do produto-pai correspondente.
 * Retorna null se o SKU não for encontrado no Bling.
 */
function buscarProdutoPai(service, sku) {
  const urlBusca = 'https://api.bling.com.br/Api/v3/produtos?codigo=' + encodeURIComponent(sku);
  const respBusca = fetchComRetry(urlBusca, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  if (respBusca.getResponseCode() !== 200) {
    throw new Error('falha na busca por código (status ' + respBusca.getResponseCode() + ')');
  }

  const dadosBusca = JSON.parse(respBusca.getContentText());
  if (!dadosBusca.data || dadosBusca.data.length === 0) {
    return null;
  }

  const id = dadosBusca.data[0].id;
  Utilities.sleep(400);

  const urlDetalhe = 'https://api.bling.com.br/Api/v3/produtos/' + id;
  const respDetalhe = fetchComRetry(urlDetalhe, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  if (respDetalhe.getResponseCode() !== 200) {
    throw new Error('falha ao buscar detalhe (status ' + respDetalhe.getResponseCode() + ')');
  }

  const detalhe = JSON.parse(respDetalhe.getContentText()).data;

  // Se o produto tem "variacao.produtoPai", ele É uma variação -> busca o pai de verdade
  if (detalhe.variacao && detalhe.variacao.produtoPai && detalhe.variacao.produtoPai.id) {
    Utilities.sleep(400);
    const urlPai = 'https://api.bling.com.br/Api/v3/produtos/' + detalhe.variacao.produtoPai.id;
    const respPai = fetchComRetry(urlPai, {
      headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
      muteHttpExceptions: true
    });

    if (respPai.getResponseCode() !== 200) {
      throw new Error('falha ao buscar produto pai (status ' + respPai.getResponseCode() + ')');
    }

    const pai = JSON.parse(respPai.getContentText()).data;
    return { codigo: pai.codigo, nome: pai.nome };
  }

  // Já é o produto pai (ou um produto simples, sem variação)
  return { codigo: detalhe.codigo, nome: detalhe.nome };
}

/**
 * TESTE MANUAL: roda o relatório imediatamente, sem esperar o gatilho das 18h.
 * Útil para testar antes de agendar, ou para gerar o relatório sob demanda.
 */
function testarGeracaoManual() {
  gerarRelatorioDiario();
}

/**
 * ==================================================
 * INVESTIGAÇÃO — PROJETO 2 (Alerta de Estoque)
 * ==================================================
 * Testa o endpoint /estoques do Bling para descobrir se ele devolve
 * o HISTÓRICO de movimentações (entradas/saídas com data) ou só o
 * saldo atual consolidado.
 */
function inspecionarEndpointEstoques() {
  const service = getBlingService();
  if (!service.hasAccess()) {
    Logger.log('Ainda não autorizado. Rode "iniciarAutorizacao" primeiro.');
    return;
  }
  // Usa o código do parafuso "SKU_EXEMPLO" que apareceu nos seus prints
  // (o mesmo componente de vários produtos compostos)
  const codigoTeste = 'CPF0-EP1';

  // Passo 1: descobrir o ID a partir do código
  const urlBusca = 'https://api.bling.com.br/Api/v3/produtos?codigo=' + encodeURIComponent(codigoTeste);
  const respBusca = UrlFetchApp.fetch(urlBusca, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  Logger.log('Status da busca: ' + respBusca.getResponseCode());
  const dadosBusca = JSON.parse(respBusca.getContentText());

  if (!dadosBusca.data || dadosBusca.data.length === 0) {
    Logger.log('Produto não encontrado com esse código: ' + codigoTeste);
    return;
  }

  const idProduto = dadosBusca.data[0].id;
  Logger.log('ID encontrado: ' + idProduto);
  Utilities.sleep(400);

  // Passo 2: testar o endpoint de estoques com esse ID
  const urlEstoque = 'https://api.bling.com.br/Api/v3/estoques/saldos?idsProdutos[]=' + idProduto;
  const respEstoque = UrlFetchApp.fetch(urlEstoque, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  Logger.log('Status do estoque: ' + respEstoque.getResponseCode());
  Logger.log('Resposta completa do endpoint /estoques/saldos:');
  Logger.log(respEstoque.getContentText());
}

/**
 * Faz um fetch com retry automático em caso de erro 429 (limite de requisições).
 * Espera de forma crescente entre tentativas (backoff exponencial: 1s, 2s, 4s, 8s, 16s).
 */
function fetchComRetry(url, options, maxTentativas) {
  maxTentativas = maxTentativas || 5;
  let tentativa = 0;
  let espera = 1000;

  while (tentativa < maxTentativas) {
    const resp = UrlFetchApp.fetch(url, options);
    const status = resp.getResponseCode();

    if (status !== 429) {
      return resp;
    }

    tentativa++;
    Logger.log('Status 429 recebido, tentativa ' + tentativa + '/' + maxTentativas + '. Aguardando ' + espera + 'ms...');
    Utilities.sleep(espera);
    espera = espera * 2;
  }

  return UrlFetchApp.fetch(url, options);
}

/**
 * ==================================================
 * ABA DE GRÁFICOS (semanal + mensal + hora + meta)
 * ==================================================
 */

/**
 * Rode esta função UMA VEZ para criar e configurar a aba "Graficos" com o
 * painel completo (semanal, mensal, por hora e meta). Se a aba já existir,
 * ela é limpa e recriada do zero — ATENÇÃO: isso zera os totais acumulados.
 */
function inicializarAbaGraficos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_GRAFICOS);
  }
  aba.clear();
  aba.getCharts().forEach(function (chart) { aba.removeChart(chart); });

  aba.getRange('A1').setValue('Gráfico Semanal (dados)');
  aba.getRange('B2:H2').setValues([DIAS_SEMANA]);
  aba.getRange('A3').setValue('Semana Anterior');
  aba.getRange('A4').setValue('Semana Atual');
  aba.getRange('B3:H4').setValue(0);

  aba.getRange('A6').setValue('Gráfico Mensal (dados)');
  aba.getRange('B7:E7').setValues([['Semana 1', 'Semana 2', 'Semana 3', 'Semana 4']]);
  aba.getRange('A8').setValue('Total');
  aba.getRange('B8:E8').setValue(0);

  aba.getRange('A10').setValue('Semana atual do ciclo (1-4)');
  aba.getRange('B10').setValue(1);
  aba.getRange('A11').setValue('Última atualização');

  aba.getRange('A13').setValue('Anúncios por Hora (hoje)');
  const rotulosHora = HORAS_TRABALHO.map(function (h) { return h + 'h'; });
  aba.getRange(14, 2, 1, rotulosHora.length).setValues([rotulosHora]);
  aba.getRange('A15').setValue('Qtd.');
  aba.getRange(15, 2, 1, rotulosHora.length).setValue(0);

  aba.getRange('A17').setValue('Anúncios Hoje');
  aba.getRange('B17').setValue(0);
  aba.getRange('A18').setValue('Anúncios Essa Semana');
  aba.getRange('B18').setValue(0);
  aba.getRange('A19').setValue('% Meta Mensal (' + META_MENSAL + ')');
  aba.getRange('B19').setValue(0);

  aba.getRange('A21').setValue('Data do gráfico de hora');
  aba.getRange('B21').setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'));

  aba.getRange('B17:B19').setFontSize(28).setFontWeight('bold').setHorizontalAlignment('center');
  aba.getRange('A17:A19').setFontWeight('bold');
  aba.setColumnWidth(1, 200);

  const graficoSemanal = aba.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(aba.getRange('A2:H4'))
    .setTransposeRowsAndColumns(true)
    .setOption('title', 'Sincronizações por dia da semana')
    .setOption('legend', { position: 'top' })
    .setPosition(2, 10, 0, 0)
    .build();
  aba.insertChart(graficoSemanal);

  const graficoMensal = aba.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(aba.getRange('A7:E8'))
    .setTransposeRowsAndColumns(true)
    .setOption('title', 'Sincronizações por semana (ciclo de 4 semanas)')
    .setPosition(20, 10, 0, 0)
    .build();
  aba.insertChart(graficoMensal);

  const graficoHorario = aba.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(aba.getRange('B14:K15'))
    .setTransposeRowsAndColumns(true)
    .setOption('title', 'Anúncios por hora (hoje)')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#f4b400'])
    .setOption('vAxis', { minValue: 0 })
    .setPosition(38, 10, 0, 0)
    .build();
  aba.insertChart(graficoHorario);

  const graficoGauge = aba.newChart()
    .setChartType(Charts.ChartType.GAUGE)
    .addRange(aba.getRange('A19:B19'))
    .setOption('title', '% Meta Mensal')
    .setOption('min', 0)
    .setOption('max', 100)
    .setOption('greenFrom', 90).setOption('greenTo', 100)
    .setOption('yellowFrom', 50).setOption('yellowTo', 90)
    .setOption('redFrom', 0).setOption('redTo', 50)
    .setPosition(56, 10, 0, 0)
    .build();
  aba.insertChart(graficoGauge);

  Logger.log('Aba "Graficos" inicializada com sucesso (com painel de horário e meta mensal).');
}

/**
 * Índice do dia da semana com Segunda = 0 ... Domingo = 6.
 */
function diaSemanaIndex(data) {
  const diaJS = data.getDay(); // 0 = Domingo ... 6 = Sábado
  return diaJS === 0 ? 6 : diaJS - 1;
}

/**
 * Soma o total sincronizado hoje na "Semana Atual", no acumulado mensal,
 * na distribuição por hora (que acumula dentro do mesmo dia e reinicia
 * quando o dia muda), e recalcula os Scorecards (hoje/semana/meta%).
 * Chamada automaticamente por gerarRelatorioDiario().
 */
function atualizarDadosGraficoHoje(totalHoje, contagemPorHora) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) {
    Logger.log('Aba "Graficos" não encontrada. Rode "inicializarAbaGraficos" primeiro.');
    return;
  }

  const indiceDia = diaSemanaIndex(new Date());
  const colunaSemana = 2 + indiceDia; // B = 2
  const celulaSemanaAtual = aba.getRange(4, colunaSemana);
  celulaSemanaAtual.setValue(celulaSemanaAtual.getValue() + totalHoje);

  const semanaDoCiclo = aba.getRange('B10').getValue();
  const colunaMensal = 1 + semanaDoCiclo; // B = 2 para semana 1
  const celulaMensal = aba.getRange(8, colunaMensal);
  celulaMensal.setValue(celulaMensal.getValue() + totalHoje);

  aba.getRange('B11').setValue(new Date());

  // Distribuição por hora (reinicia quando o dia muda, acumula dentro do mesmo dia)
  const hojeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const celulaDataHora = aba.getRange('B21');
  const dataArmazenada = celulaDataHora.getValue();

  let contagensAtuais = aba.getRange(15, 2, 1, HORAS_TRABALHO.length).getValues()[0];
  if (dataArmazenada !== hojeStr) {
    contagensAtuais = HORAS_TRABALHO.map(function () { return 0; });
    celulaDataHora.setValue(hojeStr);
  }

  const novasContagens = HORAS_TRABALHO.map(function (hora, i) {
    const acrescimo = (contagemPorHora && contagemPorHora[hora]) || 0;
    return contagensAtuais[i] + acrescimo;
  });
  aba.getRange(15, 2, 1, novasContagens.length).setValues([novasContagens]);

  // Scorecards
  aba.getRange('B17').setValue(totalHoje);

  const totalSemanaAtual = aba.getRange('B4:H4').getValues()[0]
    .reduce(function (soma, v) { return soma + v; }, 0);
  aba.getRange('B18').setValue(totalSemanaAtual);

  const totalMensal = aba.getRange('B8:E8').getValues()[0]
    .reduce(function (soma, v) { return soma + v; }, 0);
  const percentMeta = Math.min(100, Math.round((totalMensal / META_MENSAL) * 100));
  aba.getRange('B19').setValue(percentMeta);
}

/**
 * Gatilho de Segunda-feira: move "Semana Atual" para "Semana Anterior"
 * (assim ainda aparece no gráfico nesta segunda), zera a semana atual,
 * e avança o contador do ciclo mensal (resetando o mensal na 5ª semana).
 */
function gatilhoSegundaFeira() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) {
    Logger.log('Aba "Graficos" não encontrada.');
    return;
  }

  const semanaAtual = aba.getRange('B4:H4').getValues();
  aba.getRange('B3:H3').setValues(semanaAtual);
  aba.getRange('B4:H4').setValue(0);

  let semanaDoCiclo = aba.getRange('B10').getValue();
  semanaDoCiclo++;

  if (semanaDoCiclo > 4) {
    aba.getRange('B8:E8').setValue(0);
    semanaDoCiclo = 1;
    Logger.log('Novo ciclo mensal iniciado (semana 1/4).');
  }

  aba.getRange('B10').setValue(semanaDoCiclo);
  Logger.log('Gatilho de segunda-feira executado. Semana ' + semanaDoCiclo + '/4 do ciclo.');
}

/**
 * Gatilho de Terça-feira: limpa "Semana Anterior", deixando só a semana
 * atual visível no gráfico.
 */
function gatilhoTercaFeira() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) {
    Logger.log('Aba "Graficos" não encontrada.');
    return;
  }

  aba.getRange('B3:H3').setValue(0);
  Logger.log('Gatilho de terça-feira executado. Semana anterior removida do gráfico.');
}

/**
 * Rode UMA VEZ para criar os gatilhos automáticos de segunda e terça-feira.
 */
function criarGatilhosGraficos() {
  ScriptApp.newTrigger('gatilhoSegundaFeira')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();

  ScriptApp.newTrigger('gatilhoTercaFeira')
    .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(6).create();

  Logger.log('Gatilhos de segunda e terça-feira criados com sucesso.');
}

/**
 * Remove os gatilhos criados por criarGatilhosGraficos(), caso precise recriar.
 */
function removerGatilhosGraficos() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const nome = trigger.getHandlerFunction();
    if (nome === 'gatilhoSegundaFeira' || nome === 'gatilhoTercaFeira') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('Gatilhos removidos.');
}

/**
 * ==================================================
 * PAINEL WEB (página HTML servida por doGet)
 * ==================================================
 */

/**
 * Lê a aba "Graficos" e devolve os números já prontos para o painel web,
 * em formato simples (objeto JSON). Chamada pelo cliente via google.script.run.
 */
function getDadosDashboard() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) {
    return { erro: 'Aba "Graficos" não encontrada. Rode inicializarAbaGraficos primeiro.' };
  }

  const hoje = aba.getRange('B17').getValue() || 0;
  const semana = aba.getRange('B18').getValue() || 0;
  const percentMeta = aba.getRange('B19').getValue() || 0;

  const semanasValores = aba.getRange('B8:E8').getValues()[0];
  const totalMensal = semanasValores.reduce(function (soma, v) { return soma + v; }, 0);

  const horasValores = aba.getRange(15, 2, 1, HORAS_TRABALHO.length).getValues()[0];
  const horasRotulos = HORAS_TRABALHO.map(function (h) { return h + 'h'; });

  const diasValores = aba.getRange('B4:H4').getValues()[0];

  return {
    hoje: hoje,
    metaDiaria: META_DIARIA,
    semana: semana,
    metaSemanal: META_SEMANAL,
    percentMeta: percentMeta,
    totalMensal: totalMensal,
    metaMensal: META_MENSAL,
    horas: { rotulos: horasRotulos, valores: horasValores },
    dias: { rotulos: DIAS_SEMANA, valores: diasValores },
    semanas: { rotulos: ['Semana 1', 'Semana 2', 'Semana 3', 'Semana 4'], valores: semanasValores },
    atualizadoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  };
}

/**
 * Monta a página HTML completa do painel (cartões, gauge e gráficos em SVG,
 * desenhados dinamicamente no navegador a partir dos dados de getDadosDashboard).
 */
function gerarPaginaDashboard() {
  return '' +
'<!DOCTYPE html>' +
'<html><head><base target="_top">' +
'<style>' +
'  body { font-family: Arial, sans-serif; background: #f4f4f2; margin: 0; padding: 24px; color: #1a1a1a; }' +
'  h1 { font-size: 18px; font-weight: 500; margin: 0 0 20px; }' +
'  .cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 12px; }' +
'  .card { border-radius: 12px; padding: 20px; }' +
'  .card-label { font-size: 13px; margin: 0 0 6px; opacity: 0.85; }' +
'  .card-number { font-size: 40px; font-weight: 700; margin: 0; line-height: 1; }' +
'  .card-sub { font-size: 12px; margin: 8px 0 0; opacity: 0.85; }' +
'  .teal { background: #5DCAA5; color: #04342C; }' +
'  .blue { background: #85B7EB; color: #042C53; }' +
'  .amber { background: #EF9F27; color: #412402; }' +
'  .panel { background: #ffffff; border-radius: 12px; padding: 20px; margin-bottom: 12px; box-shadow: 0 0 0 1px rgba(0,0,0,0.06); }' +
'  .panel-title { font-size: 13px; color: #555; margin: 0 0 12px; }' +
'  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }' +
'  .updated { font-size: 12px; color: #888; margin-top: 16px; text-align: right; }' +
'  .gauge-wrap { display: flex; flex-direction: column; align-items: center; }' +
'</style></head><body>' +
'<h1>Painel de Sincronizações — Click Presilhas</h1>' +
'<div class="cards">' +
'  <div class="card teal">' +
'    <p class="card-label">Anúncios hoje</p>' +
'    <p class="card-number" id="numHoje">--</p>' +
'    <p class="card-sub" id="metaHoje">Meta do dia: --</p>' +
'  </div>' +
'  <div class="card blue">' +
'    <p class="card-label">Anúncios essa semana</p>' +
'    <p class="card-number" id="numSemana">--</p>' +
'    <p class="card-sub" id="metaSemana">Meta da semana: --</p>' +
'  </div>' +
'  <div class="card amber gauge-wrap">' +
'    <p class="card-label" style="align-self:flex-start;">Meta mensal</p>' +
'    <svg viewBox="0 0 140 78" style="width:130px;height:auto;margin-top:2px;">' +
'      <path d="M 12 70 A 58 58 0 0 1 128 70" fill="none" stroke-width="12" stroke-linecap="round" stroke="rgba(0,0,0,0.15)"></path>' +
'      <path id="gaugeArc" d="M 12 70 A 58 58 0 0 1 128 70" fill="none" stroke-width="12" stroke-linecap="round" stroke="#412402"></path>' +
'      <text id="gaugeTexto" x="70" y="62" text-anchor="middle" font-size="26" font-weight="700" fill="#412402">--%</text>' +
'    </svg>' +
'    <p class="card-sub" id="metaMensal">-- de --</p>' +
'  </div>' +
'</div>' +
'<div class="panel">' +
'  <p class="panel-title">Anúncios por hora (hoje)</p>' +
'  <svg id="chartHoras" viewBox="0 0 620 160" style="width:100%;height:auto;"></svg>' +
'</div>' +
'<div class="row2">' +
'  <div class="panel">' +
'    <p class="panel-title">Sincronizações por dia da semana</p>' +
'    <svg id="chartDias" viewBox="0 0 280 140" style="width:100%;height:auto;"></svg>' +
'  </div>' +
'  <div class="panel">' +
'    <p class="panel-title">Sincronizações por semana (ciclo de 4)</p>' +
'    <svg id="chartSemanas" viewBox="0 0 280 140" style="width:100%;height:auto;"></svg>' +
'  </div>' +
'</div>' +
'<p class="updated" id="atualizadoEm"></p>' +
'<script>' +
'function svgNS(tag) { return document.createElementNS("http://www.w3.org/2000/svg", tag); }' +
'function limparSvg(svg) { while (svg.firstChild) svg.removeChild(svg.firstChild); }' +
'function desenharBarras(svgId, rotulos, valores, corBarra, corDestaque) {' +
'  var svg = document.getElementById(svgId);' +
'  limparSvg(svg);' +
'  var largura = 280, altura = 140, base = 110, topo = 20;' +
'  var maxVal = Math.max.apply(null, valores.concat([1]));' +
'  var n = valores.length;' +
'  var espaco = (largura - 40) / n;' +
'  var largBarra = Math.min(30, espaco * 0.6);' +
'  var linha = svgNS("line");' +
'  linha.setAttribute("x1", 20); linha.setAttribute("y1", base);' +
'  linha.setAttribute("x2", largura - 10); linha.setAttribute("y2", base);' +
'  linha.setAttribute("stroke", "#ddd"); linha.setAttribute("stroke-width", 1);' +
'  svg.appendChild(linha);' +
'  for (var i = 0; i < n; i++) {' +
'    var x = 25 + i * espaco + (espaco - largBarra) / 2;' +
'    var alturaBarra = valores[i] > 0 ? Math.max(4, (valores[i] / maxVal) * (base - topo)) : 2;' +
'    var y = base - alturaBarra;' +
'    var rect = svgNS("rect");' +
'    rect.setAttribute("x", x); rect.setAttribute("y", y);' +
'    rect.setAttribute("width", largBarra); rect.setAttribute("height", alturaBarra);' +
'    rect.setAttribute("rx", 3);' +
'    rect.setAttribute("fill", valores[i] === maxVal && valores[i] > 0 ? corDestaque : corBarra);' +
'    svg.appendChild(rect);' +
'    if (valores[i] > 0) {' +
'      var texto = svgNS("text");' +
'      texto.setAttribute("x", x + largBarra / 2); texto.setAttribute("y", y - 6);' +
'      texto.setAttribute("text-anchor", "middle"); texto.setAttribute("font-size", 11);' +
'      texto.setAttribute("fill", "#555");' +
'      texto.textContent = valores[i];' +
'      svg.appendChild(texto);' +
'    }' +
'    var rotulo = svgNS("text");' +
'    rotulo.setAttribute("x", x + largBarra / 2); rotulo.setAttribute("y", base + 16);' +
'    rotulo.setAttribute("text-anchor", "middle"); rotulo.setAttribute("font-size", 10);' +
'    rotulo.setAttribute("fill", "#999");' +
'    rotulo.textContent = rotulos[i];' +
'    svg.appendChild(rotulo);' +
'  }' +
'}' +
'function desenharArea(svgId, rotulos, valores) {' +
'  var svg = document.getElementById(svgId);' +
'  limparSvg(svg);' +
'  var largura = 620, altura = 160, base = 130, topo = 15, esq = 30, dir = 590;' +
'  var maxVal = Math.max.apply(null, valores.concat([1]));' +
'  var n = valores.length;' +
'  var passo = (dir - esq) / (n - 1 || 1);' +
'  var pontos = [];' +
'  for (var i = 0; i < n; i++) {' +
'    var x = esq + i * passo;' +
'    var y = base - (valores[i] / maxVal) * (base - topo);' +
'    pontos.push([x, y]);' +
'  }' +
'  var poligono = svgNS("polygon");' +
'  var ptsStr = pontos.map(function(p){ return p[0] + "," + p[1]; }).join(" ");' +
'  ptsStr = esq + "," + base + " " + ptsStr + " " + dir + "," + base;' +
'  poligono.setAttribute("points", ptsStr);' +
'  poligono.setAttribute("fill", "#FAEEDA");' +
'  svg.appendChild(poligono);' +
'  var linhaPts = pontos.map(function(p){ return p[0] + "," + p[1]; }).join(" ");' +
'  var polilinha = svgNS("polyline");' +
'  polilinha.setAttribute("points", linhaPts);' +
'  polilinha.setAttribute("fill", "none");' +
'  polilinha.setAttribute("stroke", "#BA7517");' +
'  polilinha.setAttribute("stroke-width", 2);' +
'  svg.appendChild(polilinha);' +
'  var eixo = svgNS("line");' +
'  eixo.setAttribute("x1", esq); eixo.setAttribute("y1", base);' +
'  eixo.setAttribute("x2", dir); eixo.setAttribute("y2", base);' +
'  eixo.setAttribute("stroke", "#ddd"); eixo.setAttribute("stroke-width", 1);' +
'  svg.appendChild(eixo);' +
'  for (var j = 0; j < n; j++) {' +
'    var rotulo = svgNS("text");' +
'    rotulo.setAttribute("x", pontos[j][0]); rotulo.setAttribute("y", base + 18);' +
'    rotulo.setAttribute("text-anchor", "middle"); rotulo.setAttribute("font-size", 11);' +
'    rotulo.setAttribute("fill", "#999");' +
'    rotulo.textContent = rotulos[j];' +
'    svg.appendChild(rotulo);' +
'  }' +
'}' +
'function corDoGauge(pct) {' +
'  if (pct >= 90) return "#3B6D11";' +
'  if (pct >= 50) return "#854F0B";' +
'  return "#791F1F";' +
'}' +
'function renderDashboard(dados) {' +
'  if (dados.erro) { document.body.innerHTML = "<p>" + dados.erro + "</p>"; return; }' +
'  document.getElementById("numHoje").textContent = dados.hoje;' +
'  document.getElementById("metaHoje").textContent = "Meta do dia: " + dados.metaDiaria;' +
'  document.getElementById("numSemana").textContent = dados.semana;' +
'  document.getElementById("metaSemana").textContent = "Meta da semana: " + dados.metaSemanal;' +
'  document.getElementById("metaMensal").textContent = dados.totalMensal + " de " + dados.metaMensal;' +
'  var pct = Math.min(100, dados.percentMeta);' +
'  document.getElementById("gaugeTexto").textContent = pct + "%";' +
'  var arco = document.getElementById("gaugeArc");' +
'  var comprimento = 182;' +
'  arco.setAttribute("stroke-dasharray", comprimento);' +
'  arco.setAttribute("stroke-dashoffset", comprimento - (comprimento * pct / 100));' +
'  arco.setAttribute("stroke", corDoGauge(pct));' +
'  document.getElementById("gaugeTexto").setAttribute("fill", corDoGauge(pct));' +
'  desenharArea("chartHoras", dados.horas.rotulos, dados.horas.valores);' +
'  desenharBarras("chartDias", dados.dias.rotulos, dados.dias.valores, "#B4B2A9", "#378ADD");' +
'  desenharBarras("chartSemanas", dados.semanas.rotulos, dados.semanas.valores, "#B4B2A9", "#378ADD");' +
'  document.getElementById("atualizadoEm").textContent = "Atualizado em " + dados.atualizadoEm;' +
'}' +
'google.script.run.withSuccessHandler(renderDashboard).getDadosDashboard();' +
'</script>' +
'</body></html>';
}
