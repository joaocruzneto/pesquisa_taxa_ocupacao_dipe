/**
 * @OnlyCurrentDoc
 * Script de Automação de Pesquisa de Mancha
 * Versão: 1.0.3
 * Data de última atualização: 13/05/2026
 * Responsável: [João Cruz Neto / Divisão de Pesquisas]
 */

/**
 * Função executada automaticamente ao abrir a planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Pesquisa de Mancha')
    .addItem('1. Gerar Base de Dados (BD)', 'passo1_criarBD')
    .addItem('2. Criar Coluna: Nº Viagem', 'passo2_criarNumViagem')
    .addItem('3. Criar Faixa Horária', 'passo3_criarFaixaHoraria')
    .addItem('4. Criar Coluna EMP_CARRO', 'passo4_criarEmpCarro')
    .addItem('5. Criar Coluna Int_Viagens', 'passo5_criarIntViagens')
    .addSeparator() 
    .addItem('6. Padronizar Frota (MOB/CNO)', 'passo6_padronizarFrota')
    .addItem('7. Criar Coluna EMP_COD (Frota)', 'passo7_criarEmpCodFrota')
    .addSeparator()
    .addItem('8. Buscar Tipo de Veículo', 'passo8_buscarTipoVeiculo')
    .addItem('9. Buscar Cap. Pass. Sentado', 'passo9_buscarCapPassSentado')
    .addItem('10. Buscar Cap. Pass. em Pé', 'passo10_buscarCapPassPe')
    .addItem('11. Buscar Cap. Pass. Total', 'passo11_buscarCapPassTotal')
    .addSeparator()
    .addItem('12. Calcular Contagem Sentados', 'passo12_calcularContagemSentados')
    .addItem('13. Calcular Contagem em Pé', 'passo13_calcularContagemPe')
    .addItem('14. Calcular Contagem Total', 'passo14_calcularContagemTotal')
    .addSeparator()
    .addItem('15. Calcular Taxa Ocup. Sentado', 'passo15_calcularTaxaOcupSentado')
    .addItem('16. Calcular Taxa Ocup. em Pé', 'passo16_calcularTaxaOcupPe')
    .addItem('17. Calcular Taxa Ocup. Total', 'passo17_calcularTaxaOcupTotal')
    .addSeparator()
    .addItem('18. Formatar Aba BD', 'passo18_formatarAbaBD')
    .addItem('19. Criar Aba Tabelas Análises', 'passo19_criarAbaAnalises')
    .addSeparator()
    .addItem('20. Análise: Viagens por Faixa/Local', 'passo20_analiseViagensLocal')
    .addItem('21. Análise: Ocupação Total (Hierarquia)', 'passo21_analiseTaxaOcupacao')
    .addItem('22. Análise por Local, Sentido e PED', 'passo22_analisePorLocalSentidoPED')
    .addToUi();
}

/**
 * Passo 1: Verifica existência da aba "TABULAÇÃO", faz a cópia e a renomeia.
 */
function passo1_criarBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaOrigem = ss.getSheetByName('TABULAÇÃO');
  if (!abaOrigem) { ui.alert('Erro', 'A aba TABULAÇÃO não foi encontrada.', ui.ButtonSet.OK); return; }
  const abaDestino = ss.getSheetByName('BD');
  if (abaDestino) {
    if (ui.alert('Aviso', 'A aba BD já existe. Deseja substituí-la?', ui.ButtonSet.YES_NO) == ui.Button.YES) ss.deleteSheet(abaDestino);
    else return; 
  }
  const novaAba = abaOrigem.copyTo(ss); novaAba.setName('BD'); novaAba.activate();
  ss.toast('A aba BD foi gerada.', 'Passo 1 Concluído', 5);
}

/**
 * Passo 2: Criar Coluna: Nº Viagem
 * Lógica corrigida: Reinicia o contador se LINHA ou SENTIDO mudarem.
 * Processamento em Array para garantir precisão e performance.
 */
function passo2_criarNumViagem() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const ultimaLinha = sheet.getLastRow();
  const ultimaColuna = sheet.getLastColumn();
  
  if (ultimaLinha <= 1) return;

  const rangeCabecalho = sheet.getRange(1, 1, 1, ultimaColuna);
  const cabecalhos = rangeCabecalho.getValues()[0];
  
  let colLinhaIndex = -1;
  let colSentidoIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'LINHA') colLinhaIndex = i + 1;
    if (tit === 'SENTIDO') colSentidoIndex = i + 1;
  }

  if (colLinhaIndex === -1 || colSentidoIndex === -1) {
    ui.alert('Erro', 'As colunas "LINHA" e "SENTIDO" precisam existir para esta contagem.', ui.ButtonSet.OK);
    return;
  }

  // 1. Inserimos a coluna (as referências de índice mudam aqui)
  sheet.insertColumnBefore(colLinhaIndex);
  sheet.getRange(1, colLinhaIndex).setValue('Nº Viagem');

  // 2. Agora pegamos os dados atualizados das colunas de controle
  // IMPORTANTE: Como inserimos uma coluna ANTES de LINHA, o índice de LINHA continua o mesmo
  // mas o de SENTIDO pode ter mudado se ele estava à direita de LINHA.
  
  // Recalculamos os índices reais após a inserção para segurança total
  const novosCabecalhos = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxLinha = novosCabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA') + 1;
  const idxSentido = novosCabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO') + 1;

  // Pegamos todos os valores de uma vez para processar em memória
  const dadosLinha = sheet.getRange(2, idxLinha, ultimaLinha - 1, 1).getValues();
  const dadosSentido = sheet.getRange(2, idxSentido, ultimaLinha - 1, 1).getValues();
  
  const resultados = [];
  let contador = 1;

  for (let i = 0; i < dadosLinha.length; i++) {
    if (i > 0) {
      let linhaMudou = dadosLinha[i][0].toString().trim() !== dadosLinha[i-1][0].toString().trim();
      let sentidoMudou = dadosSentido[i][0].toString().trim() !== dadosSentido[i-1][0].toString().trim();
      
      if (linhaMudou || sentidoMudou) {
        contador = 1; // Reinicia se qualquer um mudar
      }
    }
    resultados.push([contador]);
    contador++;
  }

  // 3. Escrevemos todos os resultados de uma só vez na coluna Nº Viagem
  sheet.getRange(2, colLinhaIndex, ultimaLinha - 1, 1).setValues(resultados);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Contador reiniciado por Linha/Sentido com sucesso.', 'Passo 2 Concluído', 5);
}

/**
 * Passo 3: Localiza HORA, insere FAIXA-HORÁRIA e preenche.
 */
function passo3_criarFaixaHoraria() {
  const sheet = SpreadsheetApp.getActiveSheet(); const ui = SpreadsheetApp.getUi();
  const ultimaLinha = sheet.getLastRow(); const ultimaColuna = sheet.getLastColumn(); if (ultimaLinha <= 1) return;

  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  let colunaHoraIndex = -1;
  for (let i = 0; i < cabecalhos.length; i++) { if (cabecalhos[i].toString().toUpperCase().trim() === 'HORA') { colunaHoraIndex = i + 1; break; } }
  
  if (colunaHoraIndex === -1) {
    const p = ui.prompt('Não encontrada', 'Digite a LETRA da coluna HORA:', ui.ButtonSet.OK_CANCEL);
    if (p.getSelectedButton() == ui.Button.OK) { colunaHoraIndex = converterLetraParaNumero(p.getResponseText().toUpperCase().trim()); if (colunaHoraIndex < 1) return; } else return; 
  }

  const colunaDestinoIndex = colunaHoraIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= sheet.getMaxColumns()) vizinho = sheet.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  if (vizinho === 'FAIXA-HORÁRIA') { if (ui.alert('Aviso', 'A coluna "FAIXA-HORÁRIA" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return; } 
  else { sheet.insertColumnAfter(colunaHoraIndex); sheet.getRange(1, colunaDestinoIndex).setValue('FAIXA-HORÁRIA'); }

  const valoresHoras = sheet.getRange(2, colunaHoraIndex, ultimaLinha - 1, 1).getDisplayValues(); 
  const valoresFaixa = [];
  for (let i = 0; i < valoresHoras.length; i++) {
    let valor = valoresHoras[i][0].trim(); let res = "";
    if (valor !== "") {
      let partes = valor.split(':');
      if (partes.length > 0) { let h = partes[0].trim().padStart(2, '0'); if (!isNaN(parseInt(h))) res = h + ":00 - " + h + ":59"; else res = valor; }
    }
    valoresFaixa.push([res]);
  }
  sheet.getRange(2, colunaDestinoIndex, ultimaLinha - 1, 1).setValues(valoresFaixa);
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 3 Concluído', 'Sucesso', 5);
}

/**
 * Passo 4: Cria EMP_CARRO concatenando EMPRESA e CARRO (Aba BD)
 */
function passo4_criarEmpCarro() {
  const sheet = SpreadsheetApp.getActiveSheet(); const ui = SpreadsheetApp.getUi();
  const ultimaLinha = sheet.getLastRow(); const ultimaColuna = sheet.getLastColumn(); if (ultimaLinha <= 1) return;

  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  let colEmpresaIndex = -1, colCarroIndex = -1;
  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'EMPRESA') colEmpresaIndex = i + 1;
    if (tit === 'CARRO') colCarroIndex = i + 1;
  }
  if (colEmpresaIndex === -1 || colCarroIndex === -1) { ui.alert('Erro', 'Colunas EMPRESA e/ou CARRO não encontradas.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colCarroIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= sheet.getMaxColumns()) vizinho = sheet.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  if (vizinho === 'EMP_CARRO') { if (ui.alert('Aviso', 'A coluna EMP_CARRO já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return; } 
  else { sheet.insertColumnAfter(colCarroIndex); sheet.getRange(1, colunaDestinoIndex).setValue('EMP_CARRO'); }

  const valoresEmpresa = sheet.getRange(2, colEmpresaIndex, ultimaLinha - 1, 1).getDisplayValues();
  const valoresCarro = sheet.getRange(2, colCarroIndex, ultimaLinha - 1, 1).getDisplayValues();
  const res = [];
  for (let i = 0; i < valoresEmpresa.length; i++) {
    let emp = valoresEmpresa[i][0].trim(); let car = valoresCarro[i][0].trim();
    if (emp !== "" || car !== "") res.push([emp + " - " + car]); else res.push([""]); 
  }
  sheet.getRange(2, colunaDestinoIndex, ultimaLinha - 1, 1).setValues(res);
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 4 Concluído', 'Sucesso', 5);
}

/**
 * Passo 5: Cria Int_Viagens APÓS FAIXA-HORÁRIA
 */
function passo5_criarIntViagens() {
  const sheet = SpreadsheetApp.getActiveSheet(); const ui = SpreadsheetApp.getUi();
  const ultimaLinha = sheet.getLastRow(); const ultimaColuna = sheet.getLastColumn(); if (ultimaLinha <= 1) return;

  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  let colHoraIndex = -1, colNumViagemIndex = -1, colFaixaHorariaIndex = -1;
  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'HORA') colHoraIndex = i + 1;
    if (tit === 'Nº DA VIAGEM') colNumViagemIndex = i + 1;
    if (tit === 'FAIXA-HORÁRIA') colFaixaHorariaIndex = i + 1;
  }
  if (colHoraIndex === -1 || colNumViagemIndex === -1 || colFaixaHorariaIndex === -1) { ui.alert('Erro', 'Verifique colunas HORA, Nº DA VIAGEM e FAIXA-HORÁRIA.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colFaixaHorariaIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= sheet.getMaxColumns()) vizinho = sheet.getRange(1, colunaDestinoIndex).getValue().toString().trim();
  if (vizinho.toUpperCase() === 'INT_VIAGENS') { if (ui.alert('Aviso', 'A coluna Int_Viagens já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return; } 
  else { sheet.insertColumnAfter(colFaixaHorariaIndex); sheet.getRange(1, colunaDestinoIndex).setValue('Int_Viagens'); }

  const valHora = sheet.getRange(2, colHoraIndex, ultimaLinha - 1, 1).getDisplayValues();
  const valNum = sheet.getRange(2, colNumViagemIndex, ultimaLinha - 1, 1).getDisplayValues();
  const res = [];
  for (let i = 0; i < valNum.length; i++) {
    let num = valNum[i][0].toString().trim(); let horaAt = valHora[i][0].toString().trim();
    if (num === "1") res.push(["Primeira Viagem"]);
    else if (num !== "" && i > 0) {
      let horaAnt = valHora[i-1][0].toString().trim();
      let minAt = auxiliar_converterParaMinutos(horaAt); let minAnt = auxiliar_converterParaMinutos(horaAnt);
      if (minAt !== null && minAnt !== null) {
        let diff = minAt - minAnt; if (diff < 0) diff += 24 * 60; 
        let h = Math.floor(diff / 60).toString().padStart(2, '0'); let m = (diff % 60).toString().padStart(2, '0');
        res.push([h + ":" + m]);
      } else res.push([""]); 
    } else res.push([""]); 
  }
  sheet.getRange(2, colunaDestinoIndex, ultimaLinha - 1, 1).setValues(res);
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 5 Concluído', 'Sucesso', 5);
}

/**
 * Passo 6: Padroniza os códigos da frota na aba FROTA_ATUALIZADA
 */
function passo6_padronizarFrota() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!sheet) { ui.alert('Erro', 'A aba FROTA_ATUALIZADA não foi encontrada.', ui.ButtonSet.OK); return; }

  const ultimaLinha = sheet.getLastRow(); const ultimaColuna = sheet.getLastColumn(); if (ultimaLinha <= 1) return;

  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  let colEmpresaIndex = -1, colCodIndex = -1;
  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'EMPRESA') colEmpresaIndex = i + 1;
    if (tit === 'COD') colCodIndex = i + 1;
  }
  if (colEmpresaIndex === -1 || colCodIndex === -1) { ui.alert('Erro', 'Colunas EMPRESA ou COD não encontradas.', ui.ButtonSet.OK); return; }

  const valoresEmpresa = sheet.getRange(2, colEmpresaIndex, ultimaLinha - 1, 1).getDisplayValues();
  const valoresCod = sheet.getRange(2, colCodIndex, ultimaLinha - 1, 1).getDisplayValues();
  const novosCodigos = []; let alteracoes = 0;

  for (let i = 0; i < valoresEmpresa.length; i++) {
    let emp = valoresEmpresa[i][0].toString().toUpperCase().trim(); let cod = valoresCod[i][0].toString().trim();
    if (emp === 'MOB' && cod !== "") { if (!(cod.length === 4 && cod.startsWith('2'))) { cod = "2" + cod.padStart(3, '0'); alteracoes++; } } 
    else if (emp === 'CNO' && cod !== "") { if (!(cod.length === 4 && cod.startsWith('1'))) { cod = "1" + cod.padStart(3, '0'); alteracoes++; } }
    novosCodigos.push([cod]);
  }
  sheet.getRange(2, colCodIndex, ultimaLinha - 1, 1).setValues(novosCodigos);
  if (alteracoes > 0) ss.toast(alteracoes + ' registros corrigidos.', 'Passo 6 Concluído', 5);
  else ss.toast('Os códigos já estavam no padrão.', 'Passo 6 Concluído', 4);
}

/**
 * Passo 7: Cria a coluna EMP_COD na aba FROTA_ATUALIZADA
 */
function passo7_criarEmpCodFrota() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!sheet) { ui.alert('Erro', 'A aba FROTA_ATUALIZADA não encontrada.', ui.ButtonSet.OK); return; }

  const ultimaLinha = sheet.getLastRow(); const ultimaColuna = sheet.getLastColumn(); if (ultimaLinha <= 1) return;

  const cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  let colEmpresaIndex = -1, colCodIndex = -1;
  for (let i = 0; i < cabecalhos.length; i++) {
    let titulo = cabecalhos[i].toString().toUpperCase().trim();
    if (titulo === 'EMPRESA') colEmpresaIndex = i + 1;
    if (titulo === 'COD') colCodIndex = i + 1;
  }
  if (colEmpresaIndex === -1 || colCodIndex === -1) { ui.alert('Erro', 'Colunas EMPRESA ou COD não encontradas.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colCodIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= sheet.getMaxColumns()) vizinho = sheet.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  if (vizinho === 'EMP_COD') { if (ui.alert('Aviso', 'A coluna EMP_COD já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return; } 
  else { sheet.insertColumnAfter(colCodIndex); sheet.getRange(1, colunaDestinoIndex).setValue('EMP_COD'); }

  const valoresEmpresa = sheet.getRange(2, colEmpresaIndex, ultimaLinha - 1, 1).getDisplayValues();
  const valoresCod = sheet.getRange(2, colCodIndex, ultimaLinha - 1, 1).getDisplayValues();
  const res = [];
  for (let i = 0; i < valoresEmpresa.length; i++) {
    let emp = valoresEmpresa[i][0].trim(); let cod = valoresCod[i][0].trim();
    if (emp !== "" || cod !== "") res.push([emp + " - " + cod]); else res.push([""]); 
  }
  sheet.getRange(2, colunaDestinoIndex, ultimaLinha - 1, 1).setValues(res);
  ss.toast('A coluna EMP_COD foi criada com sucesso.', 'Passo 7 Concluído', 5);
}

/**
 * Passo 8: Buscar Tipo de Veículo
 */
function passo8_buscarTipoVeiculo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD"); const abaFrota = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!abaBD || !abaFrota) { ui.alert('Erro', 'Abas "BD" ou "FROTA_ATUALIZADA" não encontradas.', ui.ButtonSet.OK); return; }

  const ultLinhaFrota = abaFrota.getLastRow(); const ultColFrota = abaFrota.getLastColumn(); if (ultLinhaFrota <= 1) return;

  const cabecalhosFrota = abaFrota.getRange(1, 1, 1, ultColFrota).getValues()[0];
  let colEmpCodIndex = -1, colTipoVeiculoIndex = -1;
  for (let i = 0; i < cabecalhosFrota.length; i++) {
    let tit = cabecalhosFrota[i].toString().toUpperCase().trim();
    if (tit === 'EMP_COD') colEmpCodIndex = i + 1;
    if (tit === 'TIPO DO VEICULO' || tit === 'TIPO DO VEÍCULO') colTipoVeiculoIndex = i + 1;
  }
  if (colEmpCodIndex === -1 || colTipoVeiculoIndex === -1) { ui.alert('Erro', 'Colunas EMP_COD ou TIPO DO VEICULO não encontradas.', ui.ButtonSet.OK); return; }

  const valoresEmpCod = abaFrota.getRange(2, colEmpCodIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  const valoresTipo = abaFrota.getRange(2, colTipoVeiculoIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  
  const dicionarioFrota = {};
  for (let i = 0; i < valoresEmpCod.length; i++) {
    let chave = valoresEmpCod[i][0].toString().trim(); let valor = valoresTipo[i][0].toString().trim();
    if (chave !== "") dicionarioFrota[chave] = valor;
  }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhosBD = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colIntViagensIndex = -1, colEmpCarroIndex = -1;
  for (let i = 0; i < cabecalhosBD.length; i++) {
    let tit = cabecalhosBD[i].toString().toUpperCase().trim();
    if (tit === 'INT_VIAGENS') colIntViagensIndex = i + 1;
    if (tit === 'EMP_CARRO') colEmpCarroIndex = i + 1;
  }
  if (colIntViagensIndex === -1 || colEmpCarroIndex === -1) { ui.alert('Erro', 'Colunas Int_Viagens ou EMP_CARRO não encontradas.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colIntViagensIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= abaBD.getMaxColumns()) vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  
  if (vizinho === 'TIPO_VEÍCULO' || vizinho === 'TIPO_VEICULO') {
    if (ui.alert('Aviso', 'A coluna "tipo_veículo" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else { abaBD.insertColumnAfter(colIntViagensIndex); abaBD.getRange(1, colunaDestinoIndex).setValue('tipo_veículo'); }

  const valoresEmpCarro = abaBD.getRange(2, colEmpCarroIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const resultados = [];
  for (let i = 0; i < valoresEmpCarro.length; i++) {
    let chave = valoresEmpCarro[i][0].toString().trim();
    if (chave !== "") {
      let val = dicionarioFrota[chave];
      resultados.push([val !== undefined ? val : "#N/D"]);
    } else resultados.push([""]); 
  }
  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna tipo_veículo foi preenchida com sucesso.', 'Passo 8 Concluído', 5);
}

/**
 * Passo 9: Buscar Capacidade de Passageiros Sentados
 */
function passo9_buscarCapPassSentado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD"); const abaFrota = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!abaBD || !abaFrota) { ui.alert('Erro', 'Abas BD or FROTA_ATUALIZADA não encontradas.', ui.ButtonSet.OK); return; }

  const ultLinhaFrota = abaFrota.getLastRow(); const ultColFrota = abaFrota.getLastColumn(); if (ultLinhaFrota <= 1) return;

  const cabecalhosFrota = abaFrota.getRange(1, 1, 1, ultColFrota).getValues()[0];
  let colEmpCodIndex = -1, colCapacidadeIndex = -1;
  for (let i = 0; i < cabecalhosFrota.length; i++) {
    let tit = cabecalhosFrota[i].toString().toUpperCase().trim();
    if (tit === 'EMP_COD') colEmpCodIndex = i + 1;
    if (tit === 'CAPACIDADE_PASS_SENT') colCapacidadeIndex = i + 1;
  }
  if (colEmpCodIndex === -1 || colCapacidadeIndex === -1) { ui.alert('Erro', 'Colunas EMP_COD ou CAPACIDADE_PASS_SENT não encontradas.', ui.ButtonSet.OK); return; }

  const valoresEmpCod = abaFrota.getRange(2, colEmpCodIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  const valoresCapacidade = abaFrota.getRange(2, colCapacidadeIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  
  const dicionarioFrota = {};
  for (let i = 0; i < valoresEmpCod.length; i++) {
    let chave = valoresEmpCod[i][0].toString().trim(); let valor = valoresCapacidade[i][0].toString().trim();
    if (chave !== "") dicionarioFrota[chave] = valor;
  }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhosBD = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colTipoVeiculoIndex = -1, colEmpCarroIndex = -1;
  for (let i = 0; i < cabecalhosBD.length; i++) {
    let tit = cabecalhosBD[i].toString().toUpperCase().trim();
    if (tit === 'TIPO_VEICULO' || tit === 'TIPO_VEÍCULO') colTipoVeiculoIndex = i + 1;
    if (tit === 'EMP_CARRO') colEmpCarroIndex = i + 1;
  }
  if (colTipoVeiculoIndex === -1 || colEmpCarroIndex === -1) { ui.alert('Erro', 'Colunas tipo_veículo ou EMP_CARRO não encontradas na BD.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colTipoVeiculoIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= abaBD.getMaxColumns()) vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  
  if (vizinho === 'CAP_PASS_SENTADO') {
    if (ui.alert('Aviso', 'A coluna "cap_pass_sentado" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else { abaBD.insertColumnAfter(colTipoVeiculoIndex); abaBD.getRange(1, colunaDestinoIndex).setValue('cap_pass_sentado'); }

  const valoresEmpCarro = abaBD.getRange(2, colEmpCarroIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const resultados = [];
  for (let i = 0; i < valoresEmpCarro.length; i++) {
    let chave = valoresEmpCarro[i][0].toString().trim();
    if (chave !== "") {
      let cap = dicionarioFrota[chave];
      resultados.push([cap !== undefined ? cap : "#N/D"]);
    } else resultados.push([""]); 
  }
  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna cap_pass_sentado foi preenchida.', 'Passo 9 Concluído', 5);
}

/**
 * Passo 10: Buscar Capacidade de Passageiros em Pé
 */
function passo10_buscarCapPassPe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD"); const abaFrota = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!abaBD || !abaFrota) { ui.alert('Erro', 'As abas "BD" ou "FROTA_ATUALIZADA" não foram encontradas.', ui.ButtonSet.OK); return; }

  const ultLinhaFrota = abaFrota.getLastRow(); const ultColFrota = abaFrota.getLastColumn(); if (ultLinhaFrota <= 1) return;

  const cabecalhosFrota = abaFrota.getRange(1, 1, 1, ultColFrota).getValues()[0];
  let colEmpCodIndex = -1, colCapacidadePeIndex = -1;
  for (let i = 0; i < cabecalhosFrota.length; i++) {
    let tit = cabecalhosFrota[i].toString().toUpperCase().trim();
    if (tit === 'EMP_COD') colEmpCodIndex = i + 1;
    if (tit === 'CAPACIDADE_PASS_PE' || tit === 'CAPACIDADE_PASS_PÉ') colCapacidadePeIndex = i + 1;
  }
  if (colEmpCodIndex === -1 || colCapacidadePeIndex === -1) { ui.alert('Erro', 'Colunas EMP_COD ou CAPACIDADE_PASS_PE não encontradas.', ui.ButtonSet.OK); return; }

  const valoresEmpCod = abaFrota.getRange(2, colEmpCodIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  const valoresCapacidadePe = abaFrota.getRange(2, colCapacidadePeIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  
  const dicionarioFrota = {};
  for (let i = 0; i < valoresEmpCod.length; i++) {
    let chave = valoresEmpCod[i][0].toString().trim(); let valor = valoresCapacidadePe[i][0].toString().trim();
    if (chave !== "") dicionarioFrota[chave] = valor;
  }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhosBD = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colCapSentadoIndex = -1, colEmpCarroIndex = -1;
  for (let i = 0; i < cabecalhosBD.length; i++) {
    let tit = cabecalhosBD[i].toString().toUpperCase().trim();
    if (tit === 'CAP_PASS_SENTADO') colCapSentadoIndex = i + 1;
    if (tit === 'EMP_CARRO') colEmpCarroIndex = i + 1;
  }
  if (colCapSentadoIndex === -1 || colEmpCarroIndex === -1) { ui.alert('Erro', 'Colunas cap_pass_sentado ou EMP_CARRO não encontradas na BD.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colCapSentadoIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= abaBD.getMaxColumns()) vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  
  if (vizinho === 'CAP_PASS_PE' || vizinho === 'CAP_PASS_PÉ') {
    if (ui.alert('Aviso', 'A coluna "cap_pass_pe" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else { abaBD.insertColumnAfter(colCapSentadoIndex); abaBD.getRange(1, colunaDestinoIndex).setValue('cap_pass_pe'); }

  const valoresEmpCarro = abaBD.getRange(2, colEmpCarroIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const resultados = [];
  for (let i = 0; i < valoresEmpCarro.length; i++) {
    let chave = valoresEmpCarro[i][0].toString().trim();
    if (chave !== "") {
      let cap = dicionarioFrota[chave];
      resultados.push([cap !== undefined ? cap : "#N/D"]);
    } else resultados.push([""]); 
  }
  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna cap_pass_pe foi preenchida com sucesso.', 'Passo 10 Concluído', 5);
}

/**
 * Passo 11: Buscar Capacidade de Passageiros Total
 */
function passo11_buscarCapPassTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD"); const abaFrota = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!abaBD || !abaFrota) { ui.alert('Erro', 'As abas BD ou FROTA_ATUALIZADA não foram encontradas.', ui.ButtonSet.OK); return; }

  const ultLinhaFrota = abaFrota.getLastRow(); const ultColFrota = abaFrota.getLastColumn(); if (ultLinhaFrota <= 1) return;

  const cabecalhosFrota = abaFrota.getRange(1, 1, 1, ultColFrota).getValues()[0];
  let colEmpCodIndex = -1, colCapacidadeTotalIndex = -1;
  for (let i = 0; i < cabecalhosFrota.length; i++) {
    let tit = cabecalhosFrota[i].toString().toUpperCase().trim();
    if (tit === 'EMP_COD') colEmpCodIndex = i + 1;
    if (tit === 'CAPACIDADE_PASS_TOTAL') colCapacidadeTotalIndex = i + 1;
  }
  if (colEmpCodIndex === -1 || colCapacidadeTotalIndex === -1) { ui.alert('Erro', 'Colunas EMP_COD ou CAPACIDADE_PASS_TOTAL não encontradas.', ui.ButtonSet.OK); return; }

  const valoresEmpCod = abaFrota.getRange(2, colEmpCodIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  const valoresCapacidadeTotal = abaFrota.getRange(2, colCapacidadeTotalIndex, ultLinhaFrota - 1, 1).getDisplayValues();
  
  const dicionarioFrota = {};
  for (let i = 0; i < valoresEmpCod.length; i++) {
    let chave = valoresEmpCod[i][0].toString().trim(); let valor = valoresCapacidadeTotal[i][0].toString().trim();
    if (chave !== "") dicionarioFrota[chave] = valor;
  }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhosBD = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colCapPeIndex = -1, colEmpCarroIndex = -1;
  for (let i = 0; i < cabecalhosBD.length; i++) {
    let tit = cabecalhosBD[i].toString().toUpperCase().trim();
    if (tit === 'CAP_PASS_PE' || tit === 'CAP_PASS_PÉ') colCapPeIndex = i + 1;
    if (tit === 'EMP_CARRO') colEmpCarroIndex = i + 1;
  }
  if (colCapPeIndex === -1 || colEmpCarroIndex === -1) { ui.alert('Erro', 'Colunas cap_pass_pe ou EMP_CARRO não encontradas na BD.', ui.ButtonSet.OK); return; }

  const colunaDestinoIndex = colCapPeIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= abaBD.getMaxColumns()) vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  
  if (vizinho === 'CAP_PASS_TOTAL') {
    if (ui.alert('Aviso', 'A coluna "cap_pass_total" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else { abaBD.insertColumnAfter(colCapPeIndex); abaBD.getRange(1, colunaDestinoIndex).setValue('cap_pass_total'); }

  const valoresEmpCarro = abaBD.getRange(2, colEmpCarroIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const resultados = [];
  for (let i = 0; i < valoresEmpCarro.length; i++) {
    let chave = valoresEmpCarro[i][0].toString().trim();
    if (chave !== "") {
      let cap = dicionarioFrota[chave];
      resultados.push([cap !== undefined ? cap : "#N/D"]);
    } else resultados.push([""]); 
  }
  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna cap_pass_total foi preenchida com sucesso.', 'Passo 11 Concluído', 5);
}

/**
 * Passo 12: Calcular Contagem de Passageiros Sentados
 */
function passo12_calcularContagemSentados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colVaziaIndex = -1, colSentadoIndex = -1, colEmPeIndex = -1, colCapSentadoIndex = -1, colCapTotalIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toLowerCase().trim();
    if (tit === 'cap_pass_total') colCapTotalIndex = i + 1;
    if (tit === 'vazia') colVaziaIndex = i + 1;
    if (tit === 'sentado') colSentadoIndex = i + 1;
    if (tit === 'em pé' || tit === 'em pe') colEmPeIndex = i + 1;
    if (tit === 'cap_pass_sentado') colCapSentadoIndex = i + 1;
  }

  if (colCapTotalIndex === -1 || colCapSentadoIndex === -1) {
    ui.alert('Erro', 'Faltam colunas de capacidade (cap_pass_sentado).', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colCapTotalIndex + 1;
  let vizinho = "";
  if (colunaDestinoIndex <= abaBD.getMaxColumns()) {
    vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  }
  
  if (vizinho === 'CONTAGEM_PASS_SENTADOS') {
    if (ui.alert('Aviso', 'A coluna "contagem_pass_sentados" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else {
    abaBD.insertColumnAfter(colCapTotalIndex);
    abaBD.getRange(1, colunaDestinoIndex).setValue('contagem_pass_sentados');
  }

  const valoresVazia = colVaziaIndex !== -1 ? abaBD.getRange(2, colVaziaIndex, ultLinhaBD - 1, 1).getDisplayValues() : [];
  const valoresSentado = colSentadoIndex !== -1 ? abaBD.getRange(2, colSentadoIndex, ultLinhaBD - 1, 1).getDisplayValues() : [];
  const valoresEmPe = colEmPeIndex !== -1 ? abaBD.getRange(2, colEmPeIndex, ultLinhaBD - 1, 1).getDisplayValues() : [];
  const valoresCapSentado = abaBD.getRange(2, colCapSentadoIndex, ultLinhaBD - 1, 1).getValues();
  
  const resultados = [];

  for (let i = 0; i < valoresCapSentado.length; i++) {
    let vaziaStr = valoresVazia.length > 0 ? valoresVazia[i][0].toString().trim().toUpperCase() : "";
    let sentadoStr = valoresSentado.length > 0 ? valoresSentado[i][0].toString().trim().toUpperCase() : "";
    let emPeStr = valoresEmPe.length > 0 ? valoresEmPe[i][0].toString().trim().toUpperCase() : "";
    let capSentadoNum = parseFloat(valoresCapSentado[i][0]) || 0;

    let res = "";

    if (vaziaStr.includes("CV") || sentadoStr.includes("CV")) {
      let textoParaExtrair = vaziaStr.includes("CV") ? vaziaStr : sentadoStr;
      let numVazias = parseFloat(textoParaExtrair.replace("CV", "").trim());
      res = !isNaN(numVazias) ? Math.max(0, capSentadoNum - numVazias) : 0;
    }
    else if (vaziaStr !== "" && !isNaN(parseFloat(vaziaStr)) && parseFloat(vaziaStr) > 0) {
      res = Math.max(0, capSentadoNum - parseFloat(vaziaStr)); 
    }
    else if (emPeStr === "0") { res = 0; }
    else if (emPeStr === "LT" || emPeStr === "SL" || (!isNaN(parseFloat(emPeStr)) && parseFloat(emPeStr) > 0)) {
      res = capSentadoNum;
    }
    else if (sentadoStr === "BC") { res = capSentadoNum; } 
    else if (sentadoStr !== "" && !isNaN(parseFloat(sentadoStr))) {
      res = parseFloat(sentadoStr);
    }
    else { res = 0; }
    
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('Passo 12 Concluído', 'Sucesso', 5);
}

/**
 * Passo 13: Calcular Contagem de Passageiros em Pé
 */
function passo13_calcularContagemPe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContSentadosIndex = -1, colEmPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_SENTADOS') colContSentadosIndex = i + 1;
    if (tit === 'EM PÉ' || tit === 'EM PE') colEmPeIndex = i + 1;
  }

  if (colContSentadosIndex === -1 || colEmPeIndex === -1) {
    ui.alert('Erro', 'Execute o Passo 12 primeiro.', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colContSentadosIndex + 1;
  let vizinho = "";
  if (colunaDestinoIndex <= abaBD.getMaxColumns()) {
    vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  }
  
  if (vizinho === 'CONTAGEM_PASS_PE' || vizinho === 'CONTAGEM_PASS_PÉ') {
    if (ui.alert('Aviso', 'A aba "contagem_pass_pe" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else {
    abaBD.insertColumnAfter(colContSentadosIndex);
    abaBD.getRange(1, colunaDestinoIndex).setValue('contagem_pass_pe');
  }

  const valoresEmPe = abaBD.getRange(2, colEmPeIndex, ultLinhaBD - 1, 1).getValues();
  const resultados = [];

  for (let i = 0; i < valoresEmPe.length; i++) {
    let valorOriginal = valoresEmPe[i][0].toString().toUpperCase().trim();
    let res = (valorOriginal === "LT") ? "LOTADO" : (valorOriginal === "SL") ? "SUPERLOTADO" : (valorOriginal === "" || isNaN(parseFloat(valorOriginal))) ? 0 : parseFloat(valorOriginal);
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna contagem_pass_pe foi calculada.', 'Passo 13 Concluído', 5);
}

/**
 * Passo 14: Calcular Contagem Total (Sentados + Em Pé)
 */
function passo14_calcularContagemTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContSentadosIndex = -1, colContPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_SENTADOS') colContSentadosIndex = i + 1;
    if (tit === 'CONTAGEM_PASS_PE' || tit === 'CONTAGEM_PASS_PÉ') colContPeIndex = i + 1;
  }

  if (colContSentadosIndex === -1 || colContPeIndex === -1) {
    ui.alert('Erro', 'Execute os passos 12 e 13.', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colContPeIndex + 1;
  let vizinho = "";
  if (colunaDestinoIndex <= abaBD.getMaxColumns()) {
    vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  }
  
  if (vizinho === 'CONTAGEM_PASS_TOTAL') {
    if (ui.alert('Aviso', 'A coluna "contagem_pass_total" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else {
    abaBD.insertColumnAfter(colContPeIndex);
    abaBD.getRange(1, colunaDestinoIndex).setValue('contagem_pass_total');
  }

  const valoresSentados = abaBD.getRange(2, colContSentadosIndex, ultLinhaBD - 1, 1).getValues();
  const valoresPe = abaBD.getRange(2, colContPeIndex, ultLinhaBD - 1, 1).getValues();
  const resultados = [];

  for (let i = 0; i < valoresSentados.length; i++) {
    let s = valoresSentados[i][0];
    let p = valoresPe[i][0];
    let res = (s === "SUPERLOTADO" || p === "SUPERLOTADO") ? "SL" : (s === "LOTADO" || p === "LOTADO") ? "L" : (parseFloat(s) || 0) + (parseFloat(p) || 0);
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('Passo 14 Concluído', 'Sucesso', 5);
}

/**
 * Passo 15: Calcular Taxa de Ocupação dos Sentados
 */
function passo15_calcularTaxaOcupSentado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContagemTotalIndex = -1; let colContagemSentadosIndex = -1; let colCapSentadoIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_TOTAL') colContagemTotalIndex = i + 1;
    if (tit === 'CONTAGEM_PASS_SENTADOS') colContagemSentadosIndex = i + 1;
    if (tit === 'CAP_PASS_SENTADO') colCapSentadoIndex = i + 1;
  }

  if (colContagemTotalIndex === -1 || colContagemSentadosIndex === -1 || colCapSentadoIndex === -1) {
    ui.alert('Erro', 'Faltam colunas estruturais.', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colContagemTotalIndex + 1;
  let vizinho = ""; if (colunaDestinoIndex <= abaBD.getMaxColumns()) vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  
  if (vizinho === 'TAXA_OCUP_SENTADO') {
    if (ui.alert('Aviso', 'A coluna "taxa_ocup_sentado" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else { abaBD.insertColumnAfter(colContagemTotalIndex); abaBD.getRange(1, colunaDestinoIndex).setValue('taxa_ocup_sentado'); }

  const valoresContagemSentados = abaBD.getRange(2, colContagemSentadosIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const valoresCapSentado = abaBD.getRange(2, colCapSentadoIndex, ultLinhaBD - 1, 1).getDisplayValues();
  const resultados = [];

  for (let i = 0; i < valoresContagemSentados.length; i++) {
    let contSentStr = valoresContagemSentados[i][0].toString().trim();
    let capSentStr = valoresCapSentado[i][0].toString().trim();

    if (contSentStr.includes("Erro") || capSentStr === "#N/D") resultados.push(["Erro"]);
    else if (contSentStr === "" || capSentStr === "") resultados.push([""]);
    else {
      let numContSentados = parseFloat(contSentStr); let numCapSentado = parseFloat(capSentStr);
      if (!isNaN(numContSentados) && !isNaN(numCapSentado) && numCapSentado > 0) resultados.push([(numContSentados / numCapSentado) * 100]); 
      else resultados.push([""]);
    }
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados); 
  rangeDestino.setNumberFormat("0.00"); 
  ss.toast('Passo 15 Concluído', 'Sucesso', 5);
}

/**
 * Passo 16: Calcular Taxa de Ocupação em Pé (%)
 */
function passo16_calcularTaxaOcupPe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContPeIndex = -1, colCapPeIndex = -1, colTaxaSentadoIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_PE' || tit === 'CONTAGEM_PASS_PÉ') colContPeIndex = i + 1;
    if (tit === 'CAP_PASS_PE' || tit === 'CAP_PASS_PÉ') colCapPeIndex = i + 1;
    if (tit === 'TAXA_OCUP_SENTADO') colTaxaSentadoIndex = i + 1;
  }

  if (colContPeIndex === -1 || colCapPeIndex === -1 || colTaxaSentadoIndex === -1) {
    ui.alert('Erro', 'Verifique os passos 10, 13 e 15.', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colTaxaSentadoIndex + 1;
  let vizinho = "";
  if (colunaDestinoIndex <= abaBD.getMaxColumns()) {
    vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  }
  
  if (vizinho === 'TAXA_OCUP_PE' || vizinho === 'TAXA_OCUP_PÉ') {
    if (ui.alert('Aviso', 'A coluna "taxa_ocup_pe" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else {
    abaBD.insertColumnAfter(colTaxaSentadoIndex);
    abaBD.getRange(1, colunaDestinoIndex).setValue('taxa_ocup_pe');
  }

  const valoresContPe = abaBD.getRange(2, colContPeIndex, ultLinhaBD - 1, 1).getValues();
  const valoresCapPe = abaBD.getRange(2, colCapPeIndex, ultLinhaBD - 1, 1).getValues();
  const resultados = [];

  for (let i = 0; i < valoresContPe.length; i++) {
    let cont = valoresContPe[i][0];
    let cap = parseFloat(valoresCapPe[i][0]);
    let res = (cont === "SUPERLOTADO" || cont === "LOTADO") ? cont : (!isNaN(cap) && cap > 0) ? (parseFloat(cont) || 0) / cap * 100 : (parseFloat(cont) > 0) ? "SUPERLOTADO" : 0;
    resultados.push([res]);
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados);
  rangeDestino.setNumberFormat("0.00");
  ss.toast('Passo 16 Concluído', 'Sucesso', 5);
}

/**
 * Passo 17: Calcular Taxa de Ocupação Total (%)
 */
function passo17_calcularTaxaOcupTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContTotalIndex = -1, colCapTotalIndex = -1, colTaxaPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_TOTAL') colContTotalIndex = i + 1;
    if (tit === 'CAP_PASS_TOTAL') colCapTotalIndex = i + 1;
    if (tit === 'TAXA_OCUP_PE' || tit === 'TAXA_OCUP_PÉ') colTaxaPeIndex = i + 1;
  }

  if (colContTotalIndex === -1 || colCapTotalIndex === -1 || colTaxaPeIndex === -1) {
    ui.alert('Erro', 'Verifique os passos 11, 14 e 16.', ui.ButtonSet.OK);
    return;
  }

  const colunaDestinoIndex = colTaxaPeIndex + 1;
  let vizinho = "";
  if (colunaDestinoIndex <= abaBD.getMaxColumns()) {
    vizinho = abaBD.getRange(1, colunaDestinoIndex).getValue().toString().toUpperCase().trim();
  }
  
  if (vizinho === 'TAXA_OCUP_TOTAL') {
    if (ui.alert('Aviso', 'A coluna "taxa_ocup_total" já existe. Atualizar?', ui.ButtonSet.YES_NO) == ui.Button.NO) return;
  } else {
    abaBD.insertColumnAfter(colTaxaPeIndex);
    abaBD.getRange(1, colunaDestinoIndex).setValue('taxa_ocup_total');
  }

  const valoresContTotal = abaBD.getRange(2, colContTotalIndex, ultLinhaBD - 1, 1).getValues();
  const valoresCapTotal = abaBD.getRange(2, colCapTotalIndex, ultLinhaBD - 1, 1).getValues();
  const resultados = [];

  for (let i = 0; i < valoresContTotal.length; i++) {
    let cont = valoresContTotal[i][0];
    let cap = parseFloat(valoresCapTotal[i][0]);
    let res = (cont === "SUPERLOTADO" || cont === "LOTADO") ? cont : (!isNaN(cap) && cap > 0) ? (parseFloat(cont) || 0) / cap * 100 : (parseFloat(cont) > 0) ? "SUPERLOTADO" : 0;
    resultados.push([res]);
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados);
  rangeDestino.setNumberFormat("0.00");
  ss.toast('A taxa de ocupação total foi calculada.', 'Passo 17 Concluído', 5);
}

/**
 * Passo 18: Formatar a Aba BD inteira
 */
function passo18_formatarAbaBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultimaLinha = abaBD.getLastRow();
  const ultimaColuna = abaBD.getLastColumn();
  if (ultimaLinha <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultimaColuna).getValues()[0];

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    let colIndex = i + 1;
    let rangeColuna = abaBD.getRange(2, colIndex, ultimaLinha - 1, 1);

    if (tit === 'DATA') { rangeColuna.setNumberFormat("dd/MM/yyyy"); }
    else if (['PESQUISADORES', 'LOCAL', 'SENTIDO', 'EMPRESA', 'LINHA', 'VIA', 'CARRO', 'EMP_CARRO', 'FAIXA-HORÁRIA', 'TIPO_VEÍCULO', 'TIPO_VEICULO'].includes(tit)) {
      rangeColuna.setNumberFormat("@");
    }
    else if (['Nº DA VIAGEM', 'Nº VIAGEM', 'VAZIA', 'SENTADO', 'EM PÉ', 'EM PE', 'CAP_PASS_SENTADO', 'CAP_PASS_PE', 'CAP_PASS_PÉ', 'CAP_PASS_TOTAL', 'CONTAGEM_PASS_SENTADOS', 'CONTAGEM_PASS_PE', 'CONTAGEM_PASS_PÉ', 'CONTAGEM_PASS_TOTAL'].includes(tit)) {
      rangeColuna.setNumberFormat("0");
    }
    else if (tit === 'HORA' || tit === 'INT_VIAGENS') { rangeColuna.setNumberFormat("HH:mm"); }
    else if (['TAXA_OCUP_SENTADO', 'TAXA_OCUP_PE', 'TAXA_OCUP_PÉ', 'TAXA_OCUP_TOTAL'].includes(tit)) {
      rangeColuna.setNumberFormat("0.00%");
    }
  }

  const rangeTotal = abaBD.getRange(1, 1, ultimaLinha, ultimaColuna);
  rangeTotal.setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle");
  abaBD.getRange(1, 1, 1, ultimaColuna).setFontWeight("bold");
  ss.toast('Dados formatados.', 'Passo 18 Concluído', 5);
}

/**
 * Passo 19: Criar Aba Tabelas de Análises
 */
function passo19_criarAbaAnalises() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nomeAba = "tabelas_análises";
  let abaAnalises = ss.getSheetByName(nomeAba);

  if (abaAnalises) {
    if (ui.alert('Aviso', 'A aba já existe. Deseja substituí-la?', ui.ButtonSet.YES_NO) == ui.Button.YES) {
      ss.deleteSheet(abaAnalises);
      abaAnalises = ss.insertSheet(nomeAba);
    } else { abaAnalises.activate(); return; }
  } else { abaAnalises = ss.insertSheet(nomeAba); }

  abaAnalises.activate();
  ss.toast('Aba criada.', 'Passo 19 Concluído', 5);
}

/**
 * Passo 20: Criar Tabela de Análise (Viagens por Faixa Horária e Local)
 */
function passo20_analiseViagensLocal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");
  
  if (!abaBD) { ui.alert('Erro', 'Aba BD não encontrada.', ui.ButtonSet.OK); return; }
  if (!abaAnalises) { abaAnalises = ss.insertSheet("tabelas_análises"); }

  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) return;
  const cabecalhos = dadosBD[0];
  
  let colFaixaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'FAIXA-HORÁRIA');
  let colLinhaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA');
  let colLocalIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL');
  let colSentidoIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO');

  if (colFaixaIndex === -1 || colLocalIndex === -1 || colSentidoIndex === -1) {
    ui.alert('Erro', 'Colunas não encontradas.', ui.ButtonSet.OK);
    return;
  }

  let faixasUnicas = []; let colunasDinamicas = []; let matrizContagem = {}; 

  for (let i = 1; i < dadosBD.length; i++) {
    let faixa = dadosBD[i][colFaixaIndex].toString().trim();
    let local = dadosBD[i][colLocalIndex].toString().trim();
    let sentido = dadosBD[i][colSentidoIndex].toString().trim();
    let linhaBus = colLinhaIndex !== -1 ? dadosBD[i][colLinhaIndex].toString().trim() : "";
    if (!faixa || !local) continue; 
    let nomeColuna = linhaBus !== "" ? linhaBus + " | " + local + (sentido ? " - " + sentido : "") : local + (sentido ? " - " + sentido : "");

    if (!faixasUnicas.includes(faixa)) faixasUnicas.push(faixa);
    if (!colunasDinamicas.includes(nomeColuna)) colunasDinamicas.push(nomeColuna);
    if (!matrizContagem[faixa]) matrizContagem[faixa] = {};
    if (!matrizContagem[faixa][nomeColuna]) matrizContagem[faixa][nomeColuna] = 0;
    matrizContagem[faixa][nomeColuna]++;
  }

  faixasUnicas.sort(); colunasDinamicas.sort();
  let tabelaFinal = []; let cabecalhoTabela = ["FAIXA-HORÁRIA"];
  for (let c of colunasDinamicas) cabecalhoTabela.push(c);
  tabelaFinal.push(cabecalhoTabela);
  let totaisPorColuna = new Array(colunasDinamicas.length).fill(0);

  for (let f of faixasUnicas) {
    let linhaTabela = [f]; 
    for (let i = 0; i < colunasDinamicas.length; i++) {
      let c = colunasDinamicas[i];
      let valor = matrizContagem[f][c] || 0; 
      linhaTabela.push(valor);
      totaisPorColuna[i] += valor; 
    }
    tabelaFinal.push(linhaTabela);
  }

  let linhaTotal = ["Total geral"];
  for (let total of totaisPorColuna) { linhaTotal.push(total); }
  tabelaFinal.push(linhaTotal);

  let ultimaLinhaAnalises = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinhaAnalises === 0 ? 2 : ultimaLinhaAnalises + 3; 
  abaAnalises.getRange(linhaInicio - 1, 1).setValue("ANÁLISE 1: QUADRO DE VIAGENS POR FAIXA HORÁRIA E LOCAL").setFontWeight("bold").setFontSize(12);
  let rangeDestino = abaAnalises.getRange(linhaInicio, 1, tabelaFinal.length, tabelaFinal[0].length);
  rangeDestino.setValues(tabelaFinal).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle");
  abaAnalises.getRange(linhaInicio, 1, 1, tabelaFinal[0].length).setFontWeight("bold").setBackground("#d9ead3").setWrap(true);
  abaAnalises.getRange(linhaInicio + tabelaFinal.length - 1, 1, 1, tabelaFinal[0].length).setFontWeight("bold");

  abaAnalises.activate();
  ss.toast('Análise 1 Concluída', 'Sucesso', 5);
}

/**
 * Passo 21: Criar Tabela de Análise Hierárquica (Taxa de Ocupação Total)
 */
function passo21_analiseTaxaOcupacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");
  
  if (!abaBD) { ui.alert('Erro', 'Aba BD não encontrada.', ui.ButtonSet.OK); return; }
  if (!abaAnalises) { abaAnalises = ss.insertSheet("tabelas_análises"); }

  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) return;
  const cabecalhos = dadosBD[0];
  
  let colLinhaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA');
  let colLocalIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL');
  let colSentidoIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO');
  let colTaxaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'TAXA_OCUP_TOTAL');

  if (colLocalIndex === -1 || colTaxaIndex === -1) {
    ui.alert('Erro', 'Colunas LOCAL ou TAXA_OCUP_TOTAL não encontradas.', ui.ButtonSet.OK);
    return;
  }

  let relatorioObj = {}; let totalGeralMatriz = [0, 0, 0, 0, 0, 0]; 
  
  for (let i = 1; i < dadosBD.length; i++) {
    let taxaStr = dadosBD[i][colTaxaIndex].toString().trim().toUpperCase();
    let linha = colLinhaIndex !== -1 ? dadosBD[i][colLinhaIndex].toString().trim() : "Linha Indefinida";
    let local = dadosBD[i][colLocalIndex].toString().trim();
    let sentido = colSentidoIndex !== -1 ? dadosBD[i][colSentidoIndex].toString().trim() : "";
    let nomeLocal = local + (sentido ? " - " + sentido : "");
    if (!linha) linha = "Linha Indefinida";

    let b = -1; 
    if (taxaStr === "-") { b = 5; } else if (taxaStr === "LOTADO") { b = 4; } else {
      let taxaNum = parseFloat(taxaStr.replace('%', '').replace(',', '.'));
      if (!isNaN(taxaNum)) {
        if (taxaNum < 25) b = 0; else if (taxaNum < 50) b = 1; else if (taxaNum < 75) b = 2; else if (taxaNum < 100) b = 3; else if (taxaNum <= 115) b = 4; else b = 5;
      }
    }

    if (b !== -1) {
      if (!relatorioObj[linha]) { relatorioObj[linha] = { total: [0,0,0,0,0,0], locais: {} }; }
      if (!relatorioObj[linha].locais[nomeLocal]) { relatorioObj[linha].locais[nomeLocal] = [0,0,0,0,0,0]; }
      relatorioObj[linha].total[b]++; relatorioObj[linha].locais[nomeLocal][b]++; totalGeralMatriz[b]++;
    }
  }

  function transformarEmPercentuais(counts) {
    let somaTratada = counts.reduce((a, b) => a + b, 0);
    if (somaTratada === 0) return [0, 0, 0, 0, 0, 0, 0];
    let percentuais = counts.map(c => c / somaTratada);
    percentuais.push(1); return percentuais;
  }

  let tabelaFinal = [];
  let cabecalhoTabela = ["Linhas / Local (Parada)", "<25%", "25% - 49%", "50% - 74%", "75% - 99%", "100% - 115% (Lotado)", ">115% (Crítico)", "Total Geral"];
  tabelaFinal.push(cabecalhoTabela);

  let linhasOrdenadas = Object.keys(relatorioObj).sort();
  let indicesNegrito = []; 

  for (let l of linhasOrdenadas) {
    tabelaFinal.push([l].concat(transformarEmPercentuais(relatorioObj[l].total)));
    indicesNegrito.push(tabelaFinal.length); 
    let locaisOrdenados = Object.keys(relatorioObj[l].locais).sort();
    for (let loc of locaisOrdenados) {
      tabelaFinal.push(["    " + loc].concat(transformarEmPercentuais(relatorioObj[l].locais[loc]))); 
    }
  }

  tabelaFinal.push(["Total Geral"].concat(transformarEmPercentuais(totalGeralMatriz)));
  indicesNegrito.push(tabelaFinal.length); 

  let ultimaLinhaAnalises = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinhaAnalises === 0 ? 2 : ultimaLinhaAnalises + 3; 
  abaAnalises.getRange(linhaInicio - 1, 1).setValue("ANÁLISE 2: TAXA DE OCUPAÇÃO TOTAL POR LINHA E PARADA").setFontWeight("bold").setFontSize(12);
  let rangeDestino = abaAnalises.getRange(linhaInicio, 1, tabelaFinal.length, tabelaFinal[0].length);
  rangeDestino.setValues(tabelaFinal).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setVerticalAlignment("middle");
  abaAnalises.getRange(linhaInicio + 1, 2, tabelaFinal.length - 1, 7).setNumberFormat("0%");
  abaAnalises.getRange(linhaInicio + 1, 1, tabelaFinal.length - 1, 1).setHorizontalAlignment("left");
  abaAnalises.getRange(linhaInicio, 1, 1, tabelaFinal[0].length).setFontWeight("bold").setBackground("#efefef").setWrap(true);

  for (let idx of indicesNegrito) { abaAnalises.getRange(linhaInicio + idx - 1, 1, 1, tabelaFinal[0].length).setFontWeight("bold"); }
  abaAnalises.activate();
  ss.toast('Análise 2 Concluída', 'Sucesso', 5);
}

/**
 * Passo 22: Análise interativa por Local, Sentido, Linha e PED
 */
function passo22_analisePorLocalSentidoPED() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");

  if (!abaBD) { ui.alert("Erro", "Aba 'BD' não encontrada.", ui.ButtonSet.OK); return; }
  if (!abaAnalises) { abaAnalises = ss.insertSheet("tabelas_análises"); }

  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) return;
  const cabecalhos = dadosBD[0];

  const colunas = {
    faixa: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'FAIXA-HORÁRIA'),
    local: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL'),
    sentido: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO'),
    linha: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA'),
    ped: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'PED'),
    data: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'DATA'),
    taxaTotal: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'TAXA_OCUP_TOTAL')
  };

  if (colunas.faixa === -1 || colunas.local === -1 || colunas.taxaTotal === -1) {
    ui.alert("Erro", "Colunas essenciais não encontradas.", ui.ButtonSet.OK);
    return;
  }

  const respostaModo = ui.alert("Modo de Análise", "Detectar automaticamente as combinações?", ui.ButtonSet.YES_NO);
  let combinacoes = (respostaModo === ui.Button.YES) ? detectarCombinacoesUnicas(dadosBD, colunas) : coletarFiltrosManualmente(ui, dadosBD, colunas);
  
  if (combinacoes.length === 0) return;
  const diasUnicos = obterDiasUnicos(dadosBD, colunas.data);

  for (let combo of combinacoes) {
    const tituloTabela = montarTituloTabela(combo);
    const dadosFiltrados = filtrarDadosPorCombinacao(dadosBD, colunas, combo);
    const tabela = construirTabelaPorFaixa(dadosFiltrados, diasUnicos, colunas);
    if (tabela && tabela.length > 1) { inserirTabelaNaAbaAnalises(abaAnalises, tituloTabela, tabela); }
  }

  abaAnalises.activate();
  ss.toast(`${combinacoes.length} tabela(s) criada(s).`, "Passo 22 Concluído", 8);
}

// ==================== FUNÇÕES AUXILIARES ====================

function converterLetraParaNumero(letras) {
  let coluna = 0;
  for (let i = 0; i < letras.length; i++) { coluna += (letras.charCodeAt(i) - 64) * Math.pow(26, letras.length - i - 1); }
  return coluna;
}

function auxiliar_converterParaMinutos(horaStr) {
  let partes = horaStr.split(':');
  if (partes.length >= 2) {
    let h = parseInt(partes[0].trim(), 10); let m = parseInt(partes[1].trim(), 10);
    if (!isNaN(h) && !isNaN(m)) return (h * 60) + m;
  }
  return null;
}

function detectarCombinacoesUnicas(dadosBD, colunas) {
  const combinacoesSet = new Set();
  for (let i = 1; i < dadosBD.length; i++) {
    const linha = colunas.linha !== -1 ? dadosBD[i][colunas.linha].toString().trim() : "";
    const local = dadosBD[i][colunas.local].toString().trim();
    const sentido = colunas.sentido !== -1 ? dadosBD[i][colunas.sentido].toString().trim() : "";
    const ped = colunas.ped !== -1 ? dadosBD[i][colunas.ped].toString().trim() : "";
    if (local) combinacoesSet.add(`${linha}|${local}|${sentido}|${ped}`);
  }
  return Array.from(combinacoesSet).map(chave => {
    const [linha, local, sentido, ped] = chave.split("|");
    return { linha, local, sentido, ped };
  });
}

function coletarFiltrosManualmente(ui, dadosBD, colunas) {
  const promptFilter = (label, index) => {
    const disponiveis = obterValoresUnicos(dadosBD, index);
    if (disponiveis.length === 0) return [""];
    const resp = ui.prompt(label, `Disponíveis: ${disponiveis.join(", ")}\nDigite o desejado:`, ui.ButtonSet.OK_CANCEL);
    return resp.getSelectedButton() === ui.Button.OK ? [resp.getResponseText().trim()] : null;
  };

  const linhas = perguntarMaisDeUm(ui, "Filtro por LINHA?") ? promptFilter("Linhas", colunas.linha) : [""];
  if (!linhas) return [];
  const locais = perguntarMaisDeUm(ui, "Filtro por LOCAL?") ? promptFilter("Locais", colunas.local) : [""];
  if (!locais) return [];
  const sentidos = perguntarMaisDeUm(ui, "Filtro por SENTIDO?") ? promptFilter("Sentidos", colunas.sentido) : [""];
  if (!sentidos) return [];
  const peds = perguntarMaisDeUm(ui, "Filtro por PED?") ? promptFilter("PEDs", colunas.ped) : [""];
  if (!peds) return [];

  const combinacoes = [];
  linhas.forEach(l => locais.forEach(loc => sentidos.forEach(s => peds.forEach(p => { if(loc) combinacoes.push({linha:l, local:loc, sentido:s, ped:p}); }))));
  return combinacoes;
}

function perguntarMaisDeUm(ui, pergunta) { return ui.alert(pergunta, ui.ButtonSet.YES_NO) === ui.Button.YES; }

function obterValoresUnicos(dadosBD, colunaIndex) {
  if (colunaIndex === -1) return [];
  const valores = new Set();
  for (let i = 1; i < dadosBD.length; i++) { if (dadosBD[i][colunaIndex]) valores.add(dadosBD[i][colunaIndex].toString().trim()); }
  return Array.from(valores);
}

function obterDiasUnicos(dadosBD, colunaDataIndex) {
  const dias = obterValoresUnicos(dadosBD, colunaDataIndex);
  return dias.length > 0 ? dias : ["Dia 1"];
}

function montarTituloTabela(combo) {
  return `${combo.linha ? combo.linha + " - " : ""}${combo.local}${combo.sentido ? " - " + combo.sentido : ""}${combo.ped ? " - PED " + combo.ped : ""}`;
}

function filtrarDadosPorCombinacao(dadosBD, colunas, combo) {
  return dadosBD.slice(1).filter(row => {
    const l = colunas.linha !== -1 ? row[colunas.linha].toString().trim() : "";
    const loc = row[colunas.local].toString().trim();
    const s = colunas.sentido !== -1 ? row[colunas.sentido].toString().trim() : "";
    const p = colunas.ped !== -1 ? row[colunas.ped].toString().trim() : "";
    return (!combo.linha || l === combo.linha) && (!combo.local || loc === combo.local) && (!combo.sentido || s === combo.sentido) && (!combo.ped || p === combo.ped);
  });
}

function construirTabelaPorFaixa(dadosFiltrados, diasUnicos, colunas) {
  if (dadosFiltrados.length === 0) return [];
  const mapaFaixas = new Map();

  for (const row of dadosFiltrados) {
    const faixa = row[colunas.faixa].toString().trim();
    if (!faixa) continue;
    const data = colunas.data !== -1 ? row[colunas.data].toString().trim() : "Dia 1";
    const taxaStr = row[colunas.taxaTotal].toString().trim().toUpperCase();
    const taxaNum = parseFloat(taxaStr.replace("%", "").replace(",", "."));
    const alta = (taxaStr === "-" || taxaStr === "LOTADO" || (!isNaN(taxaNum) && taxaNum >= 100));

    if (!mapaFaixas.has(faixa)) mapaFaixas.set(faixa, { total: 0, porDia: new Map(), alta: 0 });
    const g = mapaFaixas.get(faixa);
    g.total++;
    g.porDia.set(data, (g.porDia.get(data) || 0) + 1);
    if (alta) g.alta++;
  }

  const faixasOrd = Array.from(mapaFaixas.keys()).sort((a, b) => parseInt(a) - parseInt(b));
  const dDisp = diasUnicos.slice(0, 2);
  while (dDisp.length < 2) dDisp.push(`Dia ${dDisp.length + 1}`);

  const tabela = [["Faixa Horária", `Qtd ${dDisp[0]}`, `Qtd ${dDisp[1]}`, "Média", "Qtd >=100%", "% >=100%"]];
  faixasOrd.forEach(f => {
    const g = mapaFaixas.get(f);
    const d1 = g.porDia.get(dDisp[0]) || 0;
    const d2 = g.porDia.get(dDisp[1]) || 0;
    tabela.push([f, d1, d2, ((d1 + d2) / (dDisp[1] ? 2 : 1)).toFixed(1), g.alta, g.total > 0 ? g.alta / g.total : 0]);
  });
  return tabela;
}

function inserirTabelaNaAbaAnalises(abaAnalises, titulo, tabela) {
  let ultimaLinha = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinha === 0 ? 2 : ultimaLinha + 4;
  abaAnalises.getRange(linhaInicio - 1, 1).setValue(titulo).setFontWeight("bold").setFontSize(12);
  const range = abaAnalises.getRange(linhaInicio, 1, tabela.length, tabela[0].length);
  range.setValues(tabela).setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");
  abaAnalises.getRange(linhaInicio, 1, 1, tabela[0].length).setFontWeight("bold").setBackground("#d9ead3");
  abaAnalises.getRange(linhaInicio + 1, tabela[0].length, tabela.length - 1, 1).setNumberFormat("0.00%");
  for (let i = 1; i <= tabela[0].length; i++) abaAnalises.autoResizeColumn(i);
}
