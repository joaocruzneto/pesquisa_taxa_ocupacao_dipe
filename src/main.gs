/**
 * @OnlyCurrentDoc
 * Script de Automação de Pesquisa de Mancha
 * Versão: 1.0.2
 * Data de última atualização: 01/05/2026
 * Responsável: [João Cruz Neto / Divisão de Pesquisas]
 */

/**
 * Função executada automaticamente ao abrir a planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Pesquisa de Mancha')
    .addItem('1. Gerar Base de Dados (BD)', 'passo1_criarBD')
    .addItem('2. Criar Faixa Horária', 'passo2_criarFaixaHoraria')
    .addItem('3. Criar Coluna EMP_CARRO', 'passo3_criarEmpCarro')
    .addItem('4. Criar Coluna Int_Viagens', 'passo4_criarIntViagens')
    .addSeparator() 
    .addItem('5. Padronizar Frota (MOB/CNO)', 'passo5_padronizarFrota')
    .addItem('6. Criar Coluna EMP_COD (Frota)', 'passo6_criarEmpCodFrota')
    .addSeparator()
    .addItem('7. Buscar Tipo de Veículo', 'passo7_buscarTipoVeiculo')
    .addItem('8. Buscar Cap. Pass. Sentado', 'passo8_buscarCapPassSentado')
    .addItem('9. Buscar Cap. Pass. em Pé', 'passo9_buscarCapPassPe')
    .addItem('10. Buscar Cap. Pass. Total', 'passo10_buscarCapPassTotal')
    .addSeparator()
    .addItem('11. Calcular Contagem Sentados', 'passo11_calcularContagemSentados')
    .addItem('12. Calcular Contagem em Pé', 'passo12_calcularContagemPe')
    .addItem('13. Calcular Contagem Total', 'passo13_calcularContagemTotal')
    .addSeparator()
    .addItem('14. Calcular Taxa Ocup. Sentado', 'passo14_calcularTaxaOcupSentado')
    .addItem('15. Calcular Taxa Ocup. em Pé', 'passo15_calcularTaxaOcupPe')
    .addItem('16. Calcular Taxa Ocup. Total', 'passo16_calcularTaxaOcupTotal')
    .addSeparator()
    .addItem('17. Formatar Aba BD', 'passo17_formatarAbaBD')
    .addItem('18. Criar Aba Tabelas Análises', 'passo18_criarAbaAnalises')
    .addSeparator()
    .addItem('19. Análise: Viagens por Faixa/Local', 'passo19_analiseViagensLocal')
    .addItem('20. Análise: Ocupação Total (Hierarquia)', 'passo20_analiseTaxaOcupacao')
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
 * Passo 2: Localiza HORA, insere FAIXA-HORÁRIA e preenche.
 */
function passo2_criarFaixaHoraria() {
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
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 2 Concluído', 'Sucesso', 5);
}

/**
 * Passo 3: Cria EMP_CARRO concatenando EMPRESA e CARRO (Aba BD)
 */
function passo3_criarEmpCarro() {
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
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 3 Concluído', 'Sucesso', 5);
}

/**
 * Passo 4: Cria Int_Viagens APÓS FAIXA-HORÁRIA
 */
function passo4_criarIntViagens() {
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
  SpreadsheetApp.getActiveSpreadsheet().toast('Passo 4 Concluído', 'Sucesso', 5);
}

/**
 * Passo 5: Padroniza os códigos da frota na aba FROTA_ATUALIZADA
 */
function passo5_padronizarFrota() {
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
  if (alteracoes > 0) ss.toast(alteracoes + ' registros corrigidos.', 'Passo 5 Concluído', 5);
  else ss.toast('Os códigos já estavam no padrão.', 'Passo 5 Concluído', 4);
}

/**
 * Passo 6: Cria a coluna EMP_COD na aba FROTA_ATUALIZADA
 */
function passo6_criarEmpCodFrota() {
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
  ss.toast('A coluna EMP_COD foi criada com sucesso.', 'Passo 6 Concluído', 5);
}

/**
 * Passo 7: Buscar Tipo de Veículo
 */
function passo7_buscarTipoVeiculo() {
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
  ss.toast('A coluna tipo_veículo foi preenchida com sucesso.', 'Passo 7 Concluído', 5);
}

/**
 * Passo 8: Buscar Capacidade de Passageiros Sentados
 */
function passo8_buscarCapPassSentado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD"); const abaFrota = ss.getSheetByName("FROTA_ATUALIZADA");
  if (!abaBD || !abaFrota) { ui.alert('Erro', 'Abas BD ou FROTA_ATUALIZADA não encontradas.', ui.ButtonSet.OK); return; }

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
  ss.toast('A coluna cap_pass_sentado foi preenchida.', 'Passo 8 Concluído', 5);
}

/**
 * Passo 9: Buscar Capacidade de Passageiros em Pé
 */
function passo9_buscarCapPassPe() {
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
  ss.toast('A coluna cap_pass_pe foi preenchida com sucesso.', 'Passo 9 Concluído', 5);
}

/**
 * Passo 10: Buscar Capacidade de Passageiros Total
 */
function passo10_buscarCapPassTotal() {
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
  ss.toast('A coluna cap_pass_total foi preenchida com sucesso.', 'Passo 10 Concluído', 5);
}

/**
 * Passo 11: Calcular Contagem de Passageiros Sentados
 */
function passo11_calcularContagemSentados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colVaziaIndex = -1, colSentadoIndex = -1, colEmPeIndex = -1, colCapSentadoIndex = -1, colCapTotalIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toLowerCase().trim(); // Convertendo para minúsculo para busca
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

    // --- REGRA PRIORITÁRIA: Verificação de CV (em minúsculo ou maiúsculo) ---
    // Checa tanto na coluna 'vazia' quanto na coluna 'sentado' (onde está o 7CV na sua imagem)
    if (vaziaStr.includes("CV") || sentadoStr.includes("CV")) {
      let textoParaExtrair = vaziaStr.includes("CV") ? vaziaStr : sentadoStr;
      let numVazias = parseFloat(textoParaExtrair.replace("CV", "").trim());
      if (!isNaN(numVazias)) {
        res = Math.max(0, capSentadoNum - numVazias); // Ex: 38 - 7 = 31
      } else {
        res = 0;
      }
    }
    // 1. Se VAZIA for número direto (subtrai da capacidade)
    else if (vaziaStr !== "" && !isNaN(parseFloat(vaziaStr)) && parseFloat(vaziaStr) > 0) {
      res = Math.max(0, capSentadoNum - parseFloat(vaziaStr)); 
    }
    // 2. Se "Em Pé" for 0
    else if (emPeStr === "0") {
      res = 0;
    }
    // 3. Se "Em Pé" tiver carga (LT, SL ou > 0), bancos estão cheios
    else if (emPeStr === "LT" || emPeStr === "SL" || (!isNaN(parseFloat(emPeStr)) && parseFloat(emPeStr) > 0)) {
      res = capSentadoNum;
    }
    // 4. Regra para "BC" (Banco Cheio)
    else if (sentadoStr === "BC") {
      res = capSentadoNum;
    } 
    // 5. Valor numérico direto na coluna "Sentado" (Apenas se não for CV)
    else if (sentadoStr !== "" && !isNaN(parseFloat(sentadoStr))) {
      res = parseFloat(sentadoStr);
    }
    else {
      res = 0;
    }
    
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('Cálculo corrigido: 7CV agora resulta em ' + (resultados[0] ? resultados[0][0] : ''), 'Sucesso', 5);
}

/**
 * Passo 12: Calcular Contagem de Passageiros em Pé
 */
function passo12_calcularContagemPe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContSentadosIndex = -1;
  let colEmPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_SENTADOS') colContSentadosIndex = i + 1;
    if (tit === 'EM PÉ' || tit === 'EM PE') colEmPeIndex = i + 1;
  }

  if (colContSentadosIndex === -1 || colEmPeIndex === -1) {
    ui.alert('Erro', 'Colunas "contagem_pass_sentados" e/ou "EM PÉ" não encontradas. Execute o Passo 11 primeiro.', ui.ButtonSet.OK);
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
    let res = "";

    // Lógica diferenciada para cada sigla
    if (valorOriginal === "LT") {
      res = "LOTADO";
    } 
    else if (valorOriginal === "SL") {
      res = "SUPERLOTADO";
    } 
    else if (valorOriginal === "" || isNaN(parseFloat(valorOriginal))) {
      res = 0;
    } 
    else {
      res = parseFloat(valorOriginal);
    }
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna contagem_pass_pe foi calculada com sucesso.', 'Passo 12 Concluído', 5);
}

/**
 * Passo 13: Calcular Contagem Total (Sentados + Em Pé)
 */
function passo13_calcularContagemTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContSentadosIndex = -1;
  let colContPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_SENTADOS') colContSentadosIndex = i + 1;
    if (tit === 'CONTAGEM_PASS_PE' || tit === 'CONTAGEM_PASS_PÉ') colContPeIndex = i + 1;
  }

  if (colContSentadosIndex === -1 || colContPeIndex === -1) {
    ui.alert('Erro', 'Colunas de contagem (sentados/pé) não encontradas. Execute os passos 11 e 12.', ui.ButtonSet.OK);
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
    let res = "";

    // Se qualquer um for SUPERLOTADO, o total é SUPERLOTADO
    if (s === "SUPERLOTADO" || p === "SUPERLOTADO") {
      res = "SL";
    } 
    // Se for LOTADO (e não tiver superlotado), o total é LOTADO
    else if (s === "LOTADO" || p === "LOTADO") {
      res = "L";
    }
    // Se forem números, faz a soma normal
    else {
      let numS = parseFloat(s) || 0;
      let numP = parseFloat(p) || 0;
      res = numS + numP;
    }
    resultados.push([res]);
  }

  abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1).setValues(resultados);
  ss.toast('A coluna contagem_pass_total foi calculada com sucesso.', 'Passo 13 Concluído', 5);
}

/**
 * Passo 14: Calcular Taxa de Ocupação dos Sentados
 */
function passo14_calcularTaxaOcupSentado() {
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
    ui.alert('Erro', 'Faltam colunas estruturais ("contagem_pass_total", "contagem_pass_sentados", "cap_pass_sentado").', ui.ButtonSet.OK);
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
      // Mantendo a lógica: multiplica por 100 para o valor percentual real
      if (!isNaN(numContSentados) && !isNaN(numCapSentado) && numCapSentado > 0) resultados.push([(numContSentados / numCapSentado) * 100]); 
      else resultados.push([""]);
    }
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados); 
  
  // Alterado para 2 casas decimais (0.00)
  rangeDestino.setNumberFormat("0.00"); 
  
  ss.toast('A coluna taxa_ocup_sentado foi gerada.', 'Passo 14 Concluído', 5);
}

/**
 * Passo 15: Calcular Taxa de Ocupação em Pé (%)
 */
function passo15_calcularTaxaOcupPe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContPeIndex = -1;
  let colCapPeIndex = -1;
  let colTaxaSentadoIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_PE' || tit === 'CONTAGEM_PASS_PÉ') colContPeIndex = i + 1;
    if (tit === 'CAP_PASS_PE' || tit === 'CAP_PASS_PÉ') colCapPeIndex = i + 1;
    if (tit === 'TAXA_OCUP_SENTADO') colTaxaSentadoIndex = i + 1;
  }

  if (colContPeIndex === -1 || colCapPeIndex === -1 || colTaxaSentadoIndex === -1) {
    ui.alert('Erro', 'Colunas necessárias não encontradas. Verifique os passos 9, 12 e 14.', ui.ButtonSet.OK);
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
    let res = "";

    if (cont === "SUPERLOTADO" || cont === "LOTADO") {
      res = cont;
    } 
    else {
      let numCont = parseFloat(cont) || 0;
      if (!isNaN(cap) && cap > 0) {
        // Alterado: Multiplicado por 100
        res = (numCont / cap) * 100; 
      } else if (numCont > 0 && (isNaN(cap) || cap === 0)) {
        res = "SUPERLOTADO";
      } else {
        res = 0;
      }
    }
    resultados.push([res]);
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados);
  
  // Alterado: Formato 0.00 (duas casas decimais sem %)
  rangeDestino.setNumberFormat("0.00");
  
  ss.toast('A taxa de ocupação em pé foi calculada.', 'Passo 15 Concluído', 5);
}


/**
 * Passo 16: Calcular Taxa de Ocupação Total (%)
 */
function passo16_calcularTaxaOcupTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  if (!abaBD) { ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK); return; }

  abaBD.activate();
  const ultLinhaBD = abaBD.getLastRow(); const ultColBD = abaBD.getLastColumn(); if (ultLinhaBD <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultColBD).getValues()[0];
  let colContTotalIndex = -1;
  let colCapTotalIndex = -1;
  let colTaxaPeIndex = -1;

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    if (tit === 'CONTAGEM_PASS_TOTAL') colContTotalIndex = i + 1;
    if (tit === 'CAP_PASS_TOTAL') colCapTotalIndex = i + 1;
    if (tit === 'TAXA_OCUP_PE' || tit === 'TAXA_OCUP_PÉ') colTaxaPeIndex = i + 1;
  }

  if (colContTotalIndex === -1 || colCapTotalIndex === -1 || colTaxaPeIndex === -1) {
    ui.alert('Erro', 'Colunas necessárias não encontradas. Verifique os passos 10, 13 e 15.', ui.ButtonSet.OK);
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
    let res = "";

    if (cont === "SUPERLOTADO" || cont === "LOTADO") {
      res = cont;
    } 
    else {
      let numCont = parseFloat(cont) || 0;
      if (!isNaN(cap) && cap > 0) {
        // Alterado: Multiplicado por 100
        res = (numCont / cap) * 100;
      } else if (numCont > 0 && (isNaN(cap) || cap === 0)) {
        res = "SUPERLOTADO";
      } else {
        res = 0;
      }
    }
    resultados.push([res]);
  }

  const rangeDestino = abaBD.getRange(2, colunaDestinoIndex, ultLinhaBD - 1, 1);
  rangeDestino.setValues(resultados);
  
  // Alterado: Formato 0.00 (duas casas decimais sem %)
  rangeDestino.setNumberFormat("0.00");
  
  ss.toast('A taxa de ocupação total foi calculada com sucesso.', 'Passo 16 Concluído', 5);
}

/**
 * Passo 17: Formatar a Aba BD inteira
 */
function passo17_formatarAbaBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  
  if (!abaBD) {
    ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK);
    return;
  }

  abaBD.activate();
  const ultimaLinha = abaBD.getLastRow();
  const ultimaColuna = abaBD.getLastColumn();
  
  if (ultimaLinha <= 1) return;

  const cabecalhos = abaBD.getRange(1, 1, 1, ultimaColuna).getValues()[0];

  // Identificadores de formatação para o Sheets
  const formatoTexto = "@";
  const formatoData = "dd/MM/yyyy";
  const formatoNumero = "0";
  const formatoHora = "HH:mm";
  const formatoPorcentagem = "0.00%";

  for (let i = 0; i < cabecalhos.length; i++) {
    let tit = cabecalhos[i].toString().toUpperCase().trim();
    let colIndex = i + 1;
    let rangeColuna = abaBD.getRange(2, colIndex, ultimaLinha - 1, 1);

    if (tit === 'DATA') { rangeColuna.setNumberFormat(formatoData); }
    else if (['PESQUISADORES', 'LOCAL', 'SENTIDO', 'EMPRESA', 'LINHA', 'VIA', 'CARRO', 'EMP_CARRO', 'FAIXA-HORÁRIA', 'TIPO_VEÍCULO', 'TIPO_VEICULO'].includes(tit)) {
      rangeColuna.setNumberFormat(formatoTexto);
    }
    else if (['Nº DA VIAGEM', 'VAZIA', 'SENTADO', 'EM PÉ', 'EM PE', 'CAP_PASS_SENTADO', 'CAP_PASS_PE', 'CAP_PASS_PÉ', 'CAP_PASS_TOTAL', 'CONTAGEM_PASS_SENTADOS', 'CONTAGEM_PASS_PE', 'CONTAGEM_PASS_PÉ', 'CONTAGEM_PASS_TOTAL'].includes(tit)) {
      rangeColuna.setNumberFormat(formatoNumero);
    }
    else if (tit === 'HORA') { rangeColuna.setNumberFormat(formatoHora); }
    else if (tit === 'INT_VIAGENS') { rangeColuna.setNumberFormat(formatoHora); }
    else if (['TAXA_OCUP_SENTADO', 'TAXA_OCUP_PE', 'TAXA_OCUP_PÉ', 'TAXA_OCUP_TOTAL'].includes(tit)) {
      rangeColuna.setNumberFormat(formatoPorcentagem);
    }
  }

  const rangeTotal = abaBD.getRange(1, 1, ultimaLinha, ultimaColuna);
  rangeTotal.setBorder(true, true, true, true, true, true);
  rangeTotal.setHorizontalAlignment("center");
  rangeTotal.setVerticalAlignment("middle");

  abaBD.getRange(1, 1, 1, ultimaColuna).setFontWeight("bold");

  ss.toast('Todos os dados foram formatados.', 'Passo 17 Concluído', 5);
}

/**
 * Passo 18: Criar Aba Tabelas de Análises
 */
function passo18_criarAbaAnalises() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const nomeAba = "tabelas_análises";

  let abaAnalises = ss.getSheetByName(nomeAba);

  if (abaAnalises) {
    const resposta = ui.alert(
      'Aviso', 
      'A aba "' + nomeAba + '" já existe. Deseja APAGÁ-LA e criar uma nova aba em branco? \n\n(Atenção: Isso apagará todas as tabelas e gráficos que já estiverem nela).', 
      ui.ButtonSet.YES_NO
    );
    
    if (resposta == ui.Button.YES) {
      ss.deleteSheet(abaAnalises);
      abaAnalises = ss.insertSheet(nomeAba);
    } else {
      abaAnalises.activate();
      return;
    }
  } else {
    abaAnalises = ss.insertSheet(nomeAba);
  }

  abaAnalises.activate();
  ss.toast('A aba ' + nomeAba + ' foi disponibilizada.', 'Passo 18 Concluído', 5);
}

/**
 * Passo 19: Criar Tabela de Análise (Viagens por Faixa Horária e Local)
 */
function passo19_analiseViagensLocal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");
  
  if (!abaBD) {
    ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK);
    return;
  }
  
  if (!abaAnalises) {
    abaAnalises = ss.insertSheet("tabelas_análises");
  }

  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) return;

  const cabecalhos = dadosBD[0];
  
  let colFaixaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'FAIXA-HORÁRIA');
  let colLinhaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA');
  let colLocalIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL');
  let colSentidoIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO');

  if (colFaixaIndex === -1 || colLocalIndex === -1 || colSentidoIndex === -1) {
    ui.alert('Erro', 'Colunas FAIXA-HORÁRIA, LOCAL ou SENTIDO não encontradas na aba BD.', ui.ButtonSet.OK);
    return;
  }

  let faixasUnicas = [];
  let colunasDinamicas = [];
  let matrizContagem = {}; 

  for (let i = 1; i < dadosBD.length; i++) {
    let faixa = dadosBD[i][colFaixaIndex].toString().trim();
    let local = dadosBD[i][colLocalIndex].toString().trim();
    let sentido = dadosBD[i][colSentidoIndex].toString().trim();
    let linhaBus = colLinhaIndex !== -1 ? dadosBD[i][colLinhaIndex].toString().trim() : "";
    
    if (!faixa || !local) continue; 

    let nomeColuna = "";
    if (linhaBus !== "") {
      nomeColuna = linhaBus + " | " + local + (sentido ? " - " + sentido : "");
    } else {
      nomeColuna = local + (sentido ? " - " + sentido : "");
    }

    if (!faixasUnicas.includes(faixa)) faixasUnicas.push(faixa);
    if (!colunasDinamicas.includes(nomeColuna)) colunasDinamicas.push(nomeColuna);

    if (!matrizContagem[faixa]) matrizContagem[faixa] = {};
    if (!matrizContagem[faixa][nomeColuna]) matrizContagem[faixa][nomeColuna] = 0;
    
    matrizContagem[faixa][nomeColuna]++;
  }

  faixasUnicas.sort();
  colunasDinamicas.sort();

  let tabelaFinal = [];
  
  let cabecalhoTabela = ["FAIXA-HORÁRIA"];
  for (let c of colunasDinamicas) cabecalhoTabela.push(c);
  tabelaFinal.push(cabecalhoTabela);

  let totaisPorColuna = new Array(colunasDinamicas.length).fill(0);

  for (let f of faixasUnicas) {
    let linhaTabela = [f]; 
    for (let i = 0; i < colunasDinamicas.length; i++) {
      let c = colunasDinamicas[i];
      let valor = matrizContagem[f][c] ? matrizContagem[f][c] : 0; 
      linhaTabela.push(valor);
      totaisPorColuna[i] += valor; 
    }
    tabelaFinal.push(linhaTabela);
  }

  let linhaTotal = ["Total geral"];
  for (let total of totaisPorColuna) {
    linhaTotal.push(total); 
  }
  tabelaFinal.push(linhaTotal);

  let ultimaLinhaAnalises = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinhaAnalises === 0 ? 2 : ultimaLinhaAnalises + 3; 

  abaAnalises.getRange(linhaInicio - 1, 1)
             .setValue("ANÁLISE 1: QUADRO DE VIAGENS POR FAIXA HORÁRIA E LOCAL")
             .setFontWeight("bold")
             .setFontSize(12);

  let rangeDestino = abaAnalises.getRange(linhaInicio, 1, tabelaFinal.length, tabelaFinal[0].length);
  rangeDestino.setValues(tabelaFinal);

  rangeDestino.setBorder(true, true, true, true, true, true);
  rangeDestino.setHorizontalAlignment("center");
  rangeDestino.setVerticalAlignment("middle");
  
  let rangeCabecalho = abaAnalises.getRange(linhaInicio, 1, 1, tabelaFinal[0].length);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setBackground("#d9ead3"); 
  rangeCabecalho.setWrap(true);

  let rangeTotalGeral = abaAnalises.getRange(linhaInicio + tabelaFinal.length - 1, 1, 1, tabelaFinal[0].length);
  rangeTotalGeral.setFontWeight("bold");

  abaAnalises.activate();
  ss.toast('Quadro de viagens com Total Geral gerado.', 'Análise 1 Concluída', 5);
}

/**
 * Passo 20: Criar Tabela de Análise Hierárquica (Taxa de Ocupação Total)
 */
function passo20_analiseTaxaOcupacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");
  
  if (!abaBD) {
    ui.alert('Erro', 'A aba "BD" não foi encontrada.', ui.ButtonSet.OK);
    return;
  }
  
  if (!abaAnalises) {
    abaAnalises = ss.insertSheet("tabelas_análises");
  }

  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) return;

  const cabecalhos = dadosBD[0];
  
  let colLinhaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA');
  let colLocalIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL');
  let colSentidoIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO');
  let colTaxaIndex = cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'TAXA_OCUP_TOTAL');

  if (colLocalIndex === -1 || colTaxaIndex === -1) {
    ui.alert('Erro', 'Colunas LOCAL ou TAXA_OCUP_TOTAL não encontradas na aba BD.', ui.ButtonSet.OK);
    return;
  }

  let relatorioObj = {};
  let totalGeralMatriz = [0, 0, 0, 0, 0, 0]; 
  
  for (let i = 1; i < dadosBD.length; i++) {
    // Busca a taxa
    let taxaStr = dadosBD[i][colTaxaIndex].toString().trim().toUpperCase();
    
    let linha = colLinhaIndex !== -1 ? dadosBD[i][colLinhaIndex].toString().trim() : "Linha Indefinida";
    let local = dadosBD[i][colLocalIndex].toString().trim();
    let sentido = colSentidoIndex !== -1 ? dadosBD[i][colSentidoIndex].toString().trim() : "";
    let nomeLocal = local + (sentido ? " - " + sentido : "");
    if (!linha) linha = "Linha Indefinida";

    let b = -1; 

    // Lógica para abraçar o texto string "-" inserido pelo passo 12
    if (taxaStr === "-") {
      b = 5; // Joga direto no bucket >115% Crítico
    } else if (taxaStr === "LOTADO") {
      b = 4; // Joga direto no bucket 100% - 115% Lotado
    } else {
      let taxaStrLimpa = taxaStr.replace('%', '').replace(',', '.');
      let taxaNum = parseFloat(taxaStrLimpa);
      
      if (!isNaN(taxaNum)) {
        if (taxaNum < 25) b = 0;
        else if (taxaNum < 50) b = 1;
        else if (taxaNum < 75) b = 2;
        else if (taxaNum < 100) b = 3;
        else if (taxaNum <= 115) b = 4;
        else b = 5;
      }
    }

    if (b !== -1) {
      if (!relatorioObj[linha]) {
        relatorioObj[linha] = { total: [0,0,0,0,0,0], locais: {} };
      }
      if (!relatorioObj[linha].locais[nomeLocal]) {
        relatorioObj[linha].locais[nomeLocal] = [0,0,0,0,0,0];
      }

      relatorioObj[linha].total[b]++;
      relatorioObj[linha].locais[nomeLocal][b]++;
      totalGeralMatriz[b]++;
    }
  }

  function transformarEmPercentuais(counts) {
    let somaTratada = counts.reduce((a, b) => a + b, 0);
    if (somaTratada === 0) return [0, 0, 0, 0, 0, 0, 0];
    let percentuais = counts.map(c => c / somaTratada);
    percentuais.push(1); 
    return percentuais;
  }

  let tabelaFinal = [];
  let cabecalhoTabela = [
    "Linhas / Local (Parada)", 
    "<25%", 
    "25% - 49%", 
    "50% - 74%", 
    "75% - 99%", 
    "100% - 115% (Lotado)", 
    ">115% (Crítico)", 
    "Total Geral"
  ];
  tabelaFinal.push(cabecalhoTabela);

  let linhasOrdenadas = Object.keys(relatorioObj).sort();
  let indicesNegrito = []; 

  for (let l of linhasOrdenadas) {
    let arrayLinhaMaster = transformarEmPercentuais(relatorioObj[l].total);
    tabelaFinal.push([l].concat(arrayLinhaMaster));
    indicesNegrito.push(tabelaFinal.length); 
    
    let locaisOrdenados = Object.keys(relatorioObj[l].locais).sort();
    for (let loc of locaisOrdenados) {
      let arrayLocal = transformarEmPercentuais(relatorioObj[l].locais[loc]);
      tabelaFinal.push(["   " + loc].concat(arrayLocal)); 
    }
  }

  let arrayTotalGeral = transformarEmPercentuais(totalGeralMatriz);
  tabelaFinal.push(["Total Geral"].concat(arrayTotalGeral));
  indicesNegrito.push(tabelaFinal.length); 

  let ultimaLinhaAnalises = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinhaAnalises === 0 ? 2 : ultimaLinhaAnalises + 3; 

  abaAnalises.getRange(linhaInicio - 1, 1)
             .setValue("ANÁLISE 2: TAXA DE OCUPAÇÃO TOTAL POR LINHA E PARADA")
             .setFontWeight("bold")
             .setFontSize(12);

  let rangeDestino = abaAnalises.getRange(linhaInicio, 1, tabelaFinal.length, tabelaFinal[0].length);
  rangeDestino.setValues(tabelaFinal);

  rangeDestino.setBorder(true, true, true, true, true, true);
  rangeDestino.setHorizontalAlignment("center");
  rangeDestino.setVerticalAlignment("middle");
  
  let rangeDadosPercentuais = abaAnalises.getRange(linhaInicio + 1, 2, tabelaFinal.length - 1, 7);
  rangeDadosPercentuais.setNumberFormat("0%");

  abaAnalises.getRange(linhaInicio + 1, 1, tabelaFinal.length - 1, 1).setHorizontalAlignment("left");

  let rangeCabecalho = abaAnalises.getRange(linhaInicio, 1, 1, tabelaFinal[0].length);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setBackground("#efefef"); 
  rangeCabecalho.setWrap(true);

  for (let idx of indicesNegrito) {
    abaAnalises.getRange(linhaInicio + idx - 1, 1, 1, tabelaFinal[0].length).setFontWeight("bold");
  }

  abaAnalises.activate();
  ss.toast('Ocupação hierárquica gerada com sucesso!', 'Análise 2 Concluída', 5);
}

/**
 * Funções auxiliares matemáticas
 */
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
/**
 * Passo 21: Análise interativa por Local, Sentido, Linha e PED
 * - Pergunta ao usuário se deseja detectar automaticamente ou informar manualmente
 * - Cria tabelas separadas para cada combinação (Local + Sentido + PED + Linha)
 * - Tabelas contêm: faixa horária, Qtd Dia1, Qtd Dia2, Média, Qtd >=100%, % >=100%
 */
function passo21_analisePorLocalSentidoPED() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const abaBD = ss.getSheetByName("BD");
  let abaAnalises = ss.getSheetByName("tabelas_análises");

  // Validação da aba BD
  if (!abaBD) {
    ui.alert("Erro", "A aba 'BD' não foi encontrada.", ui.ButtonSet.OK);
    return;
  }

  // Cria a aba de análises se não existir
  if (!abaAnalises) {
    abaAnalises = ss.insertSheet("tabelas_análises");
  }

  // Obtém dados da BD
  const dadosBD = abaBD.getDataRange().getDisplayValues();
  if (dadosBD.length <= 1) {
    ui.alert("Erro", "A aba 'BD' não contém dados.", ui.ButtonSet.OK);
    return;
  }

  const cabecalhos = dadosBD[0];

  // Mapeamento das colunas necessárias
  const colunas = {
    faixa: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'FAIXA-HORÁRIA'),
    local: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LOCAL'),
    sentido: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'SENTIDO'),
    linha: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'LINHA'),
    ped: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'PED'),
    data: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'DATA'),
    taxaTotal: cabecalhos.findIndex(c => c.toString().toUpperCase().trim() === 'TAXA_OCUP_TOTAL')
  };

  // Verifica se as colunas essenciais existem
  if (colunas.faixa === -1 || colunas.local === -1 || colunas.taxaTotal === -1) {
    ui.alert("Erro", "Colunas essenciais (FAIXA-HORÁRIA, LOCAL, TAXA_OCUP_TOTAL) não encontradas na BD.", ui.ButtonSet.OK);
    return;
  }

  // Pergunta ao usuário o modo de operação
  const respostaModo = ui.alert(
    "Modo de Análise",
    "Deseja que o script DETECTE AUTOMATICAMENTE todas as combinações de Local, Sentido, Linha e PED?\n\n" +
    "Clique em 'SIM' para detectar automaticamente.\n" +
    "Clique em 'NÃO' para informar manualmente cada filtro.",
    ui.ButtonSet.YES_NO
  );

  let combinacoes = [];

  if (respostaModo === ui.Button.YES) {
    // --- MODO AUTOMÁTICO: detecta todas as combinações únicas ---
    combinacoes = detectarCombinacoesUnicas(dadosBD, colunas);
    if (combinacoes.length === 0) {
      ui.alert("Aviso", "Nenhuma combinação válida encontrada para gerar tabelas.", ui.ButtonSet.OK);
      return;
    }
    ui.alert(`Detectadas ${combinacoes.length} combinações. A criação das tabelas pode levar alguns segundos.`, ui.ButtonSet.OK);
  } else {
    // --- MODO MANUAL: coleta os filtros interativamente ---
    combinacoes = coletarFiltrosManualmente(ui, dadosBD, colunas);
    if (combinacoes.length === 0) return;
  }

  // Obtém os dias únicos disponíveis na coluna DATA
  const diasUnicos = obterDiasUnicos(dadosBD, colunas.data);

  // Para cada combinação, gera uma tabela
  for (let idx = 0; idx < combinacoes.length; idx++) {
    const combo = combinacoes[idx];
    const tituloTabela = montarTituloTabela(combo);
    const dadosFiltrados = filtrarDadosPorCombinacao(dadosBD, colunas, combo);
    const tabela = construirTabelaPorFaixa(dadosFiltrados, diasUnicos, colunas);
    
    if (tabela && tabela.length > 1) {
      inserirTabelaNaAbaAnalises(abaAnalises, tituloTabela, tabela);
    }
  }

  abaAnalises.activate();
  ss.toast(`${combinacoes.length} tabela(s) criada(s) com sucesso!`, "Passo 21 Concluído", 8);
}

// ==================== FUNÇÕES AUXILIARES ====================

/**
 * Detecta automaticamente todas as combinações únicas de Linha, Local, Sentido e PED
 */
function detectarCombinacoesUnicas(dadosBD, colunas) {
  const combinacoesSet = new Set();
  
  for (let i = 1; i < dadosBD.length; i++) {
    const linha = colunas.linha !== -1 ? dadosBD[i][colunas.linha].toString().trim() : "";
    const local = dadosBD[i][colunas.local].toString().trim();
    const sentido = colunas.sentido !== -1 ? dadosBD[i][colunas.sentido].toString().trim() : "";
    const ped = colunas.ped !== -1 ? dadosBD[i][colunas.ped].toString().trim() : "";
    
    if (!local) continue;
    
    const chave = `${linha}|${local}|${sentido}|${ped}`;
    if (!combinacoesSet.has(chave)) {
      combinacoesSet.add(chave);
    }
  }
  
  const combinacoes = [];
  for (const chave of combinacoesSet) {
    const [linha, local, sentido, ped] = chave.split("|");
    combinacoes.push({ linha, local, sentido, ped });
  }
  
  return combinacoes;
}

/**
 * Coleta os filtros manualmente através de diálogos
 */
function coletarFiltrosManualmente(ui, dadosBD, colunas) {
  const combinacoes = [];
  
  // Pergunta se há mais de uma linha
  const temMultiplasLinhas = perguntarMaisDeUm(ui, "Há mais de uma LINHA pesquisada?");
  let linhas = [];
  if (temMultiplasLinhas) {
    const linhasDisponiveis = obterValoresUnicos(dadosBD, colunas.linha);
    if (linhasDisponiveis.length === 0) {
      ui.alert("Nenhuma linha encontrada nos dados.", ui.ButtonSet.OK);
      return [];
    }
    const resposta = ui.prompt("Linhas disponíveis", 
      `Linhas encontradas: ${linhasDisponiveis.join(", ")}\n\nDigite a LINHA desejada:`, 
      ui.ButtonSet.OK_CANCEL);
    if (resposta.getSelectedButton() !== ui.Button.OK) return [];
    linhas = [resposta.getResponseText().trim()];
  } else {
    linhas = [""];
  }
  
  // Pergunta se há mais de um local
  const temMultiplosLocais = perguntarMaisDeUm(ui, "Há mais de um LOCAL de pesquisa?");
  let locais = [];
  if (temMultiplosLocais) {
    const locaisDisponiveis = obterValoresUnicos(dadosBD, colunas.local);
    if (locaisDisponiveis.length === 0) {
      ui.alert("Nenhum local encontrado nos dados.", ui.ButtonSet.OK);
      return [];
    }
    const resposta = ui.prompt("Locais disponíveis", 
      `Locais encontrados: ${locaisDisponiveis.join(", ")}\n\nDigite o LOCAL desejado:`, 
      ui.ButtonSet.OK_CANCEL);
    if (resposta.getSelectedButton() !== ui.Button.OK) return [];
    locais = [resposta.getResponseText().trim()];
  } else {
    locais = [""];
  }
  
  // Pergunta se há mais de um sentido
  const temMultiplosSentidos = perguntarMaisDeUm(ui, "Há mais de um SENTIDO pesquisado?");
  let sentidos = [];
  if (temMultiplosSentidos && colunas.sentido !== -1) {
    const sentidosDisponiveis = obterValoresUnicos(dadosBD, colunas.sentido);
    if (sentidosDisponiveis.length === 0) {
      ui.alert("Nenhum sentido encontrado nos dados.", ui.ButtonSet.OK);
      return [];
    }
    const resposta = ui.prompt("Sentidos disponíveis", 
      `Sentidos encontrados: ${sentidosDisponiveis.join(", ")}\n\nDigite o SENTIDO desejado:`, 
      ui.ButtonSet.OK_CANCEL);
    if (resposta.getSelectedButton() !== ui.Button.OK) return [];
    sentidos = [resposta.getResponseText().trim()];
  } else {
    sentidos = [""];
  }
  
  // Pergunta se o PED foi informado
  const pedInformado = perguntarMaisDeUm(ui, "O PED foi informado?");
  let peds = [];
  if (pedInformado && colunas.ped !== -1) {
    const pedsDisponiveis = obterValoresUnicos(dadosBD, colunas.ped);
    if (pedsDisponiveis.length === 0) {
      ui.alert("Nenhum PED encontrado nos dados.", ui.ButtonSet.OK);
      return [];
    }
    const resposta = ui.prompt("PEDs disponíveis", 
      `PEDs encontrados: ${pedsDisponiveis.join(", ")}\n\nDigite o PED desejado:`, 
      ui.ButtonSet.OK_CANCEL);
    if (resposta.getSelectedButton() !== ui.Button.OK) return [];
    peds = [resposta.getResponseText().trim()];
  } else {
    peds = [""];
  }
  
  for (const linha of linhas) {
    for (const local of locais) {
      for (const sentido of sentidos) {
        for (const ped of peds) {
          if (local) {
            combinacoes.push({ linha, local, sentido, ped });
          }
        }
      }
    }
  }
  
  return combinacoes;
}

/**
 * Função auxiliar para perguntar "Sim/Não"
 */
function perguntarMaisDeUm(ui, pergunta) {
  const resposta = ui.alert(pergunta, "Clique em SIM ou NÃO.", ui.ButtonSet.YES_NO);
  return resposta === ui.Button.YES;
}

/**
 * Obtém valores únicos de uma coluna (ignorando cabeçalho)
 */
function obterValoresUnicos(dadosBD, colunaIndex) {
  if (colunaIndex === -1) return [];
  const valores = new Set();
  for (let i = 1; i < dadosBD.length; i++) {
    const valor = dadosBD[i][colunaIndex].toString().trim();
    if (valor) valores.add(valor);
  }
  return Array.from(valores);
}

/**
 * Obtém os dias únicos de pesquisa (coluna DATA)
 */
function obterDiasUnicos(dadosBD, colunaDataIndex) {
  if (colunaDataIndex === -1) return [1];
  const dias = new Set();
  for (let i = 1; i < dadosBD.length; i++) {
    const dataStr = dadosBD[i][colunaDataIndex].toString().trim();
    if (dataStr) dias.add(dataStr);
  }
  const diasArray = Array.from(dias);
  return diasArray.length > 0 ? diasArray : [1];
}

/**
 * Monta o título da tabela baseado na combinação
 */
function montarTituloTabela(combo) {
  let titulo = "";
  if (combo.linha) titulo += `${combo.linha} - `;
  titulo += combo.local;
  if (combo.sentido) titulo += ` - ${combo.sentido}`;
  if (combo.ped) titulo += ` - PED ${combo.ped}`;
  return titulo;
}

/**
 * Filtra os dados da BD com base na combinação fornecida
 */
function filtrarDadosPorCombinacao(dadosBD, colunas, combo) {
  const filtrados = [];
  
  for (let i = 1; i < dadosBD.length; i++) {
    const linhaVal = colunas.linha !== -1 ? dadosBD[i][colunas.linha].toString().trim() : "";
    const localVal = dadosBD[i][colunas.local].toString().trim();
    const sentidoVal = colunas.sentido !== -1 ? dadosBD[i][colunas.sentido].toString().trim() : "";
    const pedVal = colunas.ped !== -1 ? dadosBD[i][colunas.ped].toString().trim() : "";
    
    let match = true;
    if (combo.linha && linhaVal !== combo.linha) match = false;
    if (combo.local && localVal !== combo.local) match = false;
    if (combo.sentido && sentidoVal !== combo.sentido) match = false;
    if (combo.ped && pedVal !== combo.ped) match = false;
    
    if (match) {
      filtrados.push(dadosBD[i]);
    }
  }
  
  return filtrados;
}

/**
 * Constrói a tabela com faixas horárias, contagens por dia, média, >=100% e %
 */
function construirTabelaPorFaixa(dadosFiltrados, diasUnicos, colunas) {
  if (dadosFiltrados.length === 0) return [];
  
  // Agrupa por faixa horária
  const mapaFaixas = new Map();
  
  for (const linha of dadosFiltrados) {
    const faixa = linha[colunas.faixa].toString().trim();
    if (!faixa) continue;
    
    const data = colunas.data !== -1 ? linha[colunas.data].toString().trim() : "";
    const taxaStr = linha[colunas.taxaTotal].toString().trim().toUpperCase();
    
    // Determina se a ocupação é >=100% (inclui "-" e "LOTADO" e valores >=100)
    let ocupacaoAlta = false;
    if (taxaStr === "-" || taxaStr === "LOTADO") {
      ocupacaoAlta = true;
    } else {
      const taxaNum = parseFloat(taxaStr.replace("%", "").replace(",", "."));
      if (!isNaN(taxaNum) && taxaNum >= 100) {
        ocupacaoAlta = true;
      }
    }
    
    if (!mapaFaixas.has(faixa)) {
      mapaFaixas.set(faixa, {
        totalViagens: 0,
        porDia: new Map(),
        ocupacaoAltaCount: 0
      });
    }
    
    const grupo = mapaFaixas.get(faixa);
    grupo.totalViagens++;
    
    // Contagem por dia
    const chaveDia = data || "Dia 1";
    grupo.porDia.set(chaveDia, (grupo.porDia.get(chaveDia) || 0) + 1);
    
    if (ocupacaoAlta) {
      grupo.ocupacaoAltaCount++;
    }
  }
  
  // Ordena as faixas cronologicamente (ex: "00:00 - 00:59", "01:00 - 01:59"...)
  const faixasOrdenadas = Array.from(mapaFaixas.keys()).sort((a, b) => {
    const horaA = parseInt(a.split(":")[0], 10);
    const horaB = parseInt(b.split(":")[0], 10);
    return horaA - horaB;
  });
  
  // Prepara os dias disponíveis (até 2 dias, conforme solicitado)
  const diasDisponiveis = diasUnicos.slice(0, 2);
  while (diasDisponiveis.length < 2) diasDisponiveis.push(`Dia ${diasDisponiveis.length + 1}`);
  
  // Monta a tabela
  const tabela = [];
  const cabecalho = ["Faixa Horária", `Qtd ${diasDisponiveis[0]}`, `Qtd ${diasDisponiveis[1]}`, "Média", "Qtd >=100%", "% >=100%"];
  tabela.push(cabecalho);
  
  for (const faixa of faixasOrdenadas) {
    const grupo = mapaFaixas.get(faixa);
    const qtdDia1 = grupo.porDia.get(diasDisponiveis[0]) || 0;
    const qtdDia2 = diasDisponiveis[1] ? (grupo.porDia.get(diasDisponiveis[1]) || 0) : 0;
    const media = diasDisponiveis[1] ? (qtdDia1 + qtdDia2) / 2 : qtdDia1;
    const qtdAlta = grupo.ocupacaoAltaCount;
    const percentualAlta = grupo.totalViagens > 0 ? (qtdAlta / grupo.totalViagens) : 0;
    
    tabela.push([
      faixa,
      qtdDia1,
      diasDisponiveis[1] ? qtdDia2 : "",
      media.toFixed(1),
      qtdAlta,
      percentualAlta
    ]);
  }
  
  return tabela;
}

/**
 * Insere a tabela na aba de análises, pulando 4 células da última linha com conteúdo
 */
function inserirTabelaNaAbaAnalises(abaAnalises, titulo, tabela) {
  let ultimaLinha = abaAnalises.getLastRow();
  let linhaInicio = ultimaLinha === 0 ? 2 : ultimaLinha + 4;
  
  // Insere o título
  abaAnalises.getRange(linhaInicio - 1, 1)
    .setValue(titulo)
    .setFontWeight("bold")
    .setFontSize(12);
  
  // Insere a tabela
  const rangeDestino = abaAnalises.getRange(linhaInicio, 1, tabela.length, tabela[0].length);
  rangeDestino.setValues(tabela);
  rangeDestino.setBorder(true, true, true, true, true, true);
  rangeDestino.setHorizontalAlignment("center");
  rangeDestino.setVerticalAlignment("middle");
  
  // Formata cabeçalho
  const rangeCabecalho = abaAnalises.getRange(linhaInicio, 1, 1, tabela[0].length);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setBackground("#d9ead3");
  
  // Formata coluna de percentual
  const ultimaColuna = tabela[0].length;
  const rangePercentual = abaAnalises.getRange(linhaInicio + 1, ultimaColuna, tabela.length - 1, 1);
  rangePercentual.setNumberFormat("0.00%");
  
  // Ajusta largura das colunas
  for (let i = 1; i <= tabela[0].length; i++) {
    abaAnalises.autoResizeColumn(i);
  }
}
}

