async function buscarSituacao() {
  var p = SpreadsheetApp.getActiveSpreadsheet();
  var base = p.getSheetByName('Base - Situação do pedido')
  var ln = ultimolinhaColuna('Registrado', 1, base)
  if (base.getRange(ln, 1).getValue() != "") {

    var solicitante = base.getRange(ln, 2).getValue()
    var marina = base.getRange(ln, 3).getValue()
    var solicitacao = base.getRange(ln, 4).getValue()
    var descricao = base.getRange(ln, 5).getValue()
    if (base.getRange(ln, 6).getValue() == "EM COTAO") { base.getRange(ln, 6).setValue("EM COTAÇÃO") }
    var situacao = base.getRange(ln, 6).getValue()
    var comprador = base.getRange(ln, 7).getValue()
    var previsao = base.getRange(ln, 8).getValue()
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    const now = new Date();
    const from = new Date(now.getTime() + 7 * MILLIS_PER_DAY);
    var hoje = Utilities.formatDate(new Date(), "GMT-03:00", "dd/MM/yyyy");
    var validade = Utilities.formatDate(from, "GMT-03:00", "dd/MM/yyyy");
    // Solicitante	Marina	Solicitação	Descrição	Situação	Comprador	Previsao
    if (marina == "Verolme") {

      var sitmarina = p.getSheetByName('Sit_Verolme')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')

    }

    if (marina == "Piratas") {
      var sitmarina = p.getSheetByName('Sit_Piratas')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Ribeira") {
      var sitmarina = p.getSheetByName('Sit_Ribeira')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Bracuhy") {
      var sitmarina = p.getSheetByName('Sit_Bracuhy')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Itacuruca") {
      var sitmarina = p.getSheetByName('Sit_Itacuruca')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Gloria") {
      var sitmarina = p.getSheetByName('Sit_Gloria')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Paraty") {
      var sitmarina = p.getSheetByName('Sit_Paraty')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
    if (marina == "Buzios") {
      var sitmarina = p.getSheetByName('Sit_Buzios')
      for (let i = 2; i <= sitmarina.getLastRow(); i++) {
        if (sitmarina.getRange(i, 8).getValue() != "") {
          if (sitmarina.getRange(i, 8).getValue().getTime() < now.getTime() == true) {
            sitmarina.deleteRows(i, sitmarina.getRange(i, 8).getNumRows());
          }
        }
      }
      var l = 0;
      if (sitmarina.getLastRow() > 1) { l = 1 }
      var dados = sitmarina.getRange(2, 1, sitmarina.getLastRow() - l, 7).getValues();
      for (let i = 0; i < dados.length; i++) {
        if (marina + solicitacao == dados[i][1] + dados[i][2]) {
          sitmarina.getRange(2 + i, 5).setValue(situacao) // Atualizando a situação
          if (situacao == "COMPRADO") { sitmarina.getRange(2 + i, 7).setValue(previsao) }
          if (situacao == "ENTREGUE") { sitmarina.getRange(2 + i, 8).setValue(validade) }
          base.getRange(ln, 9).setValue('OK')
          return;
        }
      }
      var linha = sitmarina.getLastRow() + 1
      // Registrando quando não encontrar o pedido.
      sitmarina.getRange(linha, 1).setValue(solicitante)
      sitmarina.getRange(linha, 2).setValue(marina)
      sitmarina.getRange(linha, 3).setValue(solicitacao)
      sitmarina.getRange(linha, 4).setValue(descricao)
      sitmarina.getRange(linha, 5).setValue(situacao)
      sitmarina.getRange(linha, 6).setValue(comprador)
      if (situacao == "COMPRADO") { sitmarina.getRange(linha, 7).setValue(previsao) }
      if (situacao == "ENTREGUE") { sitmarina.getRange(linha, 8).setValue(validade) }
      base.getRange(ln, 9).setValue('OK')
    }
  }

}

// Pegar a ultima linha vazia de uma coluna especifica
function ultimolinhaColuna(x, y, z) {
  //exemplo: indiceColuna("texto a procurar","na linha","na Planilha")
  let index = z.getDataRange().getValues()[y - 1].indexOf(x);
  while (z.getRange(y, (index + 1)).getValue() != "") {
    y++
  }
  return y
}
