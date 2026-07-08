/**
 * DOJO DA TURMA — Script de integração com a planilha (v2)
 * ---------------------------------------------------------
 * Recebe os lançamentos do aplicativo e grava:
 *
 *  1. Na aba "Registros" — histórico completo (uma linha por lançamento);
 *  2. Na ABA DA TURMA — a nota entra na linha do estudante e na coluna
 *     do critério avaliativo (as mesmas colunas já existentes na planilha:
 *     "Atividades gerais / Nota", "Observação direta / Nota",
 *     "Projeto / ABA - Farmar Aura", "Cartazes / Nota", "Júri simulado /
 *     Nota", "Tarefa Bloco 1..5 / Nota", "Prova Paulista / Nota", etc.);
 *  3. Na aba "Resumo" — média por estudante e disciplina.
 *
 * Como instalar: veja o arquivo INSTRUCOES.md
 */

var CABECALHO = ["Data/Hora", "Disciplina", "Turma", "Estudante", "Critério",
                 "Tipo", "Valor/Nota", "Comentário", "Bimestre", "Professor",
                 "Lançado na aba da turma", "ID do app"];

/** Recebe os lançamentos enviados pelo aplicativo (POST). */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var dados = JSON.parse(e.postData.contents);
    var registros = dados.registros || [];
    var exclusoes = dados.exclusoes || [];
    var estruturas = dados.estruturas || [];
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaLog = obterAbaRegistros(planilha);

    // Alterações estruturais (novas turmas/disciplinas, estudantes etc.)
    // são aplicadas ANTES dos lançamentos, para que a aba nova já exista.
    var estruturasOk = [];
    if (estruturas.length > 0) estruturasOk = processarEstruturas(planilha, estruturas);

    // Exclusões pedidas pelo app (antes de gravar os novos lançamentos,
    // para que uma edição = exclusão + novo valor termine com o valor novo)
    var excluidos = 0;
    if (exclusoes.length > 0) excluidos = processarExclusoes(planilha, abaLog, exclusoes);

    // Evita duplicar lançamentos já gravados (caso o app reenvie)
    var idsExistentes = {};
    var ultimaLinha = abaLog.getLastRow();
    if (ultimaLinha > 1) {
      abaLog.getRange(2, CABECALHO.length, ultimaLinha - 1, 1).getValues()
            .forEach(function (l) { if (l[0]) idsExistentes[l[0]] = true; });
    }

    var novos = [];
    registros.forEach(function (r) {
      if (r.id && idsExistentes[r.id]) return; // já está na planilha

      // Tenta lançar direto na aba da turma (linha do estudante × coluna do critério)
      var situacao = lancarNaAbaDaTurma(planilha, r);

      novos.push([r.dataHora || "", r.disciplina || "", r.turma || "",
                  r.estudante || "", r.criterio || "", r.tipo || "",
                  (r.tipo === "Nota" ? Number(r.valor) : String(r.valor || "")),
                  r.observacao || "", r.bimestre || "", r.professor || "",
                  situacao, r.id || ""]);
    });

    if (novos.length > 0) {
      abaLog.getRange(abaLog.getLastRow() + 1, 1, novos.length, CABECALHO.length)
            .setValues(novos);
    }
    if (novos.length > 0 || excluidos > 0) atualizarResumo(planilha, abaLog);

    return resposta({ ok: true, gravados: novos.length, excluidos: excluidos,
                      estruturasOk: estruturasOk });
  } catch (erro) {
    return resposta({ ok: false, erro: String(erro) });
  } finally {
    lock.releaseLock();
  }
}

/** GET — usado pelo botão "Testar conexão" do aplicativo. */
function doGet() {
  return resposta({
    ok: true,
    planilha: SpreadsheetApp.getActiveSpreadsheet().getName()
  });
}

/* ================================================================
   EXCLUSÃO / EDIÇÃO DE LANÇAMENTOS
   ================================================================ */

/**
 * Remove da aba "Registros" os lançamentos indicados (pelo ID) e
 * recalcula a célula correspondente na aba da turma:
 *   - critério de nota  → volta ao último valor restante (ou fica vazia);
 *   - critério de texto → remove apenas o trecho excluído.
 * Devolve a quantidade de linhas excluídas.
 */
function processarExclusoes(planilha, abaLog, exclusoes) {
  var dados = abaLog.getDataRange().getValues();
  var porId = {};
  for (var i = 1; i < dados.length; i++) {
    var id = dados[i][11]; // coluna "ID do app"
    if (id) porId[String(id)] = { linha: i + 1, dados: dados[i] };
  }

  var linhasExcluir = [];
  var afetados = {}; // chave: disciplina|turma|estudante|critério|bimestre
  exclusoes.forEach(function (ex) {
    var alvo = porId[String(ex.id)];
    if (!alvo) return; // já não existe (ou nunca chegou)
    linhasExcluir.push(alvo.linha);
    var d = alvo.dados; // [dataHora,disc,turma,est,crit,tipo,valor,coment,bim,prof,célula,id]
    var chave = [d[1], d[2], d[3], d[4], d[8]].join("|");
    if (!afetados[chave]) {
      afetados[chave] = { disciplina: d[1], turma: d[2], estudante: d[3],
                          criterio: d[4], bimestre: d[8],
                          textosRemovidos: [], comentariosRemovidos: [] };
    }
    if (d[5] === "Nota") {
      if (d[7]) afetados[chave].comentariosRemovidos.push(String(d[7]));
    } else {
      afetados[chave].textosRemovidos.push(String(d[6]));
    }
  });

  // apaga de baixo para cima para não bagunçar os números das linhas
  linhasExcluir.sort(function (a, b) { return b - a; })
               .forEach(function (l) { abaLog.deleteRow(l); });

  Object.keys(afetados).forEach(function (k) {
    recomputarCelula(planilha, abaLog, afetados[k]);
  });

  return linhasExcluir.length;
}

/** Recalcula a célula da aba da turma depois de uma exclusão. */
function recomputarCelula(planilha, abaLog, info) {
  var aba = localizarAbaTurma(planilha, info.disciplina, info.turma);
  if (!aba) return;
  var col = localizarColunaCriterio(aba, info.criterio);
  if (!col) return;
  var linha = localizarLinhaEstudante(aba, info.estudante, info.bimestre);
  if (!linha) return;
  var celula = aba.getRange(linha, col);

  if (info.textosRemovidos.length > 0) {
    // critério de texto: retira só os trechos excluídos, preservando o resto
    removerFragmentos(celula, info.textosRemovidos);
  } else {
    // critério de nota: procura o último valor restante no log
    var dados = abaLog.getDataRange().getValues();
    var ultimo = null;
    for (var i = 1; i < dados.length; i++) {
      var d = dados[i];
      if (String(d[1]) === String(info.disciplina) &&
          String(d[2]) === String(info.turma) &&
          String(d[3]) === String(info.estudante) &&
          String(d[4]) === String(info.criterio) &&
          String(d[8]) === String(info.bimestre) &&
          d[5] === "Nota") ultimo = d[6];
    }
    if (ultimo === null || ultimo === "") celula.clearContent();
    else celula.setValue(Number(ultimo));
  }

  // comentários de nota que haviam sido anexados à coluna Observações
  if (info.comentariosRemovidos.length > 0) {
    var colObs = localizarColunaCriterio(aba, "Observações");
    if (colObs) removerFragmentos(aba.getRange(linha, colObs), info.comentariosRemovidos);
  }
}

/** Remove trechos específicos de uma célula de texto (separados por " / "). */
function removerFragmentos(celula, trechos) {
  var partes = String(celula.getValue() || "").split(" / ")
                .map(function (p) { return p.trim(); })
                .filter(function (p) { return p !== ""; });
  trechos.forEach(function (t) {
    var idx = partes.indexOf(String(t).trim());
    if (idx >= 0) partes.splice(idx, 1);
  });
  if (partes.length > 0) celula.setValue(partes.join(" / "));
  else celula.clearContent();
}

/* ================================================================
   LANÇAMENTO DIRETO NA ABA DA TURMA
   ================================================================ */

/**
 * Localiza a aba da turma, a linha do estudante e a coluna do critério
 * e grava o valor. Devolve um texto dizendo onde foi lançado (ou por
 * que não foi possível — nesse caso o lançamento fica só em Registros).
 */
function lancarNaAbaDaTurma(planilha, r) {
  var aba = localizarAbaTurma(planilha, r.disciplina, r.turma);
  if (!aba) return "aba da turma não encontrada";

  var coluna = localizarColunaCriterio(aba, r.criterio);
  if (!coluna) return "coluna do critério não encontrada";

  var linha = localizarLinhaEstudante(aba, r.estudante, r.bimestre);
  if (!linha) return "estudante não encontrado na aba";

  var celula = aba.getRange(linha, coluna);
  if (r.tipo === "Nota") {
    celula.setValue(Number(r.valor));
    if (r.observacao) anexarTexto(aba, linha, "Observações", r.observacao);
  } else {
    // Observações / Feedback / Frequência: acrescenta sem apagar o que já existe
    var atual = String(celula.getValue() || "").trim();
    var novo = String(r.valor || "").trim();
    celula.setValue(atual && atual.indexOf(novo) === -1 ? atual + " / " + novo : (atual || novo));
  }
  return aba.getName() + "!" + celula.getA1Notation();
}

/** Acrescenta texto à coluna indicada (se existir) na mesma linha. */
function anexarTexto(aba, linha, nomeColuna, texto) {
  var col = localizarColunaCriterio(aba, nomeColuna);
  if (!col) return;
  var celula = aba.getRange(linha, col);
  var atual = String(celula.getValue() || "").trim();
  if (atual.indexOf(texto) !== -1) return;
  celula.setValue(atual ? atual + " / " + texto : texto);
}

/** Normaliza texto: minúsculas, sem acentos, sem espaços duplicados. */
function normalizar(s) {
  return String(s || "").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ").trim();
}

/** Normalização mais forte para comparar cabeçalhos (ignora espaços e barras). */
function normalizarCabecalho(s) {
  return normalizar(s).replace(/[\s\/\-\.]/g, "");
}

/**
 * Encontra a aba correspondente à turma e disciplina.
 *   Projeto de Vida → aba "7ºA", "8ºC"… ou "9ºA PV"…
 *   História        → aba "9ºA História"…
 */
function nomeAbaTurma(disciplina, turma) {
  if (normalizar(disciplina) === "projeto de vida") return turma;
  return turma + " " + disciplina;
}

function localizarAbaTurma(planilha, disciplina, turma) {
  var abas = planilha.getSheets();

  // 1) Nome padrão exato (turmas criadas pelo app: "6ºA Geografia" etc.)
  var exato = normalizar(nomeAbaTurma(disciplina, turma));
  for (var i = 0; i < abas.length; i++)
    if (normalizar(abas[i].getName()) === exato) return abas[i];

  // 2) Heurística para as abas originais da planilha
  var alvoTurma = normalizar(turma);
  var d = normalizar(disciplina);
  var ehPV = d === "projeto de vida";
  var ehHistoria = d.indexOf("historia") !== -1;
  var candidatas = [];

  for (var i = 0; i < abas.length; i++) {
    var nome = normalizar(abas[i].getName());
    if (nome.indexOf(alvoTurma) === -1) continue;
    if (nome.indexOf("excluida") !== -1) continue; // abas marcadas como excluídas
    var temHist = nome.indexOf("historia") !== -1;
    if (ehHistoria && temHist) candidatas.push(abas[i]);
    else if (ehPV && !temHist) candidatas.push(abas[i]);
    else if (!ehPV && !ehHistoria && nome.indexOf(d) !== -1) candidatas.push(abas[i]);
  }
  if (candidatas.length === 0) return null;
  // Preferência: nome exato da turma (PV dos 7º/8º) ou com "pv" (9º anos)
  candidatas.sort(function (a, b) {
    return normalizar(a.getName()).length - normalizar(b.getName()).length;
  });
  return candidatas[0];
}

/** Encontra a coluna cujo cabeçalho corresponde ao critério. */
function localizarColunaCriterio(aba, criterio) {
  var alvo = normalizarCabecalho(criterio);
  var ultimaCol = aba.getLastColumn();
  if (ultimaCol < 1) return null;
  var cabecalhos = aba.getRange(1, 1, 1, ultimaCol).getValues()[0];
  var melhor = null;
  for (var c = 0; c < cabecalhos.length; c++) {
    var h = normalizarCabecalho(cabecalhos[c]);
    if (!h) continue;
    if (h === alvo) return c + 1;                       // igual
    if (h.indexOf(alvo) === 0 || alvo.indexOf(h) === 0) // um começa com o outro
      if (melhor === null) melhor = c + 1;
  }
  return melhor;
}

/**
 * Encontra a linha do estudante (coluna B). Se houver mais de uma linha
 * (uma por bimestre), prefere a do bimestre informado (coluna A).
 */
function localizarLinhaEstudante(aba, estudante, bimestre) {
  var alvo = normalizar(estudante);
  var alvoBim = normalizarCabecalho(bimestre);
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return null;

  var dados = aba.getRange(1, 1, ultimaLinha, 2).getValues(); // colunas A e B
  var exataBim = null, exata = null, parcialBim = null, parcial = null;

  for (var i = 1; i < dados.length; i++) {
    var nome = normalizar(dados[i][1]);
    if (!nome) continue;
    var bimLinha = normalizarCabecalho(dados[i][0]);
    var mesmaBim = alvoBim && bimLinha === alvoBim;

    if (nome === alvo) {
      if (mesmaBim && !exataBim) exataBim = i + 1;
      if (!exata) exata = i + 1;
    } else if (nome.indexOf(alvo) === 0 || alvo.indexOf(nome) === 0) {
      // tolera pequenas diferenças de nome entre abas (ex.: sobrenome faltando)
      if (mesmaBim && !parcialBim) parcialBim = i + 1;
      if (!parcial) parcial = i + 1;
    }
  }
  return exataBim || exata || parcialBim || parcial;
}

/* ================================================================
   ABA REGISTROS E RESUMO
   ================================================================ */

/** Cria (se necessário) e devolve a aba "Registros" já formatada. */
function obterAbaRegistros(planilha) {
  var aba = planilha.getSheetByName("Registros");
  if (!aba) {
    aba = planilha.insertSheet("Registros");
    aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO])
       .setFontWeight("bold").setBackground("#6c4fd8").setFontColor("#ffffff");
    aba.setFrozenRows(1);
    aba.setColumnWidths(1, CABECALHO.length, 130);
    aba.setColumnWidth(4, 260);  // Estudante
    aba.setColumnWidth(5, 240);  // Critério
    aba.setColumnWidth(8, 240);  // Comentário
  }
  return aba;
}

/** Recalcula a aba "Resumo": média das notas por estudante e disciplina. */
function atualizarResumo(planilha, abaLog) {
  var resumo = planilha.getSheetByName("Resumo");
  if (!resumo) resumo = planilha.insertSheet("Resumo");

  var ultimaLinha = abaLog.getLastRow();
  if (ultimaLinha < 2) return;
  var dados = abaLog.getRange(2, 1, ultimaLinha - 1, 9).getValues();

  var mapa = {}; // chave: disciplina|turma|estudante|bimestre
  dados.forEach(function (l) {
    var chave = l[1] + "|" + l[2] + "|" + l[3] + "|" + l[8];
    if (!mapa[chave]) mapa[chave] = { disc: l[1], turma: l[2], est: l[3],
                                      bim: l[8], soma: 0, qtd: 0, obs: 0 };
    if (l[5] === "Nota") { mapa[chave].soma += Number(l[6]) || 0; mapa[chave].qtd++; }
    else mapa[chave].obs++;
  });

  var linhas = Object.keys(mapa).sort().map(function (k) {
    var t = mapa[k];
    return [t.disc, t.turma, t.est, t.bim, t.qtd,
            t.qtd ? Math.round((t.soma / t.qtd) * 100) / 100 : "", t.obs];
  });

  resumo.clearContents();
  resumo.getRange(1, 1, 1, 7)
        .setValues([["Disciplina", "Turma", "Estudante", "Bimestre",
                     "Qtde de notas", "Média", "Observações registradas"]])
        .setFontWeight("bold").setBackground("#2eb872").setFontColor("#ffffff");
  resumo.setFrozenRows(1);
  if (linhas.length > 0) resumo.getRange(2, 1, linhas.length, 7).setValues(linhas);
  resumo.setColumnWidth(3, 260);
}

/* ================================================================
   ALTERAÇÕES ESTRUTURAIS (novas turmas, disciplinas, estudantes)
   ================================================================ */

/**
 * Aplica as alterações estruturais enviadas pelo app e devolve os IDs
 * das que deram certo (as demais ficam na fila do app e são reenviadas).
 */
function processarEstruturas(planilha, estruturas) {
  var feitos = [];
  estruturas.forEach(function (op) {
    try {
      switch (op.tipo) {
        case "novaDisciplina": break; // as abas são criadas por turma
        case "novaTurma":      criarAbaTurma(planilha, op); break;
        case "addAlunos":      adicionarEstudantesAba(planilha, op); break;
        case "editAluno":      renomearEstudanteAba(planilha, op); break;
        case "delAluno":       marcarEstudanteRemovido(planilha, op); break;
        case "delTurma":       marcarAbaExcluida(planilha, op.disciplina, op.turma); break;
        case "delDisciplina":
          (op.turmas || []).forEach(function (t) { marcarAbaExcluida(planilha, op.disciplina, t); });
          break;
      }
      feitos.push(op.id);
    } catch (e) { /* fica pendente; o app reenvia na próxima sincronização */ }
  });
  return feitos;
}

/** Cria a aba da turma com cabeçalho (Bimestre, Estudante, critérios). */
function criarAbaTurma(planilha, op) {
  if (localizarAbaTurma(planilha, op.disciplina, op.turma)) {
    // aba já existe: apenas garante os estudantes
    adicionarEstudantesAba(planilha, op);
    return;
  }
  var aba = planilha.insertSheet(nomeAbaTurma(op.disciplina, op.turma));
  var cab = ["Bimestre", "Estudante"].concat(op.criterios || []);
  aba.getRange(1, 1, 1, cab.length).setValues([cab])
     .setFontWeight("bold").setBackground("#6c4fd8").setFontColor("#ffffff");
  aba.setFrozenRows(1);
  aba.setColumnWidth(2, 260);
  var alunos = op.alunos || [];
  if (alunos.length > 0) {
    aba.getRange(2, 2, alunos.length, 1)
       .setValues(alunos.map(function (a) { return [a]; }));
  }
}

/** Acrescenta estudantes à aba da turma (ignora quem já está lá). */
function adicionarEstudantesAba(planilha, op) {
  var aba = localizarAbaTurma(planilha, op.disciplina, op.turma);
  if (!aba) throw "aba da turma não encontrada";
  (op.alunos || []).forEach(function (a) {
    if (localizarLinhaEstudante(aba, a, "")) return;
    aba.getRange(aba.getLastRow() + 1, 2).setValue(a);
  });
}

/** Renomeia o estudante na aba da turma. */
function renomearEstudanteAba(planilha, op) {
  var aba = localizarAbaTurma(planilha, op.disciplina, op.turma);
  if (!aba) throw "aba da turma não encontrada";
  var linha = localizarLinhaEstudante(aba, op.antigo, "");
  if (!linha) throw "estudante não encontrado";
  aba.getRange(linha, 2).setValue(op.novo);
}

/** Marca o estudante como removido (risca a linha — nada é apagado). */
function marcarEstudanteRemovido(planilha, op) {
  var aba = localizarAbaTurma(planilha, op.disciplina, op.turma);
  if (!aba) return;
  var linha = localizarLinhaEstudante(aba, op.estudante, "");
  if (!linha) return;
  aba.getRange(linha, 1, 1, Math.max(aba.getLastColumn(), 2))
     .setFontLine("line-through");
}

/** Marca a aba da turma como excluída (renomeia — nada é apagado). */
function marcarAbaExcluida(planilha, disciplina, turma) {
  var aba = localizarAbaTurma(planilha, disciplina, turma);
  if (!aba) return;
  var novo = "(excluída) " + aba.getName();
  if (!planilha.getSheetByName(novo)) aba.setName(novo);
}

/** Monta a resposta JSON para o aplicativo. */
function resposta(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
                       .setMimeType(ContentService.MimeType.JSON);
}
