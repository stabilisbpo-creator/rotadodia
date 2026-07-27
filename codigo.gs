/**
 * ROTA DO DIA — Backend (Google Apps Script)
 * Stabilis BPO — sistema pessoal de produtividade do João
 *
 * ABAS DA PLANILHA:
 *
 * Config (a "receita" da rotina — editável pelo painel ou direto aqui)
 *   A: ID              -> identificador único da linha (ex: "R001")
 *   B: Horario          -> "HH:MM"
 *   C: Atividade        -> texto mostrado no painel
 *   D: Alerta           -> frase imperativa da notificação (pode ficar vazia)
 *   E: Tipo             -> "Rotina" | "Profissional" | "Noite"
 *   F: Ativo            -> TRUE/FALSE (permite desligar um item sem apagar)
 *   G: IDEventoCalendar -> preenchido pela sincronização com o Google Calendar (futuro)
 *
 * Registros (histórico — uma linha por item por dia)
 *   A: Data              -> "dd/mm/aaaa"
 *   B: ConfigID           -> referência à linha da Config
 *   C: Horario
 *   D: Atividade
 *   E: Status             -> "Pendente" | "Concluido" | "NaoConcluido"
 *   F: DataHoraResposta   -> preenchido quando ele marca o status
 *   G: DiaFechado          -> TRUE/FALSE (marcado pelo botão "Salvar dia")
 */

const CONFIG_SHEET = "Config";
const REGISTROS_SHEET = "Registros";

// ============================================================
// SETUP — rodar uma vez, manualmente, pra criar as abas e a rotina validada
// ============================================================
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) config = ss.insertSheet(CONFIG_SHEET);
  config.clear();
  config.appendRow(["ID", "Horario", "Atividade", "Alerta", "Tipo", "Ativo", "IDEventoCalendar"]);

  const rotinaValidada = [
    ["R001", "05:00", "Despertador", "Acorda! Hora de levantar.", "Rotina", true, ""],
    ["R002", "05:00", "Levantar, beber água e despertar", "", "Rotina", true, ""],
    ["R003", "05:05", "Higiene pessoal", "", "Rotina", true, ""],
    ["R004", "05:20", "Devocional (Bíblia, oração e reflexão)", "Pare e vá pro devocional.", "Rotina", true, ""],
    ["R005", "05:50", "Registrar aprendizados e revisar a missão do dia", "Revisa a missão do dia.", "Rotina", true, ""],
    ["R006", "05:55", "Pegar mochila e sair de casa", "", "Rotina", true, ""],
    ["R007", "06:00", "Deslocamento até a academia", "", "Rotina", true, ""],
    ["R008", "06:15", "Academia", "Hora do treino — bora pra academia.", "Rotina", true, ""],
    ["R009", "07:10", "Banho, troca de roupa e organização na academia", "Pare e vá pro banho.", "Rotina", true, ""],
    ["R010", "07:35", "Deslocamento até o escritório", "", "Rotina", true, ""],
    ["R011", "07:45", "Ritual de Abertura (organizar mesa, revisar agenda, café, 3 prioridades)", "Início do Ritual de Abertura. Sem WhatsApp ainda.", "Profissional", true, ""],
    ["R012", "08:00", "Bloco CEO — estratégia, crescimento, IA, planejamento, indicadores", "Bloco CEO começou. WhatsApp fechado, e-mail fechado, sem reuniões.", "Profissional", true, ""],
    ["R013", "09:20", "Pausa (água, café, alongar, 1ª conferência do WhatsApp)", "Pausa — pode abrir o WhatsApp agora.", "Profissional", true, ""],
    ["R014", "09:35", "Bloco Comercial e Consultoria — reuniões, propostas, clientes, follow-up", "Bloco Comercial começou.", "Profissional", true, ""],
    ["R015", "11:40", "Encerramento da manhã (delegar, registrar decisões, CRM, organizar mesa)", "", "Profissional", true, ""],
    ["R016", "11:55", "Saída para o almoço", "Pausa pro almoço.", "Rotina", true, ""],
    ["R017", "20:45", "Separar o amanhã", "Pare e separe as coisas de amanhã.", "Noite", true, ""],
    ["R018", "21:00", "Momento de leitura", "Hora da leitura.", "Noite", true, ""]
  ];
  rotinaValidada.forEach(r => config.appendRow(r));

  let registros = ss.getSheetByName(REGISTROS_SHEET);
  if (!registros) registros = ss.insertSheet(REGISTROS_SHEET);
  registros.clear();
  registros.appendRow(["Data", "ConfigID", "Horario", "Atividade", "Status", "DataHoraResposta", "DiaFechado"]);

  SpreadsheetApp.flush();
  Logger.log("Planilha configurada com a rotina validada da manhã.");
}

// ============================================================
// WEB APP — doGet / doPost
// ============================================================
function doGet(e) {
  const action = e.parameter.action;
  if (action === "getDia") return getDia(e.parameter.data);
  if (action === "getConfig") return responder({ status: "ok", config: lerConfig() });
  if (action === "getHistorico") return getHistorico(e.parameter);
  if (action === "getDestaques") return getDestaques();
  return responder({ status: "ok" });
}

function doPost(e) {
  const dados = JSON.parse(e.postData.contents);
  if (dados.action === "marcarStatus") return marcarStatus(dados);
  if (dados.action === "fecharDia") return fecharDia(dados.data);
  if (dados.action === "salvarRotina") return salvarRotina(dados);
  return responder({ status: "erro", mensagem: "action desconhecida" });
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LEITURA
// ============================================================
function lerConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG_SHEET);
  if (!aba || aba.getLastRow() < 2) return [];
  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();
  return linhas
    .filter(l => l[5] === true) // só ativos
    .map(l => ({
      id: l[0], horario: formatarHora(l[1]), atividade: l[2],
      alerta: l[3], tipo: l[4], idEventoCalendar: l[6]
    }));
}

function lerRegistrosDoDia(dataStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(REGISTROS_SHEET);
  if (!aba || aba.getLastRow() < 2) return [];
  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();
  return linhas
    .filter(l => formatarData(l[0]) === dataStr)
    .map((l, idx) => ({
      linha: idx + 2, data: formatarData(l[0]), configId: l[1], horario: formatarHora(l[2]),
      atividade: l[3], status: l[4], dataHoraResposta: l[5], diaFechado: l[6]
    }));
}

// Junta Config + Registros do dia. Se o dia ainda não tem registros
// (ex: primeira vez que ele abre aquela data), cria como "Pendente" na hora.
function getDia(dataStr) {
  const config = lerConfig();
  const tipoPorId = {};
  config.forEach(c => { tipoPorId[c.id] = c.tipo; });

  let registros = lerRegistrosDoDia(dataStr);

  if (registros.length === 0) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(REGISTROS_SHEET);
    config.forEach(item => {
      aba.appendRow([dataStr, item.id, item.horario, item.atividade, "Pendente", "", false]);
    });
    registros = lerRegistrosDoDia(dataStr);
  }

  registros.forEach(r => { r.tipo = tipoPorId[r.configId] || "Outros"; });

  // adiciona os compromissos reais do Calendar (leitura, sem status pra marcar)
  try {
    const compromissos = lerCompromissosDoDia(dataStr);
    compromissos.forEach((c, idx) => {
      registros.push({
        linha: null, data: dataStr, configId: "CAL_" + idx,
        horario: c.horario, atividade: c.titulo, status: "Compromisso",
        dataHoraResposta: "", diaFechado: false, tipo: "Compromisso"
      });
    });
  } catch (err) {
    Logger.log("Não consegui ler compromissos do Calendar: " + err);
  }

  return responder({ status: "ok", data: dataStr, itens: registros });
}

// Retorna os registros dos últimos N dias (padrão 30), pra alimentar a aba de evolução.
function getHistorico(params) {
  const dias = parseInt(params.dias, 10) || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - dias);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(REGISTROS_SHEET);
  if (!aba || aba.getLastRow() < 2) return responder({ status: "ok", itens: [] });

  const config = lerConfig();
  const tipoPorId = {};
  config.forEach(c => { tipoPorId[c.id] = c.tipo; });

  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();
  const itens = linhas
    .map(l => ({
      data: formatarData(l[0]), configId: l[1], horario: formatarHora(l[2]),
      atividade: l[3], status: l[4], dataHoraResposta: l[5], diaFechado: l[6],
      tipo: tipoPorId[l[1]] || "Outros"
    }))
    .filter(i => {
      const d = parseDataBR(i.data);
      return d && d >= cutoff;
    });

  return responder({ status: "ok", itens: itens });
}

function parseDataBR(str) {
  if (!str) return null;
  const partes = str.split("/");
  if (partes.length !== 3) return null;
  return new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
}
function marcarStatus(dados) {
  // dados: { data, configId, status }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(REGISTROS_SHEET);
  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();

  for (let i = 0; i < linhas.length; i++) {
    if (formatarData(linhas[i][0]) === dados.data && linhas[i][1] === dados.configId) {
      const linhaPlanilha = i + 2;
      aba.getRange(linhaPlanilha, 5).setValue(dados.status);
      aba.getRange(linhaPlanilha, 6).setValue(new Date());
      return responder({ status: "ok" });
    }
  }
  return responder({ status: "erro", mensagem: "registro não encontrado" });
}

function fecharDia(dataStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(REGISTROS_SHEET);
  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();

  for (let i = 0; i < linhas.length; i++) {
    if (formatarData(linhas[i][0]) === dataStr) {
      aba.getRange(i + 2, 7).setValue(true);
    }
  }
  return responder({ status: "ok" });
}

// dados: { itens: [{id, horario, atividade, alerta, tipo}], aplicarHoje: bool }
function salvarRotina(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG_SHEET);
  const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();

  dados.itens.forEach(item => {
    let encontrado = false;
    for (let i = 0; i < linhas.length; i++) {
      if (linhas[i][0] === item.id) {
        const linhaPlanilha = i + 2;
        aba.getRange(linhaPlanilha, 2).setValue(item.horario);
        aba.getRange(linhaPlanilha, 3).setValue(item.atividade);
        aba.getRange(linhaPlanilha, 4).setValue(item.alerta || "");
        aba.getRange(linhaPlanilha, 5).setValue(item.tipo);
        encontrado = true;
        break;
      }
    }
    if (!encontrado) {
      const novoId = "R" + String(aba.getLastRow()).padStart(3, "0");
      aba.appendRow([novoId, item.horario, item.atividade, item.alerta || "", item.tipo, true, ""]);
    }
  });

  // sincroniza com o Google Calendar (rotina pessoal), sem travar o salvamento se algo falhar
  try {
    sincronizarRotinaComCalendar();
  } catch (err) {
    Logger.log("Falha ao sincronizar com o Calendar: " + err);
  }

  if (dados.aplicarHoje) {
    const hoje = formatarData(new Date());
    const regAba = ss.getSheetByName(REGISTROS_SHEET);
    const regLinhas = regAba.getRange(2, 1, regAba.getLastRow() - 1, 7).getValues();
    for (let i = regLinhas.length - 1; i >= 0; i--) {
      if (formatarData(regLinhas[i][0]) === hoje) regAba.deleteRow(i + 2);
    }
  }

  return responder({ status: "ok" });
}

// ============================================================
// GATILHO DE ALERTA — configurar pra rodar a cada 1 minuto
// (Editor de Apps Script → Gatilhos → Adicionar gatilho → checarAlertas → Baseado em tempo → Por minuto)
// ============================================================
function checarAlertas() {
  const agora = new Date();
  const horaAtual = Utilities.formatDate(agora, Session.getScriptTimeZone(), "HH:mm");
  const hojeStr = formatarData(agora);

  const config = lerConfig();
  config.forEach(item => {
    if (item.horario !== horaAtual) return;
    if (!item.alerta) return;

    const chave = "alerta_" + item.id + "_" + hojeStr;
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(chave)) return; // já disparado hoje, evita duplicar

    enviarAlerta(item.alerta);
    props.setProperty(chave, "enviado");
  });
}

// Hoje envia por e-mail (funciona sem nenhuma configuração extra).
// Quando o Firebase/PWA estiver pronto, trocar o corpo desta função
// pela chamada à API do FCM, mantendo a mesma assinatura.
function enviarAlerta(texto) {
  const destinatario = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: destinatario,
    subject: "🧭 Rota do Dia",
    body: texto
  });
}

// ============================================================
// DESTAQUES DO DIA — versículo + frase motivacional
// (busca 1x por dia e guarda em cache, pra não bater na API toda hora)
// ============================================================
function getDestaques() {
  const hojeStr = formatarData(new Date());
  const props = PropertiesService.getScriptProperties();
  const chave = "destaques_" + hojeStr;

  const salvo = props.getProperty(chave);
  if (salvo) return responder(JSON.parse(salvo));

  const resultado = {
    status: "ok",
    versiculo: buscarVersiculo(),
    frase: buscarFraseTraduzida()
  };
  props.setProperty(chave, JSON.stringify(resultado));
  return responder(resultado);
}

function buscarVersiculo() {
  try {
    const resp = UrlFetchApp.fetch("https://www.abibliadigital.com.br/api/verses/nvi/random", { muteHttpExceptions: true });
    const json = JSON.parse(resp.getContentText());
    return { texto: json.text, referencia: json.book.name + " " + json.chapter + ":" + json.verse };
  } catch (err) {
    return { texto: "Tudo o que fizerem, façam de todo o coração, como para o Senhor.", referencia: "Colossenses 3:23" };
  }
}

function buscarFraseTraduzida() {
  try {
    const respQuote = UrlFetchApp.fetch("https://zenquotes.io/api/random", { muteHttpExceptions: true });
    const quote = JSON.parse(respQuote.getContentText())[0];

    const respTrad = UrlFetchApp.fetch(
      "https://api.mymemory.translated.net/get?q=" + encodeURIComponent(quote.q) + "&langpair=en|pt-BR",
      { muteHttpExceptions: true }
    );
    const trad = JSON.parse(respTrad.getContentText());
    const texto = trad.responseData && trad.responseData.translatedText ? trad.responseData.translatedText : quote.q;

    return { texto: texto, autor: quote.a };
  } catch (err) {
    return { texto: "A disciplina é a ponte entre metas e realizações.", autor: "Jim Rohn" };
  }
}

// Rede de segurança: apaga TODOS os eventos recorrentes que a sincronização criou,
// usando os IDs salvos na coluna IDEventoCalendar. Roda manualmente pelo editor.
function apagarRotinaDoCalendar() {
  const calendario = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendario) throw new Error("Não encontrei ou não tenho acesso ao calendário: " + CALENDAR_ID);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = ss.getSheetByName(CONFIG_SHEET);
  const linhas = abaConfig.getRange(2, 1, abaConfig.getLastRow() - 1, 7).getValues();

  let apagados = 0;
  linhas.forEach((l, idx) => {
    const idEvento = l[6];
    if (!idEvento) return;
    try {
      const serie = calendario.getEventSeriesById(idEvento);
      if (serie) {
        serie.deleteEventSeries();
        apagados++;
      }
    } catch (err) {
      Logger.log("Não consegui apagar o evento da linha " + (idx + 2) + ": " + err);
    }
    abaConfig.getRange(idx + 2, 7).setValue(""); // limpa o ID na planilha de qualquer forma
  });

  Logger.log("Apagados " + apagados + " eventos recorrentes do Calendar.");
}

// ============================================================
// GOOGLE CALENDAR — jonga.chaves@gmail.com
// ============================================================
const CALENDAR_ID = "jonga.chaves@gmail.com";

// Cria/atualiza os eventos recorrentes diários da ROTINA PESSOAL no Calendar dele.
// Só sincroniza itens Tipo="Rotina" — o Bloco Profissional fica só no painel.
// A duração de cada evento é calculada até o horário do próximo item da rotina
// (o último item do dia usa 15 min como padrão).
function sincronizarRotinaComCalendar() {
  const calendario = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendario) throw new Error("Não encontrei ou não tenho acesso ao calendário: " + CALENDAR_ID);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaConfig = ss.getSheetByName(CONFIG_SHEET);
  const linhas = abaConfig.getRange(2, 1, abaConfig.getLastRow() - 1, 7).getValues();

  const rotina = linhas
    .map((l, idx) => ({
      linhaPlanilha: idx + 2, id: l[0], horario: formatarHora(l[1]),
      atividade: l[2], tipo: l[4], ativo: l[5], idEvento: l[6]
    }))
    .filter(r => r.tipo === "Rotina" && r.ativo === true)
    .sort((a, b) => a.horario.localeCompare(b.horario));

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  rotina.forEach((item, idx) => {
    // apaga a série antiga, se já existia
    if (item.idEvento) {
      try {
        const antigo = calendario.getEventSeriesById(item.idEvento);
        if (antigo) antigo.deleteEventSeries();
      } catch (e) {
        // série já não existe mais (foi apagada manualmente) — segue o baile
      }
    }

    const [h, m] = item.horario.split(":").map(Number);
    const inicio = new Date(hoje);
    inicio.setHours(h, m, 0, 0);

    let duracaoMin = 15;
    if (idx < rotina.length - 1) {
      const [h2, m2] = rotina[idx + 1].horario.split(":").map(Number);
      const calc = (h2 * 60 + m2) - (h * 60 + m);
      if (calc > 0) duracaoMin = calc;
    }
    const fim = new Date(inicio.getTime() + duracaoMin * 60000);

    const recorrencia = CalendarApp.newRecurrence().addDailyRule();
    const serie = calendario.createEventSeries(item.atividade, inicio, fim, recorrencia);

    abaConfig.getRange(item.linhaPlanilha, 7).setValue(serie.getId());
  });

  Logger.log("Rotina sincronizada: " + rotina.length + " itens.");
}

// Lê os compromissos reais (reuniões etc.) do dia, pra exibir como referência no painel.
function lerCompromissosDoDia(dataStr) {
  const calendario = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendario) throw new Error("Não encontrei ou não tenho acesso ao calendário: " + CALENDAR_ID);

  const [dd, mm, aaaa] = dataStr.split("/").map(Number);
  const dia = new Date(aaaa, mm - 1, dd);
  const eventos = calendario.getEventsForDay(dia);

  return eventos.map(ev => ({
    titulo: ev.getTitle(),
    horario: Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "HH:mm"),
    horarioFim: Utilities.formatDate(ev.getEndTime(), Session.getScriptTimeZone(), "HH:mm")
  }));
}

// ============================================================
// HELPERS (padrão já usado nos outros projetos do Stabilis)
// ============================================================
function formatarData(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const d = String(val.getDate()).padStart(2, "0");
    const m = String(val.getMonth() + 1).padStart(2, "0");
    return d + "/" + m + "/" + val.getFullYear();
  }
  const str = val.toString().trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  if (str.includes("GMT") || str.includes("UTC")) {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      const d = String(dt.getUTCDate()).padStart(2, "0");
      const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
      return d + "/" + m + "/" + dt.getUTCFullYear();
    }
  }
  return str;
}

function formatarHora(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, "0");
    const m = String(val.getMinutes()).padStart(2, "0");
    return h + ":" + m;
  }
  return val.toString().trim();
}
