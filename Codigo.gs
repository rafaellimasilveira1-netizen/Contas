/**
 * API do Fluxo de Caixa Flow
 *
 * Fica dentro da planilha DRE_Mensal_Flow e atende o painel publicado no Netlify.
 * A senha fica em Configurações do projeto > Propriedades do script > CHAVE.
 * Depois de qualquer mudança neste código: Implantar > Gerenciar implantações >
 * lápis > Versão: Nova versão > Implantar. A URL /exec continua a mesma.
 */

const PLANILHA_ID = "17gk5FXS4O09K97OjxjCJnVffouzLGmUqaqtZIjEO2Fc";
const MAX_LINHA = 1000; // a aba Lançamentos tem formatação e validação até a linha 1000

function doGet() {
  return saida_({ ok: true, msg: "API do Fluxo de Caixa no ar" });
}

function doPost(e) {
  try {
    const chave = PropertiesService.getScriptProperties().getProperty("CHAVE");
    if (!chave) return saida_({ ok: false, erro: "CHAVE não definida nas propriedades do script" });
    const b = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (b.key !== chave) return saida_({ ok: false, erro: "senha" });

    const trava = LockService.getScriptLock();
    trava.waitLock(20000);
    try {
      switch (b.action) {
        case "dados": return saida_(dados_());
        case "salvar": return saida_(salvar_(b));
        case "lote": return saida_(lote_(b));
        case "baixa": return saida_(baixa_(b));
        case "excluir": return saida_(excluir_(b));
        case "saldo": return saida_(saldo_(b));
        default: return saida_({ ok: false, erro: "ação desconhecida" });
      }
    } finally {
      trava.releaseLock();
    }
  } catch (err) {
    return saida_({ ok: false, erro: String((err && err.message) || err) });
  }
}

function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function planilha_() {
  return SpreadsheetApp.openById(PLANILHA_ID);
}

/* ---------- leitura ---------- */

function dados_() {
  const ss = planilha_();
  const tz = ss.getSpreadsheetTimeZone();

  const cfg = ss.getSheetByName("Config").getRange("B3:B4").getValues();
  const catsV = ss.getSheetByName("Categorias").getRange("A2:B60").getValues();
  const cats = [];
  catsV.forEach(function (r) {
    if (r[0] && r[1]) cats.push({ tipo: String(r[0]), cat: String(r[1]) });
  });

  const sh = ss.getSheetByName("Lançamentos");
  const ultima = Math.max(2, Math.min(sh.getLastRow(), MAX_LINHA));
  const vals = sh.getRange(2, 1, ultima - 1, 6).getValues();
  const rows = [];
  vals.forEach(function (r, i) {
    const valor = num_(r[3]);
    if (!r[0] || valor === null) return;
    rows.push({
      row: i + 2,
      tipo: String(r[0]),
      desc: String(r[1]),
      cat: String(r[2]),
      valor: valor,
      venc: iso_(r[4], tz),
      pago: iso_(r[5], tz)
    });
  });

  return {
    ok: true,
    start: iso_(cfg[0][0], tz),
    saldo: cfg[1] && cfg[1][0] !== "" ? num_(cfg[1][0]) : null,
    cats: cats,
    rows: rows
  };
}

/* ---------- gravação ---------- */

function linha_(b) {
  const tipo = b.tipo === "Saída" ? "Saída" : "Entrada";
  const valor = Number(b.valor);
  if (!(valor > 0)) throw new Error("valor inválido");
  const venc = serial_(b.venc);
  if (venc === null) throw new Error("vencimento inválido");
  const pago = b.pago ? serial_(b.pago) : "";
  if (b.pago && pago === null) throw new Error("data de pagamento inválida");
  return [tipo, texto_(b.desc, 200), texto_(b.cat, 100), valor, venc, pago];
}

function salvar_(b) {
  const sh = planilha_().getSheetByName("Lançamentos");
  const dados = linha_(b);
  let linha = b.row ? Number(b.row) : 0;
  if (linha) conferirLinha_(linha);
  else linha = proximaLinha_(sh);
  sh.getRange(linha, 1, 1, 6).setValues([dados]);
  return { ok: true, row: linha };
}

// Vários lançamentos de uma vez (recorrência). Máximo de 60.
function lote_(b) {
  const sh = planilha_().getSheetByName("Lançamentos");
  const itens = b.itens;
  if (!Array.isArray(itens) || !itens.length) throw new Error("nenhum lançamento");
  if (itens.length > 60) throw new Error("máximo de 60 lançamentos de uma vez");
  const linhas = itens.map(linha_);
  const inicio = proximaLinha_(sh);
  if (inicio + linhas.length - 1 > MAX_LINHA) throw new Error("a aba Lançamentos não tem espaço para tantos lançamentos");
  sh.getRange(inicio, 1, linhas.length, 6).setValues(linhas);
  return { ok: true, row: inicio, qtd: linhas.length };
}

function baixa_(b) {
  const sh = planilha_().getSheetByName("Lançamentos");
  const linha = Number(b.row);
  conferirLinha_(linha);
  const pago = serial_(b.pago);
  if (pago === null) throw new Error("data de pagamento inválida");
  sh.getRange(linha, 6).setValue(pago);
  return { ok: true };
}

function excluir_(b) {
  const sh = planilha_().getSheetByName("Lançamentos");
  const linha = Number(b.row);
  conferirLinha_(linha);
  sh.getRange(linha, 1, 1, 6).clearContent(); // mantém formatação e validação
  return { ok: true };
}

function saldo_(b) {
  const sh = planilha_().getSheetByName("Config");
  const data = serial_(b.start);
  if (data === null) throw new Error("data inválida");
  const valor = b.saldo === null || b.saldo === "" ? "" : Number(b.saldo);
  if (valor !== "" && isNaN(valor)) throw new Error("saldo inválido");
  sh.getRange("B3").setValue(data);
  sh.getRange("B4").setValue(valor);
  return { ok: true };
}

/* ---------- auxiliares ---------- */

function proximaLinha_(sh) {
  const vals = sh.getRange(2, 1, MAX_LINHA - 1, 6).getValues();
  let ultima = 1;
  vals.forEach(function (r, i) {
    if (r.some(function (c) { return c !== ""; })) ultima = i + 2;
  });
  const prox = ultima + 1;
  if (prox > MAX_LINHA) throw new Error("a aba Lançamentos chegou na linha " + MAX_LINHA);
  return prox;
}

function conferirLinha_(linha) {
  if (!(linha >= 2 && linha <= MAX_LINHA) || Math.floor(linha) !== linha) throw new Error("linha inválida");
}

// Texto que começa com = + - @ vira texto puro, nunca fórmula.
function texto_(v, max) {
  const t = String(v == null ? "" : v).slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function num_(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() !== "") {
    const n = parseFloat(v.replace(/[^\d,\-]/g, "").replace(",", "."));
    return isNaN(n) ? null : n;
  }
  return null;
}

function iso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  if (typeof v === "number" && v > 20000) {
    return Utilities.formatDate(new Date(Math.round((v - 25569) * 864e5)), "UTC", "yyyy-MM-dd");
  }
  if (typeof v === "string") {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v.trim());
    if (m) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  }
  return null;
}

function serial_(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
  if (!m) return null;
  return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5) + 25569;
}
