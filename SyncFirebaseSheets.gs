// ============================================================
//  PORTOEX — Sincronização Firebase → Google Sheets
//  Aba destino : DADOS ABS
//  Célula inicial: AF3565  (coluna 32, linha 3565)
//
//  INSTRUÇÕES DE INSTALAÇÃO:
//  1. Abra a planilha no Google Sheets
//  2. Extensões → Apps Script
//  3. Cole todo este código (substitua o conteúdo existente)
//  4. Salve (Ctrl+S)
//  5. Execute "configurarTrigger" UMA VEZ (clique em ▶ com a função selecionada)
//  6. Autorize as permissões solicitadas
//  A partir daí, a sincronização ocorre automaticamente a cada minuto.
// ============================================================

// ── Configurações ──────────────────────────────────────────
const FIREBASE_BASE = 'https://vagas-terceirizados-portoex-default-rtdb.firebaseio.com';
const FIREBASE_KEY  = 'AIzaSyDsKKZaK1WBo3FjNo-2o0O5a0ehV4tDy4U';
const SHEET_NAME    = 'DADOS ABS';
const START_ROW     = 3565;
const START_COL     = 32;   // Coluna AF
const NUM_COLS      = 19;   // Data → Valor Adicional

// ── Função principal ───────────────────────────────────────
function syncAcertoToSheets() {
  try {
    // 1. Buscar dados do Firebase
    const acomp     = fetchFirebase('acompanhamento');
    const empresas  = fetchFirebase('empresas');
    const motoristas = fetchFirebase('motoristas');

    if (!acomp) {
      Logger.log('Nenhum dado em acompanhamento.');
      return;
    }

    // 2. Montar linhas com a mesma lógica do _buildAcertoRows
    const rows = [];

    Object.values(acomp).forEach(function(a) {
      if (!a || !a.colNome) return;

      const emp    = (empresas && a.empresaId && empresas[a.empresaId]) ? empresas[a.empresaId] : {};
      const valores = (emp.valores && a.funcao && emp.valores[a.funcao]) ? emp.valores[a.funcao] : {};
      const valorDia = parseFloat(valores.valor) || 0;
      const tarifaHE = parseFloat(valores.he)    || 0;

      let hTrabMin = 0, heMin = 0, valorTotal = valorDia;

      if (a.chegada && a.saida) {
        hTrabMin = timeToMin(a.saida)
                 - timeToMin(a.chegada)
                 - timeToMin(a.refeicao || '01:00');
        heMin = Math.max(0, hTrabMin - 8 * 60);
        if (heMin > 0) valorTotal += (heMin / 60) * tarifaHE;
      }

      const adicional     = parseInt(a.adicional) || 0;
      const valorAdicional = valorDia * (adicional / 100);
      valorTotal += valorAdicional;

      const motNome = a.motoristaNome
        || (a.motoristaId && motoristas && motoristas[a.motoristaId]
            ? motoristas[a.motoristaId].nome : '')
        || '';

      rows.push([
        a.data        || '',            // Data
        a.colNome     || '',            // Colaborador
        a.colCPF      || '',            // CPF
        a.empresaNome || '',            // Empresa
        a.opNome      || '',            // Operação
        a.opCC        || '',            // C. Custo
        a.funcao      || '',            // Função
        a.chegada     || '',            // Chegada
        a.refeicao    || '',            // Refeição
        a.saida       || '',            // Saída
        minToHM(hTrabMin),             // H.Trabalhadas
        brl(valorTotal),               // Total
        a.md          || '',            // MD
        motNome,                        // Motorista
        minToHM(heMin),                // H.Extra
        adicional > 0 ? '+' + adicional + '%' : '—',  // Adicional
        brl(valorDia),                 // Valor Dia
        brl(tarifaHE),                 // Valor HE
        adicional > 0 ? brl(valorAdicional) : '—'     // Valor Adicional
      ]);
    });

    // 3. Ordenar por data
    rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });

    // 4. Gravar na planilha
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log('❌ Aba não encontrada: ' + SHEET_NAME);
      return;
    }

    // Limpar intervalo anterior (apenas colunas AF → AX, a partir de START_ROW)
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha >= START_ROW) {
      sheet.getRange(START_ROW, START_COL, ultimaLinha - START_ROW + 1, NUM_COLS).clearContent();
    }

    // Escrever novos dados
    if (rows.length > 0) {
      sheet.getRange(START_ROW, START_COL, rows.length, NUM_COLS).setValues(rows);
    }

    Logger.log('✅ Sincronizado: ' + rows.length + ' registros — ' + new Date().toLocaleString('pt-BR'));

  } catch (e) {
    Logger.log('❌ Erro: ' + e.message);
  }
}

// ── Configura trigger automático (1× por minuto) ───────────
function configurarTrigger() {
  // Remove triggers antigos desta função para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncAcertoToSheets') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria novo trigger: a cada 1 minuto
  ScriptApp.newTrigger('syncAcertoToSheets')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('⏱ Trigger configurado: sincronização a cada 1 minuto.');
}

// ── Remove o trigger (pausar sincronização) ────────────────
function removerTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncAcertoToSheets') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('🛑 Trigger removido. Sincronização pausada.');
}

// ── Helpers ────────────────────────────────────────────────
function fetchFirebase(path) {
  const url = FIREBASE_BASE + '/' + path + '.json?auth=' + FIREBASE_KEY;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    Logger.log('Erro ao buscar ' + path + ': HTTP ' + resp.getResponseCode());
    return null;
  }
  return JSON.parse(resp.getContentText());
}

function timeToMin(t) {
  if (!t) return 0;
  var parts = String(t).split(':').map(Number);
  return (parts[0] || 0) * 60 + (parts[1] || 0);
}

function minToHM(min) {
  if (!min || min <= 0) return '00:00';
  var abs = Math.abs(min);
  var h = Math.floor(abs / 60);
  var m = abs % 60;
  return (min < 0 ? '-' : '')
    + String(h).padStart(2, '0') + ':'
    + String(m).padStart(2, '0');
}

function brl(value) {
  return 'R$ ' + (parseFloat(value) || 0).toFixed(2).replace('.', ',');
}
