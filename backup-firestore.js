/**
 * backup-firestore.js
 * Exporta TODOS os dados do Firebase Realtime Database para um arquivo JSON com timestamp.
 *
 * Como usar:
 *   1. npm install firebase-admin
 *   2. Coloque o arquivo serviceAccountKey.json na mesma pasta
 *   3. node backup-firestore.js
 */

const { initializeApp, cert, deleteApp } = require("firebase-admin/app");
const { getDatabase } = require("firebase-admin/database");
const fs = require("fs");
const path = require("path");

// ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
const SERVICE_ACCOUNT_PATH = path.join(__dirname, "serviceAccountKey.json");
// ─────────────────────────────────────────────────────────────────────────────

function deletePreviousBackups() {
  const files = fs.readdirSync(__dirname).filter(f => f.startsWith("firestore-backup-") && f.endsWith(".json"));
  for (const file of files) {
    fs.unlinkSync(path.join(__dirname, file));
    console.log(`🗑️  Backup anterior removido: ${file}`);
  }
}

async function runBackup() {
  // Remove backups anteriores antes de criar novo
  deletePreviousBackups();

  // Lê a chave de serviço
  const serviceAccount = JSON.parse(fs.readFileSync(SERVICE_ACCOUNT_PATH, "utf8"));

  // Monta a URL do banco automaticamente a partir do project_id
  const databaseURL = `https://${serviceAccount.project_id}-default-rtdb.firebaseio.com`;

  const app = initializeApp({
    credential: cert(serviceAccount),
    databaseURL: databaseURL,
  });

  const db = getDatabase(app);

  console.log(`🔥 Conectado ao Realtime Database (${databaseURL}). Iniciando backup...\n`);

  // Exporta tudo a partir da raiz
  const snapshot = await db.ref("/").once("value");
  const data = snapshot.val();

  if (!data) {
    console.log("⚠️  O banco está vazio ou não foi encontrado.");
    await deleteApp(app);
    return;
  }

  const backup = {
    _meta: {
      project: serviceAccount.project_id,
      databaseURL: databaseURL,
      exportedAt: new Date().toISOString(),
    },
    data: data,
  };

  // Salva o arquivo com timestamp
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filename = path.join(__dirname, `firestore-backup-${timestamp}.json`);

  fs.writeFileSync(filename, JSON.stringify(backup, null, 2), "utf8");

  const sizeMB = (fs.statSync(filename).size / 1024 / 1024).toFixed(2);
  console.log(`✅ Backup salvo: firestore-backup-${timestamp}.json (${sizeMB} MB)`);

  await deleteApp(app);
}

runBackup().catch((err) => {
  console.error("❌ Erro ao fazer backup:", err.message);
  process.exit(1);
});
