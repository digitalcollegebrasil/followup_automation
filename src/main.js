import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import { readFile, writeFile, mkdir, access } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
const __filename = fileURLToPath(import.meta.url);
const __dirname  = path.dirname(__filename);

let mainWindow;

// tenta descobrir o mesmo DATA_DIR do Python chamando utils_path.app_data_dir()
async function resolvePythonDataDir() {
  try {
    const code = `
from utils_path import app_data_dir
print(str(app_data_dir()))
`;
    const py = spawn('python', ['-c', code], { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '';
    for await (const chunk of py.stdout) out += chunk.toString();
    await new Promise(r => py.on('close', r));
    const p = out.trim();
    if (p) return p;
  } catch {}
  // fallback: userData do Electron
  return app.getPath('userData');
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 800,
    backgroundColor: '#ffffff',
    webPreferences: {
      contextIsolation: true,
      preload: path.join(app.getAppPath(), 'src', 'preload.js'),
      sandbox: false
    }
  });

  await mainWindow.loadFile(path.join(process.cwd(), 'src', 'index.html'));
  // mainWindow.webContents.openDevTools(); // se quiser
}

app.whenReady().then(async () => {
  const dataDir = await resolvePythonDataDir();
  globalThis.__DATA_DIR__ = dataDir;
  await mkdir(dataDir, { recursive: true });
  await createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---------- IPC ----------
ipcMain.handle('select-xlsx', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Selecionar Planilha',
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  });
  if (canceled || !filePaths?.[0]) return null;
  const filePath = filePaths[0];
  const buf = await readFile(filePath);
  // devolve como base64 para o renderer
  return { filePath, base64: buf.toString('base64') };
});

ipcMain.handle('get-user-data-path', async () => {
  return globalThis.__DATA_DIR__;
});

ipcMain.handle('save-outputs', async (_evt, { xlsxBase64, configJSON }) => {
  const dataDir = globalThis.__DATA_DIR__;
  const TEMP_PLANILHA = path.join(dataDir, 'planilha_filtrada.xlsx');
  const CONFIG_PATH   = path.join(dataDir, 'config.json');
  const xbuf = Buffer.from(xlsxBase64, 'base64');

  await writeFile(TEMP_PLANILHA, xbuf);
  await writeFile(CONFIG_PATH, JSON.stringify(configJSON, null, 2), 'utf8');

  return { TEMP_PLANILHA, CONFIG_PATH };
});

ipcMain.handle('run-editar', async () => {
  const jsPath = path.join(app.getAppPath(), 'src', 'editar.js');
  if (!existsSync(jsPath)) throw new Error(`Não achei ${jsPath}`);

  // ache o Node "de verdade"
  const nodeBin =
    process.env.npm_node_execpath ||          // quando startado via npm
    (await which('node.exe')) ||              // Windows
    (await which('node')) ||                  // Linux/Mac
    'node';

  const env = { ...process.env, DATA_DIR: globalThis.__DATA_DIR__ };
  const args = [
    '--enable-source-maps',
    '--trace-uncaught',
    '--trace-warnings',
    '--inspect=0',
    jsPath
  ];

  const child = spawn(nodeBin, args, {
    cwd: path.dirname(jsPath),
    env,
    stdio: ['ignore', 'pipe', 'pipe'],
    windowsHide: false,
  });

  // espelha no terminal E manda pra UI
  child.stdout.on('data', d => {
    const line = d.toString();
    process.stdout.write(line);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('editar-log', line);
    }
  });
  child.stderr.on('data', d => {
    const line = d.toString();
    process.stderr.write(line);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('editar-log', line);
    }
  });
  child.on('close', code => {
    const msg = `\n[editar.js] exited with code ${code}\n`;
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('editar-exit', msg);
    }
  });
  child.on('error', e => {
    const msg = `[run-editar] spawn error: ${e.message}\n`;
    console.error(msg);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('editar-log', msg);
    }
  });

  return true;
});

// util simples tipo "which"
async function which(bin) {
  const exts = process.platform === 'win32' ? ['.exe', ''] : [''];
  const paths = (process.env.PATH || '').split(path.delimiter);
  for (const p of paths) {
    for (const ext of exts) {
      const full = path.join(p, bin.endsWith(ext) ? bin : bin + ext);
      try { await access(full); return full; } catch {}
    }
  }
  return null;
}
