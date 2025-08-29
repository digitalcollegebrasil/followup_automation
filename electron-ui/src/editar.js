// editar.js (ESM)
import 'dotenv/config';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import * as XLSX from 'xlsx/xlsx.mjs';
import { readFile } from 'node:fs/promises';
import soap from 'soap';
import { Builder, By, until } from 'selenium-webdriver';
import * as chrome from 'selenium-webdriver/chrome.js';
import chromedriver from 'chromedriver';

process.on('uncaughtException', (e) => {
  console.error('[uncaughtException]', e?.stack || e);
});
process.on('unhandledRejection', (e) => {
  console.error('[unhandledRejection]', e);
});

process.env.SELENIUM_MANAGER_LOG = process.env.SELENIUM_MANAGER_LOG || 'INFO';

const WSDL = 'https://api.sponteeducacional.net.br/WSAPIEdu.asmx?WSDL';

const CREDENCIAIS = {
  Aldeota: { codigo_cliente: '72546', token: 'QZUSqqgsLA63' },
  Sul:     { codigo_cliente: '74070', token: 'jVNLW7IIUXOh' },
};

const MESES_PT = {
  1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL',
  5: 'MAIO',    6: 'JUNHO',     7: 'JULHO', 8: 'AGOSTO',
  9: 'SETEMBRO',10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO',
};

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function readJSON(p) {
  const raw = await fs.readFile(p, 'utf8');
  return JSON.parse(raw);
}

async function getDataDir() {
  if (process.env.DATA_DIR && existsSync(process.env.DATA_DIR)) {
    console.log('[editar] DATA_DIR via env:', process.env.DATA_DIR);
    return process.env.DATA_DIR;
  }
  try {
    const { spawn } = await import('node:child_process');
    const code = `from utils_path import app_data_dir; print(str(app_data_dir()))`;
    console.log('[editar] tentando descobrir DATA_DIR via Python...');
    const out = await new Promise((resolve) => {
      const py = spawn('python', ['-c', code], { stdio: ['ignore', 'pipe', 'ignore'] });
      let s = '';
      py.stdout.on('data', (d) => (s += d.toString()));
      py.on('close', () => resolve(s.trim()));
    });
    if (out && existsSync(out)) {
      console.log('[editar] DATA_DIR via Python:', out);
      return out;
    }
  } catch (e) {
    console.log('[editar] falha ao obter DATA_DIR via Python:', e?.message || e);
  }
  const fallback = path.join(process.cwd(), '.data');
  console.log('[editar] DATA_DIR fallback:', fallback);
  return fallback;
}

async function getAlunoIdByCPF(cpf, codigo_cliente, token, soapClient) {
  const sParametrosBusca = `CPF=${cpf}`;
  try {
    console.log(`[editar] SOAP: buscando Aluno por CPF ${cpf}...`);
    const [res] = await soapClient.GetAlunosAsync({
      nCodigoCliente: codigo_cliente,
      sToken: token,
      sParametrosBusca
    });
    // (parsing mantido como antes)
    if (Array.isArray(res)) {
      const first = res[0];
      if (first && first.AlunoID) return String(first.AlunoID);
    }
    if (res && typeof res === 'object') {
      const r = res.GetAlunosResult || res.getAlunosResult || res;
      const diff = r?.diffgram || r?.Diffgram || r?.NewDataSet || r?.dataset;
      const table =
        diff?.NewDataSet?.Table ||
        diff?.DocumentElement?.Table ||
        diff?.Table ||
        r?.Table ||
        r?.Alunos || r?.Aluno;
      const arr = Array.isArray(table) ? table : (table ? [table] : []);
      for (const item of arr) {
        const id = item?.AlunoID || item?.alunoid || item?.ID;
        if (id) return String(id);
      }
      if (r?.AlunoID) return String(r.AlunoID);
    }
    console.log('[editar] SOAP: nenhum AlunoID retornado para o CPF.');
  } catch (e) {
    console.error(`Erro SOAP para CPF ${cpf}:`, e?.message || e);
  }
  return null;
}

async function buildDriver() {
  console.log('[selenium] criando Chrome driver...');
  console.log('[selenium] chromedriver.path =', chromedriver.path);

  // Garante o chromedriver no PATH (fallback se o setChromeService falhar)
  const cdDir = path.dirname(chromedriver.path);
  process.env.PATH = cdDir + path.delimiter + (process.env.PATH || '');

  const options = new chrome.Options().addArguments('--start-maximized');

  // Tente com ServiceBuilder (API v4)
  try {
    const service = new chrome.ServiceBuilder(chromedriver.path); // <- sem .build()
    const driver = await new Builder()
      .forBrowser('chrome')
      .setChromeOptions(options)
      .setChromeService(service)
      .build();
    console.log('[selenium] driver OK (ServiceBuilder).');
    return driver;
  } catch (err) {
    console.error('[selenium] falha no ServiceBuilder:', err?.stack || err);
  }

  // Fallback: sem ServiceBuilder (deixa o Selenium Manager/ PATH resolver)
  console.log('[selenium] tentando fallback sem ServiceBuilder...');
  const driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();
  console.log('[selenium] driver OK (fallback).');
  return driver;
}

async function clickElement(driver, el) {
  await driver.executeScript('arguments[0].scrollIntoView({block:"center"});', el);
  await driver.executeScript('arguments[0].click();', el);
}

async function screenshot(driver, filePath) {
  try {
    const b64 = await driver.takeScreenshot();
    await fs.writeFile(filePath, b64, 'base64');
    console.log('[screenshot]', filePath);
  } catch (e) {
    console.log('[screenshot] falhou:', e?.message || e);
  }
}

async function main() {
  console.log('=== editar.js START ===');
  const DATA_DIR = await getDataDir();
  await fs.mkdir(DATA_DIR, { recursive: true });
  const planilhaPath = path.join(DATA_DIR, 'planilha_filtrada.xlsx');
  const configPath   = path.join(DATA_DIR, 'config.json');

  console.log('[editar] configPath:', configPath);
  console.log('[editar] planilhaPath:', planilhaPath);

  if (!existsSync(configPath)) throw new Error(`Arquivo de configuração não encontrado: ${configPath}`);
  if (!existsSync(planilhaPath)) throw new Error(`Planilha não encontrada: ${planilhaPath}`);

  const config = await readJSON(configPath);
  const colunasSelecionadas = config.colunas;
  const atendenteEscolhido  = config.atendente;
  let colunaAlunoID         = config.coluna_alunoid;
  const colunaCPF           = config.coluna_cpf;

  const head_office   = config.head_office || process.env.HEAD_OFFICE;
  const emailAddress  = config.sponte_email || process.env.SPONTE_EMAIL;
  const passwordValue = config.sponte_password || process.env.SPONTE_PASSWORD;

  console.log('Sede:', head_office);
  console.log('→ Colunas:', colunasSelecionadas);
  console.log('→ Coluna AlunoID:', colunaAlunoID || '(usar CPF)');
  console.log('→ Atendente:', atendenteEscolhido);

  if (!head_office)      throw new Error('Sede (head_office) não informada no config/.env');
  if (!emailAddress || !passwordValue) throw new Error('Credenciais Sponte ausentes (email/senha).');

  console.log('[editar] lendo XLSX...');
  const buf = await readFile(planilhaPath);
  console.log('[editar] bytes XLSX:', buf.length);
  const wb = XLSX.read(buf, { type: 'buffer' });
  console.log('[editar] sheets:', wb.SheetNames);
  const sheetName = wb.SheetNames.includes('Filtrada') ? 'Filtrada' : wb.SheetNames[0];
  console.log('[editar] usando sheet:', sheetName);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: false });
  console.log('[editar] linhas:', rows.length);

  if (!rows.length) throw new Error('Planilha sem linhas.');
  const header = (rows[0] || []).map(x => String(x ?? '').trim());
  const dataRows = rows.slice(1);
  const colName = header[0];
  console.log('[editar] header[0]=', colName, ' — dataRows:', dataRows.length);

  const usarAlunoID = !!colunaAlunoID && colName === colunaAlunoID;
  const usarCPF     = !!colunaCPF && colName === colunaCPF && !usarAlunoID;
  console.log('[editar] usarAlunoID?', usarAlunoID, ' — usarCPF?', usarCPF);

  let soapClient = null;
  if (usarCPF) {
    console.log('[editar] criando SOAP client...');
    soapClient = await soap.createClientAsync(WSDL);
    console.log('[editar] SOAP client OK');
  }

  let driver;
  try {
    driver = await buildDriver();

    console.log('[selenium] abrindo Home...');
    await driver.get('https://www.sponteeducacional.net.br/Home.aspx');
    await screenshot(driver, path.join(DATA_DIR, '01_home.png'));

    console.log('[selenium] preenchendo login...');
    const email = await driver.wait(until.elementLocated(By.id('txtLogin')), 20000);
    await email.sendKeys(emailAddress);
    const password = await driver.findElement(By.id('txtSenha'));
    await password.sendKeys(passwordValue);
    const loginBtn = await driver.findElement(By.id('btnok'));
    await loginBtn.click();

    console.log('[selenium] aguardando pós-login...');
    await sleep(8000);
    await screenshot(driver, path.join(DATA_DIR, '02_pos_login.png'));

    const enterpriseSpan = await driver.findElement(By.id('ctl00_spnNomeEmpresa'));
    const enterprise = (await enterpriseSpan.getAttribute('innerText')).trim().replace(/\s+/g, '');
    console.log('[selenium] Empresa atual:', enterprise);

    const combinacoes = new Map([
      [['Aldeota', 'DIGITALCOLLEGESUL-74070'].toString(), [1, 'Acessando a sede Aldeota.']],
      [['Aldeota', 'DIGITALCOLLEGEBEZERRADEMENEZES-488365'].toString(), [1, 'Acessando a sede Aldeota.']],
      [['Sul',     'DIGITALCOLLEGEALDEOTA-72546'].toString(), [3, 'Acessando a sede Sul.']],
      [['Sul',     'DIGITALCOLLEGEBEZERRADEMENEZES-488365'].toString(), [3, 'Acessando a sede Sul.']],
      [['Bezerra', 'DIGITALCOLLEGEALDEOTA-72546'].toString(), [4, 'Acessando a sede Bezerra.']],
      [['Bezerra', 'DIGITALCOLLEGESUL-74070'].toString(), [4, 'Acessando a sede Bezerra.']],
      [['Aldeota', 'DIGITALCOLLEGEALDEOTA-72546'].toString(), [null, 'O script já está na Aldeota.']],
      [['Sul',     'DIGITALCOLLEGESUL-74070'].toString(), [null, 'O script já está no Sul.']],
      [['Bezerra', 'DIGITALCOLLEGEBEZERRADEMENEZES-488365'].toString(), [null, 'O script já está na Bezerra.']],
    ]);

    const key = [head_office, enterprise].toString();
    const [val, msg] = combinacoes.get(key) ?? [null, 'Ação não realizada: combinação não reconhecida.'];
    console.log('[selenium]', msg);

    if (val !== null) {
      console.log('[selenium] trocando sede...');
      await driver.executeScript(`$('#ctl00_hdnEmpresa').val(${val}); javascript:__doPostBack('ctl00$lnkChange','');`);
      await sleep(3000);
      await screenshot(driver, path.join(DATA_DIR, '03_troca_sede.png'));
    }

    const month = new Date().getMonth() + 1;
    const assuntoTexto = `COBRANÇA PARCELA - ${MESES_PT[month]}`;

    console.log('[editar] iniciando loop com', dataRows.length, 'linhas...');
    for (let i = 0; i < dataRows.length; i++) {
      console.log(`[loop] linha #${i + 2}`);
      try {
        let chave = String((dataRows[i] && dataRows[i][0]) ?? '').trim();
        if (!chave) {
          console.log(`[AVISO] Linha ${i + 2}: chave vazia. Pulando...`);
          continue;
        }

        let alunoId = null;
        if (usarAlunoID) {
          alunoId = chave.replace(/\.0+$/, '');
        } else if (usarCPF) {
          const cpf = chave.replace(/\D+/g, '');
          const cred = CREDENCIAIS[head_office];
          if (!cred) {
            console.log(`[AVISO] Sem credenciais da sede ${head_office}. Pulando linha ${i + 2}.`);
            continue;
          }
          alunoId = await getAlunoIdByCPF(cpf, cred.codigo_cliente, cred.token, soapClient);
          if (!alunoId) {
            console.log(`[AVISO] Não encontrei AlunoID para CPF ${cpf} (linha ${i + 2}). Pulando...`);
            continue;
          }
        }

        console.log(`--- Abrindo ficha do aluno ${alunoId} ---`);
        await driver.get(`https://www.sponteeducacional.net.br/SPCad/AlunoCadastro.aspx?cad=true&id=${alunoId}&ce=1`);
        await sleep(3000);
        await screenshot(driver, path.join(DATA_DIR, `aluno_${alunoId}_00.png`));

        const followTab = await driver.findElement(By.xpath("//*[@id='__tab_tab_TabPanel9']"));
        await clickElement(driver, followTab);
        await sleep(3000);

        const btnIncluir = await driver.findElement(By.id('tab_TabPanel9_btnIncluirFollowUp_div'));
        await clickElement(driver, btnIncluir);
        await sleep(3000);
        await screenshot(driver, path.join(DATA_DIR, `aluno_${alunoId}_01_follow.png`));

        const iframe = await driver.findElement(By.xpath("//iframe[contains(@src,'FollowUpCadastro.aspx')]"));
        await driver.switchTo().frame(iframe);
        await sleep(3000);

        await driver.wait(until.elementLocated(By.id('cmbAtendente')), 20000);
        await driver.findElement(By.id('cmbAtendente')).click();
        const optAt = await driver.findElement(
          By.xpath(`//select[@id='cmbAtendente']/option[normalize-space(.)="${atendenteEscolhido}"]`)
        );
        await optAt.click();
        await sleep(400);

        await driver.findElement(By.id('cmbTipoContato')).click();
        const optTc = await driver.findElement(
          By.xpath(`//select[@id='cmbTipoContato']/option[normalize-space(.)="WhatsApp"]`)
        );
        await optTc.click();
        await sleep(400);

        await driver.findElement(By.id('cmbTipoAgendamento')).click();
        const optTa = await driver.findElement(
          By.xpath(`//select[@id='cmbTipoAgendamento']/option[normalize-space(.)="Cobrança"]`)
        );
        await optTa.click();
        await sleep(400);

        await driver.findElement(By.id('cmbGrauInteresse')).click();
        const optGi = await driver.findElement(
          By.xpath(`//select[@id='cmbGrauInteresse']/option[normalize-space(.)="Muito Interessado"]`)
        );
        await optGi.click();
        await sleep(400);

        const assunto = await driver.findElement(By.id('txtAssunto'));
        await assunto.clear();
        await assunto.sendKeys(assuntoTexto);
        await sleep(400);
        await screenshot(driver, path.join(DATA_DIR, `aluno_${alunoId}_02_filled.png`));

        await driver.switchTo().defaultContent();
      } catch (rowErr) {
        console.error(`[ERRO] Linha ${i + 2}:`, rowErr?.message || rowErr);
        try { await driver.switchTo().defaultContent(); } catch {}
        await screenshot(driver, path.join(DATA_DIR, `erro_linha_${i + 2}.png`));
      }
    }

    console.log('[editar] loop concluído.');
  } catch (e) {
    console.error('Erro geral:', e?.message || e);
  } finally {
    try { await screenshot(driver, path.join(await getDataDir(), 'zz_final.png')); } catch {}
    try { await driver?.quit(); } catch {}
    console.log('=== editar.js END ===');
  }
}

main().catch((e) => {
  console.error('Fatal:', e);
  process.exit(1);
});
