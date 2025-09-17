const el = (id) => document.getElementById(id);

const logs = el('logs'); // <pre id="logs">
function appendLog(s) {
  if (!s) return;
  logs.textContent += String(s);
  logs.scrollTop = logs.scrollHeight;
}

function base64ToUint8Array(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

let workbook = null;
let sheetData = [];
let headerIndex = 0;

const btnPick  = el('pick');
const lblPicked= el('picked');
const selSheet = el('sheet');
const selHeader= el('header');
const selAluno = el('alunoid');
const selCPF   = el('cpf');
const selAtend = el('atendente');
const selSede  = el('sede');
const inpEmail = el('email');
const inpSenha = el('senha');
const btnGo    = el('go');
const lblStatus= el('status');

btnPick.addEventListener('click', async () => {
  // sanity check do preload:
  if (!window.api) {
    console.error('preload não disponível');
    alert('Falha ao inicializar (preload). Veja o console.');
    return;
  }

  const res = await window.api.selectXlsx();
  if (!res) return;

  lblPicked.textContent = res.filePath;

  const bytes = base64ToUint8Array(res.base64);
  workbook = window.XLSX.read(bytes, { type: 'array' }); // << use window.XLSX

  // popula abas
  selSheet.innerHTML = '';
  workbook.SheetNames.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    selSheet.appendChild(opt);
  });

  if (workbook.SheetNames.length) {
    selSheet.value = workbook.SheetNames[0];
    loadSheet();
  }
});

selSheet.addEventListener('change', loadSheet);

function previewLine(arr, maxCols = 6, maxChars = 80) {
  const v = (arr || []).slice(0, maxCols).map(v => String(v ?? '')).join(' | ');
  return v.length > maxChars ? v.slice(0, maxChars - 3) + '...' : v;
}

function loadSheet() {
  const name = selSheet.value;
  const sheet = workbook.Sheets[name];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  sheetData = rows;
  selHeader.innerHTML = '';
  rows.forEach((r, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `Linha ${i + 1}: ${previewLine(r)}`;
    selHeader.appendChild(opt);
  });
  selHeader.value = 0;
  applyHeader();
}

selHeader.addEventListener('change', applyHeader);

function applyHeader() {
  headerIndex = parseInt(selHeader.value, 10) || 0;
  const header = (sheetData[headerIndex] || []).map(x => String(x ?? '').trim());
  const cols = header.filter(Boolean);

  const fill = (select) => {
    select.innerHTML = '';
    cols.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c; opt.textContent = c;
      select.appendChild(opt);
    });
  };
  fill(selAluno);
  fill(selCPF);

  const lower = cols.map(c => c.toLowerCase());
  const idxAluno = lower.findIndex(c => c.includes('alunoid') || c === 'id' || c === 'aluno_id' || c === 'calunoid');
  if (idxAluno >= 0) selAluno.value = cols[idxAluno];
  const idxCPF = lower.findIndex(c => c === 'cpf');
  if (idxCPF >= 0) selCPF.value = cols[idxCPF];

  btnGo.disabled = false;
}

btnGo.addEventListener('click', async () => {
  try {
    lblStatus.textContent = 'Gerando arquivos...';

    const header = (sheetData[headerIndex] || []).map(x => String(x ?? '').trim());
    const rows   = sheetData.slice(headerIndex + 1);

    const colAluno = selAluno.value?.trim();
    const colCPF   = selCPF.value?.trim();
    const usarAluno = !!colAluno && header.includes(colAluno);
    const usarCPF   = !usarAluno && !!colCPF && header.includes(colCPF);

    if (!usarAluno && !usarCPF) {
      alert('Escolha a coluna de AlunoID ou, se não houver, a coluna de CPF.');
      return;
    }

    const colName = usarAluno ? colAluno : colCPF;
    const idxCol  = header.indexOf(colName);
    const values  = rows.map(r => r[idxCol]);

    const normalized = values
      .map(v => String(v ?? '').trim())
      .map(v => usarAluno ? v.replace(/\.0+$/,'') : v.replace(/\D+/g,''))
      .filter(v => v.length > 0);

    const uniq = Array.from(new Set(normalized));

    const aoa = [[colName], ...uniq.map(v => [v])];
    const ws  = window.XLSX.utils.aoa_to_sheet(aoa);
    const wb  = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Filtrada');

    const xbase64 = window.XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

    const config = {
      colunas: [colName],
      atendente: selAtend.value,
      aba: selSheet.value,
      head_office: selSede.value,
      sponte_email: inpEmail.value.trim(),
      sponte_password: inpSenha.value.trim(),
      ...(usarAluno ? { coluna_alunoid: colName } : { coluna_cpf: colName })
    };

    const saved = await window.api.saveOutputs(xbase64, config);
    lblStatus.textContent = `Salvo em: ${saved.TEMP_PLANILHA}`;

    await window.api.runEditar();
    lblStatus.textContent = `Executado editar.`;

  } catch (e) {
    console.error(e);
    alert('Falha: ' + (e?.message || e));
    lblStatus.textContent = '';
  }
});

window.api.onEditarLog((line) => appendLog(line));
window.api.onEditarExit((msg) => appendLog(msg));
