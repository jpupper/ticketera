   const statusEl = document.getElementById('status');
const ticketeadoraSelect = document.getElementById('ticketeadora');
const printerSelect = document.getElementById('printer');
const printBtn = document.getElementById('printBtn');
const resetBtn = document.getElementById('resetBtn');
const refreshHistoryBtn = document.getElementById('refreshHistoryBtn');
const previewDate = document.getElementById('previewDate');
const previewTime = document.getElementById('previewTime');
const previewNumber = document.getElementById('previewNumber');
const previewLogo = document.getElementById('previewLogo');
const previewDefault = document.getElementById('previewDefault');
const previewVeladero = document.getElementById('previewVeladero');
const previewDateVeladero = document.getElementById('previewDateVeladero');
const previewTimeVeladero = document.getElementById('previewTimeVeladero');
const previewNumberVeladero = document.getElementById('previewNumberVeladero');
const nextNumber = document.getElementById('nextNumber');
const historyBody = document.getElementById('historyBody');
const totalCount = document.getElementById('totalCount');
let apiBase = window.location.origin;
let ticketeadoras = {};

function updatePreview() {
  const now = new Date();
  const fecha = now.toLocaleDateString('es-AR', { 
    day: '2-digit', 
    month: '2-digit', 
    year: 'numeric' 
  });
  const hora = now.getHours().toString().padStart(2, '0');
  const minutos = now.getMinutes().toString().padStart(2, '0');
  const segundos = now.getSeconds().toString().padStart(2, '0');
  
  // Actualizar preview default
  previewDate.textContent = fecha;
  previewTime.textContent = `${hora}:${minutos}:${segundos}`;
  
  // Actualizar preview veladero
  previewDateVeladero.textContent = fecha;
  previewTimeVeladero.textContent = `${hora}:${minutos}:${segundos}`;
}

// Actualizar preview cada segundo
setInterval(updatePreview, 1000);
updatePreview();

async function generateTicket() {
  statusEl.textContent = '';
  printBtn.disabled = true;
  try {
    const printer = printerSelect.value;
    const ticketeadora = ticketeadoraSelect.value;
    const resp = await fetch(apiBase + '/print', { 
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ printer, ticketeadora })
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      throw new Error(data.error || 'Error al imprimir');
    }
    statusEl.textContent = `¡Ticket #${data.participantNumber} generado!`;
    statusEl.style.color = '#198754';
    const nextNum = data.participantNumber + 1;
    previewNumber.textContent = nextNum;
    previewNumberVeladero.textContent = nextNum;
    nextNumber.textContent = nextNum.toString();
    await loadHistory();
  } catch (err) {
    statusEl.textContent = 'Falló la impresión: ' + err.message;
    statusEl.style.color = '#dc3545';
  } finally {
    printBtn.disabled = false;
  }
}

printBtn.addEventListener('click', generateTicket);

// Listener para tecla U
document.addEventListener('keydown', (e) => {
  if (e.key.toLowerCase() === 'u' && !printBtn.disabled) {
    generateTicket();
  }
});

async function loadPrinters() {
  try {
    let resp = await fetch(apiBase + '/printers');
    let ct = resp.headers.get('content-type') || '';
    if (!ct.includes('application/json')) {
      resp = await fetch('http://localhost:5450/printers');
      apiBase = 'http://localhost:5450';
      ct = resp.headers.get('content-type') || '';
    }
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.error || 'No se pudieron listar impresoras');
    let { printers, defaultPrinter } = data;
    if (!Array.isArray(printers)) printers = [];
    printers.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      printerSelect.appendChild(opt);
    });
    const defName = typeof defaultPrinter === 'string' ? defaultPrinter : (defaultPrinter && (defaultPrinter.name || defaultPrinter.PrinterName || defaultPrinter.deviceId || defaultPrinter.DeviceID));
    if (defName) {
      const found = Array.from(printerSelect.options).find(o => o.value === defName);
      if (found) printerSelect.value = defName;
    }
  } catch (err) {
    statusEl.textContent = 'No se pudo cargar la lista de impresoras: ' + err.message;
    statusEl.style.color = '#dc3545';
  }
}

// Función para reiniciar el contador
async function resetCounter() {
  if (!confirm('¿Estás seguro de reiniciar el contador a 0? El historial se mantendrá.')) {
    return;
  }
  
  statusEl.textContent = '';
  resetBtn.disabled = true;
  try {
    const ticketeadora = ticketeadoraSelect.value;
    const resp = await fetch(apiBase + '/reset', { 
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ticketeadora })
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      throw new Error(data.error || 'Error al reiniciar');
    }
    statusEl.textContent = 'Contador reiniciado exitosamente';
    statusEl.style.color = '#198754';
    nextNumber.textContent = '1';
    previewNumber.textContent = '1';
    previewNumberVeladero.textContent = '1';
    await loadHistory();
  } catch (err) {
    statusEl.textContent = 'Falló el reinicio: ' + err.message;
    statusEl.style.color = '#dc3545';
  } finally {
    resetBtn.disabled = false;
  }
}

// Función para cargar el historial
async function loadHistory() {
  try {
    const ticketeadora = ticketeadoraSelect.value;
    let resp = await fetch(apiBase + '/history?ticketeadora=' + ticketeadora);
    let ct = resp.headers.get('content-type') || '';
    if (!ct.includes('application/json')) {
      resp = await fetch('http://localhost:5450/history?ticketeadora=' + ticketeadora);
      ct = resp.headers.get('content-type') || '';
    }
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.error || 'No se pudo cargar historial');
    
    const history = data.history || [];
    totalCount.textContent = history.length;
    const nextNum = (data.lastNumber + 1).toString();
    nextNumber.textContent = nextNum;
    previewNumber.textContent = nextNum;
    previewNumberVeladero.textContent = nextNum;
    
    if (history.length === 0) {
      historyBody.innerHTML = '<tr><td colspan="3" class="no-data">No hay registros aún</td></tr>';
      return;
    }
    
    // Mostrar en orden inverso (más reciente primero)
    historyBody.innerHTML = history.reverse().map(item => `
      <tr>
        <td><strong>#${item.number}</strong></td>
        <td>${item.date}</td>
        <td>${item.time}</td>
      </tr>
    `).join('');
  } catch (err) {
    historyBody.innerHTML = '<tr><td colspan="3" class="no-data">Error al cargar historial</td></tr>';
    console.error('Error cargando historial:', err);
  }
}

// Función para cargar ticketeadoras
async function loadTicketeadoras() {
  try {
    let resp = await fetch(apiBase + '/ticketeadoras');
    let ct = resp.headers.get('content-type') || '';
    if (!ct.includes('application/json')) {
      resp = await fetch('http://localhost:5450/ticketeadoras');
      ct = resp.headers.get('content-type') || '';
    }
    const data = await resp.json();
    if (data.ok && data.ticketeadoras) {
      data.ticketeadoras.forEach(t => {
        ticketeadoras[t.id] = t;
      });
    }
  } catch (err) {
    console.error('Error cargando ticketeadoras:', err);
  }
}

// Función para cambiar ticketeadora
function onTicketeadoraChange() {
  const ticketeadora = ticketeadoraSelect.value;
  
  // Mostrar el preview correcto
  if (ticketeadora === 'veladero') {
    previewDefault.style.display = 'none';
    previewVeladero.style.display = 'block';
  } else {
    previewDefault.style.display = 'block';
    previewVeladero.style.display = 'none';
  }
  
  loadHistory();
}

ticketeadoraSelect.addEventListener('change', onTicketeadoraChange);
resetBtn.addEventListener('click', resetCounter);
refreshHistoryBtn.addEventListener('click', loadHistory);

window.addEventListener('DOMContentLoaded', async () => {
  await loadTicketeadoras();
  loadPrinters();
  loadHistory();
});