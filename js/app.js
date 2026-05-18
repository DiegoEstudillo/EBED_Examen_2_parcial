const inputArchivo  = document.getElementById("inputArchivo");
const nombreArchivo = document.getElementById("nombre-archivo");
const zonaCarga     = document.getElementById("zona-carga");
const panel         = document.getElementById("panel");
const overlay       = document.getElementById("modal-overlay");
const modalTitulo   = document.getElementById("modal-titulo");
const modalCuerpo   = document.getElementById("modal-cuerpo");

let chartInstancia = null;
let datos = [];

function procesarExcel(buffer, nombre) {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet    = workbook.Sheets[workbook.SheetNames[0]];
  const rows     = XLSX.utils.sheet_to_json(sheet);

  datos = rows.map(r => ({
    peso:      parseFloat(r["Peso"]) || null,
    altura:    parseFloat(r["Altura"]) || null,
    velocidad: (r["Velocidad"] === "NA" || r["Velocidad"] === undefined) ? null : parseFloat(r["Velocidad"]),
    color:     String(r["Color"] ?? "NA").trim()
  })).filter(d => d.peso !== null);

  nombreArchivo.textContent = `✔ ${nombre} — ${datos.length} registros cargados`;
  zonaCarga.style.display = "none";
  panel.style.display = "block";
}

window.addEventListener("DOMContentLoaded", () => {
  fetch("datos.xlsx")
    .then(res => {
      if (!res.ok) throw new Error("No se pudo cargar");
      return res.arrayBuffer();
    })
    .then(buffer => procesarExcel(buffer, "datos.xlsx"))
    .catch(() => {
      nombreArchivo.textContent = "Selecciona el archivo manualmente";
    });
});

inputArchivo.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => procesarExcel(ev.target.result, file.name);
  reader.readAsArrayBuffer(file);
});

const limpiar = arr => arr.filter(x => x !== null && !isNaN(x));

function media(arr) {
  const v = limpiar(arr);
  return v.reduce((a, b) => a + b, 0) / v.length;
}

function mediana(arr) {
  const v = limpiar(arr).sort((a, b) => a - b);
  const m = Math.floor(v.length / 2);
  return v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2;
}

function moda(arr) {
  const v = limpiar(arr);
  const freq = {};
  v.forEach(x => freq[x] = (freq[x] || 0) + 1);
  const max = Math.max(...Object.values(freq));
  return Object.keys(freq).filter(k => freq[k] === max).join(", ");
}

function calcularIntervalos(arr, n = 5) {
  const v = limpiar(arr).sort((a, b) => a - b);
  const min = v[0], max = v[v.length - 1], amp = (max - min) / n;
  return Array.from({ length: n }, (_, i) => {
    const li = min + i * amp, ls = min + (i + 1) * amp;
    return {
      etiqueta: `${li.toFixed(1)}–${ls.toFixed(1)}`,
      marca:    ((li + ls) / 2).toFixed(2),
      freq:     v.filter(p => i < n - 1 ? p >= li && p < ls : p >= li && p <= ls).length
    };
  });
}

function freqColor() {
  const claves    = ["Amarillo", "Verde", "Blanco", "NA"];
  const etiquetas = ["Amarillo", "Verde", "Blanco", "Sin dato"];
  const n   = datos.length;
  const abs = claves.map(c => datos.filter(d => d.color === c).length);
  const rel = abs.map(f => (f / n * 100).toFixed(1));
  let acum  = 0;
  const acums = abs.map(f => { acum += f; return acum; });
  return { etiquetas, abs, rel, acums, n };
}

function abrirModal(titulo, html, renderChart) {
  modalTitulo.textContent = titulo;
  modalCuerpo.innerHTML   = html;
  overlay.classList.add("activo");
  if (chartInstancia) { chartInstancia.destroy(); chartInstancia = null; }
  if (renderChart) setTimeout(renderChart, 0);
}

function cerrarModal() {
  overlay.classList.remove("activo");
  if (chartInstancia) { chartInstancia.destroy(); chartInstancia = null; }
}

overlay.addEventListener("click", e => { if (e.target === overlay) cerrarModal(); });
document.getElementById("btn-cerrar").addEventListener("click", cerrarModal);

document.getElementById("btn-frec-abs").addEventListener("click", () => {
  const { etiquetas, abs, rel, acums, n } = freqColor();
  let rows = `<tr><th>Color</th><th>Frec. Absoluta</th><th>Frec. Relativa (%)</th><th>Frec. Acumulada</th></tr>`;
  etiquetas.forEach((e, i) => {
    rows += `<tr><td>${e}</td><td>${abs[i]}</td><td>${rel[i]}%</td><td>${acums[i]}</td></tr>`;
  });
  rows += `<tr class="total-row"><td>Total</td><td>${n}</td><td>100%</td><td>—</td></tr>`;
  abrirModal("Frecuencias Absolutas — Color",
    `<table>${rows}</table><br><canvas id="chart-modal" height="160"></canvas>`,
    () => {
      chartInstancia = new Chart(document.getElementById("chart-modal"), {
        type: "bar",
        data: {
          labels: etiquetas,
          datasets: [{
            data: abs,
            backgroundColor: ["#f1c40f","#2ecc71","#bdc3c7","#95a5a6"],
            borderColor:     ["#d4ac0d","#27ae60","#95a5a6","#7f8c8d"],
            borderWidth: 2
          }]
        },
        options: {
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
        }
      });
    }
  );
});

document.getElementById("btn-pastel").addEventListener("click", () => {
  const { etiquetas, abs, rel } = freqColor();
  abrirModal("Diagrama de Pastel — Frecuencia Relativa",
    `<div style="max-width:360px;margin:0 auto"><canvas id="chart-modal"></canvas></div>`,
    () => {
      chartInstancia = new Chart(document.getElementById("chart-modal"), {
        type: "pie",
        data: {
          labels: etiquetas.map((e, i) => `${e} (${rel[i]}%)`),
          datasets: [{ data: abs, backgroundColor: ["#f1c40f","#2ecc71","#bdc3c7","#95a5a6"] }]
        },
        options: { animation: false }
      });
    }
  );
});

document.getElementById("btn-acumulada").addEventListener("click", () => {
  const ints = calcularIntervalos(datos.map(d => d.peso));
  const n    = datos.length;
  let acumP  = 0;
  let rows   = `<tr><th>Intervalo</th><th>Marca</th><th>Frec. Abs.</th><th>Frec. Rel. (%)</th><th>Frec. Acum.</th></tr>`;
  ints.forEach(({ etiqueta, marca, freq }) => {
    acumP += freq;
    rows += `<tr><td>${etiqueta}</td><td>${marca}</td><td>${freq}</td><td>${(freq/n*100).toFixed(1)}%</td><td>${acumP}</td></tr>`;
  });
  rows += `<tr class="total-row"><td>Total</td><td>—</td><td>${n}</td><td>100%</td><td>—</td></tr>`;
  abrirModal("Frecuencias Acumuladas — Peso", `<table>${rows}</table>`);
});

document.getElementById("btn-poligono").addEventListener("click", () => {
  const ints = calcularIntervalos(datos.map(d => d.peso));
  abrirModal("Poligono de Frecuencias — Peso",
    `<canvas id="chart-modal" height="180"></canvas>`,
    () => {
      chartInstancia = new Chart(document.getElementById("chart-modal"), {
        type: "line",
        data: {
          labels: ints.map(i => i.marca),
          datasets: [{
            label: "Frecuencia",
            data: ints.map(i => i.freq),
            borderColor: "#1abc9c",
            backgroundColor: "rgba(26,188,156,0.15)",
            pointRadius: 5,
            fill: true,
            tension: 0.3
          }]
        },
        options: {
          animation: false,
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
        }
      });
    }
  );
});

document.getElementById("btn-estadisticas").addEventListener("click", () => {
  const variables = [
    { nombre: "Peso",      arr: datos.map(d => d.peso) },
    { nombre: "Altura",    arr: datos.map(d => d.altura) },
    { nombre: "Velocidad", arr: datos.map(d => d.velocidad) }
  ];
  let cards = "";
  variables.forEach(({ nombre, arr }) => {
    cards += `
      <div class="stat-card">
        <h3>${nombre}</h3>
        <table>
          <tr><td>Media</td><td>${media(arr).toFixed(2)}</td></tr>
          <tr><td>Mediana</td><td>${mediana(arr).toFixed(2)}</td></tr>
          <tr><td>Moda</td><td>${moda(arr)}</td></tr>
        </table>
      </div>`;
  });
  abrirModal("Media, Mediana y Moda", `<div class="stats-grid">${cards}</div>`);
});

document.getElementById("btn-datos").addEventListener("click", () => {
  let rows = `<tr><th>#</th><th>Peso</th><th>Altura</th><th>Velocidad</th><th>Color</th></tr>`;
  datos.forEach((d, i) => {
    rows += `<tr><td>${i+1}</td><td>${d.peso}</td><td>${d.altura}</td><td>${d.velocidad ?? 'NA'}</td><td>${d.color}</td></tr>`;
  });
  abrirModal("Datos Originales", `<table>${rows}</table>`);
});
