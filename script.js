const EXCEL_FILE = "Grilla 260126_Xtreme_Ultima.xlsx";
const SHEET_NAME = "MANO DE OBRA";
const TODOS = "Todos";
const ESTADOS_VACANTES = ["V", "V-BR", "V-MIX", "V-M", "V-S", "V-CS"];

let dataBase = [];
let chartTurno = null;
let chartCargo = null;
let chartEstado = null;
let chartGrupo = null;
let chartVacantesCargo = null;

function txt(v) {
  return String(v ?? "").trim();
}

function up(v) {
  return txt(v).toUpperCase();
}

function esVacante(estado) {
  return ESTADOS_VACANTES.includes(up(estado));
}

function normalizarEstado(estado, lic) {
  const est = up(estado);
  const licVal = up(lic);

  if (licVal === "L" || licVal === "LM") return "LM";
  if (est === "A") return "A";
  if (est === "P") return "P";
  if (esVacante(est)) return est;
  return est || "";
}

function esFilaValida(departamento, cargo, turno) {
  if (!txt(departamento) || !txt(cargo) || !txt(turno)) return false;

  const excluir = [
    "SUPERVISIÓN",
    "SOPORTE OPERACIÓN",
    "ADMINISTRADOR",
    "VACANTES DISPONIBLES CONTRATO ACTUAL",
    "PUESTOS ADICIONALES POR AMPLIACION"
  ];

  return !excluir.includes(up(cargo));
}

function setOptions(selectId, values, selected = TODOS) {
  const select = document.getElementById(selectId);
  select.innerHTML = "";

  const optionBase = document.createElement("option");
  optionBase.value = TODOS;
  optionBase.textContent = TODOS;
  select.appendChild(optionBase);

  [...new Set(values.filter(v => txt(v) !== ""))]
    .sort((a, b) => String(a).localeCompare(String(b), "es"))
    .forEach(value => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });

  const existe = [...select.options].some(o => o.value === selected);
  select.value = existe ? selected : TODOS;
}

function leerFiltros() {
  return {
    grupo: document.getElementById("filtroGrupo").value,
    departamento: document.getElementById("filtroDepartamento").value,
    cargo: document.getElementById("filtroCargo").value,
    turno: document.getElementById("filtroTurno").value,
    nombre: document.getElementById("filtroNombre").value,
    texto: up(document.getElementById("buscarTexto").value)
  };
}

function cumple(valor, filtro) {
  return filtro === TODOS || valor === filtro;
}

function filtrarDatos(data, filtros) {
  return data.filter(item => {
    const cumpleTexto =
      !filtros.texto ||
      up(item.nombre).includes(filtros.texto) ||
      up(item.cargo).includes(filtros.texto) ||
      up(item.departamento).includes(filtros.texto);

    return (
      cumple(item.grupo, filtros.grupo) &&
      cumple(item.departamento, filtros.departamento) &&
      cumple(item.cargo, filtros.cargo) &&
      cumple(item.turno, filtros.turno) &&
      cumple(item.nombre, filtros.nombre) &&
      cumpleTexto
    );
  });
}

function contarPorCampo(data, campo) {
  return data.reduce((acc, item) => {
    const key = item[campo] || "SIN DATO";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function contarEstado(data, estado) {
  return data.filter(x => x.estado === estado).length;
}

function contarVacantes(data) {
  return data.filter(x => esVacante(x.estado)).length;
}

function refrescarCombos() {
  const f = leerFiltros();
  const baseGrupo = dataBase.filter(item => cumple(item.grupo, f.grupo));

  setOptions("filtroGrupo", ["GRUPO 1", "GRUPO 2", "GRUPO 3", "GRUPO 4"], f.grupo);
  setOptions("filtroDepartamento", baseGrupo.map(x => x.departamento), f.departamento);
  setOptions("filtroCargo", baseGrupo.map(x => x.cargo), f.cargo);
  setOptions("filtroTurno", baseGrupo.map(x => x.turno), f.turno);
  setOptions("filtroNombre", baseGrupo.map(x => x.nombre), f.nombre);

  const vacantes = dataBase.filter(x => esVacante(x.estado));
  setOptions("filtroVacanteCargo", vacantes.map(x => x.cargo), document.getElementById("filtroVacanteCargo").value || TODOS);
}

function actualizarKPIs(data) {
  document.getElementById("kpiRegistros").textContent = data.length;
  document.getElementById("kpiAcreditados").textContent = contarEstado(data, "A");
  document.getElementById("kpiProceso").textContent = contarEstado(data, "P");
  document.getElementById("kpiVacantes").textContent = contarVacantes(data);
}

function actualizarGraficos(data) {
  const porTurno = contarPorCampo(data, "turno");
  const porCargo = contarPorCampo(data, "cargo");
  const porGrupo = contarPorCampo(data, "grupo");

  const acreditados = contarEstado(data, "A");
  const proceso = contarEstado(data, "P");
  const vacantes = contarVacantes(data);

  if (chartTurno) chartTurno.destroy();
  if (chartCargo) chartCargo.destroy();
  if (chartEstado) chartEstado.destroy();
  if (chartGrupo) chartGrupo.destroy();

  chartGrupo = new Chart(document.getElementById("chartGrupo"), {
    type: "bar",
    data: {
      labels: Object.keys(porGrupo),
      datasets: [{ label: "Dotación", data: Object.values(porGrupo) }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } }
    }
  });

  chartTurno = new Chart(document.getElementById("chartTurno"), {
    type: "bar",
    data: {
      labels: Object.keys(porTurno),
      datasets: [{ label: "Trabajadores", data: Object.values(porTurno) }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } }
    }
  });

  chartCargo = new Chart(document.getElementById("chartCargo"), {
    type: "bar",
    data: {
      labels: Object.keys(porCargo),
      datasets: [{ label: "Trabajadores", data: Object.values(porCargo) }]
    },
    options: {
      responsive: true,
      indexAxis: "y",
      plugins: { legend: { display: false } }
    }
  });

  chartEstado = new Chart(document.getElementById("chartEstado"), {
    type: "doughnut",
    data: {
      labels: ["Acreditados", "En proceso", "Vacantes"],
      datasets: [{ data: [acreditados, proceso, vacantes] }]
    },
    options: {
      responsive: true,
      plugins: { legend: { position: "bottom" } }
    }
  });
}

function renderDetalleGrupo(data) {
  const grupoDetalle = document.getElementById("detalleGrupoSelect").value;
  const tbody = document.getElementById("tablaDetalleGrupo");
  tbody.innerHTML = "";

  const rows = data.filter(x => x.grupo === grupoDetalle);

  rows.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.departamento}</td>
      <td>${item.cargo}</td>
      <td>${item.grupo}</td>
      <td>${item.nombre}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderVacantesCargo() {
  const filtroCargoVacante = document.getElementById("filtroVacanteCargo").value;
  const vacantes = dataBase.filter(x => esVacante(x.estado));
  const vacantesFiltradas = filtroCargoVacante === TODOS
    ? vacantes
    : vacantes.filter(x => x.cargo === filtroCargoVacante);

  const tbody = document.getElementById("tablaVacantesCargo");
  tbody.innerHTML = "";

  vacantesFiltradas.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.departamento}</td>
      <td>${item.cargo}</td>
      <td>${item.grupo}</td>
      <td>${item.nombre}</td>
    `;
    tbody.appendChild(tr);
  });

  const porCargoVacante = contarPorCampo(vacantesFiltradas, "cargo");

  if (chartVacantesCargo) chartVacantesCargo.destroy();

  chartVacantesCargo = new Chart(document.getElementById("chartVacantesCargo"), {
    type: "bar",
    data: {
      labels: Object.keys(porCargoVacante),
      datasets: [{ label: "Vacantes", data: Object.values(porCargoVacante) }]
    },
    options: {
      responsive: true,
      indexAxis: "y",
      plugins: { legend: { display: false } }
    }
  });
}

function renderDashboard() {
  const filtros = leerFiltros();
  const data = filtrarDatos(dataBase, filtros);

  actualizarKPIs(data);
  actualizarGraficos(data);
  renderDetalleGrupo(data);
  renderVacantesCargo();
}

function aplicarFiltros() {
  refrescarCombos();
  renderDashboard();
}

function limpiarFiltros() {
  document.getElementById("buscarTexto").value = "";

  setOptions("filtroGrupo", ["GRUPO 1", "GRUPO 2", "GRUPO 3", "GRUPO 4"], TODOS);
  setOptions("filtroDepartamento", dataBase.map(x => x.departamento), TODOS);
  setOptions("filtroCargo", dataBase.map(x => x.cargo), TODOS);
  setOptions("filtroTurno", dataBase.map(x => x.turno), TODOS);
  setOptions("filtroNombre", dataBase.map(x => x.nombre), TODOS);
  setOptions("filtroVacanteCargo", dataBase.filter(x => esVacante(x.estado)).map(x => x.cargo), TODOS);

  document.getElementById("detalleGrupoSelect").value = "GRUPO 1";

  renderDashboard();
}

async function cargarExcel() {
  try {
    const response = await fetch(EXCEL_FILE);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[SHEET_NAME];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    const resultado = [];

    for (let i = 2; i < 128 && i < rows.length; i++) {
      const row = rows[i];

      const departamento = txt(row[1]);
      const cargo = txt(row[2]);
      const turno = txt(row[3]);

      if (!esFilaValida(departamento, cargo, turno)) continue;

      const grupos = [
        { grupo: "GRUPO 1", nombre: row[4], estado: row[5], lic: "" },
        { grupo: "GRUPO 2", nombre: row[6], estado: row[7], lic: row[8] },
        { grupo: "GRUPO 3", nombre: row[9], estado: row[10], lic: row[11] },
        { grupo: "GRUPO 4", nombre: row[12], estado: row[13], lic: "" }
      ];

      grupos.forEach(g => {
        const nombre = txt(g.nombre);
        const estado = normalizarEstado(g.estado, g.lic);

        if (!nombre && !estado) return;

        resultado.push({
          departamento,
          cargo,
          turno,
          grupo: g.grupo,
          estado,
          nombre: nombre || "VACANTE"
        });
      });
    }

    dataBase = resultado;

    setOptions("filtroGrupo", ["GRUPO 1", "GRUPO 2", "GRUPO 3", "GRUPO 4"], TODOS);
    setOptions("filtroDepartamento", dataBase.map(x => x.departamento), TODOS);
    setOptions("filtroCargo", dataBase.map(x => x.cargo), TODOS);
    setOptions("filtroTurno", dataBase.map(x => x.turno), TODOS);
    setOptions("filtroNombre", dataBase.map(x => x.nombre), TODOS);
    setOptions("filtroVacanteCargo", dataBase.filter(x => esVacante(x.estado)).map(x => x.cargo), TODOS);

    document.getElementById("detalleGrupoSelect").value = "GRUPO 1";

    renderDashboard();

    document.getElementById("statusBox").textContent =
      `Conectado. Registros cargados: ${dataBase.length}`;
  } catch (error) {
    console.error(error);
    document.getElementById("statusBox").textContent =
      "Error al leer Excel. Revisar nombre de archivo, hoja o columnas.";
  }
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("btnAplicar").addEventListener("click", aplicarFiltros);
  document.getElementById("btnLimpiar").addEventListener("click", limpiarFiltros);

  ["filtroGrupo", "filtroDepartamento", "filtroCargo", "filtroTurno", "filtroNombre"].forEach(id => {
    document.getElementById(id).addEventListener("change", aplicarFiltros);
  });

  document.getElementById("buscarTexto").addEventListener("input", aplicarFiltros);
  document.getElementById("detalleGrupoSelect").addEventListener("change", renderDashboard);
  document.getElementById("filtroVacanteCargo").addEventListener("change", renderVacantesCargo);

  cargarExcel();
});