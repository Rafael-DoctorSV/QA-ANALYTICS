// Global State
let rawData = [];
let processedData = [];
let allEvaluations = []; // Nuevo array para guardar TODAS las evaluaciones (filas)
let chartInstance = null;

// DOM Elements
const loadingContainer = document.getElementById('loading-container');
const dashboardContent = document.getElementById('dashboard-content');
const dateFromInput = document.getElementById('date-from');
const dateToInput = document.getElementById('date-to');
const canalSelect = document.getElementById('canal-select');
const tipoSelect = document.getElementById('tipo-select');
const supervisorSelect = document.getElementById('supervisor-select');
const agenteSelect = document.getElementById('agente-select');

// KPI Elements
const kpiTotal = document.getElementById('kpi-total-incumplimientos');
const kpiAgentes = document.getElementById('kpi-total-agentes');

document.addEventListener('DOMContentLoaded', async () => {
    rawData = []; 
    allEvaluations = []; 
    
    const readExcelFromUrl = async (url) => {
        try {
            const response = await fetch(url);
            if (!response.ok) return null;
            const arrayBuffer = await response.arrayBuffer();
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            return XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: "" });
        } catch (e) {
            console.warn(`Could not load ${url}`, e);
            return null;
        }
    };

    // Descargar archivos automáticamente
    const dataTel = await readExcelFromUrl('./datos_telefonico.xlsx');
    if (dataTel) processRawData(dataTel, "Telefónico");
    
    const dataDig = await readExcelFromUrl('./datos_digital.xlsx');
    if (dataDig) processRawData(dataDig, "Digital");

    if (rawData.length === 0) {
        if(loadingContainer) loadingContainer.innerHTML = '<div style="text-align:center; padding: 50px;"><h3>No se encontraron los datos</h3><p>Asegúrate de que los archivos "datos_telefonico.xlsx" y "datos_digital.xlsx" existan en la misma ruta.</p></div>';
    } else {
        if(loadingContainer) loadingContainer.style.display = 'none';
        dashboardContent.style.display = 'block';
        applyFilters();
    }
});

// Filtros automáticos: Se aplican en cuanto se cambia cualquier valor
[dateFromInput, dateToInput, canalSelect, tipoSelect, supervisorSelect, agenteSelect].forEach(input => {
    if(input) {
        input.addEventListener('change', () => {
            if (typeof applyFilters === 'function') applyFilters();
        });
    }
});

// Colors for Chart
const chartColors = [
    '#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b', '#10b981', '#ef4444', '#06b6d4',
    '#6366f1', '#14b8a6', '#f43f5e', '#84cc16', '#eab308', '#d946ef', '#0ea5e9'
];

function processRawData(data, canalName) {
    if (data.length === 0) return;
    
    const keys = Object.keys(data[0]);
    
    // Find Date and Agent column dynamically
    let colFecha = keys[0]; // Usuario dijo "Puedes ocupar la fecha de la primera columna"
    let colAgente = keys.find(k => k.toLowerCase().includes("agente"));
    let colSupervisor = keys.find(k => k.toLowerCase().includes("supervisor"));
    
    // Encontrar rangos exactos por nombre de columna
    const findBoundaries = (startStr, endStr) => {
        const start = keys.findIndex(k => k.toLowerCase().includes(startStr.toLowerCase()));
        const end = keys.findIndex(k => k.toLowerCase().includes(endStr.toLowerCase()));
        if (start !== -1 && end !== -1) {
            return { min: Math.min(start, end), max: Math.max(start, end) };
        }
        return { min: -1, max: -1 };
    };

    let autoZeroStartStr = '';
    let autoZeroEndStr = '';
    let autoFailStartStr = '';
    let autoFailEndStr = '';

    if (canalName === 'Telefónico') {
        autoZeroStartStr = 'no demuestra escucha activa';
        autoZeroEndStr = 'no respeta el tiempo prudencial para finalizar la llamada tras el cierre';
        autoFailStartStr = 'habla negativamente del servicio';
        autoFailEndStr = 'agente corta la llamada';
    } else { // Digital
        autoZeroStartStr = 'lectura comprensiva y verificaci';
        autoZeroEndStr = 'no respeta el tiempo prudencial para cerrar la interacción por inactividad';
        autoFailStartStr = 'habla negativamente del servicio';
        autoFailEndStr = 'no transfiere contacto al destino correcto';
    }

    const autoZeroRange = findBoundaries(autoZeroStartStr, autoZeroEndStr);
    const autoFailRange = findBoundaries(autoFailStartStr, autoFailEndStr);
    
    // Identify criteria columns by excluding known info columns
    const excludeKeywords = ['fecha', 'id', 'agente', 'supervisor', 'analista', 'puntaje', 'conclusión', 'conclusion'];
    const criteriaColumns = keys.filter(key => {
        const lowerKey = key.toLowerCase();
        return !excludeKeywords.some(kw => lowerKey.includes(kw));
    });

    data.forEach(row => {
        const agenteRaw = row[colAgente];
        if (!agenteRaw) return; // Skip empty rows
        
        const agente = String(agenteRaw).trim();
        const supervisorRaw = colSupervisor ? row[colSupervisor] : null;
        const supervisor = supervisorRaw ? String(supervisorRaw).trim() : "Desconocido";
        let rawDate = row[colFecha];
        let dateObj = null;

        // Parse Date
        if (rawDate) {
            if (rawDate instanceof Date) {
                dateObj = new Date(rawDate.getUTCFullYear(), rawDate.getUTCMonth(), rawDate.getUTCDate());
            } else if (typeof rawDate === 'string') {
                const parts = rawDate.split(/[-/]/);
                if (parts.length >= 3) {
                    // Assuming DD/MM/YYYY
                    dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
                    if (isNaN(dateObj.getTime())) {
                        dateObj = new Date(rawDate); // Fallback
                    }
                } else {
                    dateObj = new Date(rawDate);
                }
            }
        }
        
        // Count Auto Fails for this specific evaluation based on strict column indices
        let autoFailsCount = 0;
        
        // Check for each criteria if there is a failure
        let processedTiposForRow = new Set();
        
        criteriaColumns.forEach(criteria => {
            const colIndex = keys.indexOf(criteria);
            const val = String(row[criteria] || "").trim().toUpperCase();
            
            if (val === "NO" || val === "INCUMPLE") {
                let tipoLimpio = criteria.replace(/_\d+$/, '').trim();
                
                // Omite columnas duplicadas en la misma fila (ej: Uso de lenguaje)
                if (processedTiposForRow.has(tipoLimpio)) return;
                processedTiposForRow.add(tipoLimpio);
                let clase = 'Normal'; // Comportamientos por defecto

                // Excepciones explícitas obligatorias
                const isLenguaje = tipoLimpio.toLowerCase().includes('uso de lenguaje') || tipoLimpio.toLowerCase().includes('uso de vocabulario');
                const isTMO = tipoLimpio.toLowerCase().includes('tiempo medio de operación');

                if (isLenguaje || isTMO) {
                    clase = 'Normal'; // Siempre son comportamientos
                } else if (val === "NO") {
                    clase = 'Normal'; // Comportamientos son todos los que digan "NO"
                } else if (val === "INCUMPLE") {
                    // Validar rangos
                    const isAutoZero = autoZeroRange.min !== -1 && colIndex >= autoZeroRange.min && colIndex <= autoZeroRange.max;
                    const isAutoFail = autoFailRange.min !== -1 && colIndex >= autoFailRange.min && colIndex <= autoFailRange.max;
                    
                    if (isAutoZero) {
                        clase = 'Auto Zero';
                    } else if (isAutoFail) {
                        clase = 'Auto Fail';
                    }
                }

                if (clase === 'Auto Fail') {
                    autoFailsCount++;
                }
                
                rawData.push({
                    fechaStr: rawDate,
                    fechaObj: dateObj,
                    agente: agente,
                    supervisor: supervisor,
                    canal: canalName,
                    tipo: tipoLimpio,
                    clase: clase
                });
            }
        });
        
        // Guardar la evaluación completa para los KPIs de "Total" y "Agentes Evaluados"
        allEvaluations.push({
            fechaObj: dateObj,
            agente: agente,
            supervisor: supervisor,
            canal: canalName,
            autoFails: autoFailsCount
        });
    });
}

function getFilteredData(excludeFilterName = null) {
    const fromDateStr = dateFromInput.value;
    const toDateStr = dateToInput.value;
    const selectedCanal = canalSelect.value;
    const selectedTipo = tipoSelect.value;
    const selectedSupervisor = supervisorSelect.value;
    const selectedAgente = agenteSelect.value;

    let fromDate = null;
    if (fromDateStr) {
        const parts = fromDateStr.split('-'); // El input de fecha devuelve YYYY-MM-DD
        fromDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }

    let toDate = null;
    if (toDateStr) {
        const parts = toDateStr.split('-');
        toDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        toDate.setHours(23, 59, 59, 999);
    }

    return rawData.filter(item => {
        let matchDate = true;
        if (item.fechaObj && !isNaN(item.fechaObj.getTime())) {
            if (fromDate && item.fechaObj < fromDate) matchDate = false;
            if (toDate && item.fechaObj > toDate) matchDate = false;
        }

        let matchCanal = true;
        if (excludeFilterName !== 'canal' && selectedCanal !== 'all' && item.canal !== selectedCanal) {
            matchCanal = false;
        }

        let matchTipo = true;
        if (excludeFilterName !== 'tipo' && selectedTipo !== 'all' && item.tipo !== selectedTipo) {
            matchTipo = false;
        }
        
        let matchSupervisor = true;
        if (excludeFilterName !== 'supervisor' && selectedSupervisor !== 'all' && item.supervisor !== selectedSupervisor) {
            matchSupervisor = false;
        }

        let matchAgente = true;
        if (excludeFilterName !== 'agente' && selectedAgente !== 'all' && item.agente !== selectedAgente) {
            matchAgente = false;
        }

        return matchDate && matchCanal && matchTipo && matchSupervisor && matchAgente;
    });
}

function getFilteredEvaluations(excludeFilterName = null) {
    const fromDateStr = dateFromInput.value;
    const toDateStr = dateToInput.value;
    const selectedCanal = canalSelect.value;
    const selectedSupervisor = supervisorSelect.value;
    const selectedAgente = agenteSelect.value;

    let fromDate = null;
    if (fromDateStr) {
        const parts = fromDateStr.split('-');
        fromDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }

    let toDate = null;
    if (toDateStr) {
        const parts = toDateStr.split('-');
        toDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        toDate.setHours(23, 59, 59, 999);
    }

    return allEvaluations.filter(item => {
        let matchDate = true;
        if (item.fechaObj && !isNaN(item.fechaObj.getTime())) {
            if (fromDate && item.fechaObj < fromDate) matchDate = false;
            if (toDate && item.fechaObj > toDate) matchDate = false;
        }

        let matchCanal = true;
        if (excludeFilterName !== 'canal' && selectedCanal !== 'all' && item.canal !== selectedCanal) {
            matchCanal = false;
        }
        
        let matchSupervisor = true;
        if (excludeFilterName !== 'supervisor' && selectedSupervisor !== 'all' && item.supervisor !== selectedSupervisor) {
            matchSupervisor = false;
        }

        let matchAgente = true;
        if (excludeFilterName !== 'agente' && selectedAgente !== 'all' && item.agente !== selectedAgente) {
            matchAgente = false;
        }

        return matchDate && matchCanal && matchSupervisor && matchAgente;
    });
}

function updateSelectFilter(selectElement, dataSet, defaultText) {
    const currentValue = selectElement.value;
    selectElement.innerHTML = `<option value="all">${defaultText}</option>`;
    
    Array.from(dataSet).sort().forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.textContent = item;
        selectElement.appendChild(option);
    });

    if (currentValue !== 'all' && !dataSet.has(currentValue)) {
        selectElement.value = 'all';
        return true; // El valor seleccionado ya no es válido bajo los nuevos filtros
    } else {
        selectElement.value = currentValue;
        return false;
    }
}

function updateDynamicFilters() {
    let changed = false;

    const dataForCanal = getFilteredData('canal');
    const canalesSet = new Set(dataForCanal.map(item => item.canal));
    if (updateSelectFilter(canalSelect, canalesSet, "Todos los canales")) changed = true;

    const dataForTipo = getFilteredData('tipo');
    const tiposSet = new Set(dataForTipo.map(item => item.tipo));
    if (updateSelectFilter(tipoSelect, tiposSet, "Todos los tipos")) changed = true;
    
    const dataForSupervisor = getFilteredData('supervisor');
    const supervisoresSet = new Set(dataForSupervisor.map(item => item.supervisor));
    if (updateSelectFilter(supervisorSelect, supervisoresSet, "Todos los supervisores")) changed = true;

    const dataForAgente = getFilteredData('agente');
    const agentesSet = new Set(dataForAgente.map(item => item.agente));
    if (updateSelectFilter(agenteSelect, agentesSet, "Todos los agentes")) changed = true;

    return changed;
}

function applyFilters() {
    processedData = getFilteredData();
    
    // Actualizamos las opciones de los dropdowns en base a lo filtrado
    // Si un filtro tenía seleccionado un valor que dejó de existir, se resetea a "Todos" y se re-calcula.
    if (updateDynamicFilters()) {
        processedData = getFilteredData();
    }
    
    updateDashboard();
}

function updateDashboard() {
    updateKPIs();
    updateChart();
}

function updateKPIs() {
    kpiTotal.textContent = processedData.length;

    // Utilizamos filteredEvals para que los KPIs reflejen el total de evaluaciones y agentes independientes de los incumplimientos
    const filteredEvals = getFilteredEvaluations();
    
    const agentesSet = new Set(filteredEvals.map(item => item.agente));
    kpiAgentes.textContent = agentesSet.size;
    
    const totalAutoFails = filteredEvals.reduce((sum, item) => sum + item.autoFails, 0);
    document.getElementById('kpi-auto-fail').textContent = totalAutoFails;
    
    document.getElementById('kpi-total-evaluaciones').textContent = filteredEvals.length;

    const renderTop3 = (listId, dataSubset) => {
        const counts = {};
        dataSubset.forEach(item => {
            counts[item.tipo] = (counts[item.tipo] || 0) + 1;
        });

        const listEl = document.getElementById(listId);
        const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 3);
        
        listEl.innerHTML = '';
        if (sorted.length === 0) {
            listEl.innerHTML = '<li>-</li>';
            return;
        }

        const colors = ['#8b5cf6', '#ec4899', '#f59e0b'];
        sorted.forEach((item, index) => {
            const li = document.createElement('li');
            li.style.borderLeftColor = colors[index] || colors[0];
            
            const nameSpan = document.createElement('span');
            nameSpan.className = 'top3-name';
            nameSpan.textContent = item[0];
            nameSpan.title = item[0];
            
            const countSpan = document.createElement('span');
            countSpan.className = 'count';
            countSpan.textContent = item[1];
            
            li.appendChild(nameSpan);
            li.appendChild(countSpan);
            listEl.appendChild(li);
        });
    };

    renderTop3('kpi-top3-normal-list', processedData.filter(item => item.clase === 'Normal'));
    renderTop3('kpi-top3-autofail-list', processedData.filter(item => item.clase === 'Auto Fail'));
    renderTop3('kpi-top3-autozero-list', processedData.filter(item => item.clase === 'Auto Zero'));
}

const stackedTotalPlugin = {
    id: 'stackedTotal',
    afterDatasetsDraw: (chart, args, options) => {
        const { ctx } = chart;
        chart.data.labels.forEach((label, i) => {
            let total = 0;
            let topY = chart.scales.y.bottom;
            let xPos = null;
            
            chart.data.datasets.forEach((dataset, datasetIndex) => {
                if (!chart.isDatasetVisible(datasetIndex)) return;
                const value = dataset.data[i];
                if (value > 0) {
                    total += value;
                    const meta = chart.getDatasetMeta(datasetIndex);
                    const element = meta.data[i];
                    if (element && element.y < topY) {
                        topY = element.y;
                        xPos = element.x;
                    }
                }
            });
            
            if (total > 0 && xPos !== null) {
                ctx.save();
                ctx.fillStyle = '#f8fafc'; // Color del texto
                ctx.font = 'bold 11px "Outfit", sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(total, xPos, topY - 4); // Dibujar el número 4px por encima de la barra
                ctx.restore();
            }
        });
    }
};

function updateChart() {
    const isSupervisorSelected = supervisorSelect.value !== 'all';
    const groupKey = isSupervisorSelected ? 'agente' : 'supervisor';
    const chartTitle = isSupervisorSelected ? 'Incumplimientos por Agente' : 'Incumplimientos por Supervisor';
    
    // Actualizar el título de la gráfica dinámicamente
    const chartHeaderObj = document.querySelector('.chart-header h3');
    if(chartHeaderObj) chartHeaderObj.textContent = chartTitle;

    const groupMap = {}; // e.g. { 'Juan': { 'Tipo 1': 5, 'Tipo 2': 2 } }
    const allTipos = new Set();

    processedData.forEach(item => {
        const groupName = item[groupKey] || 'Desconocido';
        if (!groupMap[groupName]) groupMap[groupName] = {};
        groupMap[groupName][item.tipo] = (groupMap[groupName][item.tipo] || 0) + 1;
        allTipos.add(item.tipo);
    });

    const labels = Object.keys(groupMap).sort();
    const tiposArray = Array.from(allTipos).sort();

    const datasets = tiposArray.map((tipo, index) => {
        const data = labels.map(g => groupMap[g][tipo] || 0);
        return {
            label: tipo,
            data: data,
            backgroundColor: chartColors[index % chartColors.length],
            borderWidth: 0,
            borderRadius: 4
        };
    });

    // Ajustar el ancho para que ocupe todo el espacio sin scroll horizontal
    const chartContainer = document.getElementById('chart-container');
    chartContainer.style.width = '100%';

    const ctx = document.getElementById('mainChart').getContext('2d');
    
    // Set default Chart.js colors for dark theme
    Chart.defaults.color = '#94a3b8';
    Chart.defaults.borderColor = 'rgba(255,255,255,0.05)';
    Chart.defaults.font.family = "'Outfit', sans-serif";

    if (chartInstance) {
        // En lugar de destruir y recrear, actualizamos los datos y forzamos el renderizado
        chartInstance.data.labels = labels;
        chartInstance.data.datasets = datasets;
        chartInstance.update();
        return;
    }

    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: { top: 25 } // Espacio extra para que no se corten los números
            },
            animation: {
                duration: 400 // Animación más rápida para los filtros
            },
            interaction: {
                mode: 'index',
                intersect: false,
            },
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 20,
                        usePointStyle: true,
                        pointStyle: 'circle'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 23, 42, 0.9)',
                    titleColor: '#f8fafc',
                    bodyColor: '#e2e8f0',
                    borderColor: 'rgba(255,255,255,0.1)',
                    borderWidth: 1,
                    padding: 12,
                    boxPadding: 6
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: {
                        display: false
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        precision: 0
                    }
                }
            }
        },
        plugins: [stackedTotalPlugin]
    });
}

// --- NUEVA PESTAÑA DETALLES DE AGENTES ---
const navDashboard = document.getElementById('nav-dashboard');
const navAgentes = document.getElementById('nav-agentes');
const dashboardContentSection = document.getElementById('dashboard-content');
const agentesContentSection = document.getElementById('agentes-content');

const agDateFromInput = document.getElementById('agentes-date-from');
const agDateToInput = document.getElementById('agentes-date-to');
const agCanalSelect = document.getElementById('agentes-canal-select');
const agSupervisorSelect = document.getElementById('agentes-supervisor-select');
const agAgenteSelect = document.getElementById('agentes-agente-select');
const agTbody = document.getElementById('agentes-tbody');
const agTotalSpan = document.getElementById('agentes-total-evaluaciones');
const agDownloadBtn = document.getElementById('btn-download-agentes');

window.switchTab = function(tabId) {
    if (tabId === 'dashboard') {
        navDashboard.classList.add('active');
        navAgentes.classList.remove('active');
        dashboardContentSection.style.display = 'block';
        agentesContentSection.style.display = 'none';
        applyFilters(); 
    } else {
        navDashboard.classList.remove('active');
        navAgentes.classList.add('active');
        dashboardContentSection.style.display = 'none';
        agentesContentSection.style.display = 'block';
        populateAgentesSelects();
        renderAgentesTable();
    }
};

function getFilteredAgentes() {
    const fromDateStr = agDateFromInput.value;
    const toDateStr = agDateToInput.value;
    const selectedCanal = agCanalSelect.value;
    const selectedSupervisor = agSupervisorSelect.value;
    const selectedAgente = agAgenteSelect.value;

    let fromDate = null;
    if (fromDateStr) {
        const parts = fromDateStr.split('-');
        fromDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }

    let toDate = null;
    if (toDateStr) {
        const parts = toDateStr.split('-');
        toDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        toDate.setHours(23, 59, 59, 999);
    }

    return allEvaluations.filter(item => {
        let matchDate = true;
        if (item.fechaObj && !isNaN(item.fechaObj.getTime())) {
            if (fromDate && item.fechaObj < fromDate) matchDate = false;
            if (toDate && item.fechaObj > toDate) matchDate = false;
        }

        let matchCanal = true;
        if (selectedCanal !== 'all' && item.canal !== selectedCanal) matchCanal = false;
        
        let matchSupervisor = true;
        if (selectedSupervisor !== 'all' && item.supervisor !== selectedSupervisor) matchSupervisor = false;

        let matchAgente = true;
        if (selectedAgente !== 'all' && item.agente !== selectedAgente) matchAgente = false;

        return matchDate && matchCanal && matchSupervisor && matchAgente;
    });
}

function renderAgentesTable() {
    const data = getFilteredAgentes();
    
    // Agrupar por Agente + Canal + Supervisor
    const agrupado = {};
    data.forEach(item => {
        const key = `${item.agente}|${item.canal}|${item.supervisor}`;
        if (!agrupado[key]) {
            agrupado[key] = {
                agente: item.agente,
                canal: item.canal,
                supervisor: item.supervisor,
                conteo: 0
            };
        }
        agrupado[key].conteo++;
    });

    // Convertir a array y ordenar por cantidad descendente
    const resultados = Object.values(agrupado).sort((a, b) => b.conteo - a.conteo);

    if (!agTbody) return;
    agTbody.innerHTML = '';
    
    // Calcular total dinámico
    const totalEvaluaciones = resultados.reduce((sum, item) => sum + item.conteo, 0);
    if(agTotalSpan) agTotalSpan.textContent = totalEvaluaciones;
    
    if (resultados.length === 0) {
        agTbody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: var(--text-secondary);">No se encontraron resultados</td></tr>';
        return;
    }

    resultados.forEach(res => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${res.agente}</td>
            <td><span class="canal-badge" style="background: ${res.canal.includes('Telef') ? 'rgba(59, 130, 246, 0.1)' : 'rgba(139, 92, 246, 0.1)'}; color: ${res.canal.includes('Telef') ? 'var(--accent-blue)' : 'var(--accent-purple)'}; padding: 4px 8px; border-radius: 4px; font-size: 0.85rem; font-weight: 500;">${res.canal}</span></td>
            <td>${res.supervisor}</td>
            <td>${res.conteo}</td>
        `;
        agTbody.appendChild(tr);
    });
}

let agentesSelectsPopulated = false;
function populateAgentesSelects() {
    if (agentesSelectsPopulated) return;
    
    const uniqueCanales = new Set();
    const uniqueSupervisores = new Set();
    const uniqueAgentes = new Set();

    allEvaluations.forEach(item => {
        uniqueCanales.add(item.canal);
        if (item.supervisor) uniqueSupervisores.add(item.supervisor);
        if (item.agente) uniqueAgentes.add(item.agente);
    });

    if(agCanalSelect) updateSelectFilter(agCanalSelect, uniqueCanales, "Todos los canales");
    if(agSupervisorSelect) updateSelectFilter(agSupervisorSelect, uniqueSupervisores, "Todos los supervisores");
    if(agAgenteSelect) updateSelectFilter(agAgenteSelect, uniqueAgentes, "Todos los agentes");
    
    agentesSelectsPopulated = true;
}

[agDateFromInput, agDateToInput, agCanalSelect, agSupervisorSelect, agAgenteSelect].forEach(input => {
    if(input) {
        input.addEventListener('change', renderAgentesTable);
    }
});

if(agDownloadBtn) {
    agDownloadBtn.addEventListener('click', () => {
        const data = getFilteredAgentes();
        const agrupado = {};
        data.forEach(item => {
            const key = `${item.agente}|${item.canal}|${item.supervisor}`;
            if (!agrupado[key]) {
                agrupado[key] = {
                    "Agente": item.agente,
                    "Canal de Atención": item.canal,
                    "Supervisor Asignado": item.supervisor,
                    "Evaluaciones Realizadas": 0
                };
            }
            agrupado[key]["Evaluaciones Realizadas"]++;
        });
        
        const resultadosExcel = Object.values(agrupado).sort((a, b) => b["Evaluaciones Realizadas"] - a["Evaluaciones Realizadas"]);
        
        if(resultadosExcel.length === 0) {
            alert('No hay datos para exportar con los filtros actuales.');
            return;
        }

        const worksheet = XLSX.utils.json_to_sheet(resultadosExcel);
        
        // Ajustar el ancho de las columnas
        const wscols = [
            {wch: 35}, // Agente
            {wch: 20}, // Canal
            {wch: 35}, // Supervisor
            {wch: 25}  // Evaluaciones
        ];
        worksheet['!cols'] = wscols;

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Resumen Agentes");
        XLSX.writeFile(workbook, "Resumen_Agentes.xlsx");
    });
}
