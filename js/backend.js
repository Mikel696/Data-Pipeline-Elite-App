/**
 * Data Pipeline Elite App - FAKE BACKEND ENGINE v3.0
 * Simula una API asíncrona cargando datos masivos desde functions_db.json
 * y los renderiza en la UI Premium.
 */

document.addEventListener('DOMContentLoaded', () => {
    // Inicializar la carga del backend
    fetchDatabase();
});

let functionDatabase = {};

async function fetchDatabase() {
    try {
        const response = await fetch('data/functions_db.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        functionDatabase = await response.json();
        renderAllGrids();
    } catch (e) {
        console.error("Backend Error: No se pudo cargar la base de datos.", e);
        document.querySelectorAll('.cards-grid').forEach(grid => {
            grid.innerHTML = '<p style="color: #ef4444;">Error de conexión con el Backend (JSON).</p>';
        });
    }
}

function renderAllGrids() {
    // Render Excel
    if(functionDatabase.excel) renderGrid('grid-excel', functionDatabase.excel, 'excel');
    // Render SQL
    if(functionDatabase.sql) renderGrid('grid-sql', functionDatabase.sql, 'sql');
    // Render Python
    if(functionDatabase.python) renderGrid('grid-python', functionDatabase.python, 'python');
    // Render Github
    if(functionDatabase.github) renderGrid('grid-github', functionDatabase.github, 'github');
    // Render Integracion
    if(functionDatabase.integration) renderGrid('grid-integration', functionDatabase.integration, 'integration');
}

function renderGrid(gridId, dataArray, moduleTheme) {
    const grid = document.getElementById(gridId);
    if (!grid) return;
    
    let html = '';
    
    dataArray.forEach(item => {
        // Escapar comillas para inyectar objeto JSON en onclick sin romper HTML
        const safeData = encodeURIComponent(JSON.stringify(item));
        
        html += `
            <div class="function-card ${moduleTheme}-hover" onclick="openFunctionModal('${safeData}', '${moduleTheme}')">
                <span class="card-category ${moduleTheme}-color">${item.categoria}</span>
                <h3 class="card-title">${item.titulo}</h3>
                <p class="card-excerpt">${item.descripcion}</p>
            </div>
        `;
    });
    
    grid.innerHTML = html;
}

/**
 * Funciones de interacción del Modal (Pop-Up)
 */
function openFunctionModal(encodedData, themeName) {
    const data = JSON.parse(decodeURIComponent(encodedData));
    
    document.getElementById('modal-title').innerText = data.titulo;
    document.getElementById('modal-category').innerText = data.categoria;
    document.getElementById('modal-desc').innerText = data.descripcion;
    document.getElementById('modal-code').innerText = data.codigo;
    
    // Aplicar el color de rama al modal
    const categorySpan = document.getElementById('modal-category');
    categorySpan.className = ''; // reset
    categorySpan.classList.add(`${themeName}-color`);
    
    const popup = document.getElementById('dynamic-modal');
    popup.classList.add('active');
}

function closeModal(e) {
    // Si se pasa evento, checar si el clic fue estrictamente en el background
    if (e && e.target !== document.getElementById('dynamic-modal') && e.target !== document.querySelector('.btn-close')) {
        return;
    }
    const popup = document.getElementById('dynamic-modal');
    if (popup) popup.classList.remove('active');
}
