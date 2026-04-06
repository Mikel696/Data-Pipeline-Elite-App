/**
 * Data Pipeline Elite App - STORAGE SCRIPT
 * Manejo de persistencia del progreso del usuario mediante localStorage.
 */

const STORAGE_KEY = 'DataPipelineEliteProgress';
const LIB_STORAGE_KEY = 'DataPipelineLibProgress';

/**
 * Retorna el progreso según el tipo ('master' o 'lib')
 */
function getProgress(type = 'master') {
    const key = type === 'master' ? STORAGE_KEY : LIB_STORAGE_KEY;
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : [];
}

function saveProgress(levelId, type = 'master') {
    const key = type === 'master' ? STORAGE_KEY : LIB_STORAGE_KEY;
    let progress = getProgress(type);
    if(!progress.includes(levelId)) {
        progress.push(levelId);
        localStorage.setItem(key, JSON.stringify(progress));
        
        // Refrescar UI global
        if (typeof updateProgressUI === 'function') updateProgressUI();
        
        // Si es de la biblioteca, refrescar los badges de las tarjetas
        if (type === 'lib' && typeof renderAllGrids === 'function') {
            renderAllGrids();
        }
    }
}

/**
 * Actualiza la UI visual de la barra de progreso (Sidebar)
 * Contabiliza 10 de Lab + 24 de Biblioteca = 34 retos
 */
function updateProgressUI() {
    const masterProgress = getProgress('master');
    const libProgress    = getProgress('lib');
    
    const totalAchieved = masterProgress.length + libProgress.length;
    const totalPossible = 34; 
    const percent = Math.round((totalAchieved / totalPossible) * 100);
    
    const bar = document.getElementById('master-progress');
    const text = document.getElementById('progress-percent');
    
    if(bar) bar.style.width = percent + '%';
    if(text) text.innerText = percent + '% Completado';
}
