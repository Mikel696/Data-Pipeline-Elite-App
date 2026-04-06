/**
 * Data Pipeline Elite App - STORAGE SCRIPT
 * Manejo de persistencia del progreso del usuario mediante localStorage.
 */

const STORAGE_KEY = 'DataPipelineEliteProgress';

/**
 * Retorna el vector de niveles superados desde localStorage
 */
function getProgress() {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
}

/**
 * Agrega un nivel superado al localStorage
 * @param {number} levelId 
 */
function saveProgress(levelId) {
    let progress = getProgress();
    if(!progress.includes(levelId)) {
        progress.push(levelId);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(progress));
        updateProgressUI();
    }
}

/**
 * Actualiza la UI visual de la barra de progreso (Sidebar)
 */
function updateProgressUI() {
    const progress = getProgress();
    const totalLevels = 10; // 10 niveles requeridos
    const percent = Math.round((progress.length / totalLevels) * 100);
    
    const bar = document.getElementById('master-progress');
    const text = document.getElementById('progress-percent');
    
    if(bar) bar.style.width = percent + '%';
    if(text) text.innerText = percent + '% Completado';
}
