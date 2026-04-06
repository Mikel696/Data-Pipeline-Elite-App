/**
 * Data Pipeline Elite App - MAIN SCRIPT
 * Manejo de la navegación global y utilidades de renderizado dinámico.
 */

document.addEventListener('DOMContentLoaded', () => {
    console.log("Data Pipeline Elite App Initialized");
    
    // Inicializar progreso visual de la Sidebar
    if(typeof updateProgressUI === 'function') {
        updateProgressUI();
    }

    // Inyectar previsualización de datos si estamos en la pestaña lab
    showDataPreview();
});

/**
 * Función principal de navegación del Sidebar.
 * @param {string} targetId - El ID de la sección (`sql`, `python`, `excel`, `master-lab`).
 * @param {HTMLElement} element - El elemento clickeado del sidebar.
 */
function appNavigate(targetId, element) {
    // 1. Quitar activo de todos los botones
    const buttons = document.querySelectorAll('.nav-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
    
    // 2. Anadir activo al botón clickeado
    if(element) element.classList.add('active');

    // 3. Ocultar todos los módulos
    const views = document.querySelectorAll('.module-view');
    views.forEach(view => {
        view.classList.remove('active-view');
        view.classList.remove('fade-in');
    });

    // 4. Mostrar solo el módulo objetivo
    const targetView = document.getElementById(targetId);
    if(targetView) {
        targetView.classList.add('active-view');
        void targetView.offsetWidth; // force relayout
        targetView.classList.add('fade-in');
    }

    // 5. Si navegamos al Lab, garantizar que la Data y el reto estén cargados
    if(targetId === 'master-lab') {
        showDataPreview();
    }
}

/**
 * Función requerida: Inyecta dinámicamente la tabla HTML con 5 filas de ejemplo
 * del dataset de Churn en el Laboratorio Maestro.
 */
function showDataPreview() {
    const previewContainer = document.getElementById('data-preview-container');
    if (!previewContainer) return; // Si no existe el contenedor, fallamos silenciosamente.

    const sampleData = [
        { id: "C001", name: "John Doe", signup: "2024-01-15", spent: "$1250.50", churn: "False" },
        { id: "C002", name: "Maria Garcia", signup: "2023-11-20", spent: "$320.00", churn: "True" },
        { id: "C003", name: "Wei Chen", signup: "2024-03-05", spent: "$890.75", churn: "False" },
        { id: "C004", name: "Ana Silva", signup: "2023-05-10", spent: "$150.20", churn: "True" },
        { id: "C005", name: "James Smith", signup: "2025-01-10", spent: "$450.00", churn: "False" }
    ];

    let html = `
        <table class="data-table">
            <thead>
                <tr>
                    <th>customer_id</th>
                    <th>name</th>
                    <th>signup_date</th>
                    <th>total_spent</th>
                    <th>is_churned</th>
                </tr>
            </thead>
            <tbody>
    `;

    sampleData.forEach(row => {
        html += `
            <tr>
                <td>${row.id}</td>
                <td>${row.name}</td>
                <td>${row.signup}</td>
                <td>${row.spent}</td>
                <td><span style="color: ${row.churn === 'True' ? '#ef4444' : '#10b981'}">${row.churn}</span></td>
            </tr>
        `;
    });

    html += `
            </tbody>
        </table>
    `;

    previewContainer.innerHTML = html;
}

/**
 * Función para alternar el estado de un acordeón en el UI inmersivo.
 * @param {HTMLElement} element - El header del acordeón clickeado.
 */
function toggleAccordion(element) {
    const item = element.parentElement;
    
    // Opcional: Cerrar otros acordeones dentro del mismo wrapper
    // const wrapper = item.parentElement;
    // wrapper.querySelectorAll('.accordion-item').forEach(acc => {
    //     if(acc !== item) acc.classList.remove('open');
    // });

    item.classList.toggle('open');
}
