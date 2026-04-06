/**
 * Data Pipeline Elite App - FAKE BACKEND ENGINE v3.0
 * Simula una API asíncrona cargando datos masivos desde functions_db.json
 * y los renderiza en la UI Premium.
 */

document.addEventListener('DOMContentLoaded', () => {
    // Inicializar la carga del backend
    fetchDatabase();
});

// En lugar de fetch (que da error CORS en file:/// locales sin servidor), 
// inyectamos nuestra Base de Datos JSON directamente en memoria.
let functionDatabase = {
    "excel": [
        { "id": "ex-1", "titulo": "Ctrl Shortcuts", "categoria": "Atajos Fundamentales", "descripcion": "Selección rápida (Ctrl+Espacio), saltos cardinales (Ctrl+Flechas) y formatos rápidos (Ctrl+Shift+$).", "codigo": "Ctrl + Espacio : Selecciona columna entera\nCtrl + 1 : Formato de Celdas" },
        { "id": "ex-2", "titulo": "Shift/Alt Shortcuts", "categoria": "Atajos Fundamentales", "descripcion": "Navegación espacial y ejecución de herramientas directas del ribbon.", "codigo": "Shift + Espacio : Fila entera\nAlt + H + A + C : Centrar contenido" },
        { "id": "ex-3", "titulo": "+, -, *, /, RAIZ()", "categoria": "Básicas y Operaciones", "descripcion": "Fundamentos aritméticos y de raíces. Base para cualquier modelado financiero.", "codigo": "=RAIZ(144) // Retorna 12" },
        { "id": "ex-4", "titulo": "SUMA & PROMEDIO", "categoria": "Matemáticas Estadísticas", "descripcion": "Funciones de agregación matriz para retornar media aritmética y cálculos de liquidación.", "codigo": "=PROMEDIO(A1:A100)" },
        { "id": "ex-5", "titulo": "CONTARA() vs CONTAR()", "categoria": "Matemáticas Estadísticas", "descripcion": "CONTAR busca estrictamente numéricos. CONTARA valida cualquier celda no vacía (ideal para contar strings).", "codigo": "=CONTARA(B2:B500) // Cuenta registros poblados" },
        { "id": "ex-6", "titulo": "SUMAR.SI & CONTAR.SI", "categoria": "Matemáticas Estadísticas", "descripcion": "Agregación bajo criterio de matching. Permite totalizar ventas de un solo agente.", "codigo": "=SUMAR.SI(RangoZona, \"Norte\", RangoVentas)" },
        { "id": "ex-7", "titulo": "SI() & SIERROR()", "categoria": "Lógicas y Limpieza", "descripcion": "Motor condicional base y control proactivo de #N/A para evitar desbordes en cascada.", "codigo": "=SI(A1>100; \"Alto\"; \"Bajo\")\n=SIERROR(VLOOKUP(...), 0)" },
        { "id": "ex-8", "titulo": "Y() / O() / NO()", "categoria": "Lógicas Avanzadas", "descripcion": "Operadores booleanos anidados típicamente dentro de la función SI() para múltiple validación.", "codigo": "=SI(Y(A1>10, B1=\"OK\"); \"Aprobado\"; \"Rechazado\")" },
        { "id": "ex-9", "titulo": "TRIM (ESPACIOS)", "categoria": "Limpieza Cadenas", "descripcion": "Poda espacios basurilla iniciales, finales y dobles. Indispensable antes de un JOIN.", "codigo": "=ESPACIOS(A2)" },
        { "id": "ex-10", "titulo": "VLOOKUP (BUSCARV)", "categoria": "Búsqueda y Referencia", "descripcion": "Cruza tablas de izquierda a derecha. Lento en datasets de millones de filas pero estándar de industria.", "codigo": "=BUSCARV(id, tabla_matriz, columna, FALSO)" },
        { "id": "ex-11", "titulo": "INDEX & MATCH", "categoria": "Búsqueda y Referencia", "descripcion": "La combinación legendaria para cruces hacia la izquierda y algoritmos que ahorran memoria caché.", "codigo": "=INDEX(VectorRetorno; MATCH(Id; VectorBusqueda; 0))" },
        { "id": "ex-12", "titulo": "XLOOKUP()", "categoria": "Matrices Dinámicas", "descripcion": "La evolución oficial de Microsoft. Busca donde sea, tiene SIERROR incorporado y devuelve múltiples columnas.", "codigo": "=XLOOKUP(Valor; VectorBusca; VectorDevuelve; \"Fallo\"; 0)" },
        { "id": "ex-13", "titulo": "FILTRAR()", "categoria": "Matrices Dinámicas", "descripcion": "Desborda un subset de datos hacia múltiples celdas basándose en la coincidencia booleana.", "codigo": "=FILTRAR(A2:C100; B2:B100=\"VIP\")" },
        { "id": "ex-14", "titulo": "UNIQUE() y ORDENAR()", "categoria": "Matrices Dinámicas", "descripcion": "Devuelve matriz limpia libre de duplicados. Se anida con ORDENAR para alfabético instantáneo.", "codigo": "=ORDENAR(UNIQUE(A2:A5000))" },
        { "id": "ex-15", "titulo": "Query Folding (Power Query)", "categoria": "Lenguaje M (ETL)", "descripcion": "Capacidad de Power Query para empujar peticiones (WHERE, GROUP) al motor SQL en lugar de bajar la data local.", "codigo": "No function; Architectural paradigm." },
        { "id": "ex-16", "titulo": "Table.Buffer()", "categoria": "Lenguaje M (ETL)", "descripcion": "Rompe el query folding pero guarda tablas de diccionario en RAM para inyectar velocidad a cruces pesados.", "codigo": "Source = Table.Buffer(Sql_Table)" },
        { "id": "ex-17", "titulo": "Slicers y Timelines", "categoria": "Tablas Dinámicas", "descripcion": "Dashboard paramétrico visual, conectando Múltiples Tablas Dinámicas (Report Connections) al modelo DAX final.", "codigo": "No code. UI Interaction" }
    ],
    "sql": [
        { "id": "sq-1", "titulo": "SELECT DISTINCT / LIMIT", "categoria": "Sintaxis Básica", "descripcion": "Trae identificadores únicos colapsados o restringe el pipeline para no destruir la memoria RAM.", "codigo": "SELECT DISTINCT user_id FROM logs LIMIT 100;" },
        { "id": "sq-2", "titulo": "Filtrado (WHERE, AND, IN)", "categoria": "Sintaxis Básica", "descripcion": "Poda primaria de registros basada en predicados. Ocurre previo al agrupamiento.", "codigo": "SELECT * FROM ventas WHERE pais IN ('CO', 'MX');" },
        { "id": "sq-3", "titulo": "INNER JOIN", "categoria": "Uniones Relacionales", "descripcion": "Devuelve la intersección absoluta de dos esquemas relacionales.", "codigo": "SELECT * FROM users u INNER JOIN orders o ON u.id = o.u_id;" },
        { "id": "sq-4", "titulo": "LEFT y RIGHT JOIN", "categoria": "Uniones Relacionales", "descripcion": "Respeta todos los registros de una matriz A, anexando coincidencias de la B. Llena vacíos con NULL.", "codigo": "SELECT * FROM users u LEFT JOIN orders o ON u.id = o.u_id;" },
        { "id": "sq-5", "titulo": "FULL OUTER JOIN", "categoria": "Uniones Relacionales", "descripcion": "Integra todo lo de la izquierda y la derecha sin excepción, generando NULLs en los cuadrantes vacíos.", "codigo": "SELECT * FROM A FULL OUTER JOIN B ON A.id = B.id;" },
        { "id": "sq-6", "titulo": "ANTI-JOIN", "categoria": "Uniones Relacionales", "descripcion": "Busca elementos huérfanos aislando mediante un filtro IS NULL después del join.", "codigo": "SELECT * FROM users LEFT JOIN orders ON... WHERE logs.id IS NULL;" },
        { "id": "sq-7", "titulo": "SUM / AVG / COUNT", "categoria": "Agregación", "descripcion": "Colapsa tablas inmensas en KPIs sumarizados absolutos.", "codigo": "SELECT region, SUM(monto) FROM ventas GROUP BY region;" },
        { "id": "sq-8", "titulo": "GROUP BY vs HAVING", "categoria": "Agregación", "descripcion": "HAVING filtra los resultados *después* de agrupar. La diferencia cardinal con WHERE.", "codigo": "GROUP BY user HAVING COUNT(*) > 10;" },
        { "id": "sq-9", "titulo": "ROW_NUMBER()", "categoria": "Window Functions", "descripcion": "Enumera vectores secuencialmente por un criterio sin comprimirlos.", "codigo": "ROW_NUMBER() OVER(PARTITION BY id ORDER BY date DESC)" },
        { "id": "sq-10", "titulo": "RANK() / DENSE_RANK()", "categoria": "Window Functions", "descripcion": "Rango estadístico (ej: Tops). DENSE() no permite saltos numerarios en caso de empates.", "codigo": "DENSE_RANK() OVER(ORDER BY sales DESC)" },
        { "id": "sq-11", "titulo": "LAG() / LEAD()", "categoria": "Window Functions", "descripcion": "Viaja al pasado o futuro 1 fila para encontrar diferencias de fechas (ej: cohortes de compra).", "codigo": "LAG(order_date, 1) OVER(PARTITION BY user ORDER BY date)" },
        { "id": "sq-12", "titulo": "NTILE()", "categoria": "Window Functions", "descripcion": "Secciona la población en cuartiles/deciles equitativos basándose en distribución estadística.", "codigo": "NTILE(4) OVER(ORDER BY revenue DESC) as Cuartil" },
        { "id": "sq-13", "titulo": "CTEs (Cláusula WITH)", "categoria": "Subconsultas y Optimizacion", "descripcion": "Modera subqueries anidados asilándolos virtualmente al inicio de la query base.", "codigo": "WITH TempData AS (SELECT * FROM a) SELECT * FROM TempData" },
        { "id": "sq-14", "titulo": "Query EXPLAIN", "categoria": "Subconsultas y Optimizacion", "descripcion": "Desnuda el plan de ejecución del motor validando Index Scans vs Full Table Scans.", "codigo": "EXPLAIN ANALYZE SELECT * FROM table;" }
    ],
    "python": [
        { "id": "py-1", "titulo": "Listas y Diccionarios", "categoria": "Estructuras Nativas", "descripcion": "Vectores Mapeados. La matriz de cualquier Data Engineer en Python. Dinámicos y mutables.", "codigo": "dict_obj = {'score': 100}\nlist_obj = [1, 2, 3]" },
        { "id": "py-2", "titulo": "Tuplas y Sets", "categoria": "Estructuras Nativas", "descripcion": "Tuplas inmutables (para proteger coords) y Sets (para encontrar valores únicos de la población).", "codigo": "unique_vals = set([1,2,2,3])" },
        { "id": "py-3", "titulo": "for, while, if/elif/else", "categoria": "Control de Flujo", "descripcion": "Corazones del iterador básico para parseo y lógica buleana transitoria.", "codigo": "for x in vector:\n  if x > 10: break" },
        { "id": "py-4", "titulo": "Funciones Lambda", "categoria": "Funciones y Decoradores", "descripcion": "Funciones asíncronas desechables (ideales dentro de .apply de pandas).", "codigo": "f = lambda x: x * 2" },
        { "id": "py-5", "titulo": "Decoradores (@)", "categoria": "Funciones y Decoradores", "descripcion": "Envuelven lógicas inyectando esteroides a la función base sin modificarla internamente.", "codigo": "@logger\ndef load_data(): ..." },
        { "id": "py-6", "titulo": "df.head() / df.describe()", "categoria": "Pandas Exploración", "descripcion": "El radar. Conteo, min, max y distribución del dataset instantánea de los tensores.", "codigo": "df.describe().T" },
        { "id": "py-7", "titulo": "df.info()", "categoria": "Pandas Exploración", "descripcion": "Imprime DTypes (objetos, fechas, float) y localiza columnas carentes o polutas de basura NaN.", "codigo": "df.info()" },
        { "id": "py-8", "titulo": "df.dropna()", "categoria": "Limpieza Wrangling", "descripcion": "Aniquila de forma nuclear toda fila o columna infractora que posea un NaN.", "codigo": "df.dropna(axis=0)" },
        { "id": "py-9", "titulo": "df.fillna()", "categoria": "Limpieza Wrangling", "descripcion": "Tratamiento de missings. Reemplaza vectores oscuros con Medias/Medianas inferidas para el Algoritmo RFM.", "codigo": "df['age'].fillna(df['age'].mean())" },
        { "id": "py-10", "titulo": "df.astype()", "categoria": "Limpieza Wrangling", "descripcion": "Transforma coercivamente la tipología del tensor (ej: String numérico a Int64 puro).", "codigo": "df['price'] = df['price'].astype('float')" },
        { "id": "py-11", "titulo": "df.drop_duplicates()", "categoria": "Limpieza Wrangling", "descripcion": "Garantiza integridad atómica purgando los registros idénticos a nivel de hash de bits.", "codigo": "df.drop_duplicates(keep='last')" },
        { "id": "py-12", "titulo": "df.groupby()", "categoria": "Transformación Tabular", "descripcion": "El equivalente a GROUP BY pero orquestado en memoria RAM en objetos Dataframes subyacentes.", "codigo": "df.groupby('cohort').mean()" },
        { "id": "py-13", "titulo": "df.merge()", "categoria": "Transformación Tabular", "descripcion": "El equivalente a un JOIN parametrizado. Permite LEFT, RIGHT o INNER interconectado.", "codigo": "df_final = pd.merge(df1, df2, on='id', how='left')" },
        { "id": "py-14", "titulo": "Seaborn y KDE Plots", "categoria": "Visualización Científica", "descripcion": "Seaborn traza distribuciones densas para encontrar variables latentes de fuga en la empresa.", "codigo": "sns.kdeplot(data=df['spent'])" }
    ],
    "github": [
        { "id": "gh-1", "titulo": "git config", "categoria": "Configuración Core", "descripcion": "Identidad. Indispensable para auditar qué desarrollador alteró el pipeline y en qué bloque CRC.", "codigo": "git config --global user.name \"Root\"" },
        { "id": "gh-2", "titulo": "git init / clone", "categoria": "Configuración Core", "descripcion": "Inicia la matriz local oculta (.git) o baja una réplica clónica exacta del ecosistema nube.", "codigo": "git clone https://repo.git" },
        { "id": "gh-3", "titulo": "git status", "categoria": "Flujo Workdir", "descripcion": "Muestra los trackers. Rojo = Modificado No Registrado. Verde = En Staging Area (Muelle de Carga).", "codigo": "git status" },
        { "id": "gh-4", "titulo": "git add", "categoria": "Flujo Workdir", "descripcion": "Sube tus archivos al Staging Area. El flag punto (.) engulle todos los cambios globalmente.", "codigo": "git add ." },
        { "id": "gh-5", "titulo": "git commit", "categoria": "Historia Inmutable", "descripcion": "Congela el stage en un Hash SHA permanente que viajará adjunto a la documentación del commit.", "codigo": "git commit -m \"Fix engine\"" },
        { "id": "gh-6", "titulo": "git push / pull", "categoria": "Nube (Origin)", "descripcion": "Propiedades vectoriales: Push lanza de local a Nube, Pull descarga lo modificado en Nube a Local.", "codigo": "git push origin main" },
        { "id": "gh-7", "titulo": "git branch / checkout", "categoria": "Seguridad Múltiple", "descripcion": "Desarrolla algoritmos en mundos aislados (ramas). Checkout transporta todo tu ecosistema entre estos mundos.", "codigo": "git checkout -b feature_x" },
        { "id": "gh-8", "titulo": "Pull Requests", "categoria": "Colaboración Corp", "descripcion": "En GitHub Cloud, pide 'permiso' para que tu rama se integre a producción (Main) pasando pruebas CI/CD.", "codigo": "Operación de Interfaz UI GitHub" },
        { "id": "gh-9", "titulo": "Forks", "categoria": "Colaboración Corp", "descripcion": "Hacer un duplicado corporativo 100% tuyo de un proyecto base. El pilar del software libre mundial.", "codigo": "Operación de Interfaz UI GitHub" }
    ],
    "integration": [
        { "id": "in-1", "titulo": "=PY() Pandas Engine", "categoria": "Python nativo Excel", "descripcion": "Contenedor Azure alojado virtualmente en celdas para inyectar Pandas y Numpy sobre Excel Tabular.", "codigo": "=PY(\n  import pandas as pd \n  return xl(\"Tabla1\").describe() \n)" },
        { "id": "in-2", "titulo": "Engine SQLAlchemy", "categoria": "ETL Cero-Fricciones", "descripcion": "Pipe directo de código Pandas hacia Bases de datos OLTP/OLAP evadiendo exportaciones de CSV mundanas.", "codigo": "engine = create_engine('sqlite:///db')\ndf = pd.read_sql('SELECT *', engine)" },
        { "id": "in-3", "titulo": "df.to_sql()", "categoria": "ETL Cero-Fricciones", "descripcion": "Comando espejo para inyectar vectores de Machine Learning de un Dataframe a tu DB de Producción automágicamente.", "codigo": "df.to_sql('predicciones', engine, if_exists='replace')" }
    ]
};

function fetchDatabase() {
    // Al ser un engine en bruto sin backend server, renderizamos la variable inmediatamente
    renderAllGrids();
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
        html += `
            <div class="function-card ${moduleTheme}-hover" onclick="openFunctionModal('${item.id}', '${moduleTheme}')">
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
function openFunctionModal(itemId, themeName) {
    // Buscar el item en la base de datos
    let data = null;
    for (const moduleObj in functionDatabase) {
        const found = functionDatabase[moduleObj].find(i => i.id === itemId);
        if (found) {
            data = found;
            break;
        }
    }
    
    if(!data) return;
    
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
