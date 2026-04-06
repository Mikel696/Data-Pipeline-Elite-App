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
        {
            "id": "ex-1",
            "titulo": "Atajos de Teclado Fundamentales",
            "categoria": "Eficiencia Operativa",
            "descripcion": "La eficiencia operativa en Excel depende del dominio de los comandos rápidos. Los atajos CTRL son para selección, edición y visualización. Los de SHIFT/ALT para navegación rápida.",
            "codigo": "Selección: Ctrl + Espacio (columna), Shift + Espacio (fila)\nCreación: Ctrl + T (crear tabla automágica)\nAnálisis: Alt + H + A + C (Remover Duplicados)"
        },
        {
            "id": "ex-2",
            "titulo": "VLOOKUP y XLOOKUP",
            "categoria": "Búsqueda y Referencia",
            "descripcion": "VLOOKUP busca un valor de forma vertical de izquierda a derecha. XLOOKUP es su evolución moderna: maneja errores nativamente de izquierda a derecha o inversa, devolviendo rangos enteros sin destruir modelos subyacentes.",
            "codigo": "=XLOOKUP(valor_buscado; array_busqueda; array_retorno; \"No Encontrado\"; 0)"
        },
        {
            "id": "ex-3",
            "titulo": "INDEX & MATCH (COINCIDIR)",
            "categoria": "Búsqueda Avanzada",
            "descripcion": "Combinación preferida para archivos grandes y pesados. El INDEX extrae el valor y el MATCH calcula dinámicamente la posición en la coordenada exacta.",
            "codigo": "=INDEX(ColumnaRetorno, MATCH(ValorBuscado, ColumnaBusqueda, 0))"
        },
        {
            "id": "ex-4",
            "titulo": "Fórmulas Estadísticas Críticas",
            "categoria": "Matemática Aplicada",
            "descripcion": "Para limpieza y recuento: CONTARA para evitar celdas vacías, SUMAR.SI para sumar condicionalmente montos basados en cohortes.",
            "codigo": "=SUMAR.SI(RangoFechas; \">\"&FECHA(2025;1;1); RangoSuma)"
        },
        {
            "id": "ex-5",
            "titulo": "Power Query (Lenguaje M)",
            "categoria": "ETL Automatizado",
            "descripcion": "Olvida hacer limpieza manual. Power Query usa 'Query Folding' para llevar cargas al origen y lenguaje funcional 'M' para documentar pasos que se ejecutarán automáticamente al presionar 'Actualizar'. Filtrar temprano y no borrar columnas ciegas garantizan el pipeline.",
            "codigo": "Table.SelectRows(Origen, each ([is_churned] <> \"True\"))"
        },
        {
            "id": "ex-6",
            "titulo": "Tablas Dinámicas y Slicers",
            "categoria": "Storytelling de Datos",
            "descripcion": "El C-Level consume inteligencia visual. Se conecta el Modelo de Datos DAX local a Pivot Tables. Agregar Segmentadores (Slicers) le permite al ejecutivo filtrar los KPIs sin corromper el cubo.",
            "codigo": "No code. UI Click: Insert -> Slicer -> Connect to Pivot."
        }
    ],
    "sql": [
        {
            "id": "sq-1",
            "titulo": "Sintaxis DML Base",
            "categoria": "Gestión Estándar",
            "descripcion": "La trinidad del SQL analítico: SELECT, WHERE (filtrado pre-agrupación), HAVING (filtrado post-agrupación) y DISTINCT para conteo único de clics.",
            "codigo": "SELECT DISTINCT user_id FROM logs WHERE status_code = 200;"
        },
        {
            "id": "sq-2",
            "titulo": "LEFT JOIN",
            "categoria": "Ingeniería Relacional",
            "descripcion": "Crucial en Churn y Analytics. Devuelve TODOS los registros de la matriz primaria (Izquierda) y pega lo encontrado a la derecha. Donde no existe actividad, devuelve NULL (bandera crítica de inactividad de usuario).",
            "codigo": "SELECT c.id, o.order_date\nFROM customers c\nLEFT JOIN orders o ON c.id = o.customer_id;"
        },
        {
            "id": "sq-3",
            "titulo": "ANTI-JOIN",
            "categoria": "Ingeniería Relacional",
            "descripcion": "Arquitectura forense: Se usa un LEFT JOIN evaluando un campo NULL para encontrar registros que existen en A pero que nunca existieron en B (Ej: Registrados sin compras).",
            "codigo": "SELECT c.id\nFROM customers c\nLEFT JOIN orders o ON c.id = o.customer_id\nWHERE o.customer_id IS NULL;"
        },
        {
            "id": "sq-4",
            "titulo": "Agregaciones & GROUP BY",
            "categoria": "Agregación",
            "descripcion": "Funciones matemáticas que condensan millones de filas. SUM, AVG, COUNT, MIN, MAX. Todo requiere empaquetado (GROUP BY) si hay dimensiones en el set.",
            "codigo": "SELECT country, AVG(amount) as ticket_promedio\nFROM orders \nGROUP BY country;"
        },
        {
            "id": "sq-5",
            "titulo": "Window Functions (ROW_NUMBER)",
            "categoria": "Analítica Avanzada",
            "descripcion": "Mantienen el dataset en su tamaño original, operando en 'ventanas' sobre los clústeres de particiones. Ideal para encontrar 'La primera compra de cada cliente'.",
            "codigo": "ROW_NUMBER() OVER (PARTITION BY customer_id ORDER BY order_date ASC)"
        },
        {
            "id": "sq-6",
            "titulo": "LAG y LEAD",
            "categoria": "Analítica de Series Temporales",
            "descripcion": "Desplazan las filas hacia atrás o adelante sin cambiar su identidad real. Usado drásticamente para calcular 'Días desde la compra anterior' por usuario.",
            "codigo": "LAG(order_date, 1) OVER (PARTITION BY customer_id ORDER BY order_date)"
        },
        {
            "id": "sq-7",
            "titulo": "CTEs (WITH)",
            "categoria": "Optimización Analítica",
            "descripcion": "Common Table Expressions. Reemplaza la monstruosidad de los 'Subqueries Anidados' generando vistas temporales hermosas para la carga de lógica.",
            "codigo": "WITH ActiveUsers AS (\n  SELECT * FROM users WHERE active=1\n)\nSELECT * FROM ActiveUsers;"
        }
    ],
    "python": [
        {
            "id": "py-1",
            "titulo": "Listas y Diccionarios",
            "categoria": "Estructuras de Datos Core",
            "descripcion": "Los diccionarios son pares de clave:valor para búsquedas al nivel Hash O(1). Las listas son colecciones dinámicas y mutables en Python. Los pilares de los JSON web.",
            "codigo": "payload = {'users': [1,2,3], 'active': True}"
        },
        {
            "id": "py-2",
            "titulo": "Decoradores y Lambdas",
            "categoria": "Pythonic Foundations",
            "descripcion": "Lambda provee funciones asíncronas desechables (ideales dentro de .apply de pandas). Decoradores envuelven lógicas (como medir el tiempo de un proceso) inyectando esteroides.",
            "codigo": "@timer\ndef scrape_data():\n  ...\ndf['tag'] = df['amount'].apply(lambda x: 'VIP' if x > 1000 else 'Std')"
        },
        {
            "id": "py-3",
            "titulo": "df.dropna & df.fillna",
            "categoria": "Pandas Data Wrangling",
            "descripcion": "Un algoritmo M.L jamás entrena con vacíos. Pandas purga nulos (dropna) o los estandariza (fillna) inyectando la media/mediana inteligente, preservando vectores tácticos.",
            "codigo": "df['spent'].fillna(df['spent'].mean(), inplace=True)"
        },
        {
            "id": "py-4",
            "titulo": "Agregación (groupby + agg)",
            "categoria": "Pandas Data Wrangling",
            "descripcion": "La traducción literal del mundo relacional al tensorial. Almacena data bajo diccionarios agrupados para recrear métricas corporativas, núcleo del Modelo RFM.",
            "codigo": "df.groupby('id').agg({'compra': 'count', 'monto': 'sum'})"
        },
        {
            "id": "py-5",
            "titulo": "Visualización: Seaborn",
            "categoria": "Machine Learning Prep",
            "descripcion": "Seaborn es una armadura de alto nivel estético sobre Matplotlib. Los gráficos de densidad (KDE Plots) o de correlación (Heatmaps) dictan los insights al Data Scientist.",
            "codigo": "import seaborn as sns\nsns.heatmap(df.corr(), annot=True, cmap='coolwarm')"
        }
    ],
    "github": [
        {
            "id": "gh-1",
            "titulo": "Configuración (git config)",
            "categoria": "Ingeniería DevOps",
            "descripcion": "Establece las credenciales del desarrollador localmente. Obligatorio para generar los trazos criptográficos históricos (commits).",
            "codigo": "git config --global user.name \"Miguel\"\ngit config --global user.email \"m.@.m\""
        },
        {
            "id": "gh-2",
            "titulo": "Staging (git add)",
            "categoria": "Flujo Repositorio Local",
            "descripcion": "Seleccionar explícitamente cuáles cambios editados en el ordenador local están listos para ser sellados en el próximo bloque de código.",
            "codigo": "git add ."
        },
        {
            "id": "gh-3",
            "titulo": "Confirmación (git commit)",
            "categoria": "Flujo Repositorio Local",
            "descripcion": "Es el sellado criogénico. Firma en bloque los archivos en un estado inmutable temporal, dándole contexto explicativo para el resto del equipo CI/CD.",
            "codigo": "git commit -m \"Fix: Parcheado de Null Pointers RFM\""
        },
        {
            "id": "gh-4",
            "titulo": "Colaboración (git push/pull)",
            "categoria": "Sincronía en la Nube",
            "descripcion": "Si Git opera en tu Desktop, GitHUB hospeda en la nube. Push eleva los commits, Pull los absorbe si tu compañero modificó master en Japón mientras dormías.",
            "codigo": "git push origin main"
        },
        {
            "id": "gh-5",
            "titulo": "Ramas y Pull Requests",
            "categoria": "Trabajo Open Source",
            "descripcion": "Nunca corras sentencias a main. Creas una rama 'fix-issue'. Empujas todo ahí. Levantas un 'Pull Request' (PR). El Sr. Developer audita tu labor y, si pasa QA, otorga 'Merge'.",
            "codigo": "git checkout -b fix-issue\n-- work work --\ngit push origin fix-issue"
        }
    ],
    "integration": [
        {
            "id": "in-1",
            "titulo": "Python in Excel =PY()",
            "categoria": "Arquitectura Híbrida",
            "descripcion": "La magia tecnológica de 2026. Aísla librerías avanzadas como Scikit-Learn en Azure, corriendo nativamente en la tabla pivotante Excel. Excel es un lienzo, Python es el motor.",
            "codigo": "=PY(\n  df = xl(\"Table1[#All]\", headers=True)\n  return df.describe()\n)"
        },
        {
            "id": "in-2",
            "titulo": "Engine SQLAlchemy",
            "categoria": "Modelado Cero-Toques",
            "descripcion": "ORM de Python que se comunica por puertos directamente con la base de datos viva. Pandas succiona vía 'read_sql' transformando TB de data relacional directamente a Arrays NumPy en memoria virtual.",
            "codigo": "from sqlalchemy import create_engine\nengine = create_engine('postgresql://user:pass@host/db')\ndf = pd.read_sql('SELECT * FROM churn_db', engine)"
        }
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
