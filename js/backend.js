/**
 * Data Pipeline Elite App - FAKE BACKEND ENGINE v4.5
 * Excel Sandbox Simulator + Gamification Engine
 */

document.addEventListener('DOMContentLoaded', () => {
    fetchDatabase();
});

// ─────────────────────────────────────────────────────────────────────────────
// BASE DE DATOS EN MEMORIA (Fake Backend)
// ─────────────────────────────────────────────────────────────────────────────
const functionDatabase = {
    "excel": [
        { "id": "ex-1",  "titulo": "Ctrl Shortcuts",          "categoria": "Atajos Fundamentales",    "descripcion": "Selección rápida de columna entera (Ctrl+Espacio), saltos al final del rango (Ctrl+Flechas), abrir formato de celdas (Ctrl+1).",    "codigo": "Ctrl + Espacio  →  Selecciona columna entera\nCtrl + 1        →  Abre diálogo Formato de Celdas\nCtrl + Fin      →  Última celda con datos" },
        { "id": "ex-2",  "titulo": "Shift/Alt Shortcuts",     "categoria": "Atajos Fundamentales",    "descripcion": "Selección de fila entera (Shift+Espacio), centrar contenido (Alt+H+A+C).",                                                          "codigo": "Shift + Espacio         →  Fila entera\nAlt + H + A + C         →  Centrar contenido" },
        { "id": "ex-3",  "titulo": "+, -, *, /, RAIZ()",      "categoria": "Básicas y Operaciones",   "descripcion": "Fundamentos aritméticos en la celda. RAIZ devuelve la raíz cuadrada de un número positivo.",                                          "codigo": "=RAIZ(144)    // Retorna 12\n=2^8          // Retorna 256" },
        { "id": "ex-4",  "titulo": "SUMA & PROMEDIO",         "categoria": "Matemáticas Estadísticas","descripcion": "SUMA agrega valores numéricos. PROMEDIO retorna la media aritmética del rango especificado.",                                          "codigo": "=SUMA(C1:C100)\n=PROMEDIO(A1:A100)" },
        { "id": "ex-5",  "titulo": "CONTARA() vs CONTAR()",   "categoria": "Matemáticas Estadísticas","descripcion": "CONTAR solo cuenta numéricos. CONTARA cuenta cualquier celda no vacía, ideal para strings.",                                          "codigo": "=CONTAR(B2:B500)    // Solo números\n=CONTARA(B2:B500)   // Cualquier valor" },
        { "id": "ex-6",  "titulo": "SUMAR.SI & CONTAR.SI",    "categoria": "Matemáticas Estadísticas","descripcion": "Suma o cuenta condicionalmente los valores que cumplan un criterio dado.",                                                             "codigo": "=SUMAR.SI(A2:A100, \"Norte\", C2:C100)\n=CONTAR.SI(B2:B500, \">1000\")" },
        { "id": "ex-7",  "titulo": "SI() & SIERROR()",        "categoria": "Lógicas y Limpieza",      "descripcion": "SI evalúa una condición. SIERROR captura errores (#N/A, #¡VALOR!) devolviendo un valor alternativo.",                               "codigo": "=SI(A1>100, \"Alto\", \"Bajo\")\n=SIERROR(BUSCARV(A1,M:N,2,0), \"No encontrado\")" },
        { "id": "ex-8",  "titulo": "Y() / O() / NO()",        "categoria": "Lógicas Avanzadas",       "descripcion": "Operadores booleanos que se anidan dentro de SI() para evaluar múltiples condiciones al mismo tiempo.",                               "codigo": "=SI(Y(A1>10, B1=\"OK\"), \"Aprobado\", \"Rechazado\")\n=SI(O(A1<0, A1>100), \"Fuera de rango\", \"OK\")" },
        { "id": "ex-9",  "titulo": "TRIM (ESPACIOS)",         "categoria": "Limpieza Cadenas",        "descripcion": "ESPACIOS purga espacios iniciales, finales y dobles. Esencial antes de cualquier BUSCARV para evitar coincidencias fallidas.",       "codigo": "=ESPACIOS(A2)\n=SUSTITUIR(A2, \" \", \"\")" },
        { "id": "ex-10", "titulo": "VLOOKUP (BUSCARV)",       "categoria": "Búsqueda y Referencia",   "descripcion": "Busca un valor en la primera columna de una tabla y devuelve el valor de la columna indicada. Siempre usar FALSO para coincidir exacto.", "codigo": "=BUSCARV(A2, $E$1:$G$100, 2, FALSO)" },
        { "id": "ex-11", "titulo": "INDEX & MATCH",           "categoria": "Búsqueda y Referencia",   "descripcion": "INDEX extrae el valor. MATCH calcula la posición dinámica. Permite búsquedas hacia la izquierda, imposibles con BUSCARV.",           "codigo": "=INDICE(B2:B100, COINCIDIR(\"Target\", A2:A100, 0))" },
        { "id": "ex-12", "titulo": "XLOOKUP()",               "categoria": "Matrices Dinámicas",      "descripcion": "Evolución de BUSCARV. Busca en cualquier dirección, maneja errores nativamente y devuelve múltiples columnas.",                       "codigo": "=BUSCARX(A2, E2:E100, F2:F100, \"No encontrado\", 0)" },
        { "id": "ex-13", "titulo": "FILTRAR()",               "categoria": "Matrices Dinámicas",      "descripcion": "Devuelve un subconjunto del rango que cumpla una condición. El resultado \"desborda\" hacia celdas adyacentes automáticamente.",      "codigo": "=FILTRAR(A2:C100, B2:B100=\"VIP\")" },
        { "id": "ex-14", "titulo": "UNIQUE() y ORDENAR()",    "categoria": "Matrices Dinámicas",      "descripcion": "UNIQUE elimina duplicados del rango. ORDENAR clasifica el resultado alfabética o numéricamente.",                                     "codigo": "=ORDENAR(UNICOS(A2:A5000))" },
        { "id": "ex-15", "titulo": "Query Folding",           "categoria": "Lenguaje M (ETL)",        "descripcion": "Power Query delega WHERE/GROUP al motor origen (SQL Server) en lugar de bajar toda la data localmente. Crítico para tablas de millones de filas.", "codigo": "// Principio arquitectónico:\n// Filtrar filas ANTES de expandir columnas\nTable.SelectRows(Source, each [Pais] = \"CO\")" },
        { "id": "ex-16", "titulo": "Table.Buffer()",          "categoria": "Lenguaje M (ETL)",        "descripcion": "Fuerza que la tabla se evalúe y se almacene en memoria RAM, acelerando los cruces repetitivos.",                                      "codigo": "let\n  Src = Table.Buffer(BaseDatos),\n  Resultado = Table.Join(Src, ...)\nin Resultado" },
        { "id": "ex-17", "titulo": "Slicers y Timelines",     "categoria": "Tablas Dinámicas",        "descripcion": "Controles visuales interactivos que filtran Tablas Dinámicas sin alterar los datos subyacentes. Se conectan a múltiples pivots vía Report Connections.", "codigo": "// Insertar > Segmentación de datos\n// Conectar a múltiples Tablas Dinámicas:\n// Clic derecho > Conexiones de informe" }
    ],
    "sql": [
        { "id": "sq-1",  "titulo": "SELECT DISTINCT / LIMIT", "categoria": "Sintaxis Básica",             "descripcion": "DISTINCT colapsa duplicados. LIMIT restringe el resultado para no reventar la memoria del motor.", "codigo": "SELECT DISTINCT user_id FROM logs LIMIT 100;" },
        { "id": "sq-2",  "titulo": "Filtrado WHERE, AND, IN", "categoria": "Sintaxis Básica",             "descripcion": "Poda primaria PRE-agrupación. IN acepta listas de valores, evitando múltiples OR anidados.",       "codigo": "SELECT * FROM ventas WHERE pais IN ('CO','MX') AND monto > 0;" },
        { "id": "sq-3",  "titulo": "INNER JOIN",              "categoria": "Uniones Relacionales",         "descripcion": "Retorna únicamente los registros que tienen coincidencia en AMBAS tablas. Excluye huérfanos.",    "codigo": "SELECT u.nombre, o.total\nFROM users u\nINNER JOIN orders o ON u.id = o.user_id;" },
        { "id": "sq-4",  "titulo": "LEFT JOIN",               "categoria": "Uniones Relacionales",         "descripcion": "Retorna TODOS los de la izquierda + los que coincidan a la derecha. NULL donde no hay match.", "codigo": "SELECT u.nombre, o.total\nFROM users u\nLEFT JOIN orders o ON u.id = o.user_id;" },
        { "id": "sq-5",  "titulo": "FULL OUTER JOIN",         "categoria": "Uniones Relacionales",         "descripcion": "Combina todo de ambos lados, rellenando NULL donde no hay coincidencia.",                        "codigo": "SELECT * FROM A FULL OUTER JOIN B ON A.id = B.id;" },
        { "id": "sq-6",  "titulo": "ANTI-JOIN",               "categoria": "Uniones Relacionales",         "descripcion": "Encuentra registros en A que NO existen en B. Clave para detectar clientes sin pedidos.",       "codigo": "SELECT u.id FROM users u\nLEFT JOIN orders o ON u.id = o.user_id\nWHERE o.user_id IS NULL;" },
        { "id": "sq-7",  "titulo": "SUM / AVG / COUNT",       "categoria": "Agregación",                   "descripcion": "Funciones de agregación que colapsan millones de filas en KPIs relevantes.",                   "codigo": "SELECT region, SUM(monto), AVG(monto), COUNT(*)\nFROM ventas GROUP BY region;" },
        { "id": "sq-8",  "titulo": "GROUP BY vs HAVING",      "categoria": "Agregación",                   "descripcion": "HAVING filtra POST-agrupación. WHERE filtra PRE-agrupación. Error clásico de confundirlos.",   "codigo": "SELECT user_id, COUNT(*) as pedidos\nFROM orders\nGROUP BY user_id\nHAVING COUNT(*) > 5;" },
        { "id": "sq-9",  "titulo": "ROW_NUMBER()",            "categoria": "Window Functions",             "descripcion": "Enumera filas dentro de cada partición sin comprimir el dataset.",                             "codigo": "SELECT *, ROW_NUMBER() OVER(PARTITION BY user_id ORDER BY order_date DESC) as rn\nFROM orders;" },
        { "id": "sq-10", "titulo": "RANK() / DENSE_RANK()",   "categoria": "Window Functions",             "descripcion": "RANK salta números en empates; DENSE_RANK no. Útil para rankings de ventas.",                  "codigo": "SELECT nombre, ventas, DENSE_RANK() OVER(ORDER BY ventas DESC) as ranking FROM rep;" },
        { "id": "sq-11", "titulo": "LAG() / LEAD()",          "categoria": "Window Functions",             "descripcion": "LAG trae el valor de la fila anterior; LEAD el de la siguiente. Ideal para diferencias entre fechas.", "codigo": "SELECT user_id, order_date,\n  LAG(order_date) OVER(PARTITION BY user_id ORDER BY order_date) as prev_date\nFROM orders;" },
        { "id": "sq-12", "titulo": "NTILE()",                 "categoria": "Window Functions",             "descripcion": "Divide la población en N bloques de igual tamaño. Perfecto para cuartiles de RFM.",           "codigo": "SELECT user_id, monto, NTILE(4) OVER(ORDER BY monto DESC) as cuartil FROM ventas;" },
        { "id": "sq-13", "titulo": "CTEs (WITH)",             "categoria": "Subconsultas y Optimización",  "descripcion": "Nombra subqueries como tablas temporales, eliminando el horror de los subqueries anidados.",   "codigo": "WITH activos AS (\n  SELECT * FROM users WHERE activo = 1\n)\nSELECT * FROM activos WHERE pais = 'CO';" },
        { "id": "sq-14", "titulo": "EXPLAIN ANALYZE",         "categoria": "Subconsultas y Optimización",  "descripcion": "Muestra el plan de ejecución real del motor. Revela Index Scans vs Full Table Scans costosos.", "codigo": "EXPLAIN ANALYZE\nSELECT * FROM orders WHERE user_id = 42;" }
    ],
    "python": [
        { "id": "py-1",  "titulo": "Listas y Diccionarios",     "categoria": "Estructuras Nativas",       "descripcion": "Vectores mutables y mapas clave-valor. Columna vertebral de cualquier pipeline de datos.",    "codigo": "clientes = [{'id': 1, 'nombre': 'Ana'}, {'id': 2, 'nombre': 'Luis'}]\nprint(clientes[0]['nombre'])  # Ana" },
        { "id": "py-2",  "titulo": "Tuplas y Sets",             "categoria": "Estructuras Nativas",       "descripcion": "Tuplas: inmutables (coordenadas, configs). Sets: colecciones únicas para deduplicar IDs.",  "codigo": "coords = (4.6097, -74.0817)\nunique_ids = set([1, 2, 2, 3])  # {1, 2, 3}" },
        { "id": "py-3",  "titulo": "for / while / if-elif",    "categoria": "Control de Flujo",          "descripcion": "Estructuras de iteración y decisión. Base de cualquier pipeline de transformación.",        "codigo": "for cliente in clientes:\n    if cliente['gasto'] > 1000:\n        print('VIP')\n    elif cliente['gasto'] > 500:\n        print('Medio')" },
        { "id": "py-4",  "titulo": "Funciones Lambda",          "categoria": "Funciones y Decoradores",   "descripcion": "Funciones anónimas desechables. Ideales para .apply() de Pandas sin definir funciones completas.", "codigo": "df['segmento'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Estándar')" },
        { "id": "py-5",  "titulo": "Decoradores (@)",           "categoria": "Funciones y Decoradores",   "descripcion": "Envuelven una función añadiendo lógica antes/después sin modificar su cuerpo.",              "codigo": "@timer\ndef cargar_datos():\n    df = pd.read_csv('ventas.csv')\n    return df" },
        { "id": "py-6",  "titulo": "df.head() / df.describe()", "categoria": "Pandas Exploración",        "descripcion": "Exploración rápida: head muestra primeras filas, describe genera estadísticas de distribución.", "codigo": "df.head(10)\ndf.describe().T  # Transpuesto para legibilidad" },
        { "id": "py-7",  "titulo": "df.info()",                 "categoria": "Pandas Exploración",        "descripcion": "Muestra dtypes, conteo de no-nulos por columna y uso de memoria. Primer paso de auditoría.", "codigo": "df.info()\n# Resultado: columanas, dtypes, non-null counts, memory usage" },
        { "id": "py-8",  "titulo": "df.dropna()",               "categoria": "Limpieza Wrangling",        "descripcion": "Elimina filas/columnas con NaN. axis=0 borra filas; axis=1 borra columnas.",                "codigo": "df_limpio = df.dropna(axis=0)         # Elimina filas con cualquier NaN\ndf_limpio = df.dropna(subset=['email'])  # Solo si email es nulo" },
        { "id": "py-9",  "titulo": "df.fillna()",               "categoria": "Limpieza Wrangling",        "descripcion": "Imputa valores faltantes con media, mediana, o un valor constante para preservar el volumen del dataset.", "codigo": "df['edad'].fillna(df['edad'].median(), inplace=True)\ndf['pais'].fillna('Desconocido', inplace=True)" },
        { "id": "py-10", "titulo": "df.astype()",               "categoria": "Limpieza Wrangling",        "descripcion": "Convierte el dtype de una columna. Crítico para parsear fechas, enteros o categóricas.",  "codigo": "df['precio'] = df['precio'].astype('float64')\ndf['fecha'] = pd.to_datetime(df['fecha'])" },
        { "id": "py-11", "titulo": "df.drop_duplicates()",      "categoria": "Limpieza Wrangling",        "descripcion": "Elimina filas duplicadas. keep='last' conserva la más reciente; keep=False elimina todas.", "codigo": "df = df.drop_duplicates(subset=['email'], keep='last')" },
        { "id": "py-12", "titulo": "df.groupby()",              "categoria": "Transformación Tabular",    "descripcion": "Agrupa el DataFrame por una clave y aplica funciones de agregación. Equivalente a GROUP BY.", "codigo": "rfm = df.groupby('cliente_id').agg(\n    frecuencia=('orden_id', 'count'),\n    monetario=('monto', 'sum')\n)" },
        { "id": "py-13", "titulo": "df.merge()",                "categoria": "Transformación Tabular",    "descripcion": "JOIN entre DataFrames. Soporta how='left','right','inner','outer'.",                       "codigo": "resultado = pd.merge(clientes, ordenes, on='cliente_id', how='left')" },
        { "id": "py-14", "titulo": "Seaborn: heatmap / kdeplot","categoria": "Visualización Científica",  "descripcion": "heatmap visualiza correlaciones entre variables. kdeplot muestra distribución de densidad.",  "codigo": "import seaborn as sns\nsns.heatmap(df.corr(), annot=True, cmap='coolwarm')\nsns.kdeplot(data=df, x='gasto')" }
    ],
    "github": [
        { "id": "gh-1", "titulo": "git config",          "categoria": "Configuración Core",    "descripcion": "Define tu identidad en Git. Obligatorio para que los commits queden firmados con tu nombre y email.", "codigo": "git config --global user.name \"Tu Nombre\"\ngit config --global user.email \"tu@email.com\"" },
        { "id": "gh-2", "titulo": "git init / clone",    "categoria": "Configuración Core",    "descripcion": "init crea un repositorio local vacío. clone descarga una copia completa de un repositorio remoto.",  "codigo": "git init\ngit clone https://github.com/usuario/repo.git" },
        { "id": "gh-3", "titulo": "git status",          "categoria": "Flujo Workdir",         "descripcion": "Muestra el estado de tus archivos: modificados (rojo), en staging (verde), sin seguimiento.",       "codigo": "git status" },
        { "id": "gh-4", "titulo": "git add",             "categoria": "Flujo Workdir",         "descripcion": "Mueve cambios al Staging Area (muelle de carga). El punto agrega todos los archivos modificados.",  "codigo": "git add .\ngit add js/backend.js   # Solo un archivo específico" },
        { "id": "gh-5", "titulo": "git commit",          "categoria": "Historia Inmutable",    "descripcion": "Congela el staging en un hash SHA permanente. El mensaje debe ser descriptivo y en presente.",       "codigo": "git commit -m \"feat: Add Excel sandbox simulator v4.5\"" },
        { "id": "gh-6", "titulo": "git push / pull",     "categoria": "Nube (Origin)",         "descripcion": "push sube commits locales al servidor. pull descarga y fusiona los cambios remotos.",               "codigo": "git push origin main\ngit pull origin main" },
        { "id": "gh-7", "titulo": "git branch / checkout","categoria": "Seguridad Múltiple",   "descripcion": "branch lista o crea ramas. checkout te mueve entre ramas. Nunca trabajes directo en main.",       "codigo": "git branch feature/sandbox\ngit checkout feature/sandbox\n# O en una línea:\ngit checkout -b feature/sandbox" },
        { "id": "gh-8", "titulo": "Pull Requests",       "categoria": "Colaboración Corp",     "descripcion": "Mecanismo de revisión de código en GitHub. Tu rama solicita fusionarse con main tras revisión.",   "codigo": "# 1. Sube tu rama:\ngit push origin feature/sandbox\n# 2. En GitHub > Compare & Pull Request\n# 3. Describe cambios y pide revisión" },
        { "id": "gh-9", "titulo": "Forks",               "categoria": "Colaboración Corp",     "descripcion": "Copia completa de un repositorio en tu cuenta. Base del Open Source: modificas sin afectar el original.", "codigo": "# En GitHub > Botón 'Fork'\n# Luego clona TU fork:\ngit clone https://github.com/TU_USUARIO/repo.git" }
    ],
    "integration": [
        { "id": "in-1", "titulo": "=PY() Pandas Engine",   "categoria": "Python nativo Excel", "descripcion": "Microsoft 365 ejecuta Python en Azure, directamente en celdas Excel. Accede al rango con xl().",          "codigo": "=PY(\nimport pandas as pd\ndf = xl(\"Tabla1[#Todos]\", headers=True)\nreturn df.describe()\n)" },
        { "id": "in-2", "titulo": "Engine SQLAlchemy",     "categoria": "ETL Cero-Fricciones", "descripcion": "ORM que conecta Pandas directamente a cualquier base de datos relacional sin exportar CSV.",            "codigo": "from sqlalchemy import create_engine\nimport pandas as pd\nengine = create_engine('postgresql://user:pass@host/db')\ndf = pd.read_sql('SELECT * FROM churn_view', engine)" },
        { "id": "in-3", "titulo": "df.to_sql()",           "categoria": "ETL Cero-Fricciones", "descripcion": "Escribe un DataFrame directamente a una tabla SQL. if_exists='replace' sobreescribe; 'append' agrega.", "codigo": "df_predicciones.to_sql(\n    'ml_predictions',\n    engine,\n    if_exists='replace',\n    index=False\n)" }
    ]
};

// ─────────────────────────────────────────────────────────────────────────────
// MOTOR DE VALIDADORES ESPECÍFICOS POR EJERCICIO
// ─────────────────────────────────────────────────────────────────────────────
const EXERCISE_MAP = {
    // EXCEL — HARDWARE (atajos físicos de teclado)
    "ex-1": { type:"hardware", key:"ctrl+space",  instruccion:"Presiona Ctrl + Espacio en tu teclado para seleccionar una columna entera.", exito:"¡Ctrl + Espacio detectado! Columna seleccionada." },
    "ex-2": { type:"hardware", key:"shift+space", instruccion:"Presiona Shift + Espacio en tu teclado para seleccionar una fila entera.", exito:"¡Shift + Espacio detectado! Fila seleccionada." },

    // EXCEL — FÓRMULAS (se escribe en la barra fx)
    "ex-3":  { type:"formula", regex:/^=RAIZ\(144\)$/i,                  instruccion:"Escribe en la barra fx la fórmula para calcular la raíz cuadrada de 144.", exito:"¡=RAIZ(144) correcto! Resultado: 12" },
    "ex-4":  { type:"formula", regex:/^=PROMEDIO\(A1:A100\)$/i,          instruccion:"Escribe =PROMEDIO() del rango A1:A100 en la barra fx.", exito:"¡=PROMEDIO correcto!" },
    "ex-5":  { type:"formula", regex:/^=CONTARA\(B2:B500\)$/i,           instruccion:"Cuenta todas las celdas no vacías del rango B2:B500 con CONTARA.", exito:"¡=CONTARA correcto! Cuenta cualquier tipo de valor." },
    "ex-6":  { type:"formula", regex:/^=SUMAR\.SI\(.+,.+,.+\)$/i,        instruccion:"Escribe =SUMAR.SI() con 3 argumentos: rango zona, criterio y rango suma.", exito:"¡=SUMAR.SI correcto! Suma condicional validada." },
    "ex-7":  { type:"formula", regex:/^=SI\(.+,.+,.+\)$/i,               instruccion:"Escribe =SI(A1>100, \"Alto\", \"Bajo\") en la barra fx.", exito:"¡=SI() correcto! Función condicional validada." },
    "ex-8":  { type:"formula", regex:/^=SI\(\s*Y\(.+\)|=SI\(\s*O\(.+\)/i,instruccion:"Escribe =SI(Y(A1>10, B1=\"OK\"), \"Aprobado\", \"Rechazado\") para lógica múltiple.", exito:"¡Operadores booleanos Y/O correctos!" },
    "ex-9":  { type:"formula", regex:/^=ESPACIOS\(A\d+\)$/i,             instruccion:"Aplica ESPACIOS() sobre la celda A2 para limpiar espacios extra.", exito:"¡=ESPACIOS correcto! Cadena limpiada." },
    "ex-10": { type:"formula", regex:/^=BUSCARV\(.+\)$/i,                instruccion:"Escribe un BUSCARV completo: =BUSCARV(A2, $E$1:$G$100, 2, FALSO)", exito:"¡=BUSCARV correcto! Cruce a la derecha validado." },
    "ex-11": { type:"formula", regex:/^=INDICE\(.+,\s*COINCIDIR\(.+\)\)$/i,instruccion:"Escribe =INDICE(B2:B100, COINCIDIR(\"Target\", A2:A100, 0))", exito:"¡INDICE+COINCIDIR correcto! Cruce bidireccional." },
    "ex-12": { type:"formula", regex:/^=BUSCARX\(.+\)$/i,                instruccion:"Escribe =BUSCARX(A2, E2:E100, F2:F100, \"No encontrado\", 0)", exito:"¡=BUSCARX correcto! El poder de XLOOKUP validado." },
    "ex-13": { type:"formula", regex:/^=FILTRAR\(.+\)$/i,                instruccion:"Escribe =FILTRAR(A2:C100, B2:B100=\"VIP\") para extraer solo los clientes VIP.", exito:"¡=FILTRAR correcto! Subconjunto dinámico validado." },
    "ex-14": { type:"formula", regex:/^=ORDENAR\(UNICOS\(.+\)\)|^=UNICOS\(.+\)/i,instruccion:"Escribe =ORDENAR(UNICOS(A2:A5000)) para obtener únicos ordenados.", exito:"¡ORDENAR+UNICOS correcto! Lista deduplicada y ordenada." },
    "ex-15": { type:"terminal",regex:/Table\.SelectRows|Table\.Buffer|query.fold/i,instruccion:"Escribe la función M: Table.SelectRows(Source, each [Pais] = \"CO\")", exito:"¡Query Folding M correcto! Filtrado delegado al origen." },
    "ex-16": { type:"terminal",regex:/Table\.Buffer\(/i,                  instruccion:"Escribe la expresión M: Source = Table.Buffer(BaseDatos)", exito:"¡Table.Buffer() correcto! Tabla cargada en RAM." },
    "ex-17": { type:"terminal",regex:/slicer|segmentacion|timeline|pivot/i,instruccion:"Describe el flujo: cómo insertas un Slicer y lo conectas a plusieurs Tablas Dinámicas.", exito:"¡Concepto de Slicer validado! Dashboard interactivo." },

    // SQL
    "sq-1":  { type:"terminal", regex:/SELECT\s+DISTINCT/i,              instruccion:"Escribe SELECT DISTINCT user_id FROM logs LIMIT 100;", exito:"¡SELECT DISTINCT correcto!" },
    "sq-2":  { type:"terminal", regex:/WHERE.+(IN\s*\(|AND|OR)/i,        instruccion:"Filtra con WHERE pais IN ('CO','MX') en una consulta.", exito:"¡WHERE + IN correcto!" },
    "sq-3":  { type:"terminal", regex:/INNER\s+JOIN/i,                   instruccion:"Escribe un INNER JOIN entre users y orders.", exito:"¡INNER JOIN correcto! Solo registros coincidentes." },
    "sq-4":  { type:"terminal", regex:/LEFT\s+JOIN/i,                    instruccion:"Escribe un LEFT JOIN de users a orders.", exito:"¡LEFT JOIN correcto! Todos los users incluidos." },
    "sq-5":  { type:"terminal", regex:/FULL\s+(OUTER\s+)?JOIN/i,         instruccion:"Escribe FULL OUTER JOIN entre tabla A y B.", exito:"¡FULL OUTER JOIN correcto!" },
    "sq-6":  { type:"terminal", regex:/LEFT\s+JOIN.+IS\s+NULL/is,        instruccion:"Escribe el ANTI-JOIN: LEFT JOIN + WHERE columna IS NULL.", exito:"¡ANTI-JOIN correcto! Huérfanos identificados." },
    "sq-7":  { type:"terminal", regex:/SUM\(|AVG\(|COUNT\(/i,            instruccion:"Escribe SELECT region, SUM(monto) FROM ventas GROUP BY region;", exito:"¡Agregación SUM/AVG/COUNT correcta!" },
    "sq-8":  { type:"terminal", regex:/HAVING/i,                         instruccion:"Escribe una query con GROUP BY y HAVING COUNT(*) > 5.", exito:"¡HAVING correcto! Filtro post-agrupación." },
    "sq-9":  { type:"terminal", regex:/ROW_NUMBER\(\)\s+OVER/i,          instruccion:"Escribe ROW_NUMBER() OVER(PARTITION BY user_id ORDER BY order_date DESC).", exito:"¡ROW_NUMBER() correcto! Window function operativa." },
    "sq-10": { type:"terminal", regex:/DENSE_RANK\(\)|RANK\(\)/i,        instruccion:"Escribe DENSE_RANK() OVER(ORDER BY ventas DESC) para el ranking.", exito:"¡DENSE_RANK correcto!" },
    "sq-11": { type:"terminal", regex:/LAG\(|LEAD\(/i,                   instruccion:"Usa LAG(order_date) OVER(PARTITION BY user_id ORDER BY order_date) para fecha anterior.", exito:"¡LAG/LEAD correcto! Desplazamiento temporal." },
    "sq-12": { type:"terminal", regex:/NTILE\(\d+\)/i,                   instruccion:"Escribe NTILE(4) OVER(ORDER BY monto DESC) para cuartiles.", exito:"¡NTILE correcto! Cuartiles generados." },
    "sq-13": { type:"terminal", regex:/WITH\s+\w+\s+AS\s*\(/i,           instruccion:"Escribe un CTE: WITH activos AS (SELECT * FROM users WHERE activo=1) SELECT * FROM activos;", exito:"¡CTE correcto! Subquery nombrado." },
    "sq-14": { type:"terminal", regex:/EXPLAIN(\s+ANALYZE)?/i,           instruccion:"Escribe EXPLAIN ANALYZE SELECT * FROM orders WHERE user_id = 42;", exito:"¡EXPLAIN correcto! Plan de ejecución visible." },

    // PYTHON
    "py-1":  { type:"terminal", regex:/\[|\{/,                           instruccion:"Crea un diccionario Python: clientes = {'id': 1, 'nombre': 'Ana'}", exito:"¡Diccionario correcto! Estructura de datos lista." },
    "py-2":  { type:"terminal", regex:/set\(|tuple\(/i,                  instruccion:"Crea un set de IDs únicos: ids = set([1,2,2,3])", exito:"¡Set/Tupla correcto! Uniqueness validada." },
    "py-3":  { type:"terminal", regex:/for\s+\w+\s+in|while\s+/,         instruccion:"Escribe un bucle for iterando sobre una lista de clientes.", exito:"¡Bucle for correcto!" },
    "py-4":  { type:"terminal", regex:/lambda/,                          instruccion:"Escribe: df['seg'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Std')", exito:"¡Lambda correcto! Función anónima aplicada." },
    "py-5":  { type:"terminal", regex:/@\w+/,                            instruccion:"Define un decorador con @timer sobre una función de carga de datos.", exito:"¡Decorador correcto!" },
    "py-6":  { type:"terminal", regex:/df\.(head|describe)\(/,           instruccion:"Escribe df.describe().T para explorar el DataFrame.", exito:"¡df.describe() correcto! EDA iniciada." },
    "py-7":  { type:"terminal", regex:/df\.info\(\)/,                    instruccion:"Escribe df.info() para inspeccionar los dtypes y nulos.", exito:"¡df.info() correcto! Auditoría completada." },
    "py-8":  { type:"terminal", regex:/df\.dropna\(/,                    instruccion:"Escribe df.dropna(axis=0) o df.dropna(subset=['email']) para purgar nulos.", exito:"¡df.dropna() correcto! NaNs eliminados." },
    "py-9":  { type:"terminal", regex:/df\[.+\]\.fillna\(/,              instruccion:"Escribe df['edad'].fillna(df['edad'].median(), inplace=True)", exito:"¡df.fillna() correcto! Missings imputados." },
    "py-10": { type:"terminal", regex:/df\[.+\]\.astype\(/,              instruccion:"Escribe df['precio'] = df['precio'].astype('float64')", exito:"¡.astype() correcto! Dtype convertido." },
    "py-11": { type:"terminal", regex:/drop_duplicates\(/,               instruccion:"Escribe df.drop_duplicates(subset=['email'], keep='last')", exito:"¡drop_duplicates correcto! Duplicados purgados." },
    "py-12": { type:"terminal", regex:/groupby\(/,                       instruccion:"Escribe df.groupby('cliente_id').agg({'orden_id': 'count', 'monto': 'sum'})", exito:"¡groupby correcto! Agregación RFM completada." },
    "py-13": { type:"terminal", regex:/pd\.merge\(/,                     instruccion:"Escribe pd.merge(clientes, ordenes, on='cliente_id', how='left')", exito:"¡pd.merge correcto! JOIN entre DataFrames." },
    "py-14": { type:"terminal", regex:/heatmap\(|kdeplot\(/,             instruccion:"Escribe sns.heatmap(df.corr(), annot=True) para visualizar correlaciones.", exito:"¡Seaborn correcto! Visualización list." },

    // GITHUB
    "gh-1":  { type:"terminal", regex:/git\s+config\s+--global/,       instruccion:"Escribe git config --global user.name \"TuNombre\"", exito:"¡git config correcto! Identidad registrada." },
    "gh-2":  { type:"terminal", regex:/git\s+(init|clone)/,             instruccion:"Escribe git clone https://github.com/usuario/repo.git", exito:"¡git clone correcto! Repositorio clonado." },
    "gh-3":  { type:"terminal", regex:/git\s+status/,                   instruccion:"Escribe git status para ver el estado del working directory.", exito:"¡git status correcto! Estado del repo visible." },
    "gh-4":  { type:"terminal", regex:/git\s+add/,                      instruccion:"Escribe git add . para agregar todos los cambios al staging.", exito:"¡git add correcto! Archivos en staging." },
    "gh-5":  { type:"terminal", regex:/git\s+commit\s+-m/,              instruccion:"Escribe git commit -m \"feat: descripción del cambio\"", exito:"¡git commit correcto! Cambio sellado en historia." },
    "gh-6":  { type:"terminal", regex:/git\s+(push|pull)/,              instruccion:"Escribe git push origin main para subir a GitHub.", exito:"¡git push correcto! Commits en la nube." },
    "gh-7":  { type:"terminal", regex:/git\s+(branch|checkout)/,        instruccion:"Escribe git checkout -b feature/mi-rama para crear y entrar a una rama.", exito:"¡git branch correcto! Rama aislada creada." },
    "gh-8":  { type:"terminal", regex:/git\s+push.+origin|pull.request|PR/i, instruccion:"Escribe git push origin feature/mi-rama (primer paso para abrir un PR).", exito:"¡PR flow correcto! Rama publicada, listo para revisar." },
    "gh-9":  { type:"terminal", regex:/fork|clone.+github/i,            instruccion:"Describe el flujo de Fork: clone el fork y usa git remote add upstream ...", exito:"¡Fork workflow correcto!" },

    // INTEGRACIONES
    "in-1":  { type:"formula", regex:/^=PY\(/i,                          instruccion:"Escribe =PY( para iniciar una celda Python en Excel.", exito:"¡=PY() correcto! Motor Python/Azure activado." },
    "in-2":  { type:"terminal", regex:/create_engine\(/,                 instruccion:"Escribe from sqlalchemy import create_engine y llama create_engine('postgresql://...')", exito:"¡SQLAlchemy correcto! Pipe DB-Python establecido." },
    "in-3":  { type:"terminal", regex:/\.to_sql\(/,                      instruccion:"Escribe df_predicciones.to_sql('tabla', engine, if_exists='replace', index=False)", exito:"¡to_sql() correcto! DataFrame escrito en DB." }
};

// ─────────────────────────────────────────────────────────────────────────────
// ESTADO GLOBAL
// ─────────────────────────────────────────────────────────────────────────────
let currentExercise = null;  // { type, regex/key, instruccion, exito }

// ─────────────────────────────────────────────────────────────────────────────
// INICIALIZACIÓN
// ─────────────────────────────────────────────────────────────────────────────
function fetchDatabase() {
    renderAllGrids();
}

function renderAllGrids() {
    for (const module of ['excel','sql','python','github','integration']) {
        if (functionDatabase[module]) renderGrid(`grid-${module}`, functionDatabase[module], module);
    }
}

function renderGrid(gridId, dataArray, moduleTheme) {
    const grid = document.getElementById(gridId);
    if (!grid) return;
    grid.innerHTML = dataArray.map(item => `
        <div class="function-card ${moduleTheme}-hover" onclick="openFunctionModal('${item.id}','${moduleTheme}')">
            <span class="card-category ${moduleTheme}-color">${item.categoria}</span>
            <h3 class="card-title">${item.titulo}</h3>
            <p class="card-excerpt">${item.descripcion}</p>
        </div>`).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// MODAL: ABRIR
// ─────────────────────────────────────────────────────────────────────────────
function openFunctionModal(itemId, themeName) {
    let data = null;
    for (const module in functionDatabase) {
        data = functionDatabase[module].find(i => i.id === itemId);
        if (data) break;
    }
    if (!data) return;

    // Rellenar Teoría
    document.getElementById('modal-title').innerText    = data.titulo;
    document.getElementById('modal-category').innerText = data.categoria;
    document.getElementById('modal-desc').innerText     = data.descripcion;
    document.getElementById('modal-code').innerText     = data.codigo;

    // Obtener definición de ejercicio
    currentExercise = EXERCISE_MAP[itemId] || {
        type: 'terminal',
        regex: /.*/,
        instruccion: `Demuestra el uso de ${data.titulo}.`,
        exito: `¡${data.titulo} correcto!`
    };

    // Rellenar descripción de misión
    document.getElementById('modal-exercise-desc').innerText = currentExercise.instruccion;

    // Elegir sandbox según tipo
    const useExcel = (currentExercise.type === 'formula' || currentExercise.type === 'hardware');
    const termBox  = document.getElementById('terminal-sandbox');
    const excelBox = document.getElementById('excel-sandbox');

    if (useExcel) {
        termBox.classList.add('hidden');
        excelBox.classList.remove('hidden');
        resetExcelSandbox(currentExercise.type === 'hardware');
    } else {
        excelBox.classList.add('hidden');
        termBox.classList.remove('hidden');
        resetTerminalSandbox();
    }

    // Color de categoría
    const catSpan = document.getElementById('modal-category');
    catSpan.className = `${themeName}-color`;

    switchModalTab('teoria');
    document.getElementById('dynamic-modal').classList.add('active');
}

function resetExcelSandbox(isHardware) {
    const bar     = document.getElementById('excel-formula-bar');
    const overlay = document.getElementById('key-listener-overlay');
    const fb      = document.getElementById('excel-feedback');
    bar.value         = '';
    bar.disabled      = isHardware;
    fb.innerHTML      = isHardware
        ? '> Modo Hardware activo — Presiona las teclas indicadas en tu teclado físico.'
        : '> Escribe tu fórmula en la barra fx y pulsa ✓ o Enter.';
    fb.style.color    = '#a5b4fc';
    overlay.classList.toggle('hidden', !isHardware);
}

function resetTerminalSandbox() {
    const inp = document.getElementById('modal-code-input');
    const fb  = document.getElementById('modal-feedback');
    inp.value        = '';
    inp.placeholder  = 'Escribe el comando/sentencia aquí...';
    fb.innerHTML     = '> Esperando input...';
    fb.style.color   = '#a5b4fc';
}

// ─────────────────────────────────────────────────────────────────────────────
// MODAL: TABS
// ─────────────────────────────────────────────────────────────────────────────
function switchModalTab(tabId) {
    document.querySelectorAll('.modal-tab').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + tabId).classList.add('active');
    document.getElementById('modal-body-teoria').classList.add('hidden');
    document.getElementById('modal-body-practica').classList.add('hidden');
    document.getElementById('modal-body-' + tabId).classList.remove('hidden');
}

// ─────────────────────────────────────────────────────────────────────────────
// VALIDADORES
// ─────────────────────────────────────────────────────────────────────────────

// Terminal (SQL / Python / GitHub / Power Query conceptual)
function verifyMicroLab() {
    if (!currentExercise) return;
    const val = document.getElementById('modal-code-input').value.trim();
    const fb  = document.getElementById('modal-feedback');
    if (!val) { fb.innerHTML = '> Error: terminal vacío.'; fb.style.color = '#ef4444'; return; }

    let ok;
    try { ok = currentExercise.regex.test(val); } catch(e) { ok = true; }
    if (ok) {
        fb.innerHTML   = `> ✅ [SUCCESS] ${currentExercise.exito}`;
        fb.style.color = '#10b981';
    } else {
        fb.innerHTML   = `> ❌ [FAILED] Sintaxis incorrecta. Intenta de nuevo.`;
        fb.style.color = '#ef4444';
    }
}

// Excel — Fórmula (barra fx)
function verifyExcelLab() {
    if (!currentExercise || currentExercise.type !== 'formula') return;
    const val = document.getElementById('excel-formula-bar').value.trim();
    const fb  = document.getElementById('excel-feedback');
    if (!val) { fb.innerHTML = '> Celda vacía.'; fb.style.color = '#ef4444'; return; }

    let ok;
    try { ok = currentExercise.regex.test(val); } catch(e) { ok = true; }
    if (ok) {
        fb.innerHTML   = `> ✅ [CELL SUCCESS] ${currentExercise.exito}`;
        fb.style.color = '#10b981';
        // Escribir resultado visual en celda C3
        const c3 = document.querySelector('.highlight-cell');
        if (c3) c3.innerText = '✓';
    } else {
        fb.innerHTML   = `> ❌ Fórmula incorrecta. Ejemplo esperado: ${currentExercise.instruccion}`;
        fb.style.color = '#ef4444';
    }
}

// Excel — Hardware (teclado físico)
document.addEventListener('keydown', function(e) {
    const modal = document.getElementById('dynamic-modal');
    if (!modal || !modal.classList.contains('active')) return;
    if (!currentExercise || currentExercise.type !== 'hardware') return;
    // No interceptar cuando el usuario escribe en inputs
    const tag = document.activeElement.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    e.preventDefault(); // evitar scroll del modal con Espacio

    const fb = document.getElementById('excel-feedback');
    let ok = false;
    if (currentExercise.key === 'ctrl+space'  && e.ctrlKey  && e.code === 'Space') ok = true;
    if (currentExercise.key === 'shift+space' && e.shiftKey && e.code === 'Space') ok = true;

    if (ok) {
        fb.innerHTML = `> ✅ [HARDWARE SUCCESS] ${currentExercise.exito}`;
        fb.style.color = '#10b981';
        document.getElementById('key-listener-overlay').classList.add('hidden');
    } else {
        fb.innerHTML = `> 🎹 Tecla '${e.key}' detectada — sigue intentando...`;
        fb.style.color = '#fbbf24';
    }
});

// Selección visual de celda
function selectCell(element, cellId) {
    document.querySelectorAll('.cell').forEach(c => c.classList.remove('selected'));
    element.classList.add('selected');
    document.getElementById('excel-formula-bar').focus();
}

// ─────────────────────────────────────────────────────────────────────────────
// MODAL: CERRAR
// ─────────────────────────────────────────────────────────────────────────────
function closeModal(e) {
    if (e && e.target !== document.getElementById('dynamic-modal') &&
             e.target !== document.querySelector('.btn-close')) return;
    document.getElementById('dynamic-modal').classList.remove('active');
}
