/**
 * Data Pipeline Elite App — BACKEND ENGINE v4.5 Live
 * Formula Evaluator + Context-Aware Grid + Hardware Keydown
 */

document.addEventListener('DOMContentLoaded', fetchDatabase);

// ─────────────────────────────────────────────────────────────────────────────
// DATABASE
// ─────────────────────────────────────────────────────────────────────────────
const functionDatabase = {
    "excel": [
        { "id":"ex-1",  "titulo":"Ctrl Shortcuts",           "categoria":"Atajos Fundamentales",     "descripcion":"Selección de columna entera (Ctrl+Espacio), saltos al fin del rango (Ctrl+Flechas), abrir Formato de Celdas (Ctrl+1).",            "codigo":"Ctrl + Espacio  →  Selecciona columna entera\nCtrl + 1        →  Formato de Celdas\nCtrl + Fin      →  Última celda con datos" },
        { "id":"ex-2",  "titulo":"Shift/Alt Shortcuts",      "categoria":"Atajos Fundamentales",     "descripcion":"Shift+Espacio selecciona la fila entera. Alt+H+A+C centra el contenido de la celda activa desde el Ribbon.",                         "codigo":"Shift + Espacio  →  Fila entera\nAlt + H + A + C →  Centrar contenido" },
        { "id":"ex-3",  "titulo":"+, -, *, /, RAIZ()",       "categoria":"Básicas y Operaciones",    "descripcion":"Operaciones aritméticas elementales. RAIZ() calcula la raíz cuadrada positiva. Siempre inicia la celda con = para activar el motor de cálculo.", "codigo":"=RAIZ(144)  →  12\n=2^8        →  256\n=(A1+B1)*C1 →  Operación combinada" },
        { "id":"ex-4",  "titulo":"SUMA & PROMEDIO",          "categoria":"Matemáticas Estadísticas", "descripcion":"SUMA() agrega el rango. PROMEDIO() retorna la media aritmética. Ambas ignoran celdas vacías y textos.",                              "codigo":"=SUMA(B2:B6)\n=PROMEDIO(B2:B6)" },
        { "id":"ex-5",  "titulo":"CONTARA() vs CONTAR()",    "categoria":"Matemáticas Estadísticas", "descripcion":"CONTAR() cuenta solo numéricos. CONTARA() cuenta cualquier celda no vacía, incluyendo texto.",                                        "codigo":"=CONTAR(B2:B6)    →  Cuenta números\n=CONTARA(A2:A6)   →  Cuenta no-vacíos" },
        { "id":"ex-6",  "titulo":"SUMAR.SI & CONTAR.SI",     "categoria":"Matemáticas Estadísticas", "descripcion":"Suma o cuenta solo los registros que cumplen un criterio. Muy usada para segmentar ventas por región o categoría.",                    "codigo":"=SUMAR.SI(A2:A6,\"Norte\",B2:B6)\n=CONTAR.SI(A2:A6,\"Norte\")" },
        { "id":"ex-7",  "titulo":"SI() & SIERROR()",         "categoria":"Lógicas y Limpieza",       "descripcion":"SI() evalúa condición y devuelve uno de dos resultados. SIERROR() captura errores devolviendo un valor alternativo seguro.",           "codigo":"=SI(B2>500,\"Alto\",\"Bajo\")\n=SIERROR(BUSCARV(A2,E:F,2,0),\"N/A\")" },
        { "id":"ex-8",  "titulo":"Y() / O() / NO()",         "categoria":"Lógicas Avanzadas",        "descripcion":"Y() exige que TODAS las condiciones sean verdad. O() basta con que UNA lo sea. Siempre anidados dentro de SI().",                      "codigo":"=SI(Y(B2>100,C2=\"OK\"),\"Aprobado\",\"Rechazado\")" },
        { "id":"ex-9",  "titulo":"TRIM (ESPACIOS)",          "categoria":"Limpieza Cadenas",         "descripcion":"ESPACIOS() elimina espacios al inicio, al final y duplicados internos. Clave antes de cualquier BUSCARV.",                            "codigo":"=ESPACIOS(A2)" },
        { "id":"ex-10", "titulo":"VLOOKUP (BUSCARV)",        "categoria":"Búsqueda y Referencia",    "descripcion":"Busca un valor en la primera columna de una tabla y devuelve la columna indicada. FALSO obliga coincidencia exacta.",                 "codigo":"=BUSCARV(A2,$E$2:$G$6,2,FALSO)" },
        { "id":"ex-11", "titulo":"INDEX & MATCH",            "categoria":"Búsqueda y Referencia",    "descripcion":"INDICE extrae el valor exacto, COINCIDIR calcula la posición. Permite búsquedas hacia la izquierda imposibles con BUSCARV.",          "codigo":"=INDICE(C2:C6,COINCIDIR(A9,A2:A6,0))" },
        { "id":"ex-12", "titulo":"XLOOKUP (BUSCARX)",        "categoria":"Matrices Dinámicas",       "descripcion":"Evolución de BUSCARV en Microsoft 365. Busca en cualquier dirección, maneja errores nativamente y devuelve múltiples columnas.",       "codigo":"=BUSCARX(A9,A2:A6,C2:C6,\"No encontrado\",0)" },
        { "id":"ex-13", "titulo":"FILTRAR()",                "categoria":"Matrices Dinámicas",       "descripcion":"Devuelve un subconjunto dinámico del rango. El resultado 'desborda' a celdas adyacentes automáticamente (Spill).",                   "codigo":"=FILTRAR(A2:C6,C2:C6>400)" },
        { "id":"ex-14", "titulo":"UNIQUE() y ORDENAR()",     "categoria":"Matrices Dinámicas",       "descripcion":"UNIQUE() elimina duplicados del rango. ORDENAR() clasifica el resultado.",                                                           "codigo":"=ORDENAR(UNICOS(A2:A6))" },
        { "id":"ex-15", "titulo":"Query Folding",            "categoria":"Lenguaje M (ETL)",         "descripcion":"Power Query delega WHERE/GROUP al motor de origen (SQL) en lugar de bajar toda la data localmente.",                                  "codigo":"Table.SelectRows(Source, each [Region] = \"Norte\")" },
        { "id":"ex-16", "titulo":"Table.Buffer()",           "categoria":"Lenguaje M (ETL)",         "descripcion":"Carga la tabla en RAM eliminando re-evaluaciones costosas. Rompe el query folding pero acelera cruces repetitivos.",                  "codigo":"let\n  Src = Table.Buffer(BaseDatos),\n  Res = Table.Join(Src,...)\nin Res" },
        { "id":"ex-17", "titulo":"Slicers y Timelines",      "categoria":"Tablas Dinámicas",         "descripcion":"Controles visuales que filtran Tablas Dinámicas. Se conectan via 'Conexiones de informe' a múltiples pivots.",                       "codigo":"Insertar → Segmentación de datos\n→ Clic derecho → Conexiones de informe" }
    ],
    "sql": [
        { "id":"sq-1",  "titulo":"SELECT DISTINCT / LIMIT",  "categoria":"Sintaxis Básica",           "descripcion":"DISTINCT colapsa duplicados. LIMIT restringe el número de filas para no destruir la memoria del motor.",      "codigo":"SELECT DISTINCT user_id FROM logs LIMIT 100;" },
        { "id":"sq-2",  "titulo":"Filtrado WHERE, AND, IN",  "categoria":"Sintaxis Básica",           "descripcion":"Poda primaria PRE-agrupación. IN reemplaza múltiples OR.",                                                  "codigo":"SELECT * FROM ventas\nWHERE pais IN ('CO','MX') AND monto > 0;" },
        { "id":"sq-3",  "titulo":"INNER JOIN",               "categoria":"Uniones Relacionales",      "descripcion":"Retorna únicamente los registros con coincidencia en AMBAS tablas.",                                        "codigo":"SELECT u.nombre, o.total\nFROM users u\nINNER JOIN orders o ON u.id = o.user_id;" },
        { "id":"sq-4",  "titulo":"LEFT JOIN",                "categoria":"Uniones Relacionales",      "descripcion":"Todos los de la izquierda + coincidencias de la derecha. NULL donde no hay match.",                         "codigo":"SELECT u.nombre, o.total\nFROM users u\nLEFT JOIN orders o ON u.id = o.user_id;" },
        { "id":"sq-5",  "titulo":"FULL OUTER JOIN",          "categoria":"Uniones Relacionales",      "descripcion":"Combina todo de ambos lados, rellenando NULL donde no hay coincidencia.",                                   "codigo":"SELECT * FROM A FULL OUTER JOIN B ON A.id = B.id;" },
        { "id":"sq-6",  "titulo":"ANTI-JOIN",                "categoria":"Uniones Relacionales",      "descripcion":"Encuentra registros en A sin correspondencia en B. Detecta clientes sin pedidos.",                          "codigo":"SELECT u.id FROM users u\nLEFT JOIN orders o ON u.id = o.user_id\nWHERE o.user_id IS NULL;" },
        { "id":"sq-7",  "titulo":"SUM / AVG / COUNT",        "categoria":"Agregación",                "descripcion":"Colapsan millones de filas en KPIs. Siempre requieren GROUP BY si hay dimensiones.",                        "codigo":"SELECT region, SUM(monto), AVG(monto)\nFROM ventas GROUP BY region;" },
        { "id":"sq-8",  "titulo":"GROUP BY vs HAVING",       "categoria":"Agregación",                "descripcion":"HAVING filtra POST-agregación; WHERE filtra PRE-agrupación.",                                              "codigo":"SELECT user_id, COUNT(*) as n\nFROM orders\nGROUP BY user_id HAVING COUNT(*) > 5;" },
        { "id":"sq-9",  "titulo":"ROW_NUMBER()",             "categoria":"Window Functions",          "descripcion":"Enumera filas dentro de cada partición sin comprimir el dataset.",                                          "codigo":"ROW_NUMBER() OVER(PARTITION BY user_id ORDER BY date DESC)" },
        { "id":"sq-10", "titulo":"RANK() / DENSE_RANK()",    "categoria":"Window Functions",          "descripcion":"DENSE_RANK no salta números en empates. Ideal para Top-N rankings.",                                       "codigo":"DENSE_RANK() OVER(ORDER BY ventas DESC) as ranking" },
        { "id":"sq-11", "titulo":"LAG() / LEAD()",           "categoria":"Window Functions",          "descripcion":"LAG trae valor de fila anterior; LEAD de la siguiente. Para calcular días entre compras.",                 "codigo":"LAG(order_date) OVER(PARTITION BY user_id ORDER BY order_date)" },
        { "id":"sq-12", "titulo":"NTILE()",                  "categoria":"Window Functions",          "descripcion":"Divide la población en N bloques iguales. Perfecto para cuartiles RFM.",                                   "codigo":"NTILE(4) OVER(ORDER BY monto DESC) as cuartil" },
        { "id":"sq-13", "titulo":"CTEs (WITH)",              "categoria":"Subconsultas",              "descripcion":"Nombra subqueries como tablas temporales. Elimina los horribles subqueries anidados.",                     "codigo":"WITH activos AS (\n  SELECT * FROM users WHERE activo=1\n)\nSELECT * FROM activos;" },
        { "id":"sq-14", "titulo":"EXPLAIN ANALYZE",          "categoria":"Optimización",              "descripcion":"Muestra el plan de ejecución real. Revela Index Scans vs Full Table Scans.",                              "codigo":"EXPLAIN ANALYZE SELECT * FROM orders WHERE user_id = 42;" }
    ],
    "python": [
        { "id":"py-1",  "titulo":"Listas y Diccionarios",    "categoria":"Estructuras Nativas",      "descripcion":"Vectores mutables y mapas clave-valor. Columna vertebral de cualquier pipeline.",                          "codigo":"clientes = [{'id':1,'nombre':'Ana'},{'id':2,'nombre':'Luis'}]" },
        { "id":"py-2",  "titulo":"Tuplas y Sets",            "categoria":"Estructuras Nativas",      "descripcion":"Tuplas inmutables para coordenadas/configs. Sets para deduplicar IDs.",                                   "codigo":"coords = (4.61, -74.08)\nunique_ids = set([1,2,2,3])  # {1,2,3}" },
        { "id":"py-3",  "titulo":"for / while / if-elif",   "categoria":"Control de Flujo",         "descripcion":"Iteradores y condicionales. Base de cualquier pipeline de transformación.",                                "codigo":"for c in clientes:\n    if c['gasto'] > 1000: print('VIP')" },
        { "id":"py-4",  "titulo":"Funciones Lambda",         "categoria":"Funciones y Decoradores",  "descripcion":"Funciones anónimas desechables. Ideales para .apply() sobre columnas Pandas.",                           "codigo":"df['seg'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Std')" },
        { "id":"py-5",  "titulo":"Decoradores (@)",          "categoria":"Funciones y Decoradores",  "descripcion":"Envuelven una función añadiendo lógica (ej. timing, logging) sin modificar su cuerpo.",                  "codigo":"@timer\ndef cargar(): df = pd.read_csv('ventas.csv'); return df" },
        { "id":"py-6",  "titulo":"df.head() / describe()",   "categoria":"Pandas Exploración",       "descripcion":"head() muestra primeras filas. describe() genera estadísticas: count, mean, std, min, max.",              "codigo":"df.head(10)\ndf.describe().T" },
        { "id":"py-7",  "titulo":"df.info()",                "categoria":"Pandas Exploración",       "descripcion":"Audita dtypes, conteo de no-nulos y uso de memoria por columna.",                                        "codigo":"df.info()" },
        { "id":"py-8",  "titulo":"df.dropna()",              "categoria":"Limpieza Wrangling",       "descripcion":"Elimina filas/columnas con NaN. axis=0: filas; axis=1: columnas; subset para columnas específicas.",     "codigo":"df.dropna(axis=0)\ndf.dropna(subset=['email'])" },
        { "id":"py-9",  "titulo":"df.fillna()",              "categoria":"Limpieza Wrangling",       "descripcion":"Imputa NaN con media, mediana o constante para preservar el volumen del dataset.",                       "codigo":"df['edad'].fillna(df['edad'].median(), inplace=True)" },
        { "id":"py-10", "titulo":"df.astype()",              "categoria":"Limpieza Wrangling",       "descripcion":"Convierte dtype de la columna. Crítico para parsear fechas, enteros o categóricas.",                     "codigo":"df['precio'] = df['precio'].astype('float64')" },
        { "id":"py-11", "titulo":"df.drop_duplicates()",     "categoria":"Limpieza Wrangling",       "descripcion":"Elimina filas duplicadas. keep='last' conserva la más reciente.",                                        "codigo":"df.drop_duplicates(subset=['email'], keep='last')" },
        { "id":"py-12", "titulo":"df.groupby()",             "categoria":"Transformación Tabular",   "descripcion":"Agrupa y agrega. Equivalente al GROUP BY de SQL en memoria RAM.",                                        "codigo":"rfm = df.groupby('id').agg(freq=('orden','count'),mon=('monto','sum'))" },
        { "id":"py-13", "titulo":"df.merge()",               "categoria":"Transformación Tabular",   "descripcion":"JOIN entre DataFrames. how='left','right','inner','outer'.",                                             "codigo":"pd.merge(clientes, ordenes, on='id', how='left')" },
        { "id":"py-14", "titulo":"Seaborn: heatmap/kdeplot", "categoria":"Visualización Científica", "descripcion":"heatmap muestra correlaciones. kdeplot muestra distribución de densidad.",                              "codigo":"sns.heatmap(df.corr(), annot=True, cmap='coolwarm')" }
    ],
    "github": [
        { "id":"gh-1", "titulo":"git config",          "categoria":"Configuración Core",   "descripcion":"Define identidad del desarrollador. Obligatorio para firmar commits.",                                        "codigo":"git config --global user.name \"Tu Nombre\"\ngit config --global user.email \"tu@email.com\"" },
        { "id":"gh-2", "titulo":"git init / clone",    "categoria":"Configuración Core",   "descripcion":"init crea repositorio local. clone descarga copia completa del remoto.",                                     "codigo":"git init\ngit clone https://github.com/usuario/repo.git" },
        { "id":"gh-3", "titulo":"git status",          "categoria":"Flujo Workdir",        "descripcion":"Estado del working directory. Rojo = sin staging; Verde = en staging.",                                     "codigo":"git status" },
        { "id":"gh-4", "titulo":"git add",             "categoria":"Flujo Workdir",        "descripcion":"Mueve cambios al Staging Area. El punto (.) agrega todo.",                                                  "codigo":"git add .\ngit add js/backend.js" },
        { "id":"gh-5", "titulo":"git commit",          "categoria":"Historia Inmutable",   "descripcion":"Congela staging en hash SHA permanente. El mensaje debe describir el cambio en presente.",                  "codigo":"git commit -m \"feat: Add Excel sandbox v4.5\"" },
        { "id":"gh-6", "titulo":"git push / pull",     "categoria":"Nube (Origin)",        "descripcion":"push sube commits locales al remoto. pull descarga y fusiona cambios remotos.",                             "codigo":"git push origin main\ngit pull origin main" },
        { "id":"gh-7", "titulo":"git branch / checkout","categoria":"Seguridad Múltiple",  "descripcion":"branch crea mundos aislados. checkout viaja entre ramas.",                                                  "codigo":"git checkout -b feature/sandbox\ngit checkout main" },
        { "id":"gh-8", "titulo":"Pull Requests",       "categoria":"Colaboración Corp",    "descripcion":"Solicitud formal de fusión. Tu rama pide integrarse a main tras revisión y CI/CD.",                        "codigo":"git push origin feature/sandbox\n# Luego en GitHub: Compare & Pull Request" },
        { "id":"gh-9", "titulo":"Forks",               "categoria":"Colaboración Corp",    "descripcion":"Duplicado del repositorio en tu cuenta. Base del Open Source.",                                            "codigo":"# GitHub: Botón Fork\ngit clone https://github.com/TU/repo.git" }
    ],
    "integration": [
        { "id":"in-1", "titulo":"=PY() Pandas Engine", "categoria":"Python nativo Excel",  "descripcion":"Microsoft 365 ejecuta Python en Azure desde celdas Excel. xl() accede al rango.",                          "codigo":"=PY(\nimport pandas as pd\ndf = xl(\"Tabla1[#Todos]\", headers=True)\nreturn df.describe()\n)" },
        { "id":"in-2", "titulo":"Engine SQLAlchemy",   "categoria":"ETL Cero-Fricciones",  "descripcion":"Conecta Pandas a cualquier RDBMS. read_sql() carga queries directamente.",                                "codigo":"from sqlalchemy import create_engine\nengine = create_engine('postgresql://user:pass@host/db')\ndf = pd.read_sql('SELECT * FROM churn', engine)" },
        { "id":"in-3", "titulo":"df.to_sql()",         "categoria":"ETL Cero-Fricciones",  "descripcion":"Escribe DataFrame a tabla SQL. if_exists='replace' sobreescribe; 'append' agrega.",                       "codigo":"df.to_sql('ml_predictions', engine, if_exists='replace', index=False)" }
    ]
};

// ─────────────────────────────────────────────────────────────────────────────
// GRIDS CONTEXTUALES POR EJERCICIO
// Cada ejercicio tiene sus propios datos de contexto y celda objetivo
// ─────────────────────────────────────────────────────────────────────────────
const GRID_CONTEXTS = {
    // Atajos → datos genéricos, modo hardware sin grid editable
    "ex-1": null,
    "ex-2": null,
    // Fórmulas matemáticas
    "ex-3": {
        cols: ["A","B"],
        headers: ["Número","Resultado"],
        rows: [["144",""],["256",""],["81",""],["36",""],["Formula →",""]],
        target: {row:5,col:2},
        hint: "Escribe =RAIZ(144) en la celda marcada en amarillo"
    },
    "ex-4": {
        cols: ["A","B"],
        headers: ["Vendedor","Ventas"],
        rows: [["Ana","380"],["Luis","520"],["María","290"],["Pedro","610"],["Carlos","455"],["TOTAL →",""]],
        target: {row:6,col:2},
        hint: "Escribe =SUMA(B2:B6) o =PROMEDIO(B2:B6)"
    },
    "ex-5": {
        cols: ["A","B"],
        headers: ["Cliente","Email"],
        rows: [["Ana","ana@mail.com"],["Luis",""],["María","maria@c.co"],["Pedro",""],["Carlos","pc@g.com"],["Conteo →",""]],
        target: {row:6,col:2},
        hint: "Escribe =CONTARA(B2:B6) para contar emails no vacíos"
    },
    "ex-6": {
        cols: ["A","B","C"],
        headers: ["Región","Ventas","Criterio"],
        rows: [["Norte","380","Norte"],["Sur","520",""],["Norte","290",""],["Sur","610",""],["Norte","455",""],["SUMAR.SI →","",""]],
        target: {row:6,col:2},
        hint: "Escribe =SUMAR.SI(A2:A6,\"Norte\",B2:B6)"
    },
    "ex-7": {
        cols: ["A","B"],
        headers: ["Monto","Segmento"],
        rows: [["380",""],["620",""],["150",""],["890",""],["480",""],["Formula →",""]],
        target: {row:6,col:2},
        hint: "=SI(A2>500,\"Alto\",\"Bajo\") — escríbela para la primera fila"
    },
    "ex-8": {
        cols: ["A","B","C"],
        headers: ["Monto","Estado","Resultado"],
        rows: [["380","OK",""],["620","OK",""],["150","Pendiente",""],["890","OK",""],["480","Pendiente",""],["Formula →","",""]],
        target: {row:6,col:3},
        hint: "=SI(Y(A2>500,B2=\"OK\"),\"Aprobado\",\"Rechazado\")"
    },
    "ex-9": {
        cols: ["A","B"],
        headers: ["Nombre Bruto","Nombre Limpio"],
        rows: [["  Ana García ",""],["Luis  López",""],["  María  ",""],["Pedro Ruiz",""],["  Carlos B ",""],["Formula →",""]],
        target: {row:6,col:2},
        hint: "=ESPACIOS(A2) — elimina espacios extra"
    },
    "ex-10": {
        cols: ["A","B","C","D","E"],
        headers: ["ID Buscar","","","ID","Nombre"],
        rows: [["C003","→ Nombre:","","C001","Ana"],["","","","C002","Luis"],["","","","C003","María"],["","","","C004","Pedro"],["","","","C005","Carlos"]],
        target: {row:1,col:2},
        hint: "=BUSCARV(A2,$D$2:$E$6,2,FALSO) — busca el nombre del ID en A2"
    },
    "ex-11": {
        cols: ["A","B","C","D"],
        headers: ["Producto","Precio","","Buscar"],
        rows: [["Laptop","1200","","Teclado"],["Mouse","25","→ Precio:",""],["Teclado","45","",""],["Monitor","320","",""],["Webcam","80","",""]],
        target: {row:2,col:3},
        hint: "=INDICE(B2:B6,COINCIDIR(D2,A2:A6,0))"
    },
    "ex-12": {
        cols: ["A","B","C","D"],
        headers: ["Producto","Precio","","Buscar"],
        rows: [["Laptop","1200","","Mouse"],["Mouse","25","→ Precio:",""],["Teclado","45","",""],["Monitor","320","",""],["Webcam","80","",""]],
        target: {row:2,col:3},
        hint: "=BUSCARX(D2,A2:A6,B2:B6,\"No encontrado\",0)"
    },
    "ex-13": {
        cols: ["A","B","C"],
        headers: ["Cliente","Región","Ventas"],
        rows: [["Ana","Norte","380"],["Luis","Sur","620"],["María","Norte","290"],["Pedro","Norte","890"],["Carlos","Sur","455"],["→ FILTRAR Norte","",""]],
        target: {row:6,col:1},
        hint: "=FILTRAR(A2:C6,B2:B6=\"Norte\") — filtra clientes del Norte"
    },
    "ex-14": {
        cols: ["A","B"],
        headers: ["Ciudades","Únicas Ordenadas"],
        rows: [["Bogotá",""],["Medellín",""],["Bogotá",""],["Cali",""],["Medellín","→"],["Formula →",""]],
        target: {row:6,col:2},
        hint: "=ORDENAR(UNICOS(A2:A6))"
    },
    "ex-15": {
        cols: ["A"],
        headers: ["Código M"],
        rows: [["Table.SelectRows(Source, each [Region] = \"Norte\")"],["// Filtra en origen, no en Excel"],[""],[""],[" "]],
        target: null,
        hint: "Escribe la función M en la terminal (no en la celda)"
    },
    "ex-16": {
        cols: ["A"],
        headers: ["Código M"],
        rows: [["let"],["  Src = Table.Buffer(BaseDatos),"],["  Res = Table.Join(Src, ...)"],[" in Res"],[" "]],
        target: null,
        hint: "Escribe Table.Buffer() en la terminal"
    },
    "ex-17": {
        cols: ["A","B","C"],
        headers: ["Trimestre","Región","Ventas"],
        rows: [["Q1","Norte","380"],["Q1","Sur","620"],["Q2","Norte","290"],["Q2","Sur","890"],["Q3","Norte","455"]],
        target: null,
        hint: "Insertar → Segmentación → conectar a tabla dinámica"
    }
};

// ─────────────────────────────────────────────────────────────────────────────
// MAPA DE EJERCICIOS — Validadores y tipos
// ─────────────────────────────────────────────────────────────────────────────
const EXERCISE_MAP = {
    "ex-1":  { type:"hardware", key:"ctrl+space",   instruccion:"Presiona Ctrl + Espacio para seleccionar la columna entera.", exito:"¡Ctrl + Espacio detectado! Columna seleccionada." },
    "ex-2":  { type:"hardware", key:"shift+space",  instruccion:"Presiona Shift + Espacio para seleccionar la fila entera.", exito:"¡Shift + Espacio detectado! Fila seleccionada." },
    "ex-3":  { type:"formula",  regex:/^=RAIZ\(144\)$/i,                    instruccion:"Haz clic en la celda amarilla y escribe =RAIZ(144)", exito:"¡Resultado: 12! Raíz cuadrada de 144.", compute: () => 12 },
    "ex-4":  { type:"formula",  regex:/^=SUMA\(B2:B6\)|^=PROMEDIO\(B2:B6\)/i, instruccion:"Escribe =SUMA(B2:B6) o =PROMEDIO(B2:B6) en la celda marcada.", exito:"¡Correcto!", compute: (formula) => formula.match(/PROMEDIO/i) ? "( 431 )" : "2255" },
    "ex-5":  { type:"formula",  regex:/^=CONTARA\(B2:B6\)|^=CONTAR\(B2:B6\)/i, instruccion:"Escribe =CONTARA(B2:B6) para contar emails no vacíos.", exito:"¡3 emails encontrados!", compute: (f) => f.match(/CONTARA/i) ? "3" : "0" },
    "ex-6":  { type:"formula",  regex:/^=SUMAR\.SI\(A2:A6,"Norte",B2:B6\)|^=SUMAR\.SI\(A2:A6,C2,B2:B6\)/i, instruccion:"Escribe =SUMAR.SI(A2:A6,\"Norte\",B2:B6)", exito:"¡Total Norte: 1125!", compute: () => "1.125" },
    "ex-7":  { type:"formula",  regex:/^=SI\(.+,.+,.+\)$/i,                instruccion:"Escribe =SI(A2>500,\"Alto\",\"Bajo\") en la celda.", exito:"¡Función condicional correcta!", compute: () => "Bajo" },
    "ex-8":  { type:"formula",  regex:/^=SI\(\s*Y\(.+\)/i,                 instruccion:"Escribe =SI(Y(A2>500,B2=\"OK\"),\"Aprobado\",\"Rechazado\")", exito:"¡Lógica booleana validada!", compute: () => "Rechazado" },
    "ex-9":  { type:"formula",  regex:/^=ESPACIOS\(A\d+\)$/i,              instruccion:"Escribe =ESPACIOS(A2) para limpiar la celda.", exito:"¡Cadena limpiada!", compute: () => "Ana García" },
    "ex-10": { type:"formula",  regex:/^=BUSCARV\(.+\)$/i,                 instruccion:"Escribe =BUSCARV(A2,$D$2:$E$6,2,FALSO)", exito:"¡BUSCARV correcto! Resultado: María", compute: () => "María" },
    "ex-11": { type:"formula",  regex:/^=INDICE\(.+,COINCIDIR\(.+\)\)$/i,  instruccion:"Escribe =INDICE(B2:B6,COINCIDIR(D2,A2:A6,0))", exito:"¡INDICE+COINCIDIR correcto! Resultado: 45", compute: () => "45" },
    "ex-12": { type:"formula",  regex:/^=BUSCARX\(.+\)$/i,                 instruccion:"Escribe =BUSCARX(D2,A2:A6,B2:B6,\"No encontrado\",0)", exito:"¡BUSCARX correcto! Resultado: 25", compute: () => "25" },
    "ex-13": { type:"formula",  regex:/^=FILTRAR\(.+\)$/i,                 instruccion:"Escribe =FILTRAR(A2:C6,B2:B6=\"Norte\")", exito:"¡3 registros Norte filtrados!", compute: () => "Ana\nMaría\nPedro" },
    "ex-14": { type:"formula",  regex:/^=ORDENAR\(UNICOS\(.+\)\)|^=UNICOS\(.+\)/i, instruccion:"Escribe =ORDENAR(UNICOS(A2:A6))", exito:"¡Únicos ordenados: Bogotá, Cali, Medellín!", compute: () => "Bogotá, Cali, Medellín" },
    "ex-15": { type:"terminal", regex:/Table\.SelectRows/i,                 instruccion:"Escribe: Table.SelectRows(Source, each [Region] = \"Norte\")",  exito:"¡Query Folding correcto! Filtrado delegado." },
    "ex-16": { type:"terminal", regex:/Table\.Buffer\(/i,                   instruccion:"Escribe: Source = Table.Buffer(BaseDatos)",                     exito:"¡Table.Buffer() correcto! Tabla en RAM." },
    "ex-17": { type:"terminal", regex:/slicer|segmentacion|segmentación|insertar|pivot/i, instruccion:"Describe el flujo para insertar Slicers.",        exito:"¡Concepto Slicer validado!" },
    "sq-1":  { type:"terminal", regex:/SELECT\s+DISTINCT/i,             instruccion:"Escribe SELECT DISTINCT user_id FROM logs LIMIT 100;", exito:"¡SELECT DISTINCT correcto!" },
    "sq-2":  { type:"terminal", regex:/WHERE.+(IN\s*\(|AND|OR)/i,       instruccion:"Filtra con WHERE pais IN ('CO','MX').", exito:"¡WHERE + IN correcto!" },
    "sq-3":  { type:"terminal", regex:/INNER\s+JOIN/i,                  instruccion:"Escribe un INNER JOIN entre users y orders.", exito:"¡INNER JOIN correcto!" },
    "sq-4":  { type:"terminal", regex:/LEFT\s+JOIN/i,                   instruccion:"Escribe un LEFT JOIN de users a orders.", exito:"¡LEFT JOIN correcto!" },
    "sq-5":  { type:"terminal", regex:/FULL\s+(OUTER\s+)?JOIN/i,        instruccion:"Escribe FULL OUTER JOIN entre tabla A y B.", exito:"¡FULL OUTER JOIN correcto!" },
    "sq-6":  { type:"terminal", regex:/LEFT\s+JOIN[\s\S]+IS\s+NULL/i,   instruccion:"ANTI-JOIN: LEFT JOIN + WHERE columna IS NULL.", exito:"¡ANTI-JOIN correcto!" },
    "sq-7":  { type:"terminal", regex:/SUM\(|AVG\(|COUNT\(/i,           instruccion:"SELECT region, SUM(monto) FROM ventas GROUP BY region;", exito:"¡Agregación correcta!" },
    "sq-8":  { type:"terminal", regex:/HAVING/i,                        instruccion:"GROUP BY ... HAVING COUNT(*) > 5", exito:"¡HAVING correcto!" },
    "sq-9":  { type:"terminal", regex:/ROW_NUMBER\(\)\s*OVER/i,         instruccion:"ROW_NUMBER() OVER(PARTITION BY user_id ORDER BY date DESC)", exito:"¡ROW_NUMBER correcto!" },
    "sq-10": { type:"terminal", regex:/DENSE_RANK\(\)|RANK\(\)/i,       instruccion:"DENSE_RANK() OVER(ORDER BY ventas DESC)", exito:"¡DENSE_RANK correcto!" },
    "sq-11": { type:"terminal", regex:/LAG\(|LEAD\(/i,                  instruccion:"LAG(order_date) OVER(PARTITION BY user_id ORDER BY date)", exito:"¡LAG correcto!" },
    "sq-12": { type:"terminal", regex:/NTILE\(\d+\)/i,                  instruccion:"NTILE(4) OVER(ORDER BY monto DESC)", exito:"¡NTILE correcto!" },
    "sq-13": { type:"terminal", regex:/WITH\s+\w+\s+AS\s*\(/i,          instruccion:"WITH activos AS (SELECT * FROM users WHERE activo=1) SELECT * FROM activos;", exito:"¡CTE correcto!" },
    "sq-14": { type:"terminal", regex:/EXPLAIN(\s+ANALYZE)?/i,          instruccion:"EXPLAIN ANALYZE SELECT * FROM orders WHERE user_id = 42;", exito:"¡EXPLAIN correcto!" },
    "py-1":  { type:"terminal", regex:/[\[{]/,                          instruccion:"Crea: clientes = [{'id':1,'nombre':'Ana'}]", exito:"¡Diccionario correcto!" },
    "py-2":  { type:"terminal", regex:/set\(|tuple\(/i,                 instruccion:"ids = set([1,2,2,3])", exito:"¡Set correcto!" },
    "py-3":  { type:"terminal", regex:/for\s+\w+\s+in|while\s+/,        instruccion:"Escribe un bucle for sobre una lista.", exito:"¡Bucle correcto!" },
    "py-4":  { type:"terminal", regex:/lambda/,                         instruccion:"df['seg'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Std')", exito:"¡Lambda correcto!" },
    "py-5":  { type:"terminal", regex:/@\w+/,                           instruccion:"@timer sobre una función de carga.", exito:"¡Decorador correcto!" },
    "py-6":  { type:"terminal", regex:/df\.(head|describe)\(/,          instruccion:"df.describe().T", exito:"¡df.describe correcto!" },
    "py-7":  { type:"terminal", regex:/df\.info\(\)/,                   instruccion:"df.info()", exito:"¡df.info correcto!" },
    "py-8":  { type:"terminal", regex:/df\.dropna\(/,                   instruccion:"df.dropna(axis=0) o df.dropna(subset=['email'])", exito:"¡dropna correcto!" },
    "py-9":  { type:"terminal", regex:/df\[.+\]\.fillna\(/,             instruccion:"df['edad'].fillna(df['edad'].median(), inplace=True)", exito:"¡fillna correcto!" },
    "py-10": { type:"terminal", regex:/\.astype\(/,                     instruccion:"df['precio'].astype('float64')", exito:"¡astype correcto!" },
    "py-11": { type:"terminal", regex:/drop_duplicates\(/,              instruccion:"df.drop_duplicates(subset=['email'], keep='last')", exito:"¡drop_duplicates correcto!" },
    "py-12": { type:"terminal", regex:/groupby\(/,                      instruccion:"df.groupby('id').agg({'orden':'count','monto':'sum'})", exito:"¡groupby correcto!" },
    "py-13": { type:"terminal", regex:/pd\.merge\(/,                    instruccion:"pd.merge(clientes, ordenes, on='id', how='left')", exito:"¡pd.merge correcto!" },
    "py-14": { type:"terminal", regex:/heatmap\(|kdeplot\(/,            instruccion:"sns.heatmap(df.corr(), annot=True)", exito:"¡Seaborn correcto!" },
    "gh-1":  { type:"terminal", regex:/git\s+config\s+--global/,        instruccion:"git config --global user.name \"TuNombre\"", exito:"¡git config correcto!" },
    "gh-2":  { type:"terminal", regex:/git\s+(init|clone)/,             instruccion:"git clone https://github.com/usuario/repo.git", exito:"¡git clone correcto!" },
    "gh-3":  { type:"terminal", regex:/git\s+status/,                   instruccion:"git status", exito:"¡git status correcto!" },
    "gh-4":  { type:"terminal", regex:/git\s+add/,                      instruccion:"git add .", exito:"¡git add correcto!" },
    "gh-5":  { type:"terminal", regex:/git\s+commit\s+-m/,              instruccion:"git commit -m \"feat: descripción\"", exito:"¡git commit correcto!" },
    "gh-6":  { type:"terminal", regex:/git\s+(push|pull)/,              instruccion:"git push origin main", exito:"¡git push correcto!" },
    "gh-7":  { type:"terminal", regex:/git\s+(branch|checkout)/,        instruccion:"git checkout -b feature/mi-rama", exito:"¡git branch correcto!" },
    "gh-8":  { type:"terminal", regex:/git\s+push.+origin|pull.request/i, instruccion:"git push origin feature/mi-rama", exito:"¡PR flow correcto!" },
    "gh-9":  { type:"terminal", regex:/fork|clone.+github/i,            instruccion:"Describe el flujo de Fork + clone.", exito:"¡Fork correcto!" },
    "in-1":  { type:"formula",  regex:/^=PY\(/i,                        instruccion:"Escribe =PY( para iniciar celda Python en Excel.", exito:"¡=PY() correcto!", compute: () => "DataFrame(5×8)" },
    "in-2":  { type:"terminal", regex:/create_engine\(/,                instruccion:"create_engine('postgresql://user:pass@host/db')", exito:"¡SQLAlchemy correcto!" },
    "in-3":  { type:"terminal", regex:/\.to_sql\(/,                     instruccion:"df.to_sql('tabla', engine, if_exists='replace')", exito:"¡to_sql correcto!" }
};

// ─────────────────────────────────────────────────────────────────────────────
// ESTADO GLOBAL
// ─────────────────────────────────────────────────────────────────────────────
let currentExercise = null;
let currentItemId   = null;
let selectedCell    = null; // referencia DOM a la celda activa

// ─────────────────────────────────────────────────────────────────────────────
// INICIALIZACIÓN + RENDER DE GRIDS
// ─────────────────────────────────────────────────────────────────────────────
function fetchDatabase() { renderAllGrids(); }

function renderAllGrids() {
    for (const mod of ['excel','sql','python','github','integration']) {
        if (functionDatabase[mod]) renderGrid(`grid-${mod}`, functionDatabase[mod], mod);
    }
}

function renderGrid(gridId, dataArray, theme) {
    const grid = document.getElementById(gridId);
    if (!grid) return;
    grid.innerHTML = dataArray.map(item => `
        <div class="function-card ${theme}-hover" onclick="openFunctionModal('${item.id}','${theme}')">
            <span class="card-category ${theme}-color">${item.categoria}</span>
            <h3 class="card-title">${item.titulo}</h3>
            <p class="card-excerpt">${item.descripcion}</p>
        </div>`).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// MODAL: ABRIR
// ─────────────────────────────────────────────────────────────────────────────
function openFunctionModal(itemId, themeName) {
    let data = null;
    for (const mod in functionDatabase) {
        data = functionDatabase[mod].find(i => i.id === itemId);
        if (data) break;
    }
    if (!data) return;

    currentItemId   = itemId;
    currentExercise = EXERCISE_MAP[itemId] || { type:'terminal', regex:/.+/, instruccion:`Demuestra ${data.titulo}.`, exito:`¡${data.titulo} correcto!` };

    // Rellenar Teoría
    document.getElementById('modal-title').innerText    = data.titulo;
    document.getElementById('modal-category').innerText = data.categoria;
    document.getElementById('modal-desc').innerText     = data.descripcion;
    document.getElementById('modal-code').innerText     = data.codigo;

    // Misión
    document.getElementById('modal-exercise-desc').innerText = currentExercise.instruccion;

    // Elegir sandbox
    const useExcel    = currentExercise.type === 'formula' || currentExercise.type === 'hardware';
    const termBox     = document.getElementById('terminal-sandbox');
    const excelBox    = document.getElementById('excel-sandbox');

    if (useExcel) {
        termBox.classList.add('hidden');
        excelBox.classList.remove('hidden');
        buildExcelGrid(itemId);
        resetExcelSandbox(currentExercise.type === 'hardware');
    } else {
        excelBox.classList.add('hidden');
        termBox.classList.remove('hidden');
        resetTerminalSandbox(data);
    }

    // Color categoría
    const cat = document.getElementById('modal-category');
    cat.className = `${themeName}-color`;

    switchModalTab('teoria');
    document.getElementById('dynamic-modal').classList.add('active');
}

// ─────────────────────────────────────────────────────────────────────────────
// CONSTRUIR GRID CONTEXTUAL
// ─────────────────────────────────────────────────────────────────────────────
function buildExcelGrid(itemId) {
    const ctx = GRID_CONTEXTS[itemId];
    const table = document.getElementById('excel-grid-table');
    if (!ctx || !ctx.rows) {
        // Grid genérico si no hay contexto
        table.innerHTML = `<thead><tr><th></th><th>A</th><th>B</th><th>C</th></tr></thead>
        <tbody>
            <tr><td class="row-num">1</td><td class="cell" onclick="selectCell(this,'A1')">Dato 1</td><td class="cell" onclick="selectCell(this,'B1')">100</td><td class="cell target-cell" id="result-cell" onclick="selectCell(this,'C1',true)"></td></tr>
            <tr><td class="row-num">2</td><td class="cell" onclick="selectCell(this,'A2')">Dato 2</td><td class="cell" onclick="selectCell(this,'B2')">200</td><td class="cell" onclick="selectCell(this,'C2')"></td></tr>
            <tr><td class="row-num">3</td><td class="cell" onclick="selectCell(this,'A3')">Dato 3</td><td class="cell" onclick="selectCell(this,'B3')">300</td><td class="cell" onclick="selectCell(this,'C3')"></td></tr>
        </tbody>`;
        return;
    }

    const cols = ctx.cols;
    const headers = ctx.headers;
    const rows    = ctx.rows;
    const target  = ctx.target;

    let thead = `<thead><tr><th></th>${cols.map((c,i) => `<th>${c}<br><small style="color:#888">${headers[i]||''}</small></th>`).join('')}</tr></thead>`;
    let tbody = '<tbody>';
    rows.forEach((row, ri) => {
        tbody += `<tr><td class="row-num">${ri+1}</td>`;
        cols.forEach((_, ci) => {
            const isTarget = target && ri + 1 === target.row && ci + 1 === target.col;
            const cellId   = `${cols[ci]}${ri+1}`;
            const cls      = isTarget ? 'cell target-cell' : 'cell';
            const id       = isTarget ? 'id="result-cell"' : '';
            const val      = row[ci] !== undefined ? row[ci] : '';
            tbody += `<td class="${cls}" ${id} onclick="selectCell(this,'${cellId}',${isTarget})">${val}</td>`;
        });
        tbody += '</tr>';
    });
    tbody += '</tbody>';

    table.innerHTML = thead + tbody;

    // Auto-seleccionar la celda objetivo
    const tCell = document.getElementById('result-cell');
    if (tCell) {
        selectCell(tCell, `${cols[(target.col-1)]}${target.row}`, true);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SELECCIONAR CELDA → actualiza barra de nombre + formula bar preview
// ─────────────────────────────────────────────────────────────────────────────
function selectCell(element, cellId, isTarget) {
    document.querySelectorAll('.cell').forEach(c => c.classList.remove('selected'));
    element.classList.add('selected');
    selectedCell = element;

    const nameBar = document.getElementById('excel-cell-name');
    if (nameBar) nameBar.value = cellId;

    const fBar = document.getElementById('excel-formula-bar');
    if (fBar && !fBar.disabled) fBar.focus();
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMULA LIVE PREVIEW → mientras el usuario tipea, muestra en la celda
// ─────────────────────────────────────────────────────────────────────────────
function onFormulaInput(value) {
    const cell = selectedCell || document.getElementById('result-cell');
    if (!cell) return;
    if (value.startsWith('=')) {
        cell.innerText = value;           // muestra la fórmula en edición
        cell.style.color = '#1e40af';     // azul fórmula
        cell.style.fontStyle = 'normal';
    } else {
        cell.innerText = value;
        cell.style.color = '#1e293b';
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// RESET SANDBOXES
// ─────────────────────────────────────────────────────────────────────────────
function resetExcelSandbox(isHardware) {
    const bar     = document.getElementById('excel-formula-bar');
    const overlay = document.getElementById('key-listener-overlay');
    const fb      = document.getElementById('excel-feedback');
    const hwDisp  = document.getElementById('hw-keys-display');

    bar.value    = '';
    bar.disabled = isHardware;
    fb.innerHTML = isHardware
        ? '> 🎹 Modo Hardware — presiona las teclas en tu teclado físico.'
        : `> ${currentExercise ? currentExercise.instruccion : 'Escribe tu fórmula...'}`;
    fb.style.color = '#a5b4fc';

    overlay.classList.toggle('hidden', !isHardware);

    if (isHardware && hwDisp) {
        // Mostrar las teclas como badges
        const keyMap = { 'ctrl+space':['Ctrl','Espacio'], 'shift+space':['Shift','Espacio'] };
        const keys = keyMap[currentExercise.key] || [];
        hwDisp.innerHTML = keys.map(k => `<span class="hw-key-badge">${k}</span>`).join('<span style="color:#217346;font-size:1.4rem;margin:0 4px;">+</span>');
    }
}

function resetTerminalSandbox(data) {
    const inp = document.getElementById('modal-code-input');
    const fb  = document.getElementById('modal-feedback');
    inp.value       = '';
    inp.placeholder = currentExercise ? currentExercise.instruccion : `Escribe el comando...`;
    fb.innerHTML    = '> Esperando input...';
    fb.style.color  = '#a5b4fc';
}

// ─────────────────────────────────────────────────────────────────────────────
// VALIDADOR: FÓRMULA EXCEL
// ─────────────────────────────────────────────────────────────────────────────
function verifyExcelLab() {
    if (!currentExercise || currentExercise.type !== 'formula') return;
    const val  = document.getElementById('excel-formula-bar').value.trim();
    const fb   = document.getElementById('excel-feedback');
    const cell = selectedCell || document.getElementById('result-cell');

    if (!val) { fb.innerHTML = '> Celda vacía.'; fb.style.color = '#ef4444'; return; }

    let ok;
    try { ok = currentExercise.regex.test(val); } catch(e) { ok = true; }

    if (ok) {
        // Calcular resultado y mostrarlo en la celda como Excel real
        const result = currentExercise.compute ? currentExercise.compute(val) : '✓';
        if (cell) {
            cell.innerText   = result;
            cell.style.color = '#1e293b';
            cell.style.fontStyle = 'normal';
            cell.classList.add('result-cell');
            cell.classList.remove('target-cell');
        }
        fb.innerHTML   = `> ✅ [CELL SUCCESS] ${currentExercise.exito}  →  Valor calculado: ${result}`;
        fb.style.color = '#217346';
        // Update formula bar to show computed (like Excel pressing Enter)
        document.getElementById('excel-formula-bar').value = val;
    } else {
        if (cell) {
            cell.innerText = '#¡VALOR!';
            cell.classList.add('error-cell');
        }
        fb.innerHTML   = `> ❌ #ERROR  →  ${currentExercise.instruccion}`;
        fb.style.color = '#ef4444';
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// VALIDADOR: TERMINAL (SQL, Python, GitHub)
// ─────────────────────────────────────────────────────────────────────────────
function verifyMicroLab() {
    if (!currentExercise) return;
    const val = document.getElementById('modal-code-input').value.trim();
    const fb  = document.getElementById('modal-feedback');
    if (!val) { fb.innerHTML = '> Terminal vacío.'; fb.style.color = '#ef4444'; return; }

    let ok;
    try { ok = currentExercise.regex.test(val); } catch(e) { ok = true; }
    if (ok) {
        fb.innerHTML   = `> ✅ [SUCCESS] ${currentExercise.exito}`;
        fb.style.color = '#10b981';
    } else {
        fb.innerHTML   = `> ❌ Sintaxis incorrecta. Pista: ${currentExercise.instruccion}`;
        fb.style.color = '#ef4444';
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// VALIDADOR: HARDWARE (atajos de teclado)
// ─────────────────────────────────────────────────────────────────────────────
document.addEventListener('keydown', function(e) {
    const modal = document.getElementById('dynamic-modal');
    if (!modal || !modal.classList.contains('active')) return;
    if (!currentExercise || currentExercise.type !== 'hardware') return;
    const tag = document.activeElement.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    e.preventDefault();

    const fb = document.getElementById('excel-feedback');
    let ok = false;
    if (currentExercise.key === 'ctrl+space'  && e.ctrlKey  && e.code === 'Space') ok = true;
    if (currentExercise.key === 'shift+space' && e.shiftKey && e.code === 'Space') ok = true;

    if (ok) {
        fb.innerHTML = `> ✅ [HARDWARE SUCCESS] ${currentExercise.exito}`;
        fb.style.color = '#217346';
        document.getElementById('key-listener-overlay').classList.add('hidden');
        // Resaltar toda la columna/fila visualmente
        const cells = document.querySelectorAll('.cell');
        cells.forEach(c => { c.classList.add('result-cell'); });
        setTimeout(() => cells.forEach(c => c.classList.remove('result-cell')), 1200);
    } else {
        fb.innerHTML = `> 🎹 Tecla '${e.key}' detectada. Sigue intentando...`;
        fb.style.color = '#fbbf24';
    }
});

// ─────────────────────────────────────────────────────────────────────────────
// TABS DEL MODAL
// ─────────────────────────────────────────────────────────────────────────────
function switchModalTab(tabId) {
    document.querySelectorAll('.modal-tab').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-'+tabId).classList.add('active');
    document.getElementById('modal-body-teoria').classList.add('hidden');
    document.getElementById('modal-body-practica').classList.add('hidden');
    document.getElementById('modal-body-'+tabId).classList.remove('hidden');
}

// ─────────────────────────────────────────────────────────────────────────────
// CERRAR MODAL
// ─────────────────────────────────────────────────────────────────────────────
function closeModal(e) {
    if (e && e.target !== document.getElementById('dynamic-modal') &&
             !e.target.classList.contains('btn-close')) return;
    document.getElementById('dynamic-modal').classList.remove('active');
    currentExercise = null;
    selectedCell    = null;
}
