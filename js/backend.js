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
        // ATAJOS
        { "id":"ex-1","titulo":"Ctrl Shortcuts","categoria":"Atajos Fundamentales","descripcion":"Ctrl+Espacio=columna entera. Ctrl+1=Formato de Celdas. Ctrl+Shift+$=moneda. Ctrl+Fin=última celda con dato.","codigo":"Ctrl+Espacio → Columna entera\nCtrl+1       → Formato Celdas\nCtrl+Shift+$ → Formato moneda\nCtrl+Fin     → Última celda" },
        { "id":"ex-2","titulo":"Shift / Alt Shortcuts","categoria":"Atajos Fundamentales","descripcion":"Shift+Espacio=fila entera. Alt+=inserta SUMA automática. F2=editar celda. Alt+H+A+C=centrar.","codigo":"Shift+Espacio → Fila entera\nAlt+=         → Auto-SUMA\nF2            → Editar celda\nAlt+H+A+C    → Centrar" },
        // MATEMÁTICAS
        { "id":"ex-3","titulo":"SUMA & PROMEDIO","categoria":"Matemáticas Estadísticas","descripcion":"SUMA agrega el rango completo. PROMEDIO calcula la media aritmética ignorando celdas vacías. Base de cualquier reporte financiero.","codigo":"=SUMA(B2:B6)\n=PROMEDIO(B2:B6)" },
        { "id":"ex-4","titulo":"SUMAR.SI.CONJUNTO","categoria":"Matemáticas Estadísticas","descripcion":"Suma registros que cumplen MÚLTIPLES criterios simultáneos. Más poderosa que SUMAR.SI para reportes complejos por región+producto+estado.","codigo":"=SUMAR.SI.CONJUNTO(C2:C6, A2:A6,\"Norte\", B2:B6,\"Laptop\")" },
        { "id":"ex-5","titulo":"CONTAR.SI.CONJUNTO","categoria":"Matemáticas Estadísticas","descripcion":"Cuenta registros con múltiples filtros activos. Ejemplo real: contar clientes VIP activos de Colombia con compras mayores a $500.","codigo":"=CONTAR.SI.CONJUNTO(A2:A6,\"CO\",B2:B6,\"VIP\",C2:C6,\">500\")" },
        { "id":"ex-6","titulo":"MAX.SI.CONJUNTO & MIN.SI","categoria":"Matemáticas Estadísticas","descripcion":"Máximo o mínimo dentro de un subconjunto filtrado. Clave para rankings de ventas por categoría sin alterar el dataset.","codigo":"=MAX.SI.CONJUNTO(C2:C6, A2:A6,\"Norte\")\n=MIN.SI.CONJUNTO(C2:C6, B2:B6,\"Laptop\")" },
        { "id":"ex-7","titulo":"REDONDEAR & ENTERO","categoria":"Matemáticas Estadísticas","descripcion":"REDONDEAR controla decimales para presentaciones. ENTERO trunca al entero inferior. Esencial en facturas, precios y KPIs.","codigo":"=REDONDEAR(A2*1.19, 2)   → IVA con 2 decimales\n=ENTERO(A2/1000)*1000    → Redondeo a miles" },
        // LÓGICAS
        { "id":"ex-8","titulo":"SI() anidado","categoria":"Lógicas y Condicionales","descripcion":"SI() anidado crea árboles de decisión con 3 o más ramas. Fundamental para clasificar clientes (VIP/Medio/Básico) o estados de pago.","codigo":"=SI(B2>1000,\"VIP\",SI(B2>500,\"Medio\",\"Básico\"))" },
        { "id":"ex-9","titulo":"SI.CONJUNTO (IFS)","categoria":"Lógicas y Condicionales","descripcion":"Reemplazo moderno y legible de SI() anidado. Evalúa condiciones en orden, devuelve el primer resultado verdadero. Disponible Office 2019+.","codigo":"=SI.CONJUNTO(B2>1000,\"VIP\",B2>500,\"Medio\",B2>0,\"Básico\",VERDADERO,\"Sin ventas\")" },
        { "id":"ex-10","titulo":"SIERROR & SINO.DISPONIBLE","categoria":"Lógicas y Condicionales","descripcion":"SIERROR captura cualquier error. SINO.DISPONIBLE captura solo #N/A dejando pasar otros errores para detectarlos. Protegen los modelos en producción.","codigo":"=SIERROR(BUSCARV(A2,Tabla,2,0),\"No encontrado\")\n=SINO.DISPONIBLE(BUSCARX(A2,E:E,F:F),0)" },
        // TEXTO
        { "id":"ex-11","titulo":"IZQUIERDA / DERECHA / EXTRAE","categoria":"Funciones de Texto","descripcion":"Extraen substrings por posición. Clave para parsear códigos como 'CO-2024-001', separar nombres/apellidos o extraer prefijos de IDs.","codigo":"=IZQUIERDA(A2,2)    → Prefijo 2 chars\n=DERECHA(A2,3)      → Sufijo 3 chars\n=EXTRAE(A2,4,4)     → 4 chars desde pos 4" },
        { "id":"ex-12","titulo":"CONCATENAR / UNIRCADENAS","categoria":"Funciones de Texto","descripcion":"CONCATENAR une textos y celdas. UNIRCADENAS (TEXTJOIN) añade delimitador y puede ignorar vacíos, ideal para listas separadas por coma.","codigo":"=CONCATENAR(A2,\" \",B2)              → \"Ana García\"\n=UNIRCADENAS(\", \",VERDADERO,A2:A6)  → \"Ana, Luis, María\"" },
        { "id":"ex-13","titulo":"TEXTO() — Formatear valores","categoria":"Funciones de Texto","descripcion":"Convierte números o fechas en texto con formato personalizado. Esencial en reportes donde el número debe aparecer como moneda, % o fecha legible.","codigo":"=TEXTO(A2,\"$#,##0.00\")    → \"$1,250.00\"\n=TEXTO(B2,\"dd/mm/aaaa\") → \"06/04/2026\"\n=TEXTO(C2,\"0.0%\")        → \"85.5%\"" },
        { "id":"ex-14","titulo":"LARGO / HALLAR / SUSTITUIR","categoria":"Funciones de Texto","descripcion":"LARGO cuenta caracteres. HALLAR localiza posición de un substring. SUSTITUIR reemplaza ocurrencias. Cadena de limpieza para datos de sistemas ERP.","codigo":"=LARGO(A2)             → 15 caracteres\n=HALLAR(\"@\",A2)         → Posición del @\n=SUSTITUIR(A2,\"-\",\"\")   → Elimina guiones" },
        // FECHAS
        { "id":"ex-15","titulo":"FECHA / MES / AÑO / DIA","categoria":"Funciones de Fecha","descripcion":"Extraen componentes de fechas para agrupar por período en pivots. FECHA() construye fechas desde partes. Base de cualquier análisis temporal.","codigo":"=AÑO(A2)    → 2026\n=MES(A2)    → 4\n=DIA(A2)    → 6\n=FECHA(AÑO(A2),MES(A2),1)  → Primer día del mes" },
        { "id":"ex-16","titulo":"HOY / DIAS.LAB / SIFECHA","categoria":"Funciones de Fecha","descripcion":"HOY() devuelve fecha dinámica. DIAS.LAB cuenta días hábiles entre fechas. SIFECHA calcula antigüedad en días, meses o años.","codigo":"=HOY()                  → Hoy dinámico\n=DIAS.LAB(A2,B2)        → Días hábiles\n=SIFECHA(A2,HOY(),\"m\") → Meses antigüedad" },
        // BÚSQUEDA
        { "id":"ex-17","titulo":"BUSCARV (VLOOKUP)","categoria":"Búsqueda y Referencia","descripcion":"Busca en la primera columna de una tabla y devuelve la columna indicada. FALSO=coincidencia exacta. El estándar histórico de los cruces de datos.","codigo":"=BUSCARV(A2,$E$2:$G$6,2,FALSO)" },
        { "id":"ex-18","titulo":"INDICE & COINCIDIR","categoria":"Búsqueda y Referencia","descripcion":"Combina INDICE (extrae valor) con COINCIDIR (calcula posición). Permite buscar hacia la izquierda y es más rápida en tablas grandes que BUSCARV.","codigo":"=INDICE(C2:C6,COINCIDIR(A9,A2:A6,0))" },
        { "id":"ex-19","titulo":"BUSCARX (XLOOKUP)","categoria":"Búsqueda y Referencia","descripcion":"El BUSCARV evolucionado de M365. Busca en cualquier dirección, maneja errores nativamente, puede devolver múltiples columnas de una sola vez.","codigo":"=BUSCARX(A9,A2:A6,C2:C6,\"No encontrado\",0)" },
        // MATRICES DINÁMICAS
        { "id":"ex-20","titulo":"FILTRAR()","categoria":"Matrices Dinámicas","descripcion":"Devuelve un subconjunto dinámico que cumple la condición. Resultado 'desborda' automáticamente (Spill). Reemplaza filtros manuales en miles de filas.","codigo":"=FILTRAR(A2:C6,C2:C6>500,\"Sin resultados\")" },
        { "id":"ex-21","titulo":"UNIQUE() y ORDENAR()","categoria":"Matrices Dinámicas","descripcion":"UNICOS() elimina duplicados dinámicamente. ORDENAR() clasifica el resultado. Combinados reemplazan horas de trabajo manual de deduplicación.","codigo":"=ORDENAR(UNICOS(A2:A100))" },
        { "id":"ex-22","titulo":"LET() — Variables en fórmulas","categoria":"Matrices Dinámicas","descripcion":"Declara variables internas en la fórmula evitando repetir cálculos. Mejora la legibilidad y el rendimiento de fórmulas complejas. Solo M365.","codigo":"=LET(tasa,1.19,base,SUMA(B2:B6),base*tasa)" },
        // POWER QUERY
        { "id":"ex-23","titulo":"Power Query: Lenguaje M","categoria":"Power Query (ETL)","descripcion":"Transformaciones repetibles y auditables. Table.SelectRows filtra en el origen (Query Folding). Table.Buffer carga en RAM para cruces rápidos.","codigo":"Table.SelectRows(Source, each [Region]=\"Norte\")\nTable.Buffer(TablaBase)" },
        // PIVOT
        { "id":"ex-24","titulo":"Tablas Dinámicas + Slicers","categoria":"Tablas Dinámicas","descripcion":"Resumen millones de filas en segundos. Los Slicers permiten filtrado interactivo visual. Usar 'Conexiones de informe' para controlar múltiples pivots.","codigo":"1. Insertar → Tabla Dinámica\n2. Insertar → Segmentación de datos\n3. Clic derecho → Conexiones de informe" }
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
    "ex-1": null,  // hardware
    "ex-2": null,  // hardware
    // ex-3: SUMA & PROMEDIO — tabla de ventas por vendedor
    "ex-3": {
        cols:["A","B"], headers:["Vendedor","Ventas"],
        rows:[["Ana","380"],["Luis","520"],["María","290"],["Pedro","610"],["Carlos","455"],["TOTAL →",""]],
        target:{row:6,col:2}, hint:"=SUMA(B2:B6) o =PROMEDIO(B2:B6)"
    },
    // ex-4: SUMAR.SI.CONJUNTO — ventas por región + producto
    "ex-4": {
        cols:["A","B","C","D"], headers:["Región","Producto","Ventas","Resultado"],
        rows:[["Norte","Laptop","720",""],["Sur","Mouse","180",""],["Norte","Laptop","730",""],["Sur","Laptop","620",""],["Norte","Teclado","190",""],["SUMAR.SI.CONJ →","","",""]],
        target:{row:6,col:4}, hint:"=SUMAR.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\",B2:B6,\"Laptop\")"
    },
    // ex-5: CONTAR.SI.CONJUNTO — clientes VIP de CO con compras >500
    "ex-5": {
        cols:["A","B","C","D"], headers:["País","Segmento","Compra","Resultado"],
        rows:[["CO","VIP","780",""],["MX","VIP","320",""],["CO","Std","200",""],["CO","VIP","650",""],["MX","Std","900",""],["CONTAR.SI.CONJ →","","",""]],
        target:{row:6,col:4}, hint:"=CONTAR.SI.CONJUNTO(A2:A6,\"CO\",B2:B6,\"VIP\",C2:C6,\">500\")"
    },
    // ex-6: MAX.SI.CONJUNTO — máximo de ventas por región
    "ex-6": {
        cols:["A","B","C"], headers:["Región","Vendedor","Ventas"],
        rows:[["Norte","Ana","380"],["Sur","Luis","620"],["Norte","María","890"],["Sur","Pedro","310"],["Norte","Carlos","550"],["MAX Norte →","",""]],
        target:{row:6,col:3}, hint:"=MAX.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\")"
    },
    // ex-7: REDONDEAR — cálculo de IVA
    "ex-7": {
        cols:["A","B","C"], headers:["Producto","Precio Base","Precio+IVA"],
        rows:[["Laptop","380",""],["Mouse","45",""],["Teclado","22",""],["Monitor","320",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=REDONDEAR(B2*1.19, 2)"
    },
    // ex-8: SI() anidado — clasificación de clientes
    "ex-8": {
        cols:["A","B","C"], headers:["Cliente","Compra","Segmento"],
        rows:[["Ana","380",""],["Luis","1200",""],["María","650",""],["Pedro","90",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=SI(B2>1000,\"VIP\",SI(B2>500,\"Medio\",\"Básico\"))"
    },
    // ex-9: SI.CONJUNTO — clasificación multi-rama
    "ex-9": {
        cols:["A","B","C"], headers:["Cliente","Venta","Nivel"],
        rows:[["Ana","380",""],["Luis","1500",""],["María","720",""],["Pedro","0",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=SI.CONJUNTO(B2>1000,\"VIP\",B2>500,\"Medio\",B2>0,\"Básico\",VERDADERO,\"Sin ventas\")"
    },
    // ex-10: SIERROR — proteger un BUSCARV
    "ex-10": {
        cols:["A","B","C","D","E"], headers:["ID Buscar","Nombre","","ID","País"],
        rows:[["X999","→ Resultado:","","C001","Colombia"],["","","","C002","México"],["","","","C003","Brasil"],["","","","",""],["","","","",""]],
        target:{row:1,col:2}, hint:"=SIERROR(BUSCARV(A2,$D$2:$E$4,2,FALSO),\"No encontrado\")"
    },
    // ex-11: IZQUIERDA / DERECHA / EXTRAE — parsear código
    "ex-11": {
        cols:["A","B","C"], headers:["Código SKU","Prefijo (2)","Resultado"],
        rows:[["CO-2024-001","→ Extraer:",""],["MX-2024-002","",""],["BR-2025-001","",""],["AR-2024-003","",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=IZQUIERDA(A2,2) → extrae el prefijo del país"
    },
    // ex-12: CONCATENAR — unir nombre + apellido
    "ex-12": {
        cols:["A","B","C"], headers:["Nombre","Apellido","Nombre Completo"],
        rows:[["Ana","García",""],["Luis","López",""],["María","Torres",""],["Pedro","Ruiz",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=CONCATENAR(A2,\" \",B2)"
    },
    // ex-13: TEXTO() — formatear precios
    "ex-13": {
        cols:["A","B","C"], headers:["Precio Raw","Formato","Resultado"],
        rows:[["1250","$#,##0.00",""],["0.855","0.0%",""],["44922","dd/mm/aaaa",""],["380","$ #,##0",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=TEXTO(A2,B2) — aplica el formato de la columna B"
    },
    // ex-14: LARGO / HALLAR — analizar email
    "ex-14": {
        cols:["A","B"], headers:["Email","Long. / Pos @"],
        rows:[["ana@empresa.com",""],["luis.lopez@corp.co",""],["maria@gmail.com",""],["pedro@hotmail.com",""],["Formula →",""]],
        target:{row:5,col:2}, hint:"=LARGO(A2) o =HALLAR(\"@\",A2)"
    },
    // ex-15: Fechas — extraer componentes
    "ex-15": {
        cols:["A","B","C","D"], headers:["Fecha Pedido","Año","Mes","Día"],
        rows:[["06/04/2026","","",""],["15/01/2025","","",""],["28/11/2024","","",""],["03/07/2026","","",""],["Formula →","","",""]],
        target:{row:5,col:2}, hint:"=AÑO(A2) en col B, =MES(A2) en col C"
    },
    // ex-16: HOY / DIAS.LAB — calcular antigüedad
    "ex-16": {
        cols:["A","B","C"], headers:["Contrato","Hoy/Fin","Días Hábiles"],
        rows:[["01/01/2026","06/04/2026",""],["15/02/2026","06/04/2026",""],["01/03/2026","06/04/2026",""],["","",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=DIAS.LAB(A2,B2) — días hábiles entre fechas"
    },
    // ex-17: BUSCARV — cruce de clientes
    "ex-17": {
        cols:["A","B","C","D","E"], headers:["ID Buscar","Nombre","","ID","Nombre"],
        rows:[["C003","→ Nombre:","","C001","Ana"],["","","","C002","Luis"],["","","","C003","María"],["","","","C004","Pedro"],["","","","C005","Carlos"]],
        target:{row:1,col:2}, hint:"=BUSCARV(A2,$D$2:$E$6,2,FALSO)"
    },
    // ex-18: INDICE & COINCIDIR
    "ex-18": {
        cols:["A","B","C","D"], headers:["Producto","Precio","","Buscar"],
        rows:[["Laptop","1200","","Teclado"],["Mouse","25","→ Precio:",""],["Teclado","45","",""],["Monitor","320","",""],["Webcam","80","",""]],
        target:{row:2,col:3}, hint:"=INDICE(B2:B6,COINCIDIR(D2,A2:A6,0))"
    },
    // ex-19: BUSCARX
    "ex-19": {
        cols:["A","B","C","D"], headers:["Producto","Precio","","Buscar"],
        rows:[["Laptop","1200","","Mouse"],["Mouse","25","→ Precio:",""],["Teclado","45","",""],["Monitor","320","",""],["Webcam","80","",""]],
        target:{row:2,col:3}, hint:"=BUSCARX(D2,A2:A6,B2:B6,\"No encontrado\",0)"
    },
    // ex-20: FILTRAR — ventas > 500
    "ex-20": {
        cols:["A","B","C"], headers:["Vendedor","Región","Ventas"],
        rows:[["Ana","Norte","380"],["Luis","Sur","620"],["María","Norte","290"],["Pedro","Norte","890"],["Carlos","Sur","455"],["→ FILTRAR >500","",""]],
        target:{row:6,col:1}, hint:"=FILTRAR(A2:C6,C2:C6>500,\"Sin resultados\")"
    },
    // ex-21: UNICOS + ORDENAR
    "ex-21": {
        cols:["A","B"], headers:["Ciudades","Ordenadas"],
        rows:[["Bogotá",""],["Medellín",""],["Bogotá",""],["Cali",""],["Medellín","→"],["Formula →",""]],
        target:{row:6,col:2}, hint:"=ORDENAR(UNICOS(A2:A6))"
    },
    // ex-22: LET()
    "ex-22": {
        cols:["A","B","C"], headers:["Vendedor","Ventas","Total+IVA"],
        rows:[["Ana","380",""],["Luis","520",""],["María","290",""],["Pedro","610",""],["Formula →","",""]],
        target:{row:5,col:3}, hint:"=LET(tasa,1.19,base,SUMA(B2:B5),base*tasa)"
    },
    // ex-23 y ex-24: Terminales (Power Query y Pivot)
    "ex-23": null,
    "ex-24": null
};


// ─────────────────────────────────────────────────────────────────────────────
// MAPA DE EJERCICIOS — Validadores y tipos
// ─────────────────────────────────────────────────────────────────────────────
const EXERCISE_MAP = {
    "ex-1":  { type:"hardware", key:"ctrl+space",   instruccion:"Presiona Ctrl + Espacio para seleccionar la columna entera.", exito:"¡Ctrl + Espacio detectado! Columna seleccionada." },
    "ex-2":  { type:"hardware", key:"shift+space",  instruccion:"Presiona Shift + Espacio para seleccionar la fila entera.", exito:"¡Shift + Espacio detectado! Fila seleccionada." },
    // Matemáticas
    "ex-3":  { type:"formula", regex:/^=SUMA\(B2:B6\)|^=PROMEDIO\(B2:B6\)/i,  instruccion:"Escribe =SUMA(B2:B6) o =PROMEDIO(B2:B6) en la celda marcada.", exito:"¡Correcto!", compute:(f) => f.match(/PROMEDIO/i) ? "431" : "2.255" },
    "ex-4":  { type:"formula", regex:/^=SUMAR\.SI\.CONJUNTO\(.+\)$/i,          instruccion:"Escribe =SUMAR.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\",B2:B6,\"Laptop\")", exito:"¡SUMAR.SI.CONJUNTO correcto! Total: 1.450", compute:() => "1.450" },
    "ex-5":  { type:"formula", regex:/^=CONTAR\.SI\.CONJUNTO\(.+\)$/i,         instruccion:"Escribe =CONTAR.SI.CONJUNTO(A2:A6,\"CO\",B2:B6,\"VIP\",C2:C6,\">500\")", exito:"¡CONTAR.SI.CONJUNTO correcto! 2 registros.", compute:() => "2" },
    "ex-6":  { type:"formula", regex:/^=MAX\.SI\.CONJUNTO\(.+\)|^=MIN\.SI\.CONJUNTO\(.+\)/i, instruccion:"Escribe =MAX.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\")", exito:"¡MAX condicional correcto! Máximo Norte: 890", compute:() => "890" },
    "ex-7":  { type:"formula", regex:/^=REDONDEAR\(.+,.+\)|^=ENTERO\(.+\)/i,   instruccion:"Escribe =REDONDEAR(A2*1.19, 2) para calcular IVA redondeado.", exito:"¡REDONDEAR correcto!", compute:(f) => { try { return String(Math.round(380*1.19*100)/100); } catch(e){ return "452.20"; } } },
    // Lógicas
    "ex-8":  { type:"formula", regex:/^=SI\(.+,SI\(.+\)\)/i,                   instruccion:"Escribe =SI(B2>1000,\"VIP\",SI(B2>500,\"Medio\",\"Básico\"))", exito:"¡SI anidado correcto!", compute:() => "Básico" },
    "ex-9":  { type:"formula", regex:/^=SI\.CONJUNTO\(.+\)$/i,                  instruccion:"Escribe =SI.CONJUNTO(B2>1000,\"VIP\",B2>500,\"Medio\",B2>0,\"Básico\",VERDADERO,\"Sin ventas\")", exito:"¡SI.CONJUNTO correcto!", compute:() => "Básico" },
    "ex-10": { type:"formula", regex:/^=SIERROR\(.+\)|^=SINO\.DISPONIBLE\(.+\)/i, instruccion:"Escribe =SIERROR(BUSCARV(A2,Tabla,2,0),\"No encontrado\")", exito:"¡SIERROR correcto! Errores capturados.", compute:() => "No encontrado" },
    // Texto
    "ex-11": { type:"formula", regex:/^=IZQUIERDA\(.+\)|^=DERECHA\(.+\)|^=EXTRAE\(.+\)/i, instruccion:"Escribe =IZQUIERDA(A2,2) para extraer el prefijo del código.", exito:"¡Texto extraído!", compute:() => "CO" },
    "ex-12": { type:"formula", regex:/^=CONCATENAR\(.+\)|^=UNIRCADENAS\(.+\)/i, instruccion:"Escribe =CONCATENAR(A2,\" \",B2) para unir nombre y apellido.", exito:"¡Cadenas unidas!", compute:() => "Ana García" },
    "ex-13": { type:"formula", regex:/^=TEXTO\(.+,.+\)$/i,                      instruccion:"Escribe =TEXTO(A2,\"$#,##0.00\") para formatear el precio.", exito:"¡TEXTO correcto!", compute:() => "$1,250.00" },
    "ex-14": { type:"formula", regex:/^=LARGO\(.+\)|^=HALLAR\(.+\)|^=SUSTITUIR\(.+\)/i, instruccion:"Escribe =LARGO(A2) para contar los caracteres del email.", exito:"¡LARGO correcto!", compute:() => "15" },
    // Fechas
    "ex-15": { type:"formula", regex:/^=AÑO\(.+\)|^=MES\(.+\)|^=DIA\(.+\)|^=FECHA\(.+\)/i, instruccion:"Escribe =AÑO(A2) para extraer el año de la fecha de pedido.", exito:"¡Función de fecha correcta!", compute:() => "2026" },
    "ex-16": { type:"formula", regex:/^=HOY\(\)|^=DIAS\.LAB\(.+\)|^=SIFECHA\(.+\)/i,       instruccion:"Escribe =DIAS.LAB(A2,B2) para calcular días hábiles entre fechas.", exito:"¡Días hábiles calculados!", compute:() => "18 días hábiles" },
    // Búsqueda
    "ex-17": { type:"formula", regex:/^=BUSCARV\(.+\)$/i,                       instruccion:"Escribe =BUSCARV(A2,$D$2:$E$6,2,FALSO) para buscar el nombre del cliente.", exito:"¡BUSCARV correcto! Resultado: María", compute:() => "María" },
    "ex-18": { type:"formula", regex:/^=INDICE\(.+,COINCIDIR\(.+\)\)$/i,        instruccion:"Escribe =INDICE(B2:B6,COINCIDIR(D2,A2:A6,0))", exito:"¡INDICE+COINCIDIR correcto! Resultado: 45", compute:() => "45" },
    "ex-19": { type:"formula", regex:/^=BUSCARX\(.+\)$/i,                       instruccion:"Escribe =BUSCARX(D2,A2:A6,B2:B6,\"No encontrado\",0)", exito:"¡BUSCARX correcto! Resultado: 25", compute:() => "25" },
    // Matrices
    "ex-20": { type:"formula", regex:/^=FILTRAR\(.+\)$/i, instruccion:"Escribe =FILTRAR(A2:C6,C2:C6>500,\"Sin resultados\")", exito:"¡FILTRAR dinámico correcto!", compute:() => "Luis (620)\nPedro (890)" },
    "ex-21": { type:"formula", regex:/^=ORDENAR\(UNICOS\(.+\)\)|^=UNICOS\(.+\)/i, instruccion:"Escribe =ORDENAR(UNICOS(A2:A6)) para ciudades únicas.", exito:"¡UNICOS+ORDENAR correcto!", compute:() => "Bogotá, Cali, Medellín" },
    "ex-22": { type:"formula", regex:/^=LET\(.+\)$/i, instruccion:"Escribe =LET(tasa,1.19,base,SUMA(B2:B5),base*tasa)", exito:"¡Variable LET calculada!", compute:() => "2,082.50" },
    // Power Query + Pivot (terminal)
    "ex-23": { type:"terminal", regex:/Table\.(SelectRows|Buffer)\(/i, instruccion:"Escribe: Table.SelectRows(Source, each [Region]=\"Norte\")", exito:"¡Power Query M correcto!" },
    "ex-24": { type:"terminal", regex:/slicer|segmentacion|segmentación|tabla dinámica|pivot/i, instruccion:"Describe: 1. Insertar Tabla Dinámica  2. Insertar Segmentador  3. Conexiones de informe", exito:"¡Flujo Pivot+Slicer correcto!" },

    "sq-1":  { type:"terminal", regex:/SELECT\s+DISTINCT/i, instruccion:"Escribe SELECT DISTINCT user_id FROM logs LIMIT 100;", exito:"¡SELECT DISTINCT correcto!" },
    "sq-2":  { type:"terminal", regex:/WHERE.+(IN\s*\(|AND|OR)/i, instruccion:"Filtra con WHERE pais IN ('CO','MX').", exito:"¡WHERE + IN correcto!" },
    "sq-3":  { type:"terminal", regex:/INNER\s+JOIN/i, instruccion:"Escribe un INNER JOIN entre users y orders.", exito:"¡INNER JOIN correcto!" },
    "sq-4":  { type:"terminal", regex:/LEFT\s+JOIN/i, instruccion:"Escribe un LEFT JOIN de users a orders.", exito:"¡LEFT JOIN correcto!" },
    "sq-5":  { type:"terminal", regex:/FULL\s+(OUTER\s+)?JOIN/i, instruccion:"Escribe FULL OUTER JOIN entre tabla A y B.", exito:"¡FULL OUTER JOIN correcto!" },
    "sq-6":  { type:"terminal", regex:/LEFT\s+JOIN[\s\S]+IS\s+NULL/i, instruccion:"ANTI-JOIN: LEFT JOIN + WHERE columna IS NULL.", exito:"¡ANTI-JOIN correcto!" },
    "sq-7":  { type:"terminal", regex:/SUM\(|AVG\(|COUNT\(/i, instruccion:"SELECT region, SUM(monto) FROM ventas GROUP BY region;", exito:"¡Agregación correcta!" },
    "sq-8":  { type:"terminal", regex:/HAVING/i, instruccion:"GROUP BY ... HAVING COUNT(*) > 5", exito:"¡HAVING correcto!" },
    "sq-9":  { type:"terminal", regex:/ROW_NUMBER\(\)\s*OVER/i, instruccion:"ROW_NUMBER() OVER(PARTITION BY user_id ORDER BY date DESC)", exito:"¡ROW_NUMBER correcto!" },
    "sq-10": { type:"terminal", regex:/DENSE_RANK\(\)|RANK\(\)/i, instruccion:"DENSE_RANK() OVER(ORDER BY ventas DESC)", exito:"¡DENSE_RANK correcto!" },
    "sq-11": { type:"terminal", regex:/LAG\(|LEAD\(/i, instruccion:"LAG(order_date) OVER(PARTITION BY user_id ORDER BY date)", exito:"¡LAG correcto!" },
    "sq-12": { type:"terminal", regex:/NTILE\(\d+\)/i, instruccion:"NTILE(4) OVER(ORDER BY monto DESC)", exito:"¡NTILE correcto!" },
    "sq-13": { type:"terminal", regex:/WITH\s+\w+\s+AS\s*\(/i, instruccion:"WITH activos AS (SELECT * FROM users WHERE activo=1) SELECT * FROM activos;", exito:"¡CTE correcto!" },
    "sq-14": { type:"terminal", regex:/EXPLAIN(\s+ANALYZE)?/i, instruccion:"EXPLAIN ANALYZE SELECT * FROM orders WHERE user_id = 42;", exito:"¡EXPLAIN correcto!" },
    "py-1":  { type:"terminal", regex:/[\[{]/, instruccion:"Crea: clientes = [{'id':1,'nombre':'Ana'}]", exito:"¡Diccionario correcto!" },
    "py-2":  { type:"terminal", regex:/set\(|tuple\(/i, instruccion:"ids = set([1,2,2,3])", exito:"¡Set correcto!" },
    "py-3":  { type:"terminal", regex:/for\s+\w+\s+in|while\s+/, instruccion:"Escribe un bucle for sobre una lista.", exito:"¡Bucle correcto!" },
    "py-4":  { type:"terminal", regex:/lambda/, instruccion:"df['seg'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Std')", exito:"¡Lambda correcto!" },
    "py-5":  { type:"terminal", regex:/@\w+/, instruccion:"@timer sobre una función de carga.", exito:"¡Decorador correcto!" },
    "py-6":  { type:"terminal", regex:/df\.(head|describe)\(/, instruccion:"df.describe().T", exito:"¡df.describe correcto!" },
    "py-7":  { type:"terminal", regex:/df\.info\(\)/, instruccion:"df.info()", exito:"¡df.info correcto!" },
    "py-8":  { type:"terminal", regex:/df\.dropna\(/, instruccion:"df.dropna(axis=0) o df.dropna(subset=['email'])", exito:"¡dropna correcto!" },
    "py-9":  { type:"terminal", regex:/df\[.+\]\.fillna\(/, instruccion:"df['edad'].fillna(df['edad'].median(), inplace=True)", exito:"¡fillna correcto!" },
    "py-10": { type:"terminal", regex:/\.astype\(/, instruccion:"df['precio'].astype('float64')", exito:"¡astype correcto!" },
    "py-11": { type:"terminal", regex:/drop_duplicates\(/, instruccion:"df.drop_duplicates(subset=['email'], keep='last')", exito:"¡drop_duplicates correcto!" },
    "py-12": { type:"terminal", regex:/groupby\(/, instruccion:"df.groupby('id').agg({'orden':'count','monto':'sum'})", exito:"¡groupby correcto!" },
    "py-13": { type:"terminal", regex:/pd\.merge\(/, instruccion:"pd.merge(clientes, ordenes, on='id', how='left')", exito:"¡pd.merge correcto!" },
    "py-14": { type:"terminal", regex:/heatmap\(|kdeplot\(/, instruccion:"sns.heatmap(df.corr(), annot=True)", exito:"¡Seaborn correcto!" },
    "gh-1":  { type:"terminal", regex:/git\s+config\s+--global/, instruccion:"git config --global user.name \"TuNombre\"", exito:"¡git config correcto!" },
    "gh-2":  { type:"terminal", regex:/git\s+(init|clone)/, instruccion:"git clone https://github.com/usuario/repo.git", exito:"¡git clone correcto!" },
    "gh-3":  { type:"terminal", regex:/git\s+status/, instruccion:"git status", exito:"¡git status correcto!" },
    "gh-4":  { type:"terminal", regex:/git\s+add/, instruccion:"git add .", exito:"¡git add correcto!" },
    "gh-5":  { type:"terminal", regex:/git\s+commit\s+-m/, instruccion:"git commit -m \"feat: descripción\"", exito:"¡git commit correcto!" },
    "gh-6":  { type:"terminal", regex:/git\s+(push|pull)/, instruccion:"git push origin main", exito:"¡git push correcto!" },
    "gh-7":  { type:"terminal", regex:/git\s+(branch|checkout)/, instruccion:"git checkout -b feature/mi-rama", exito:"¡git branch correcto!" },
    "gh-8":  { type:"terminal", regex:/git\s+push.+origin|pull.request/i, instruccion:"git push origin feature/mi-rama", exito:"¡PR flow correcto!" },
    "gh-9":  { type:"terminal", regex:/fork|clone.+github/i, instruccion:"Describe el flujo de Fork + clone.", exito:"¡Fork correcto!" },
    "in-1":  { type:"formula", regex:/^=PY\(/i, instruccion:"Escribe =PY( para iniciar celda Python en Excel.", exito:"¡=PY() correcto!", compute: () => "DataFrame(5×8)" },
    "in-2":  { type:"terminal", regex:/create_engine\(/, instruccion:"create_engine('postgresql://user:pass@host/db')", exito:"¡SQLAlchemy correcto!" },
    "in-3":  { type:"terminal", regex:/\.to_sql\(/, instruccion:"df.to_sql('tabla', engine, if_exists='replace')", exito:"¡to_sql correcto!" }
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
    
    // Obtener progreso de la biblioteca
    const libProgress = typeof getProgress === 'function' ? getProgress('lib') : [];
    
    grid.innerHTML = dataArray.map(item => {
        const isDone = libProgress.includes(item.id);
        const doneBadge = isDone ? '<div class="done-badge">✓</div>' : '';
        
        return `
        <div class="function-card ${theme}-hover ${isDone ? 'card-done' : ''}" 
             onclick="openFunctionModal('${item.id}','${theme}')">
            ${doneBadge}
            <span class="card-category ${theme}-color">${item.categoria}</span>
            <h3 class="card-title">${item.titulo}</h3>
            <p class="card-excerpt">${item.descripcion}</p>
        </div>`;
    }).join('');
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
        
        saveProgress(currentItemId, 'lib'); // Guardar progreso en la biblioteca
        if (typeof triggerSuccess === 'function') triggerSuccess();
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
        saveProgress(currentItemId, 'lib'); // Guardar progreso en la biblioteca
        if (typeof triggerSuccess === 'function') triggerSuccess();
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
