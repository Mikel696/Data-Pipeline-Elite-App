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
        { "id":"sq-1",  "titulo":"SELECT DISTINCT",          "categoria":"Sintaxis Básica",           "descripcion":"DISTINCT colapsa duplicados en el set de resultados, crucial para hallar valores únicos como IDs de usuarios activos. LIMIT restringe el número de filas.", "codigo":"SELECT DISTINCT user_id\nFROM logs\nLIMIT 100;" },
        { "id":"sq-3",  "titulo":"INNER JOIN",               "categoria":"Uniones Relacionales",      "descripcion":"Retorna únicamente los registros con coincidencia exacta en las llaves de ambas tablas. Es el más eficiente para cruzar entidades (ej. Usuarios con Compras).", "codigo":"SELECT u.nombre, o.total\nFROM users u\nINNER JOIN orders o ON u.id = o.user_id;" },
        { "id":"sq-7",  "titulo":"Agregaciones (SUM/AVG)",   "categoria":"Agregación",                "descripcion":"Funciones que colapsan millones de filas en KPIs de negocio. Siempre requieren GROUP BY si se mezclan con columnas dimensionales.", "codigo":"SELECT region, SUM(monto)\nFROM ventas\nGROUP BY region;" },
        { "id":"sq-13", "titulo":"CTEs (WITH)",              "categoria":"Subconsultas",              "descripcion":"Las Common Table Expressions (CTE) mejoran la legibilidad y permiten organizar queries jerárquicamente. Reemplazan subqueries anidadas.", "codigo":"WITH activos AS (\n  SELECT * FROM users WHERE status='active'\n)\nSELECT * FROM activos;" }
    ],
    "python": [
        { "id":"py-1",  "titulo":"Listas y Diccionarios",    "categoria":"Estructuras Nativas",      "descripcion":"Vectores mutables (list) y mapas hash (dict). Son las unidades fundamentales de transferencia de datos en APIs REST y pipelines locales.", "codigo":"item = {'id':1,'val':380}\nlista = ['Ana','Luis']" },
        { "id":"py-4",  "titulo":"Funciones Lambda",         "categoria":"Funciones y Decoradores",  "descripcion":"Funciones anónimas de una sola línea. En análisis de datos, se usan casi exclusivamente dentro de `.apply()` para transformaciones columna a columna rápidas.", "codigo":"df['seg'] = df['gasto'].apply(lambda x: 'VIP' if x > 1000 else 'Std')" },
        { "id":"py-6",  "titulo":"Pandas: Exploración",       "categoria":"Pandas Core",              "descripcion":"`.head()` permite verificar la correcta carga del dataset, mientras que `.describe()` arroja un resumen estadístico (cuartiles, media, std) de todas las columnas numéricas.", "codigo":"df.head(10)\ndf.describe().T" },
        { "id":"py-8",  "titulo":"Limpieza de NaNs",         "categoria":"Data Wrangling",           "descripcion":"`.dropna()` elimina ruidos que podrían sesgar cálculos estadísticos. Se puede filtrar por filas (`axis=0`) o columnas (`axis=1`).", "codigo":"df.dropna(subset=['email'], inplace=True)" },
        { "id":"py-12", "titulo":"Group & Aggregate",       "categoria":"Transformación Tabular",   "descripcion":"Indispensable para reportes RFM o agregaciones temporales. `.groupby()` bloquea dimensiones y `.agg()` aplica funciones matemáticas personalizadas sobre ellas.", "codigo":"df.groupby('id').agg({'monto':'sum','user':'count'})" },
        { "id":"py-13", "titulo":"Pandas Merge (JOIN)",      "categoria":"Transformación Tabular",   "descripcion":"Equivalente al JOIN de SQL en RAM. Permite conectar fuentes heterogéneas (ej: un CSV de ventas con un JSON de perfiles) mediante llaves comunes.", "codigo":"pd.merge(clientes, ordenes, on='id', how='left')" },
        { "id":"py-14", "titulo":"Seaborn: Correlación",      "categoria":"Visualización Científica", "descripcion":"`heatmap` visualiza la matriz de correlación entre variables. Clave para hallar predictores de Churn antes de entrenar un modelo de Machine Learning.", "codigo":"sns.heatmap(df.corr(), annot=True, cmap='viridis')" }
    ],
    "github": [
        { "id":"gh-1", "titulo":"Git Config & Auth",    "categoria":"Configuración Core",   "descripcion":"Establece tu firma digital en el historial. Sin una identidad configurada, Git rechazará tus intentos de commit para preservar la trazabilidad.", "codigo":"git config --global user.name \"Name\"\ngit config --global user.email \"mail\"" },
        { "id":"gh-4", "titulo":"Git Add (Staging)",    "categoria":"Flujo Workdir",        "descripcion":"Mueve cambios del directorio de trabajo al área de preparación (Staging). Es la 'sala de espera' antes de confirmar una versión inmutable.", "codigo":"git add .\ngit add script.py" },
        { "id":"gh-5", "titulo":"Git Commit (Atomic)",  "categoria":"Historia Inmutable",   "descripcion":"Crea un punto de restauración permanente en la historia del proyecto. Los commits deben ser 'atómicos': un solo cambio lógico por cada mensaje.", "codigo":"git commit -m \"feat: add churn model\"" },
        { "id":"gh-6", "titulo":"Git Push / Pull",      "categoria":"Nube (Origin)",        "descripcion":"`push` sincroniza tus commits locales con el servidor remoto (GitHub). `pull` descarga y fusiona los cambios de tus compañeros en tu rama actual.", "codigo":"git push origin main\ngit pull origin main" },
        { "id":"gh-8", "titulo":"Pull Requests (PR)",    "categoria":"Colaboración",         "descripcion":"El mecanismo de revisión de código por excelencia. Permite proponer cambios a una rama protegida, habilitando discusiones y pruebas automáticas antes del merge.", "codigo":"1. Push rama local\n2. Open PR en GitHub\n3. Review & Merge" }
    ],
    "integration": [
        { "id":"in-1", "titulo":"Pandas xl() Engine",   "categoria":"Python en Excel",      "descripcion":"Accede directamente a rangos y tablas de Excel como objetos DataFrame nativos. Permite usar el poder de limpieza de Pandas sin salir del ecosistema MS.", "codigo":"df = xl(\"Table1[All]\", headers=True)\nreturn df.describe()" },
        { "id":"in-2", "titulo":"Pandas to SQL",        "categoria":"ETL Pipeline",         "descripcion":"Cierra el ciclo de datos escribiendo los resultados de un análisis Python directamente en una base de datos corporativa para su consumo en dashboards BI.", "codigo":"df.to_sql('reporte_ventas', engine,\nif_exists='replace')" }
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
// CONTEXTOS DE MICRO-LAB (SQL, Python, etc.)
// ─────────────────────────────────────────────────────────────────────────────
const MICRO_LAB_CONTEXTS = {
    // SQL
    "sq-1":  { title: "Tabla: User_Logs", html: "<table class='micro-lab-table'><thead><tr><th>user_id</th><th>action</th><th>timestamp</th></tr></thead><tbody><tr><td>101</td><td>login</td><td>2026-04-01</td></tr><tr><td>102</td><td>click</td><td>2026-04-01</td></tr><tr><td>101</td><td>logout</td><td>2026-04-01</td></tr></tbody></table>" },
    "sq-3":  { title: "Schema: Users & Orders", html: "<div style='display:grid; grid-template-columns:1fr 1fr; gap:10px;'><table class='micro-lab-table'><thead><tr><th>u.id</th><th>u.nombre</th></tr></thead><tbody><tr><td>1</td><td>Ana</td></tr><tr><td>2</td><td>Luis</td></tr></tbody></table><table class='micro-lab-table'><thead><tr><th>o.user_id</th><th>o.total</th></tr></thead><tbody><tr><td>1</td><td>500</td></tr><tr><td>1</td><td>200</td></tr></tbody></table></div>" },
    "sq-7":  { title: "Tabla: Ventas_Global", html: "<table class='micro-lab-table'><thead><tr><th>region</th><th>monto</th><th>vendedor</th></tr></thead><tbody><tr><td>Norte</td><td>1500</td><td>Ana</td></tr><tr><td>Sur</td><td>1200</td><td>Luis</td></tr><tr><td>Norte</td><td>800</td><td>Pedro</td></tr></tbody></table>" },
    "sq-13": { title: "Contexto: CTE 'activos'", html: "<div class='micro-lab-obj'>-- Vista lógica en memoria --\nactivos = SELECT * FROM users WHERE status='active'</div>" },
    
    // Python
    "py-1":  { title: "Variable: clientes", html: "<div class='micro-lab-obj'>clientes = [\n  {'id': 1, 'nombre': 'Ana'},\n  {'id': 2, 'nombre': 'Luis'}\n]</div>" },
    "py-4":  { title: "DataFrame: df (Pre-apply)", html: "<table class='micro-lab-table'><thead><tr><th>id</th><th>gasto</th></tr></thead><tbody><tr><td>0</td><td>1200</td></tr><tr><td>1</td><td>450</td></tr></tbody></table>" },
    "py-6":  { title: "DataFrame: df", html: "<table class='micro-lab-table'><thead><tr><th>id</th><th>gasto</th><th>segmento</th></tr></thead><tbody><tr><td>0</td><td>1200</td><td>VIP</td></tr><tr><td>1</td><td>450</td><td>Std</td></tr><tr><td>2</td><td>890</td><td>Mid</td></tr></tbody></table>" },
    "py-8":  { title: "DataFrame: df_with_nas", html: "<table class='micro-lab-table'><thead><tr><th>id</th><th>email</th><th>edad</th></tr></thead><tbody><tr><td>0</td><td>ana@mail.com</td><td>25</td></tr><tr><td>1</td><td>NaN</td><td>32</td></tr><tr><td>2</td><td>luis@mail.com</td><td>NaN</td></tr></tbody></table>" },
    "py-12": { title: "DataFrame: logs", html: "<table class='micro-lab-table'><thead><tr><th>user_id</th><th>monto</th></tr></thead><tbody><tr><td>101</td><td>500</td></tr><tr><td>101</td><td>1200</td></tr><tr><td>102</td><td>450</td></tr></tbody></table>" },
    "py-13": { title: "Datasets: Clientes + Ordenes", html: "<div style='display:grid; grid-template-columns:1fr 1fr; gap:10px;'><div class='micro-lab-obj'>df_c: [id, nombre]</div><div class='micro-lab-obj'>df_o: [user_id, total]</div></div>" },
    "py-14": { title: "Matrix: df.corr()", html: "<div class='micro-lab-obj'>[ [1.0, 0.85], [0.85, 1.0] ]\n(Sesgo detectado en Churn)</div>" },
    
    // GitHub
    "gh-1":  { title: "Terminal Config", html: "<div class='micro-lab-obj'>$ git config --list\n(Identity missing or incorrect)</div>" },
    "gh-4":  { title: "Git Status: Staging Area", html: "<div class='micro-lab-obj'>Changes not staged for commit:\n  modified:   script.py\n  (use \"git add ...\" to stage)</div>" },
    "gh-5":  { title: "Atomic Commit History", html: "<div class='micro-lab-obj'>[main 55a2b3] feat: add churn model\n1 file changed, 45 insertions(+)</div>" },
    "gh-8":  { title: "GitHub Flow Visual", html: "<div class='micro-lab-obj'>Main [Protected] <--- PR [Feature]\n(Check: unit-tests PASSED)</div>" },
    
    // Integración
    "in-1":  { title: "Excel Object: xl()", html: "<div class='micro-lab-obj'># Representación del rango A1:C6\n[[Ana, Norte, 380], [Luis, Sur, 620], ...]\n# Pandas cargará esto automáticamente.</div>" },
    "in-2":  { title: "Pipeline: df.to_sql()", html: "<div class='micro-lab-obj'>Connection: postgresql://db_admin\nStatus: Waiting for write command...</div>" }
};


// ─────────────────────────────────────────────────────────────────────────────
// MAPA DE EJERCICIOS — Validadores y tipos
// ─────────────────────────────────────────────────────────────────────────────
const EXERCISE_MAP = {
    // Excel (Hardware)
    "ex-1":  { type:"hardware", key:"ctrl+space",   instruccion:"Presiona Ctrl + Espacio para seleccionar la columna entera.", exito:"¡Ctrl + Espacio detectado! Columna seleccionada." },
    "ex-2":  { type:"hardware", key:"shift+space",  instruccion:"Presiona Shift + Espacio para seleccionar la fila entera.", exito:"¡Shift + Espacio detectado! Fila seleccionada." },
    // Excel (Fórmulas)
    "ex-3":  { type:"formula", regex:/^=SUMA\(B2:B6\)|^=PROMEDIO\(B2:B6\)/i, instruccion:"Escribe =SUMA(B2:B6) o =PROMEDIO(B2:B6).", exito:"¡Correcto!", compute:(f) => f.match(/PROMEDIO/i) ? "431" : "2.255" },
    "ex-4":  { type:"formula", regex:/^=SUMAR\.SI\.CONJUNTO\(.+\)$/i, instruccion:"Escribe =SUMAR.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\",B2:B6,\"Laptop\")", exito:"¡SUMAR.SI.CONJUNTO correcto!", compute:() => "1.450" },
    "ex-5":  { type:"formula", regex:/^=CONTAR\.SI\.CONJUNTO\(.+\)$/i, instruccion:"Escribe =CONTAR.SI.CONJUNTO(A2:A6,\"CO\",B2:B6,\"VIP\",C2:C6,\">500\")", exito:"¡CONTAR.SI.CONJUNTO correcto!", compute:() => "2" },
    "ex-6":  { type:"formula", regex:/^=MAX\.SI\.CONJUNTO\(.+\)|^=MIN\.SI\.CONJUNTO\(.+\)/i, instruccion:"Escribe =MAX.SI.CONJUNTO(C2:C6,A2:A6,\"Norte\")", exito:"¡MAX correcto!", compute:() => "890" },
    "ex-7":  { type:"formula", regex:/^=REDONDEAR\(.+,.+\)|^=ENTERO\(.+\)/i, instruccion:"Escribe =REDONDEAR(B2*1.19, 2).", exito:"¡REDONDEAR correcto!", compute:() => "452.20" },
    "ex-8":  { type:"formula", regex:/^=SI\(.+,SI\(.+\)\)/i, instruccion:"Escribe =SI(B2>1000,\"VIP\",SI(B2>500,\"Medio\",\"Básico\"))", exito:"¡SI anidado correcto!", compute:() => "Básico" },
    "ex-9":  { type:"formula", regex:/^=SI\.CONJUNTO\(.+\)$/i, instruccion:"Escribe =SI.CONJUNTO(B2>1000,\"VIP\",B2>500,\"Medio\",B2>0,\"Básico\",VERDADERO,\"Sin ventas\")", exito:"¡SI.CONJUNTO correcto!", compute:() => "Básico" },
    "ex-10": { type:"formula", regex:/^=SIERROR\(.+\)|^=SINO\.DISPONIBLE\(.+\)/i, instruccion:"Escribe =SIERROR(BUSCARV(A2,Tabla,2,0),\"No encontrado\")", exito:"¡SIERROR correcto!", compute:() => "No encontrado" },
    "ex-11": { type:"formula", regex:/^=IZQUIERDA\(.+\)|^=DERECHA\(.+\)|^=EXTRAE\(.+\)/i, instruccion:"Escribe =IZQUIERDA(A2,2).", exito:"¡Texto extraído!", compute:() => "CO" },
    "ex-12": { type:"formula", regex:/^=CONCATENAR\(.+\)|^=UNIRCADENAS\(.+\)/i, instruccion:"Escribe =CONCATENAR(A2,\" \",B2).", exito:"¡Cadenas unidas!", compute:() => "Ana García" },
    "ex-13": { type:"formula", regex:/^=TEXTO\(.+,.+\)$/i, instruccion:"Escribe =TEXTO(A2,\"$#,##0.00\").", exito:"¡TEXTO correcto!", compute:() => "$1,250.00" },
    "ex-14": { type:"formula", regex:/^=LARGO\(.+\)|^=HALLAR\(.+\)|^=SUSTITUIR\(.+\)/i, instruccion:"Escribe =LARGO(A2).", exito:"¡LARGO correcto!", compute:() => "15" },
    "ex-15": { type:"formula", regex:/^=AÑO\(.+\)|^=MES\(.+\)|^=DIA\(.+\)|^=FECHA\(.+\)/i, instruccion:"Escribe =AÑO(A2).", exito:"¡Año extraído!", compute:() => "2026" },
    "ex-20": { type:"formula", regex:/^=FILTRAR\(.+\)$/i, instruccion:"Escribe =FILTRAR(A2:C6,C2:C6>500,\"Sin resultados\")", exito:"¡FILTRAR correcto!", compute:() => "Luis, Pedro" },
    "ex-21": { type:"formula", regex:/^=ORDENAR\(UNICOS\(.+\)\)|^=UNICOS\(.+\)/i, instruccion:"Escribe =ORDENAR(UNICOS(A2:A6)).", exito:"¡UNICOS+ORDENAR correcto!", compute:() => "Bogotá, Cali, Medellín" },
    "ex-22": { type:"formula", regex:/^=LET\(.+\)$/i, instruccion:"Escribe =LET(tasa,1.19,base,SUMA(B2:B5),base*tasa)", exito:"¡LET correcto!", compute:() => "2,082.50" },
    
    // BI (Terminal)
    "ex-23": { type:"terminal", regex:/Table\.(SelectRows|Buffer)\(/i, instruccion:"Escribe: Table.SelectRows(Source, each [Region]=\"Norte\")", exito:"¡Power Query M correcto!" },
    "ex-24": { type:"terminal", regex:/slicer|segmentacion|segmentación|tabla dinámica|pivot/i, instruccion:"Explica el flujo: 1. Pivot  2. Slicer  3. Conexiones", exito:"¡BI Workflow correcto!" },

    // SQL Micro-Labs
    "sq-1":  { type:"terminal", regex:/SELECT\s+DISTINCT\s+user_id\s+FROM\s+logs/i, instruccion:"SELECT DISTINCT user_id FROM logs LIMIT 100;", exito:"¡Correcto!", compute: () => "Result: [101, 102] (2 rows)" },
    "sq-3":  { type:"terminal", regex:/INNER\s+JOIN\s+orders/i, instruccion:"SELECT u.nombre, o.total FROM users u INNER JOIN orders o ON u.id = o.user_id;", exito:"¡JOIN correcto!", compute: () => "Result: [Ana: 700, Luis: 0]" },
    "sq-7":  { type:"terminal", regex:/SUM\(\w+\).+GROUP\s+BY\s+region/i, instruccion:"SELECT region, SUM(monto) FROM ventas GROUP BY region;", exito:"¡Agregación correcta!", compute: () => "Result: [Norte: 2300, Sur: 1200]" },
    "sq-13": { type:"terminal", regex:/WITH\s+activos\s+AS/i, instruccion:"Escribe un CTE 'activos' filtrando status='active'.", exito:"¡CTE correcto!", compute: () => "CTE 'activos' definida." },

    // Python Micro-Labs (Completo)
    "py-1":  { type:"terminal", regex:/clientes\s*=\s*[\[{]/, instruccion:"Define la lista 'clientes' con IDs 1 y 2.", exito:"¡Estructura definida!", compute: () => "Memory: clientes <list> en cache." },
    "py-4":  { type:"terminal", regex:/\.apply\(lambda/, instruccion:"Usa .apply(lambda x: 'VIP' if x > 1000 else 'Std').", exito:"¡Lambda segmentada!", compute: () => "df['seg'] = ['VIP', 'Std']" },
    "py-6":  { type:"terminal", regex:/df\.describe\(\)/, instruccion:"Ejecuta df.describe() para ver la estadística.", exito:"¡Describe generado!", compute: () => "count: 3.0, mean: 846.6, std: 377.2" },
    "py-8":  { type:"terminal", regex:/df\.dropna\(subset=\['email'\]\)/, instruccion:"df.dropna(subset=['email'], inplace=True)", exito:"¡NaAs eliminados!", compute: () => "Rows dropped: 1. Clean dataset ready." },
    "py-12": { type:"terminal", regex:/df\.groupby\(.*\)\.agg\(.*\)/, instruccion:"df.groupby('user_id').agg({'monto':'sum'})", exito:"¡Agregación Pandas OK!", compute: () => "Result: {101: 1700, 102: 450}" },
    "py-13": { type:"terminal", regex:/pd\.merge\(.*\)/, instruccion:"pd.merge(clientes, ordenes, on='id', how='left')", exito:"¡Datasets unidos!", compute: () => "Combined Dataframe (X rows)" },
    "py-14": { type:"terminal", regex:/sns\.heatmap\(.*\)/, instruccion:"sns.heatmap(df.corr(), annot=True)", exito:"¡Mapa de calor generado!", compute: () => "Render: Matriz de Correlación [vibrante]" },
    
    // GitHub Master (Completo)
    "gh-1":  { type:"terminal", regex:/git\s+config\s+--global/i, instruccion:"git config --global user.name \"Tu Nombre\"", exito:"¡Identidad establecida!" },
    "gh-4":  { type:"terminal", regex:/git\s+add\s+\.?/, instruccion:"git add . para pasar archivos a Staging.", exito:"¡Archivos preparados!" },
    "gh-5":  { type:"terminal", regex:/git\s+commit\s+-m/i, instruccion:"git commit -m \"feat: mensaje atómico\"", exito:"¡Commit inmutable creado!" },
    "gh-6":  { type:"terminal", regex:/git\s+(push|pull)/i, instruccion:"git push origin main para sincronizar la nube.", exito:"¡Sincronización exitosa con Origin!" },
    "gh-8":  { type:"terminal", regex:/PR|Merge|Draft/i, instruccion:"Explica el flujo PR: Push → Open PR → Review → Merge.", exito:"¡Flujo de colaboración dominado!" },
    
    // Integración (Completo)
    "in-1":  { type:"terminal", regex:/xl\(.*\)|describe/i, instruccion:"Carga Tabla1 con xl() y aplica .describe().", exito:"¡Excel-Python interoperable!", compute: () => "Pandas read range OK. Statistics generated." },
    "in-2":  { type:"terminal", regex:/\.to_sql\(.*\)/, instruccion:"df.to_sql('reporte', engine, if_exists='replace')", exito:"¡ETL Pipeline cerrado!", compute: () => "Database WRITE success. 350 rows affected." }
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
        
        // Inyectar contexto de Micro-Lab si existe
        const preview = document.getElementById('modal-context-preview');
        const ctx = MICRO_LAB_CONTEXTS[itemId];
        
        if (ctx) {
            preview.innerHTML = `<span class="context-title">${ctx.title}</span>${ctx.html}`;
            preview.classList.remove('hidden');
        } else {
            preview.innerHTML = '';
            preview.classList.add('hidden');
        }
        
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
        const result = currentExercise.compute ? currentExercise.compute(val) : '';
        fb.innerHTML   = `> ✅ [SUCCESS] ${currentExercise.exito}${result ? `\n> ${result}` : ''}`;
        fb.style.color = '#10b981';
        saveProgress(currentItemId, 'lib'); 
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
