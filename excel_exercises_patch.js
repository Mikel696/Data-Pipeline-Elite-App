// EXCEL EXERCISES — Replacement module for backend.js
// Paste this as the "excel" array key inside functionDatabase

const EXCEL_DB = [
    // ── ATAJOS ─────────────────────────────────────────────────────────────
    { "id":"ex-1",  "titulo":"Ctrl Shortcuts", "categoria":"Atajos Fundamentales",
      "descripcion":"Ctrl+Espacio selecciona la columna entera. Ctrl+Fin salta a la última celda con dato. Ctrl+1 abre Formato de Celdas. Ctrl+Shift+$ aplica formato moneda.",
      "codigo":"Ctrl + Espacio  → Columna entera\nCtrl + 1        → Formato de Celdas\nCtrl + Shift+$  → Formato moneda\nCtrl + Fin      → Última celda con datos" },

    { "id":"ex-2",  "titulo":"Shift / Alt Shortcuts", "categoria":"Atajos Fundamentales",
      "descripcion":"Shift+Espacio selecciona la fila entera. Alt+= inserta SUMA automática. Alt+H+A+C centra celdas. F2 entra en modo edición de celda activa.",
      "codigo":"Shift + Espacio  → Fila entera\nAlt + =          → Auto-SUMA\nF2               → Editar celda\nAlt+H+A+C        → Centrar" },

    // ── MATEMÁTICAS ─────────────────────────────────────────────────────────
    { "id":"ex-3",  "titulo":"SUMA & PROMEDIO", "categoria":"Matemáticas Estadísticas",
      "descripcion":"SUMA() agrega todos los valores del rango. PROMEDIO() calcula la media aritmética ignorando celdas vacías. Son las funciones más usadas en finanzas y control.",
      "codigo":"=SUMA(B2:B6)        → Total\n=PROMEDIO(B2:B6)   → Media" },

    { "id":"ex-4",  "titulo":"SUMAR.SI.CONJUNTO", "categoria":"Matemáticas Estadísticas",
      "descripcion":"Suma celdas que cumplen MÚLTIPLES criterios simultáneos. Imprescindible para cruzar región + producto + estado en un solo cálculo sin tablas dinámicas.",
      "codigo":"=SUMAR.SI.CONJUNTO(C2:C6, A2:A6, \"Norte\", B2:B6, \"Laptop\")" },

    { "id":"ex-5",  "titulo":"CONTAR.SI.CONJUNTO", "categoria":"Matemáticas Estadísticas",
      "descripcion":"Cuenta registros que cumplen varios criterios al mismo tiempo. Ejemplo real: contar clientes VIP activos en Colombia.",
      "codigo":"=CONTAR.SI.CONJUNTO(A2:A6,\"CO\", B2:B6,\"VIP\", C2:C6,\">500\")" },

    { "id":"ex-6",  "titulo":"MAX.SI.CONJUNTO & MIN.SI", "categoria":"Matemáticas Estadísticas",
      "descripcion":"Encuentran el máximo o mínimo dentro de un subconjunto filtrado. Clave para rankings por categoría sin alterar el dataset.",
      "codigo":"=MAX.SI.CONJUNTO(C2:C6, A2:A6, \"Norte\")\n=MIN.SI.CONJUNTO(C2:C6, B2:B6, \"Laptop\")" },

    { "id":"ex-7",  "titulo":"REDONDEAR & ENTERO", "categoria":"Matemáticas Estadísticas",
      "descripcion":"REDONDEAR controla decimales para presentaciones. ENTERO trunca hacia el entero inferior. Esencial en facturas, precios y KPIs financieros.",
      "codigo":"=REDONDEAR(A2*1.19, 2)    → IVA redondeado a 2 decimales\n=ENTERO(A2/1000)*1000   → Redondeo a miles" },

    // ── LÓGICAS ─────────────────────────────────────────────────────────────
    { "id":"ex-8",  "titulo":"SI() anidado", "categoria":"Lógicas y Condicionales",
      "descripcion":"SI() anidado crea árboles de decisión. Fundamental para clasificar clientes (VIP/Medio/Básico) o categorizar estados de pago.",
      "codigo":"=SI(B2>1000,\"VIP\",SI(B2>500,\"Medio\",\"Básico\"))" },

    { "id":"ex-9",  "titulo":"SI.CONJUNTO (IFS)", "categoria":"Lógicas y Condicionales",
      "descripcion":"Reemplazo moderno de SI() anidados. Evalúa múltiples condiciones en orden, devolve el primer resultado verdadero. Más legible que SI() anidado a 4 niveles.",
      "codigo":"=SI.CONJUNTO(B2>1000,\"VIP\", B2>500,\"Medio\", B2>0,\"Básico\", VERDADERO,\"Sin ventas\")" },

    { "id":"ex-10", "titulo":"SIERROR & SINO.DISPONIBLE", "categoria":"Lógicas y Condicionales",
      "descripcion":"SIERROR captura cualquier error (#N/A, #¡VALOR!, #¡DIV/0!). SINO.DISPONIBLE (IFNA) captura solo #N/A, dejando pasar otros errores para detectarlos.",
      "codigo":"=SIERROR(BUSCARV(A2,Tabla,2,0), \"No encontrado\")\n=SINO.DISPONIBLE(BUSCARX(A2,E:E,F:F), 0)" },

    // ── TEXTO ───────────────────────────────────────────────────────────────
    { "id":"ex-11", "titulo":"IZQUIERDA / DERECHA / EXTRAE", "categoria":"Funciones de Texto",
      "descripcion":"Extraen substrings por posición. IZQUIERDA extrae desde el inicio, DERECHA desde el final, EXTRAE desde cualquier posición. Clave para parsear códigos como 'CO-2024-001'.",
      "codigo":"=IZQUIERDA(A2,2)        → Las 2 primeras letras\n=DERECHA(A2,3)          → Las 3 últimas letras\n=EXTRAE(A2,4,4)         → 4 chars desde posición 4" },

    { "id":"ex-12", "titulo":"CONCATENAR / UNIRCADENAS", "categoria":"Funciones de Texto",
      "descripcion":"CONCATENAR une celdas y texto. UNIRCADENAS (TEXTJOIN) agrega un delimitador y puede ignorar vacíos, ideal para combinar listas de ciudades o nombres.",
      "codigo":"=CONCATENAR(A2,\" \",B2)           → \"Ana García\"\n=UNIRCADENAS(\", \",VERDADERO,A2:A6)  → \"Ana, Luis, María\"" },

    { "id":"ex-13", "titulo":"TEXTO() — Formatear valores", "categoria":"Funciones de Texto",
      "descripcion":"Convierte un número o fecha en texto con formato personalizado. Esencial para crear reportes donde el número debe aparecer como moneda, porcentaje o fecha legible.",
      "codigo":"=TEXTO(A2,\"$#,##0.00\")     → \"$1,250.00\"\n=TEXTO(B2,\"dd/mm/aaaa\")  → \"06/04/2026\"\n=TEXTO(C2,\"0.0%\")        → \"85.5%\"" },

    { "id":"ex-14", "titulo":"LARGO / HALLAR / SUSTITUIR", "categoria":"Funciones de Texto",
      "descripcion":"LARGO cuenta caracteres. HALLAR encuentra posición de un substring. SUSTITUIR reemplaza ocurrencias. Cadena completa de limpieza para datos importados de sistemas.",
      "codigo":"=LARGO(A2)                     → 15 (caracteres)\n=HALLAR(\"@\",A2)                → 5 (posición del @)\n=SUSTITUIR(A2,\"-\",\"\")          → Elimina guiones" },

    // ── FECHAS ──────────────────────────────────────────────────────────────
    { "id":"ex-15", "titulo":"FECHA / MES / AÑO / DIA", "categoria":"Funciones de Fecha",
      "descripcion":"Extraen componentes de una fecha para agrupar por mes o año en tablas dinámicas. FECHA() construye fechas desde partes independientes.",
      "codigo":"=AÑO(A2)                       → 2026\n=MES(A2)                       → 4\n=DIA(A2)                       → 6\n=FECHA(2026,AÑO(A2),MES(A2))  → Fecha construida" },

    { "id":"ex-16", "titulo":"DIAS.LAB / SIFECHA / HOY", "categoria":"Funciones de Fecha",
      "descripcion":"DIAS.LAB cuenta días hábiles entre fechas (excluye fines de semana). SIFECHA calcula diferencias en días, meses o años. HOY() devuelve la fecha actual dinámica.",
      "codigo":"=HOY()                        → Fecha de hoy dinámica\n=DIAS.LAB(A2,B2)             → Días hábiles entre fechas\n=SIFECHA(A2,HOY(),\"m\")       → Meses de antigüedad" },

    // ── BÚSQUEDA AVANZADA ────────────────────────────────────────────────────
    { "id":"ex-17", "titulo":"BUSCARV (VLOOKUP)", "categoria":"Búsqueda y Referencia",
      "descripcion":"Busca un valor en la primera columna de una tabla y devuelve el de la columna indicada. FALSO obliga coincidencia exacta. El estándar histórico de los cruces de datos.",
      "codigo":"=BUSCARV(A2,$E$2:$G$6,2,FALSO)" },

    { "id":"ex-18", "titulo":"INDICE & COINCIDIR", "categoria":"Búsqueda y Referencia",
      "descripcion":"Combinación superior a BUSCARV: permite buscar hacia la izquierda, es más rápida en tablas grandes y no necesita conocer el número de columna.", 
      "codigo":"=INDICE(C2:C6,COINCIDIR(A9,A2:A6,0))" },

    { "id":"ex-19", "titulo":"BUSCARX (XLOOKUP)", "categoria":"Búsqueda y Referencia",
      "descripcion":"El BUSCARV evolucionado de Microsoft 365. Busca en cualquier dirección, maneja errores nativamente, puede devolver múltiples columnas y hacer búsqueda aproximada avanzada.",
      "codigo":"=BUSCARX(A9,A2:A6,C2:C6,\"No encontrado\",0)" },

    // ── MATRICES DINÁMICAS ───────────────────────────────────────────────────
    { "id":"ex-20", "titulo":"FILTRAR()", "categoria":"Matrices Dinámicas",
      "descripcion":"Devuelve un subconjunto dinámico del rango que cumple la condición. El resultado 'desborda' automáticamente a celdas adyacentes (comportamiento Spill de Excel 365).",
      "codigo":"=FILTRAR(A2:C6,C2:C6>500,\"Sin resultados\")" },

    { "id":"ex-21", "titulo":"UNIQUE() y ORDENAR()", "categoria":"Matrices Dinámicas",
      "descripcion":"UNICOS() elimina duplicados y devuelve valores únicos. ORDENAR() clasifica el resultado. La combinación reemplaza trabajos de horas con tablas dinámicas manuales.",
      "codigo":"=ORDENAR(UNICOS(A2:A100))\n=UNICOS(A2:A6,VERDAD)    → Únicos por fila" },

    { "id":"ex-22", "titulo":"LET() — Variables en fórmulas", "categoria":"Matrices Dinámicas",
      "descripcion":"Permite declarar variables internas dentro de una fórmula para evitar repetición y mejorar la legibilidad de fórmulas complejas. Disponible en Microsoft 365.",
      "codigo":"=LET(\n  tasa, 1.19,\n  base, SUMA(B2:B6),\n  base*tasa\n)" },

    // ── POWER QUERY + PIVOT ──────────────────────────────────────────────────
    { "id":"ex-23", "titulo":"Power Query: Transformaciones M", "categoria":"Power Query (ETL)",
      "descripcion":"El lenguaje M de Power Query permite transformaciones repetibles y auditables. Filtrar en el origen evita bajar millones de filas innecesariamente.",
      "codigo":"// Filtro en origen (Query Folding):\nTable.SelectRows(Source, each [Region] = \"Norte\")\n// Buffer en RAM:\nTable.Buffer(TablaBase)" },

    { "id":"ex-24", "titulo":"Tablas Dinámicas + Slicers", "categoria":"Tablas Dinámicas",
      "descripcion":"Las Tablas Dinámicas resumen millones de filas en segundos. Los Slicers (Segmentadores) permiten al C-Level filtrar interactivamente sin tocar los datos subyacentes.",
      "codigo":"1. Insertar → Tabla Dinámica\n2. Insertar → Segmentación de datos\n3. Clic derecho → Conexiones de informe\n   (conectar a múltiples pivots)" }
];
