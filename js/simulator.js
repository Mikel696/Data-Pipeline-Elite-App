/**
 * Data Pipeline Elite App - SIMULATOR CEREBRO
 * Motor de simulación del Laboratorio Maestro
 */

// Estado del simulador
let currentLevel = 0;

// Definición de los Retos
const MASTER_CHALLENGES = [
    // --- SQL CHALLENGES (5) ---
    {
        id: 1,
        modulo: "SQL",
        instruccion: "Nivel 1: Selecciona todos los clientes (tabla `customers`) que se registraron después del 1 de enero de 2024.",
        pista: "Usa SELECT * FROM cust... WHERE signup_date > ...",
        validator: (input) => {
            const regex = /SELECT\s+\*\s+FROM\s+customers\s+WHERE\s+signup_date\s*>\s*['"]2024-01-01['"]/i;
            return regex.test(input);
        }
    },
    {
        id: 2,
        modulo: "SQL",
        instruccion: "Nivel 2: Haz un LEFT JOIN entre `customers` (alias c) y `orders` (alias o) usando la llave `customer_id` y devuelve c.name y o.total_amount.",
        pista: "Asegúrate de usar LEFT JOIN orders o ON c.customer_id = o.customer_id",
        validator: (input) => {
            const regex = /LEFT\s+JOIN\s+orders\s+o\s+ON\s+c\.customer_id\s*=\s*o\.customer_id/i;
            return regex.test(input) && input.toLowerCase().includes('c.name') && input.toLowerCase().includes('o.total_amount');
        }
    },
    {
        id: 3,
        modulo: "SQL",
        instruccion: "Nivel 3: Encuentra clientes inactivos. Busca aquellos donde la diferencia entre '2025-01-01' y su última order_date sea mayor a 365 días usando DATEDIFF.",
        pista: "DATEDIFF('2025-01-01', order_date) > 365",
        validator: (input) => {
            const regex = /DATEDIFF\s*\(\s*['"]2025-01-01['"]\s*,\s*order_date\s*\)\s*>\s*365/i;
            return regex.test(input);
        }
    },
    {
        id: 4,
        modulo: "SQL",
        instruccion: "Nivel 4: Usa una Window Function para asignar un número de fila único (ROW_NUMBER) particionado por `country` y ordenado por `signup_date` DESC.",
        pista: "ROW_NUMBER() OVER (PARTITION BY country ORDER BY signup_date DESC)",
        validator: (input) => {
            const regex = /ROW_NUMBER\s*\(\s*\)\s*OVER\s*\(\s*PARTITION\s+BY\s+country\s+ORDER\s+BY\s+signup_date\s+DESC\s*\)/i;
            return regex.test(input);
        }
    },
    {
        id: 5,
        modulo: "SQL",
        instruccion: "Nivel 5: Calcula el gasto total (SUM) por país agrupando por `country`.",
        pista: "SELECT country, SUM(...) ... GROUP BY country",
        validator: (input) => {
            const regex = /GROUP\s+BY\s+country/i;
            return regex.test(input) && input.toLowerCase().includes('sum(');
        }
    },
    
    // --- PYTHON CHALLENGES (3) ---
    {
        id: 6,
        modulo: "Python",
        instruccion: "Nivel 6: Carga la librería pandas como pd y lee el archivo 'datasets_ejemplo.csv' en un DataFrame llamado df.",
        pista: "import pandas as pd \\n df = pd.read_csv(...)",
        validator: (input) => {
            const str = input.toLowerCase();
            return str.includes("import pandas as pd") && str.includes("pd.read_csv('datasets_ejemplo.csv')") && str.includes("df =");
        }
    },
    {
        id: 7,
        modulo: "Python",
        instruccion: "Nivel 7: Limpia los valores nulos en la columna 'total_spent' reemplazándolos por la media (mean) de esa misma columna.",
        pista: "df['total_spent'].fillna(df['total_spent'].mean(), inplace=True)",
        validator: (input) => {
            const str = input.replace(/\s+/g, '').toLowerCase();
            return str.includes("fillna(df['total_spent'].mean()");
        }
    },
    {
        id: 8,
        modulo: "Python",
        instruccion: "Nivel 8: Aplica un filtro para crear un 'df_churn' solo con los clientes donde 'is_churned' es True.",
        pista: "df_churn = df[df['is_churned'] == True]",
        validator: (input) => {
            const str = input.replace(/\s+/g, '').toLowerCase();
            return str.includes("df[df['is_churned']==true]") || str.includes('df[df["is_churned"]==true]');
        }
    },
    
    // --- EXCEL CHALLENGES (2) ---
    {
        id: 9,
        modulo: "Excel",
        instruccion: "Nivel 9: En lenguaje M (Power Query), filtra la tabla para remover líneas donde 'status' es nulo usando Table.SelectRows.",
        pista: "Table.SelectRows(Origen, each ([status] <> null))",
        validator: (input) => {
            const str = input.toLowerCase().replace(/\s+/g, '');
            return str.includes("table.selectrows") && str.includes("[status]<>null");
        }
    },
    {
        id: 10,
        modulo: "Excel",
        instruccion: "Nivel 10: Escribe una fórmula DAX básica para crear la medida 'Total Ventas', sumando la columna 'total_spent' de la tabla 'Customers'.",
        pista: "Total Ventas = SUM(Customers[total_spent])",
        validator: (input) => {
            const str = input.toLowerCase().replace(/\s+/g, '');
            return str.includes("sum(customers[total_spent])");
        }
    }
];

function loadChallenge(index) {
    if (index >= MASTER_CHALLENGES.length) {
        document.getElementById('mission-title').innerText = "¡Pipeline Completado!";
        document.getElementById('mission-desc').innerText = "¡Has dominado el Ecosistema! Eres un Data Architect Elite.";
        document.getElementById('code-editor').style.display = 'none';
        document.getElementById('submit-btn').style.display = 'none';
        return;
    }
    
    currentLevel = index;
    const challenge = MASTER_CHALLENGES[index];
    
    document.getElementById('mission-title').innerHTML = `<span class="${challenge.modulo.toLowerCase()}-color">[${challenge.modulo}]</span> Misión ${index + 1}/10`;
    document.getElementById('mission-desc').innerText = challenge.instruccion;
    document.getElementById('code-editor').value = "";
    document.getElementById('terminal-feedback').innerHTML = "";
    document.getElementById('terminal-feedback').className = "terminal-feedback";
}

function checkAnswer() {
    const inputCode = document.getElementById('code-editor').value;
    const challenge = MASTER_CHALLENGES[currentLevel];
    const feedback = document.getElementById('terminal-feedback');
    
    if(challenge.validator(inputCode)) {
        // Success
        feedback.innerText = "> [✓] Ejecución exitosa. Validado.";
        feedback.className = "terminal-feedback success";
        saveProgress(challenge.id);
        triggerSuccess();
        
        setTimeout(() => {
            loadChallenge(currentLevel + 1);
        }, 3000);
    } else {
        // Error
        feedback.innerText = `> [x] Syntax Error o Lógica Incorrecta.\n> Pista del sistema: ${challenge.pista}`;
        feedback.className = "terminal-feedback error";
    }
}

function triggerSuccess() {
    const overlay = document.getElementById('success-overlay');
    if(overlay) {
        overlay.classList.remove('hidden');
        setTimeout(() => {
            overlay.classList.add('hidden');
        }, 2500);
    }
}

// Iniciar cargando el nivel adecuado (el primero no completado)
document.addEventListener('DOMContentLoaded', () => {
    // Si la pestana del lab esta activa
    const progress = getProgress();
    loadChallenge(progress.length);
});
