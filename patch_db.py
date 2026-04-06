import json
import re

with open('g:/Office/LEARN/EXCEL/Data-Pipeline-Elite-App/js/backend.js', 'r', encoding='utf-8') as f:
    content = f.read()

# Extract JSON
match = re.search(r'let functionDatabase = (\{.*?\});\n\nfunction fetchDatabase', content, flags=re.DOTALL)
if not match:
    print("Could not find db")
    exit(1)

db_str = match.group(1)
# fix some trailing commas or keys if needed
db = json.loads(db_str)

# Inject exercices
for cat in db:
    for item in db[cat]:
        # Generate exercise
        cmd = item['codigo'].split('\n')[0]
        # Basic smart extraction for regex
        regex_val = ".*"
        if cat == 'excel':
            regex_val = r"(?i)" + re.escape(cmd.split('(')[0].replace('=', ''))
            if '=' not in cmd: regex_val = r"(?i).*"
        elif cat == 'sql':
            regex_val = r"(?i)" + cmd.split(' ')[0]
        elif cat == 'python':
            if '.' in cmd: regex_val = r"(?i)" + re.escape(cmd.split('.')[0])
            else: regex_val = r"(?i).*"
        elif cat == 'github':
            regex_val = r"git\s+[a-z]+"
        
        item['ejercicio'] = {
            "instruccion": f"El líder técnico ha solicitado implementar arquitectura para: {item['titulo']}. Demuestra la sintaxis correcta en la terminal inferior.",
            "validador_regex": regex_val,
            "mensaje_exito": f"¡Pipeline validado! {item['titulo']} operando a máxima capacidad."
        }
        
        # overriding specifically some for wow-factor
        if item['id'] == 'gh-2':
            item['ejercicio'] = {"instruccion": "Clona el repositorio 'https://repo.git' en tu máquina local.", "validador_regex": r"git clone https://repo\.git", "mensaje_exito": "¡Repo clonado satisfactoriamente!"}
        if item['id'] == 'gh-4':
            item['ejercicio'] = {"instruccion": "Agrega TODOS los cambios modificados al Muelle de Carga (Staging Area).", "validador_regex": r"git add \.", "mensaje_exito": "¡Trackers activados!"}
        if item['id'] == 'ex-10':
            item['ejercicio'] = {"instruccion": "Haz un BUSCARV buscando la celda A2, en matriz Rango, devolviendo la 2da columna, coincidencia FALSO.", "validador_regex": r"(?i)=\s*BUSCARV\s*\(\s*A2\s*,\s*Rango\s*,\s*2\s*,\s*FALSO\s*\)", "mensaje_exito": "¡VLOOKUP ejecutado!"}
        if item['id'] == 'sq-1':
            item['ejercicio'] = {"instruccion": "Escribe una consulta para extraer registros DISTINTOS de 'user_id' en 'logs'.", "validador_regex": r"(?i)SELECT\s+DISTINCT\s+user_id\s+FROM\s+logs", "mensaje_exito": "¡Matriz de usuarios purgada con DISTINCT!"}
        if item['id'] == 'py-8':
            item['ejercicio'] = {"instruccion": "Elimina las filas nulas del DataFrame 'df'.", "validador_regex": r"(?i)df\.dropna", "mensaje_exito": "¡Missing values purgados en vector cruzado!"}

new_db_str = json.dumps(db, ensure_ascii=False, indent=4)
new_content = content.replace(db_str, new_db_str)

with open('g:/Office/LEARN/EXCEL/Data-Pipeline-Elite-App/js/backend.js', 'w', encoding='utf-8') as f:
    f.write(new_content)

print("DB Patched with micro-labs!")
