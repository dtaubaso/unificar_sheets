# Unificar Sheets
Sencillo script para unificar varias planillas de Google Sheets en una misma planilla, para graficar o analizar datos.

Útil para quienes tienen muchas planillas de Google Sheets iguales de distintos clientes y necesitan unificar todas en una misma,
evita tener que copiar y pegar todas en una nueva planilla.

## Requisitos:

- Generar un Sheet nuevo en blanco donde se va a pegar este código, en Extensiones/Apps Script
- Generar un Sheet con una columna con el *nombre del cliente* o un identificador, y el id de su correspondiente sheet 
en la siguiente columna (sin comillas).
  + El ID de Google Sheets está en la url, entre dos barras "/": 
por ejemplo, si la url es https://docs.google.com/spreadsheets/d/AAABBBCCC_123456/edit#gid=0, entonces "**AAABBBCCC_123456**" es el ID.
- Los sheets que se van a unificar tienen que tener todos las mismas columnas en el mismo orden, y estar en pestañas con el mismo nombre.
- Cambiar las variables donde corresponda en el script (están comentadas como "Cambiar"), y ejecutar.
