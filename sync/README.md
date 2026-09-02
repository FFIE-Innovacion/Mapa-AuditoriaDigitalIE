# Sincronización de avances (Matriz FFIE → dashboards)

`sync_avances.py` cruza la **última visita ejecutada** de cada sede (Sheet de visitas, pestaña `BASE_PM`)
con el **% de avance de obra** de la Matriz FFIE en SharePoint (pestaña `BASE_PM`, fila del mes de esa visita)
y escribe `data/avances.json`. El mapa y el panel gerencial leen ese JSON; si no está disponible,
vuelven al % del Sheet.

El workflow `.github/workflows/sync-avances.yml` lo ejecuta todos los días a las 6:00 am (Bogotá)
y cuando se pulsa "Run workflow". Secrets necesarios: `VISITAS_CSV_URL` y `MATRIZ_URL`
(enlace anónimo de solo lectura de la Matriz). Alternativas: `MATRIZ_CSV_URL` o subir `data/matriz_ffie.xlsx`.

Prueba local:
```
pip install openpyxl
python sync/sync_avances.py --visitas "Matriz FFIE.xlsx" --matriz "Matriz FFIE.xlsx" --out data/avances.json
```
