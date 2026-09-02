# Sincronización de avances (Matriz FFIE → dashboards)

`sync_avances.py` cruza la **última visita ejecutada** de cada sede (Sheet de visitas, pestaña `BASE_PM`)
con el **% de avance de obra** de la Matriz FFIE en SharePoint (pestaña `BASE_PM`, fila del mes de esa visita)
y escribe `data/avances.json`. El mapa y el panel gerencial leen ese JSON; si no está disponible,
vuelven al % del Sheet.

El workflow `.github/workflows/sync-avances.yml` lo ejecuta todos los días a las 6:00 am (Bogotá)
y cuando se pulsa "Run workflow".

Cómo llega la Matriz al job (en orden de preferencia):
1. **Entra ID** (`descargar_matriz.py`): secrets `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`
   y `SP_FILE_PATH` (o `SP_FILE_NAME`). Lo que hay que pedirle a TI está en `SOLICITUD_TI_EntraID.md`.
2. Enlace anónimo de SharePoint: secret `MATRIZ_URL`.
3. Pestaña en Google Sheets: secret `MATRIZ_CSV_URL`.
4. Archivo `data/matriz_ffie.xlsx` subido al repo.

`VISITAS_CSV_URL` es opcional (por defecto usa el CSV publicado del Sheet).

Prueba local:
```
pip install openpyxl
python sync/sync_avances.py --visitas "Matriz FFIE.xlsx" --matriz "Matriz FFIE.xlsx" --out data/avances.json
```
