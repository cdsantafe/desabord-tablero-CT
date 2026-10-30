# Dashboard acumulado – Notificación al cliente y adherencia al radio (BEES Deliver)

Lee el export acumulado (Excel, hoja `Export`, o CSV) y muestra los indicadores del dashboard diario
más información operativa.

## Archivos
- `app.py` – la aplicación Streamlit
- `requirements.txt` – dependencias
- `datos/` – carpeta con los Excel/CSV del export; **todos los archivos de esta carpeta se acumulan automáticamente**
  (ya incluye `Septiembre.xlsx`)

## Ejecutar
    pip install -r requirements.txt
    streamlit run app.py

## Desplegar en Streamlit Community Cloud
1. Subí `app.py`, `requirements.txt`, `README.md` y la carpeta `datos/` a un repositorio de GitHub.
2. En share.streamlit.io creá la app apuntando a `app.py`.

## Agregar más meses
- **Fijo (recomendado):** copiá el nuevo Excel en `datos/` (nombre libre, por ejemplo `2026-10.xlsx`),
  subilo a GitHub y la app lo suma sola. Streamlit Cloud se actualiza al hacer el commit.
- **Temporal:** cargalo desde el dashboard; se suma a los de la carpeta mientras dure la sesión.
- Los archivos se leen en orden alfabético y los cargados desde el dashboard van al final. Si una visita
  aparece en dos archivos, queda la del último; conviene nombrarlos con fecha (`2026-09.xlsx`, `2026-10.xlsx`).

## Reglas de cálculo
- Se evalúan solo visitas con hora de arribo (`FECHA_INICIO_VISITA`).
- Notificación > 2 min / > 5 min: arribo − `FECHA_NOTIFICACION` ≥ 120 s / ≥ 300 s (sin notificación = no cumple).
- Adherencia al radio: `FLAG_RADIO_INICIO` = 1, solo con ruta iniciada (`FECHA_INICIO_VIAJE`) y arribo.
- Tiempo de visita: `FECHA_FIN_VISITA` − `FECHA_INICIO_VISITA`; se puede ver por chofer y por cliente (con búsqueda) y se excluyen las visitas por encima de un tope en minutos configurable.
- Chofer = `CHOFER BEES`. Las horas del export se toman como hora local.
- Si una visita (planilla + cliente) aparece en varios archivos, se usa la del último.
