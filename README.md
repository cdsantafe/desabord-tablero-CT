# Dashboard acumulado – Notificación al cliente y adherencia al radio (BEES Deliver)

Lee el export acumulado (Excel, hoja `Export`, o CSV) y muestra los indicadores del dashboard diario
más información operativa.

## Archivos
- `app.py` – la aplicación Streamlit
- `requirements.txt` – dependencias
- `Septiembre.xlsx` – archivo por defecto (si no cargás otro, se usa este)

## Ejecutar
    pip install -r requirements.txt
    streamlit run app.py

## Desplegar en Streamlit Community Cloud
1. Subí los tres archivos a la raíz de un repositorio de GitHub.
2. En share.streamlit.io creá la app apuntando a `app.py`.
3. Para otros meses, cargá los Excel desde el dashboard (se pueden subir varios a la vez y se acumulan).

## Reglas de cálculo
- Se evalúan solo visitas con hora de arribo (`FECHA_INICIO_VISITA`).
- Notificación > 2 min / > 5 min: arribo − `FECHA_NOTIFICACION` ≥ 120 s / ≥ 300 s (sin notificación = no cumple).
- Adherencia al radio: `FLAG_RADIO_INICIO` = 1, solo con ruta iniciada (`FECHA_INICIO_VIAJE`) y arribo.
- Tiempo de visita: `FECHA_FIN_VISITA` − `FECHA_INICIO_VISITA`; se puede ver por chofer y por cliente (con búsqueda) y se excluyen las visitas por encima de un tope en minutos configurable.
- Chofer = `CHOFER BEES`. Las horas del export se toman como hora local.
- Si una visita (planilla + cliente) aparece en varios archivos, se usa la del último.
