# CONTROL FASES SFCO211 — PRO v9

Novedades:
- Nueva pestaña "Comparación semanal".
- Registro de avance todos los viernes.
- Guarda: Avance General + Fase 1 + Fase 2 + Fase 3 + Fase 4 + usuario.
- Gráfico de líneas para comparar evolución semanal.
- Comparación contra semana anterior en puntos porcentuales.
- Historial almacenado en hoja HISTORIAL_SEMANAL del Excel de trabajo.
- Administrador puede registrar manualmente fuera del viernes para pruebas/regularización.
- Diseño mejorado para celular.
- Se mantiene login con usuarios y claves.
- Se mantienen Dashboard, gráficos por fase, edición de avances y Ruta Crítica.

IMPORTANTE:
En Streamlit Community Cloud el archivo de trabajo vive dentro de la sesión.
Para que múltiples usuarios compartan un historial permanente sin depender de descargar/subir Excel,
la siguiente evolución recomendada es conectar una base de datos en línea (por ejemplo Supabase/PostgreSQL).

Archivos a reemplazar en GitHub:
- app.py
- requirements.txt
- CONTROL_FASES_SFCO211.xlsx
- logo_san_francisco.png
- README.md

La carpeta .streamlit contiene solo una plantilla de Secrets; las claves reales deben configurarse
directamente en Streamlit Community Cloud.
