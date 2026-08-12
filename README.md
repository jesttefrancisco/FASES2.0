# CONTROL FASES SFCO211 — PRO v8

Cambios:
- Eliminadas del menú: Terminaciones, Resumen y Avance semanal.
- Ruta Crítica revisada y normalizada a porcentaje 0%–100%.
- Ruta Crítica incluye avance promedio, gráfico por piso, gráfico por torre y tabla completa.
- Los valores 0–1 de Ruta Crítica se convierten a porcentaje automáticamente.
- Inicio de sesión con usuarios y claves mediante Streamlit Secrets.
- Incluye ejemplo `.streamlit/secrets.toml.example`.
- NO se guardan contraseñas reales en GitHub.

Usuarios:
Configura tus usuarios en Streamlit Community Cloud → App settings → Secrets.
Usa el contenido de `.streamlit/secrets.toml.example` como plantilla.

Archivos principales a reemplazar:
- app.py
- requirements.txt
- CONTROL_FASES_SFCO211.xlsx
- logo_san_francisco.png
- README.md

Además puedes subir:
- .streamlit/secrets.toml.example (solo ejemplo, no claves reales)
