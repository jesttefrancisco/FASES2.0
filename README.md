# CONTROL FASES SFCO211 — APP PROFESIONAL

Esta versión reemplaza completamente la base anterior y usa `CONTROL_FASES_SFCO211.xlsx`.

## Cambios principales

- Dashboard profesional.
- Fases detectadas: FASE1, FASE2, FASE3 y FASE4.
- Todos los avances se muestran como porcentaje 0–100%.
- FASE1 se conserva internamente en escala 0–100.
- FASE2, FASE3 y FASE4 se conservan internamente en escala 0–1 para respetar el Excel.
- Editor simple por Torre → Piso → Departamento → Partida.
- El `% Avance Real Depto` se recalcula automáticamente como promedio de las partidas.
- Incluye Resumen, Avance Semanal, Ruta Crítica y Terminaciones.
- Permite descargar el Excel actualizado.

## Actualizar GitHub / Streamlit

Reemplaza en la raíz del repositorio:
- `app.py`
- `requirements.txt`
- `CONTROL_FASES_SFCO211.xlsx`
- `README.md`

En Streamlit deja `app.py` como Main file path.

- Logo oficial de Edificio San Francisco 211 / PAZ integrado en menú lateral y cabecera.
