# CONTROL FASES SFCO211 — PRO v10 SUPABASE

Base online compartida para:
- avances editados de FASE1–FASE4,
- usuario que realizó el cambio,
- historial semanal de los viernes,
- Avance General + Fase 1 + Fase 2 + Fase 3 + Fase 4.

Configuración:
1. Crear proyecto en Supabase.
2. Ejecutar `supabase_setup.sql` en SQL Editor.
3. Copiar Project URL y service_role key.
4. En Streamlit Community Cloud > App settings > Secrets pegar la plantilla
   `.streamlit/secrets.toml.example` y reemplazar los valores.
5. Guardar y reiniciar la app.

No publiques claves reales en GitHub.


## v11 - Corrección Ruta Crítica
- Ruta Crítica ya no depende de la columna calculada del Excel.
- Calcula el avance directamente desde las partidas de SEGUIMIENTO R.CRITICA.
- Convierte valores 0–1 y 0–100 a una sola escala 0%–100%.
- Muestra KPI de avance, pisos, departamentos y departamentos con avance.
- Incluye gráficos por piso, torre y partida.
- Incluye filtros Torre/Piso y detalle por departamento.


## v12 - Diagnóstico Supabase
- Valida formato de URL.
- Valida que la clave comience con sb_secret_.
- Realiza una consulta real a phase_updates.
- Muestra el error exacto dentro de la app si la conexión falla.


## v13
Ruta Crítica corregida sobre la versión con Supabase funcionando.


## v14 - Comparación de avance de los viernes
- Gráfico de líneas con un punto por cada viernes registrado.
- Series: Avance General, Fase 1, Fase 2, Fase 3 y Fase 4.
- Último viernes mostrado como KPI.
- Comparación con viernes anterior en puntos porcentuales.
- Tabla histórica con fecha, porcentajes y usuario.
- Usa el historial online de Supabase cuando está conectado.


## v15 - Avance oficial de fases
Se corrige el cálculo general para coincidir con la hoja RESUMEN:
- Fase 1 ≈ 13.4%
- Fase 2 ≈ 4.0%
- Fase 3 ≈ 2.0%
- Fase 4 = 0.0%
- Avance General = promedio de las cuatro fases oficiales.

La comparación semanal de los viernes guarda estos mismos porcentajes oficiales.
