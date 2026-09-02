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


## v16 - Porcentajes exactos de la planilla
- Fase 1 se muestra como 13%.
- Fase 2 se muestra como 4%.
- Fase 3 se muestra como 2%.
- Fase 4 se muestra como 0%.
- El Avance General se recalcula desde esos cuatro valores oficiales.
- Resultado general esperado: 4,8%.
- La comparación semanal guarda estos mismos valores oficiales.


## v17 - Eliminación de filas A-B / Todos
Se retiraron de FASE1, FASE2, FASE3 y FASE4 todas las filas resumen donde:
- Torre = A-B
- Departamento = Todos

Porcentajes recalculados desde los registros individuales:
- Fase 1: 13%
- Fase 2: 4%
- Fase 3: 2%
- Fase 4: 0%
- Avance General: 4,8%


## v18 - Perfiles por fase
- Administrador: acceso total.
- Editor FASE1: solo Fase 1.
- Editor FASE2: solo Fase 2.
- Editor FASE3: solo Fase 3.
- Editor FASE4: solo Fase 4.
- Visor: todas las fases, solo lectura.

Las claves se cambian cuando quieras en Streamlit → Manage app → Secrets.


## v19 - Gráficos uniformes y menú claro
- Todos los gráficos principales usan una altura uniforme de 400 px.
- Se evita el crecimiento/zoom visual al cambiar entre pestañas.
- Menú lateral actualizado a un azul más claro.
- Se mantienen perfiles de acceso, Supabase, comparación semanal, ruta crítica y porcentajes.


## v20 - Exportación conservando formato original
Al exportar:
- se usa el mismo CONTROL_FASES_SFCO211.xlsx como plantilla;
- se conservan hojas, formato, colores, tamaños y fórmulas;
- se incorporan los avances actuales almacenados en Supabase;
- Excel queda configurado para recalcular fórmulas al abrir.


## v21 - Excel depurado y Avance Semanal vinculado
Se eliminaron las pestañas:
- IMPRIMIR
- Hoja2
- PROJECT TA Y TB KARINA
- EXEL KARINA
- PROJECT OBRA TA Y TB
- GRAFICOS TA Y TB
- PROG TERMINACIONES

Además:
- FASE1–FASE4 se muestran uniformemente como porcentajes al descargar.
- RESUMEN se muestra uniformemente como porcentaje.
- AVANCE SEMANAL 1 toma el avance de cada partida desde la Fase correspondiente y por Piso.
- Las partidas antiguas sin equivalente actual quedan en blanco en vez de mostrar errores.


## v22 - Exportación estable
- Fuerza una plantilla depurada al cambiar de versión.
- Solo conserva: AVANCE SEMANAL 1, RESUMEN, FASE1, FASE2, FASE3, FASE4 y SEGUIMIENTO R.CRITICA.
- Elimina componentes de tablas dinámicas que ya no tienen hojas asociadas.
- Exporta aplicando Supabase directamente dentro del paquete XLSX, sin regrabar todo el libro.
- Evita el mensaje de Microsoft Excel indicando que encontró contenido con problemas.

## v26 – Sincronización, Ruta Crítica visual y Gantt oficial
- Sincronización automática de avances desde Supabase entre computadores (TTL corto + botón manual de sincronización).
- Ruta Crítica en matriz por fase/partida y Piso 1–9, alimentada desde Fases completas.
- Carta Gantt con fechas verificadas contra Libro2.xlsx enviado por el usuario.
- Semáforo de 3 tonos: verde 100%, amarillo 50–99%, naranjo 0–49%.
- Vista de semáforo adicional para Fases completas sin afectar la edición de datos.

## v27 - Corrección avance oficial y ruta crítica completa
- Los KPI principales usan el último registro disponible de Comparación semanal como corte oficial.
- La pantalla Comparación semanal conserva un bloque separado de avance actual en vivo.
- Ruta crítica muestra todas las partidas de cada fase, sin selección parcial: FASE1=20, FASE2=15, FASE3=21, FASE4=12.
- El detalle por partida sigue leyendo los cambios de la base online/Supabase.

## v30 - Porcentajes y atraso dinámico
- Comparación semanal formateada explícitamente como porcentaje.
- Variación mostrada en puntos porcentuales (pp).
- Carta Gantt incorpora Fecha de revisión.
- Días de atraso = días calendario posteriores al término planificado cuando el progreso es menor a 100%.
- Estado de plazo: Completado / En plazo / No iniciado / Atrasado.


## v32
- Carta Gantt: Fase 4 / Piso 9 termina el 30-08-2027.
- Paneles Ruta Crítica y Carta Gantt ampliados para pantallas de escritorio.

## v33 — Corrección correlativa Carta Gantt
- Fase 1 mantiene las fechas base confirmadas.
- Fases 2, 3 y 4 corrigen el desfase de 2 días según la tabla de programación compartida.
- La secuencia se aplica a todos los pisos 1–9.
- Fase 4 / Piso 9 conserva término 30-08-2027 (inicio 02-08-2027, 28 días según la lógica actual de la app).

## v35 · Autoguardado verificado
- Cada cambio de avance se guarda automáticamente en Supabase.
- El guardado se verifica leyendo nuevamente el registro persistido.
- La app muestra confirmación o error con fecha/hora y actividad.
- Fases completas ya no requiere botón Guardar: cada celda modificada se autoguarda.
- Se eliminó el caché de carga de fases para reflejar cambios de otros computadores con mayor rapidez.
- Se incluye botón para verificar conexión y último cambio guardado por fase.


## v40 – corrección de lectura completa
- `phase_updates` se lee con paginación completa (no solo el primer bloque de Supabase).
- El avance se reconstruye por `phase + excel_row + activity`.
- El histórico semanal y la auditoría también usan lectura paginada.
- Se muestra un diagnóstico con el total de registros online cargados.
- No requiere volver a ingresar los avances del sábado: si existen en `phase_updates`, se aplican automáticamente.


## v41 – Alternativa 2
- Ruta Crítica y Carta Gantt se muestran una debajo de la otra, a todo el ancho.
- Orden fijo: Fase 1 Piso 1–9, Fase 2 Piso 1–9, Fase 3 Piso 1–9, Fase 4 Piso 1–9.
- Cada piso termina el último día hábil del mismo mes de inicio.
- Fase 1 Piso 1 inicia 01-07-2026.

## v42 - Avance por Departamento
- Nueva pestaña `🏙️ Avance por Departamento`.
- Elevación dinámica por Torre y Piso.
- Cada departamento muestra F1, F2, F3 y F4 calculadas desde Supabase.
- Semáforo: verde 100%, amarillo 50-99.9%, naranjo 0-49.9%.
- No modifica ni duplica los avances existentes.


## v44
Corrección de la vista Avance por Departamento: Torre B, Piso 7 ahora incluye departamentos 701 a 721. Los avances continúan vinculados por fila/partida a Supabase.

## v46 · Base de recuperación Excel → Supabase
- Incluye como plantilla oficial el Excel `CONTROL_FASES_SFCO211_RECUPERADO_BASE_ONLINE.xlsx` modificado por el usuario.
- En **Importar / Exportar** el Administrador puede validar y sincronizar todas las partidas FASE1–FASE4 hacia `phase_updates`.
- La sincronización usa `upsert` con la clave única `(phase, excel_row, activity)`, por lo que sobrescribe el porcentaje oficial sin duplicar registros.
- No se vacía `phase_updates` antes de escribir; primero se guarda todo y luego se verifica contra una lectura paginada completa.
- Las celdas vacías de partidas se consideran 0% al convertir el Excel en base oficial.
- `updated_by` queda identificado como `RECUPERACION EXCEL · <usuario>`.

## v47 - Sincronización Excel → Supabase visible
- La confirmación y el botón de actualización aparecen inmediatamente debajo del bloque de recuperación.
- La validación del Excel se ejecuta solo al pulsar el botón.
- El botón queda deshabilitado hasta marcar la confirmación.
- El proceso sigue usando upsert y verificación contra Supabase; no suma ni duplica porcentajes.
