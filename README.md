# CONTROL FASES SFCO211 — PRO v7

Corrección crítica:
- El avance ya NO depende de la fórmula/caché de "% Avance Real Depto" del Excel.
- Se calcula directamente desde todas las partidas visibles en "Fases completas".
- Primero se calcula el avance de cada departamento.
- Luego se obtiene el avance general de cada Fase.
- El Dashboard y todos los gráficos usan esos mismos datos.
- FASE1, FASE2, FASE3 y FASE4 quedan en escala visible 0%–100%.
- Al editar y guardar una partida, el avance se recalcula inmediatamente.
- Se mantienen logo, gráficos por piso, torre y partida.

Archivos a reemplazar en GitHub:
- app.py
- requirements.txt
- CONTROL_FASES_SFCO211.xlsx
- logo_san_francisco.png
- README.md
