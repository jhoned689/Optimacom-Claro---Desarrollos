# Optimacom Claro — Desarrollos

Seguimiento de automatizaciones y entregas del equipo de desarrollo Optimacom / Claro.

**Cadencia de entregas:** miércoles y viernes de cada semana.

## Vista rápida — junio 2026

| % | Proyecto | Estado |
|---|----------|--------|
| 90% | Gestión firmas actas liquidación | App lista — pendiente aval producción |
| 85% | Auditoría PACC-FACC Servicios | Pruebas en curso |
| 50% | Auditoría PACC-FACC Hardware | Demo + pruebas |
| 15% | Validación HW línea a línea (IA) | Investigación / modelo |

## Archivos

| Archivo | Descripción |
|---------|-------------|
| [`PBI Desarrollos.xlsx`](PBI%20Desarrollos.xlsx) | Fuente principal (alimenta Power BI) |
| [`proyectos/`](proyectos/README.md) | **Una carpeta por desarrollo** con actividades dentro |
| [`data/ejecucion-desarrollo.csv`](data/ejecucion-desarrollo.csv) | Detalle plano (todas las filas) |
| [`data/proyectos.json`](data/proyectos.json) | Proyectos + actividades en JSON |
| [`docs/PROXIMOS-ENTREGABLES.md`](docs/PROXIMOS-ENTREGABLES.md) | Próximas entregas mié/vie |
| [`docs/GITHUB.md`](docs/GITHUB.md) | Guía de actualización en GitHub |

## Actualizar cronograma

```powershell
cd "Desarrollo - Cronograma y ejecucion"
python actualizar_cronograma_jun2026.py
git add .
git commit -m "Actualización cronograma"
git push
```

Esto actualiza el Excel, regenera los CSV y mantiene fechas en miércoles/viernes.

## Estructura del Excel

- **EJECUCION DESAROLLO** — actividades por fase, fecha, estado y % avance
- **DESARROLLOS LISTOS** — resumen para tablero PBI
- **DATA DESARROLLO** — registro de solicitudes
- **LB** — estados del ciclo de vida

## Responsable

Jhon Medina — desarrollo y automatización

Última actualización: **24 de junio de 2026**
