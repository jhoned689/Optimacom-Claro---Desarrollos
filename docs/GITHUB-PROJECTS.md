# Guía: GitHub Projects + ejecución por desarrollo

## Por qué Projects está vacío

[GitHub Projects](https://github.com/jhoned689/Optimacom-Claro---Desarrollos/projects) **no lee** las carpetas `proyectos/` del repositorio.

| Qué hicimos en el repo | Qué ve Projects |
|------------------------|-----------------|
| Carpetas + README + CSV | ❌ No aparece |
| **Issues** (1 por actividad) | ✅ Sí — son las tarjetas del tablero |
| **Milestones** (1 por desarrollo) | ✅ Agrupa actividades por proyecto |

## Modelo recomendado

```
GitHub Project (tablero único)
    │
    ├── Milestone: REVISION ACTA PACC - FACC SERVICIOS   ← el "proyecto"
    │       ├── Issue: [Pruebas] Socialización inicio...  ← actividad
    │       ├── Issue: [Cambios] Ajustes post-prueba...
    │       └── ...
    │
    ├── Milestone: APP PENALIDADES
    │       ├── Issue: [Desarrollo] Implementación...
    │       └── ...
    └── ...
```

**Vista del tablero:** `Group by → Milestone` → cada desarrollo muestra sus actividades.

**Vista tabla:** columnas Fecha (due date), Estado (labels), % en el cuerpo del issue.

## Paso 1 — Autenticarte (una sola vez)

```powershell
gh auth login
```

Elige: GitHub.com → HTTPS → Login with browser.

Alternativa con token:

```powershell
$env:GH_TOKEN = "ghp_tu_token_aqui"
```

Token en: GitHub → Settings → Developer settings → Personal access tokens → permiso **repo**.

## Paso 2 — Sincronizar Issues desde el cronograma

```powershell
cd "C:\Users\edwin\OneDrive\Desktop\Desarollo\Desarrollo - Cronograma y ejecucion"

# Ver qué se crearía (sin tocar GitHub)
python sync_github_issues.py --dry-run

# Crear ~140 issues + 40 milestones
python sync_github_issues.py

# Crear tablero Project y vincular issues
python sync_github_issues.py --project
```

## Paso 3 — Configurar la vista en Projects (manual, 2 min)

1. Abre [Projects](https://github.com/jhoned689/Optimacom-Claro---Desarrollos/projects)
2. Abre el project **"Ejecución Desarrollos Optimacom"** (o créalo: **New project → Board**)
3. **Add item →** busca issues con label `cronograma` (o importa desde repo)
4. En la vista:
   - **Layout → Group by → Milestone** (cada desarrollo = grupo)
   - O **Group by → Status** con columnas: HOY | PROCESO | PENDIENTE | REALIZADO
5. Filtro útil: `label:cronograma`

## Labels que crea el script

| Label | Significado |
|-------|-------------|
| `cronograma` | Todas las actividades sincronizadas |
| `estado/hoy` | Entrega hoy |
| `estado/proceso` | En curso |
| `estado/pendiente` | Por hacer |
| `estado/realizado` | Cerrado (issue closed) |
| `desarrollo/nombre-slug` | Filtrar por desarrollo |

## Rutina miércoles / viernes

```powershell
python actualizar_cronograma_jun2026.py   # Excel + proyectos/ + JSON
python sync_github_issues.py              # actualiza Issues en GitHub
git add . && git commit -m "Cronograma" && git push origin HEAD:main
```

## Relación repo ↔ Projects

| En el repo | En GitHub Projects |
|------------|-------------------|
| `proyectos/revision-acta-pacc-facc-servicios/` | Milestone del mismo nombre |
| `actividades.csv` fila 4 | Issue `[Pruebas] Socialización...` |
| Estado HOY / PENDIENTE | Label + abierto/cerrado |
| Fecha 24/06/2026 | Due date del issue |

## Si prefieres un Project por desarrollo

No es recomendable (40 projects). Mejor **un solo Project** agrupado por Milestone.

Si necesitas aislar uno: filtro `label:desarrollo/revision-acta-pacc-facc-servicios`.
