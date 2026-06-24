# Publicar en GitHub

## 1. Crear el repositorio en GitHub

1. Entra a [github.com/new](https://github.com/new)
2. Nombre: `Optimacom-Claro---Desarrollos` (ya creado)
3. URL: https://github.com/jhoned689/Optimacom-Claro---Desarrollos

## 2. Subir desde tu PC

Abre PowerShell en esta carpeta y ejecuta (cambia `TU_USUARIO`):

```powershell
cd "C:\Users\edwin\OneDrive\Desktop\Desarollo\Desarrollo - Cronograma y ejecucion"

git init
git add .
git commit -m "Cronograma desarrollos Optimacom — jun 2026"

git branch -M main
git remote add origin https://github.com/jhoned689/Optimacom-Claro---Desarrollos.git
git push -u origin main
```

## 3. Actualizar después de cada miércoles o viernes

```powershell
python actualizar_cronograma_jun2026.py
git add .
git commit -m "Actualización cronograma — fecha"
git push
```

## 4. Vista más profesional (opcional)

En GitHub → **Settings** → **Pages**:

- Source: **Deploy from branch**
- Branch: `main` / carpeta `/docs`

Luego edita `docs/index.md` o usa los `.md` existentes como landing.

## 5. Instalar GitHub CLI (opcional)

Para crear el repo desde terminal:

```powershell
winget install GitHub.cli
gh auth login
gh repo create jhoned689/Optimacom-Claro---Desarrollos --private --source=. --push
```

## Qué se publica

| Incluido | Motivo |
|----------|--------|
| README, docs/, data/*.csv | Vista profesional y legible |
| PBI Desarrollos.xlsx | Fuente del tablero |
| Script Python | Reproducir actualizaciones |

Si prefieres **no** subir el Excel, agrégalo al `.gitignore` y deja solo los CSV en GitHub.
