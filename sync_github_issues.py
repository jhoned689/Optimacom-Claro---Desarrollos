# -*- coding: utf-8 -*-
"""
Sincroniza el cronograma (proyectos.json) con GitHub Issues + Milestones.

Modelo:
  - 1 Milestone  = 1 desarrollo (proyecto)
  - 1 Issue      = 1 actividad del desarrollo
  - GitHub Project (tablero) agrupa Issues por Milestone → ves ejecución por desarrollo

Requisito: gh auth login  O  variable GH_TOKEN con permiso repo.

Uso:
  python sync_github_issues.py          # crear/actualizar issues
  python sync_github_issues.py --dry-run
  python sync_github_issues.py --project  # además vincula issues al Project board
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
JSON_PATH = ROOT / "data" / "proyectos.json"
REPO = "jhoned689/Optimacom-Claro---Desarrollos"
PROJECT_TITLE = "Ejecución Desarrollos Optimacom"

MARKER_PREFIX = "cronograma-sync:"

ESTADO_LABELS = {
    "HOY": "estado/hoy",
    "PROCESO": "estado/proceso",
    "PENDIENTE": "estado/pendiente",
    "REALIZADO": "estado/realizado",
}


def gh_api(method: str, endpoint: str, payload: dict | None = None) -> dict | list | None:
    args = ["api", endpoint, "-X", method]
    if payload is not None:
        args.extend(["--input", "-"])
    proc = subprocess.run(
        ["gh", *args],
        input=json.dumps(payload) if payload else None,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    if not proc.stdout.strip():
        return None
    return json.loads(proc.stdout)


def gh(*args: str) -> dict | list | None:
    proc = subprocess.run(
        ["gh", *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    if not proc.stdout.strip():
        return None
    return json.loads(proc.stdout)


def check_auth() -> None:
    r = subprocess.run(["gh", "auth", "status"], capture_output=True, text=True)
    if r.returncode != 0:
        print(
            "GitHub CLI no autenticado.\n"
            "Ejecuta:  gh auth login\n"
            "O define:  set GH_TOKEN=ghp_xxxx  (permiso repo)\n",
            file=sys.stderr,
        )
        sys.exit(1)


def sync_id(slug: str, index: int) -> str:
    return f"{MARKER_PREFIX}{slug}/{index:03d}"


def parse_due(fecha: str) -> str | None:
    if not fecha or fecha == "—":
        return None
    try:
        d = datetime.strptime(fecha, "%d/%m/%Y")
        return d.strftime("%Y-%m-%dT12:00:00Z")
    except ValueError:
        return None


def ensure_labels(dry_run: bool) -> None:
    labels = set(ESTADO_LABELS.values())
    labels.add("cronograma")
    existing: set[str] = set()
    if not dry_run:
        existing = {
            lb["name"]
            for lb in (gh("api", f"repos/{REPO}/labels", "--paginate") or [])
        }
    for name in labels:
        if name in existing:
            continue
        color = {"estado/hoy": "d73a4a", "estado/proceso": "fbca04",
                 "estado/pendiente": "0e8a16", "estado/realizado": "6f42c1",
                 "cronograma": "0366d6"}.get(name, "ededed")
        if dry_run:
            print(f"  [dry] crear label: {name}")
            continue
        gh_api("POST", f"repos/{REPO}/labels", {
            "name": name,
            "color": color,
        })


def load_existing_issues() -> dict[str, dict]:
    """Mapa sync_id → issue."""
    issues = gh("api", f"repos/{REPO}/issues", "--paginate",
                "-f", "state=all", "-f", "per_page=100") or []
    out: dict[str, dict] = {}
    for issue in issues:
        if "pull_request" in issue:
            continue
        body = issue.get("body") or ""
        m = re.search(rf"<!--\s*({MARKER_PREFIX}[^\s]+)\s*-->", body)
        if m:
            out[m.group(1)] = issue
    return out


def load_milestones() -> dict[str, dict]:
    ms = gh("api", f"repos/{REPO}/milestones", "--paginate") or []
    return {m["title"]: m for m in ms}


def milestone_body(p: dict) -> str:
    return (
        f"**Desarrollo:** {p['nombre']}\n\n"
        f"- Avance: **{int(p['avance']*100) if p['avance'] <= 1 else p['avance']}%**\n"
        f"- Estado: {p['estado']}\n"
        f"- Actividades: {p['total_actividades']} "
        f"({p['realizadas']} realizadas, {p['en_curso']} en curso, {p['pendientes']} pendientes)\n"
        f"- Próxima entrega: {p['proxima_fecha']}\n\n"
        f"Carpeta: [`proyectos/{p['slug']}`](../blob/main/proyectos/{p['slug']}/README.md)"
    )


def issue_body(p: dict, act: dict, sid: str) -> str:
    avance = act.get("avance")
    pct = f"{int(avance*100)}%" if isinstance(avance, (int, float)) and avance <= 1 else (str(avance) if avance else "—")
    return f"""<!-- {sid} -->

## Desarrollo
**{p['nombre']}**

| Campo | Valor |
|-------|-------|
| Fase | {act['fase']} |
| Entregable | {act['entregable']} |
| Fecha | {act['fecha']} |
| Estado | **{act['estado']}** |
| % Avance | {pct} |

{act.get('notas') or ''}

---
_Sincronizado desde cronograma Optimacom_
"""


def ensure_milestone(p: dict, milestones: dict[str, dict], dry_run: bool) -> int | None:
    title = p["nombre"]
    if title in milestones:
        mid = milestones[title]["number"]
        if not dry_run:
            gh_api("PATCH", f"repos/{REPO}/milestones/{mid}", {
                "description": milestone_body(p),
                "state": "closed" if p["estado"] == "REALIZADO" else "open",
            })
        return mid

    if dry_run:
        print(f"  [dry] crear milestone: {title}")
        return 0

    created = gh_api("POST", f"repos/{REPO}/milestones", {
        "title": title,
        "description": milestone_body(p),
    })
    milestones[title] = created
    return created["number"]


def sync_issues(proyectos: list, dry_run: bool) -> tuple[int, int, int]:
    ensure_labels(dry_run)
    milestones = load_milestones() if not dry_run else {}
    existing = load_existing_issues() if not dry_run else {}

    created = updated = skipped = 0

    for p in proyectos:
        mid = ensure_milestone(p, milestones, dry_run)
        slug = p["slug"]

        for i, act in enumerate(p["actividades"], start=1):
            sid = sync_id(slug, i)
            title = f"[{act['fase']}] {act['actividades']}"
            labels = ["cronograma", ESTADO_LABELS.get(act["estado"], "estado/pendiente"),
                      f"desarrollo/{slug}"]
            body = issue_body(p, act, sid)
            state = "closed" if act["estado"] == "REALIZADO" else "open"
            due = parse_due(act.get("fecha", ""))

            if dry_run:
                print(f"  [dry] {sid} | {title[:60]} | {act['estado']}")
                continue

            payload: dict = {
                "title": title,
                "body": body,
                "labels": labels,
                "state": state,
            }
            if mid:
                payload["milestone"] = mid
            if due:
                payload["due_on"] = due

            if sid in existing:
                num = existing[sid]["number"]
                gh_api("PATCH", f"repos/{REPO}/issues/{num}", payload)
                updated += 1
            else:
                gh_api("POST", f"repos/{REPO}/issues", payload)
                created += 1

    return created, updated, skipped


def ensure_desarrollo_labels(proyectos: list, dry_run: bool) -> None:
    existing = {
        lb["name"]
        for lb in (gh("api", f"repos/{REPO}/labels", "--paginate") or [])
    } if not dry_run else set()

    for p in proyectos:
        name = f"desarrollo/{p['slug']}"
        if name in existing or dry_run:
            if dry_run and name not in existing:
                print(f"  [dry] label: {name}")
            continue
        gh_api("POST", f"repos/{REPO}/labels", {
            "name": name,
            "color": "c5def5",
            "description": f"Actividades del desarrollo: {p['nombre']}",
        })


def link_to_project(dry_run: bool) -> None:
    """Crea Project v2 (si no existe) y agrega todos los issues del cronograma."""
    if dry_run:
        print("[dry] vincular issues al GitHub Project")
        return

    owner = REPO.split("/")[0]
    projects = gh("project", "list", "--owner", owner, "--format", "json") or []
    project = next((p for p in projects if p.get("title") == PROJECT_TITLE), None)

    if not project:
        project = gh("project", "create", "--owner", owner,
                       "--title", PROJECT_TITLE, "--format", "json")
        print(f"Project creado: {project['url']}")

    project_id = project["id"]
    issues = gh("api", f"repos/{REPO}/issues", "--paginate",
                "-f", "labels=cronograma", "-f", "state=all") or []

    for issue in issues:
        url = issue["html_url"]
        try:
            gh("project", "item-add", str(project["number"]),
               "--owner", owner, "--url", url)
        except RuntimeError as e:
            if "already exists" not in str(e).lower():
                pass  # item-add idempotente en algunas versiones

    print(f"Project: {project.get('url', project_id)}")
    print("Configura vista: Group by → Milestone (cada desarrollo = columna/grupo)")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--project", action="store_true", help="Crear/vincular GitHub Project")
    args = parser.parse_args()

    if not JSON_PATH.exists():
        print("Ejecuta primero: python exportar_proyectos.py", file=sys.stderr)
        sys.exit(1)

    proyectos = json.loads(JSON_PATH.read_text(encoding="utf-8"))

    if not args.dry_run:
        check_auth()

    print(f"Repo: {REPO}")
    print(f"Proyectos: {len(proyectos)}")
    print(f"Actividades: {sum(len(p['actividades']) for p in proyectos)}")

    if not args.dry_run:
        ensure_desarrollo_labels(proyectos, args.dry_run)

    c, u, s = sync_issues(proyectos, args.dry_run)
    print(f"Issues — creados: {c}, actualizados: {u}")

    if args.project:
        link_to_project(args.dry_run)

    if args.dry_run:
        print("\nQuita --dry-run para publicar en GitHub.")


if __name__ == "__main__":
    main()
