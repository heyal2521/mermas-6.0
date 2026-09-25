"""API de histórico compartido para TOP MERMAS V59.

Puede añadirse a la aplicación Flask existente con:
    from history_api_v59 import register_history_api
    register_history_api(app)
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, request


FORMAT = "TOP_MERMAS_CENTRAL_V59"
LOCAL_PATH = Path(os.getenv("TOP_MERMAS_HISTORY_FILE", "/tmp/top_mermas_v59_history.json"))
GITHUB_REPOSITORY = os.getenv("GITHUB_REPOSITORY") or os.getenv("GITHUB_REPO", "heyal2521/mermas-6.0")
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main")
GITHUB_PATH = os.getenv("GITHUB_HISTORY_PATH", "historico/TOP_MERMAS_CENTRAL_V59.json")
GITHUB_ARCHIVE_BRANCH = os.getenv("GITHUB_ARCHIVE_BRANCH", "gh-pages")
GITHUB_ARCHIVE_DIRECTORY = os.getenv("GITHUB_ARCHIVE_DIRECTORY", "historico/archivos")
MAX_SHARED_FILE_BYTES = int(os.getenv("TOP_MERMAS_MAX_FILE_BYTES", str(15 * 1024 * 1024)))
GITHUB_TOKEN = os.getenv("GITHUB_HISTORY_TOKEN") or os.getenv("GITHUB_TOKEN")
ALLOWED_ORIGINS = {
    value.strip()
    for value in os.getenv(
        "TOP_MERMAS_ALLOWED_ORIGINS",
        "https://heyal2521.github.io,https://top-mermas.onrender.com",
    ).split(",")
    if value.strip()
}
LOCK = threading.Lock()


def _empty_payload() -> dict:
    return {"format": FORMAT, "version": 1, "updatedAt": None, "records": []}


def _github_request(method: str, body: dict | None = None, path: str | None = None, branch: str | None = None) -> dict:
    quoted_path = urllib.parse.quote(path or GITHUB_PATH, safe="/")
    url = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/contents/{quoted_path}"
    if method == "GET":
        url += "?ref=" + urllib.parse.quote(branch or GITHUB_BRANCH)
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "top-mermas-v59",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if GITHUB_TOKEN:
        headers["Authorization"] = f"Bearer {GITHUB_TOKEN}"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def _load() -> tuple[dict, str | None]:
    if GITHUB_TOKEN:
        try:
            result = _github_request("GET")
            content = base64.b64decode(result["content"]).decode("utf-8")
            payload = json.loads(content)
            return payload, result.get("sha")
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                raise
            return _empty_payload(), None

    if not LOCAL_PATH.exists():
        return _empty_payload(), None
    return json.loads(LOCAL_PATH.read_text(encoding="utf-8")), None


def _save(payload: dict, sha: str | None) -> None:
    encoded = base64.b64encode(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ).decode("ascii")
    if GITHUB_TOKEN:
        body = {
            "message": "Actualizar histórico compartido TOP MERMAS V59",
            "content": encoded,
            "branch": GITHUB_BRANCH,
        }
        if sha:
            body["sha"] = sha
        _github_request("PUT", body)
        return

    LOCAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    temp_path = LOCAL_PATH.with_suffix(".tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    temp_path.replace(LOCAL_PATH)


def _clean_record(raw: dict) -> dict:
    kind = str(raw.get("kind", "")).strip().lower()
    if kind not in {"original", "generated"}:
        raise ValueError("Tipo de fichero no válido.")
    year = int(raw.get("year"))
    week = int(raw.get("week"))
    cycle = int(raw.get("cycle"))
    if not (2020 <= year <= 2100 and 1 <= week <= 53 and 1 <= cycle <= 9):
        raise ValueError("Año, semana o ciclo fuera de rango.")

    rows = []
    seen = set()
    for item in raw.get("rows", []):
        mc = str(item.get("mc", "")).strip().upper().replace(" ", "")
        if not mc or mc in seen:
            continue
        seen.add(mc)
        rows.append({"mc": mc[:40], "family": str(item.get("family", ""))[:100]})
    if not rows:
        raise ValueError("El fichero no contiene modelos/calidades.")

    compositions = []
    for item in raw.get("compositionIndex", []):
        mc = str(item.get("mc", "")).strip().upper().replace(" ", "")
        values = []
        for value in item.get("compositions", []):
            label = str(value.get("label", ""))[:60]
            if label:
                values.append({"label": label, "percentage": str(value.get("percentage", ""))[:30]})
        if mc:
            compositions.append(
                {"mc": mc, "year": year, "week": str(week), "cycle": str(cycle), "compositions": values}
            )

    return {
        "key": f"{kind}||{year}||{week}||{cycle}",
        "kind": kind,
        "year": year,
        "week": str(week),
        "cycle": str(cycle),
        "fileName": str(raw.get("fileName", "fichero.xlsx"))[:180],
        "fileSize": int(raw.get("fileSize") or 0),
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "rows": rows,
        "compositionIndex": compositions,
    }


def register_history_api(app: Flask) -> None:
    @app.after_request
    def add_v59_cors(response):
        origin = request.headers.get("Origin")
        if origin in ALLOWED_ORIGINS or (origin == "null" and request.method in {"GET", "POST", "OPTIONS"}):
            response.headers["Access-Control-Allow-Origin"] = origin
            response.headers["Vary"] = "Origin"
            response.headers["Access-Control-Allow-Headers"] = "Content-Type"
            response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        return response

    @app.route("/api/v59/history", methods=["GET", "POST", "OPTIONS"])
    def v59_history():
        if request.method == "OPTIONS":
            return ("", 204)
        try:
            with LOCK:
                payload, sha = _load()
                if request.method == "GET":
                    return jsonify(payload)

                body = request.get_json(silent=True) or {}
                raw_records = body.get("records")
                is_batch = isinstance(raw_records, list)
                if is_batch:
                    if not raw_records:
                        raise ValueError("No se han recibido registros para sincronizar.")
                    if len(raw_records) > 250:
                        raise ValueError("La sincronización admite un máximo de 250 registros.")
                    clean_records = [_clean_record(item or {}) for item in raw_records]
                else:
                    clean_records = [_clean_record(body.get("record") or {})]

                records = payload.setdefault("records", [])
                indexes = {item.get("key"): i for i, item in enumerate(records)}
                inserted = 0
                replaced = 0
                for record in clean_records:
                    index = indexes.get(record["key"])
                    if index is None:
                        indexes[record["key"]] = len(records)
                        records.append(record)
                        inserted += 1
                    else:
                        records[index] = record
                        replaced += 1
                records.sort(key=lambda item: (item.get("year", 0), int(item.get("week", 0)), int(item.get("cycle", 0)), item.get("kind", "")))
                payload["updatedAt"] = datetime.now(timezone.utc).isoformat()
                _save(payload, sha)
                if is_batch:
                    return jsonify({"ok": True, "inserted": inserted, "replaced": replaced, "processed": len(clean_records), "total": len(records)})
                action = "inserted" if inserted else "replaced"
                return jsonify({"ok": True, "action": action, "record": clean_records[0], "total": len(records)})
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception as exc:
            app.logger.exception("Error en histórico V59")
            return jsonify({"ok": False, "error": "No se pudo actualizar el histórico compartido."}), 500


    @app.route("/api/v59/files", methods=["GET", "POST", "OPTIONS"])
    def v59_files():
        if request.method == "OPTIONS":
            return ("", 204)
        if not GITHUB_TOKEN:
            return jsonify({"ok": False, "error": "El almacenamiento público no está configurado en el servidor."}), 503
        try:
            with LOCK:
                payload, history_sha = _load()
                files = payload.setdefault("files", [])
                if request.method == "GET":
                    files.sort(key=lambda item: item.get("uploadedAt", ""), reverse=True)
                    return jsonify({"ok": True, "files": files, "total": len(files)})

                uploaded = request.files.get("file")
                if uploaded is None or not uploaded.filename:
                    raise ValueError("Selecciona un fichero Excel.")
                original_name = uploaded.filename.replace("\\", "/").split("/")[-1].strip()[:180]
                extension = os.path.splitext(original_name)[1].lower()
                if extension not in {".xlsx", ".xlsm", ".xls"}:
                    raise ValueError("Solo se admiten ficheros .xlsx, .xlsm o .xls.")
                data = uploaded.stream.read(MAX_SHARED_FILE_BYTES + 1)
                if not data:
                    raise ValueError("El fichero está vacío.")
                if len(data) > MAX_SHARED_FILE_BYTES:
                    raise ValueError(f"El fichero supera el límite de {MAX_SHARED_FILE_BYTES // (1024 * 1024)} MB.")
                digest = hashlib.sha256(data).hexdigest()
                stem = os.path.splitext(original_name)[0]
                safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")[:90] or "fichero"
                archive_name = f"{safe_stem}-{digest[:16]}{extension}"
                archive_path = f"{GITHUB_ARCHIVE_DIRECTORY.rstrip('/')}/{archive_name}"
                existing_sha = None
                try:
                    existing = _github_request("GET", path=archive_path, branch=GITHUB_ARCHIVE_BRANCH)
                    existing_sha = existing.get("sha")
                except urllib.error.HTTPError as exc:
                    if exc.code != 404:
                        raise
                if existing_sha is None:
                    body = {
                        "message": f"Archivar Excel público: {original_name}",
                        "content": base64.b64encode(data).decode("ascii"),
                        "branch": GITHUB_ARCHIVE_BRANCH,
                    }
                    _github_request("PUT", body, path=archive_path)
                public_url = (
                    f"https://raw.githubusercontent.com/{GITHUB_REPOSITORY}/"
                    f"{urllib.parse.quote(GITHUB_ARCHIVE_BRANCH, safe='')}/"
                    f"{urllib.parse.quote(archive_path, safe='/')}"
                )
                entry = {
                    "id": digest,
                    "name": original_name,
                    "size": len(data),
                    "sha256": digest,
                    "url": public_url,
                    "uploadedAt": datetime.now(timezone.utc).isoformat(),
                }
                by_id = {item.get("id"): item for item in files}
                by_id[digest] = entry
                payload["files"] = list(by_id.values())
                payload["updatedAt"] = datetime.now(timezone.utc).isoformat()
                _save(payload, history_sha)
                return jsonify({"ok": True, "file": entry, "total": len(payload["files"])})
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception:
            app.logger.exception("Error en archivos públicos V59")
            return jsonify({"ok": False, "error": "No se pudo guardar el Excel en el archivo público."}), 500


app = Flask(__name__)
register_history_api(app)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "10000")))
