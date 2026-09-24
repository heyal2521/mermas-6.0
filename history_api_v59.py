"""API de histórico compartido para TOP MERMAS V59.

Puede añadirse a la aplicación Flask existente con:
    from history_api_v59 import register_history_api
    register_history_api(app)
"""

from __future__ import annotations

import base64
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


def _github_request(method: str, body: dict | None = None) -> dict:
    quoted_path = urllib.parse.quote(GITHUB_PATH, safe="/")
    url = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/contents/{quoted_path}"
    if method == "GET":
        url += "?ref=" + urllib.parse.quote(GITHUB_BRANCH)
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
                record = _clean_record(body.get("record") or {})
                records = payload.setdefault("records", [])
                index = next((i for i, item in enumerate(records) if item.get("key") == record["key"]), None)
                action = "inserted"
                if index is None:
                    records.append(record)
                else:
                    records[index] = record
                    action = "replaced"
                records.sort(key=lambda item: (item.get("year", 0), int(item.get("week", 0)), int(item.get("cycle", 0)), item.get("kind", "")))
                payload["updatedAt"] = datetime.now(timezone.utc).isoformat()
                _save(payload, sha)
                return jsonify({"ok": True, "action": action, "record": record, "total": len(records)})
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception as exc:
            app.logger.exception("Error en histórico V59")
            return jsonify({"ok": False, "error": "No se pudo actualizar el histórico compartido."}), 500


app = Flask(__name__)
register_history_api(app)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "10000")))
