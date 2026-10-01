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

from flask import jsonify, request

FORMAT = "TOP_MERMAS_HISTORICO"
REPO = os.getenv("GITHUB_REPO") or os.getenv("GITHUB_REPOSITORY", "heyal2521/mermas-6.0")
BRANCH = os.getenv("GITHUB_BRANCH", "main")
TOKEN = (os.getenv("GITHUB_HISTORY_TOKEN") or os.getenv("GITHUB_TOKEN") or "").strip()
WRITE_KEY = (os.getenv("TOP_MERMAS_WRITE_KEY") or os.getenv("SHARED_DASHBOARD_WRITE_KEY") or "").strip()
ALLOWED = {
    x.strip()
    for x in os.getenv(
        "TOP_MERMAS_ALLOWED_ORIGINS",
        "https://heyal2521.github.io,https://mermas-6-0.onrender.com,https://top-mermas.onrender.com",
    ).split(",")
    if x.strip()
}

MANIFEST_PATH = "historico/TOP_MERMAS_SAFE_MANIFEST_V60.json"
ITEM_DIR = "historico/v60_items"
LEGACY_RAW_URL = (
    "https://raw.githubusercontent.com/"
    "heyal2521/mermas-6.0/"
    "e1070757b9cf8ea05a1b1ce2cd65412745eadd94/"
    "historico/TOP_MERMAS_HISTORICO_SHARED_V60.json"
)
LOCK = threading.Lock()
MAX_BYTES = 15 * 1024 * 1024


def now():
    return datetime.now(timezone.utc).isoformat()


def gh_contents(method, path, body=None, ref=None):
    encoded_path = urllib.parse.quote(path, safe="/")
    url = f"https://api.github.com/repos/{REPO}/contents/{encoded_path}"
    if method == "GET" and ref:
        url += "?ref=" + urllib.parse.quote(ref, safe="")
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "top-mermas-v60-safe",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if TOKEN:
        headers["Authorization"] = f"Bearer {TOKEN}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def read_json_file(path, ref=None):
    result = gh_contents("GET", path, ref=ref or BRANCH)
    raw = base64.b64decode(result["content"].replace("\n", ""))
    return json.loads(raw.decode("utf-8")), result.get("sha")


def write_json_file(path, payload, sha=None):
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if not raw or raw[:1] != b"{" or raw[-1:] != b"}":
        raise RuntimeError("El contenido generado no es un JSON válido.")
    body = {
        "message": "Actualizar histórico compartido TOP MERMAS v60",
        "content": base64.b64encode(raw).decode("ascii"),
        "branch": BRANCH,
    }
    if sha:
        body["sha"] = sha
    return gh_contents("PUT", path, body=body)


def load_manifest():
    try:
        payload, sha = read_json_file(MANIFEST_PATH)
        if payload.get("format") != FORMAT or not isinstance(payload.get("items"), list):
            raise RuntimeError("Manifiesto central no válido.")
        return payload, sha
    except urllib.error.HTTPError as exc:
        if exc.code != 404:
            raise
        payload = {"format": FORMAT, "version": 3, "updatedAt": now(), "items": []}
        result = write_json_file(MANIFEST_PATH, payload)
        return payload, result.get("content", {}).get("sha") if isinstance(result.get("content"), dict) else None


def load_legacy_items():
    try:
        req = urllib.request.Request(
            LEGACY_RAW_URL,
            headers={"User-Agent": "top-mermas-v60-safe"},
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
        if payload.get("format") != FORMAT or not isinstance(payload.get("items"), list):
            return []
        return payload["items"]
    except Exception:
        return []


def item_path(item_hash):
    return f"{ITEM_DIR}/{item_hash}.json"


def load_all_items():
    legacy = load_legacy_items()
    manifest, _ = load_manifest()

    by_hash = {}
    for item in legacy:
        digest = item.get("sha256")
        if not digest and item.get("dataBase64"):
            try:
                digest = hashlib.sha256(base64.b64decode(item["dataBase64"])).hexdigest()
            except Exception:
                digest = None
        if digest:
            item["sha256"] = digest
            by_hash[digest] = item

    for entry in manifest.get("items", []):
        path = entry.get("path")
        digest = entry.get("sha256")
        if not path:
            continue
        try:
            item, _ = read_json_file(path)
            if digest:
                item["sha256"] = digest
            by_hash[item.get("sha256") or digest] = item
        except Exception:
            continue

    items = list(by_hash.values())
    items.sort(key=lambda x: x.get("addedAt", ""), reverse=True)
    return items, manifest


def validate_item(raw):
    data_b64 = str(raw.get("dataBase64") or "")
    if not data_b64:
        raise ValueError("Falta dataBase64.")
    try:
        data = base64.b64decode(data_b64, validate=True)
    except Exception as exc:
        raise ValueError("dataBase64 no válido.") from exc
    if not data:
        raise ValueError("Fichero vacío.")
    if len(data) > MAX_BYTES:
        raise ValueError("Fichero superior a 15 MB.")

    kind = str(raw.get("kind") or "original").lower()
    if kind not in {"original", "generated"}:
        raise ValueError("Tipo de fichero no válido.")

    digest = hashlib.sha256(data).hexdigest()
    return {
        "name": str(raw.get("name") or "FICHERO.xlsx")[:180],
        "size": len(data),
        "type": str(raw.get("type") or "application/octet-stream")[:150],
        "kind": kind,
        "sourceName": str(raw.get("sourceName") or "")[:180],
        "addedAt": str(raw.get("addedAt") or now()),
        "compositionIndex": raw.get("compositionIndex"),
        "dataBase64": base64.b64encode(data).decode("ascii"),
        "sha256": digest,
    }


def upsert_item(raw):
    item = validate_item(raw)
    path = item_path(item["sha256"])

    try:
        _, sha = read_json_file(path)
    except urllib.error.HTTPError as exc:
        if exc.code != 404:
            raise
        sha = None
    except Exception as read_exc:
        # Si el objeto ya existe pero su contenido está vacío/corrupto,
        # necesitamos únicamente su SHA para sobrescribirlo. La API Contents
        # devuelve ese SHA aunque el contenido no sea JSON válido.
        try:
            raw_meta = gh_contents("GET", path, ref=BRANCH)
            sha = raw_meta.get("sha")
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                raise
            sha = None

    write_json_file(path, item, sha=sha)

    # Verifica que GitHub devuelve el fichero íntegro antes de anunciarlo.
    verified, _ = read_json_file(path)
    if verified.get("sha256") != item["sha256"] or not verified.get("dataBase64"):
        raise RuntimeError("La verificación del fichero guardado ha fallado.")

    manifest, manifest_sha = load_manifest()
    entries = manifest.setdefault("items", [])
    existing = {e.get("sha256"): e for e in entries}

    entry = {
        "sha256": item["sha256"],
        "path": path,
        "name": item["name"],
        "kind": item["kind"],
        "sourceName": item["sourceName"],
        "size": item["size"],
        "addedAt": item["addedAt"],
    }
    existing[item["sha256"]] = entry

    # Para TOP_GENERADO sustituimos la referencia anterior del mismo origen,
    # pero nunca borramos el fichero físico antiguo.
    if item["kind"] == "generated" and item["sourceName"]:
        entries = [
            e for e in entries
            if not (
                e.get("kind") == "generated"
                and e.get("sourceName") == item["sourceName"]
            )
        ]
        entries.append(entry)
    else:
        entries = list({e.get("sha256"): e for e in entries if e.get("sha256")}.values())
        if not any(e.get("sha256") == item["sha256"] for e in entries):
            entries.append(entry)

    manifest["items"] = sorted(entries, key=lambda e: e.get("addedAt", ""), reverse=True)
    manifest["updatedAt"] = now()

    # El manifiesto es pequeño; se valida antes de sustituirlo.
    probe = json.dumps(manifest, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if not probe or probe[:1] != b"{" or probe[-1:] != b"}":
        raise RuntimeError("No se pudo validar el manifiesto.")

    write_json_file(MANIFEST_PATH, manifest, sha=manifest_sha)
    return item


def register_shared_history_api_safe(app):
    @app.after_request
    def cors_safe(resp):
        origin = request.headers.get("Origin")
        if origin in ALLOWED or origin == "null":
            resp.headers["Access-Control-Allow-Origin"] = origin
            resp.headers["Vary"] = "Origin"
            resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Top-Mermas-Key"
            resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
            resp.headers["Access-Control-Max-Age"] = "600"
        return resp

    @app.route("/api/history", methods=["GET", "POST", "OPTIONS"])
    def history_safe():
        if request.method == "OPTIONS":
            return ("", 204)
        try:
            if request.method == "GET":
                items, _ = load_all_items()
                return jsonify({
                    "format": FORMAT,
                    "version": 3,
                    "updatedAt": now(),
                    "items": items,
                })

            if not WRITE_KEY:
                return jsonify({"ok": False, "error": "Código de escritura no configurado en Render."}), 503

            body = request.get_json(silent=True)
            if body is None:
                body = json.loads(request.get_data(as_text=True) or "{}")

            key = str(body.get("_writeKey") or request.headers.get("X-Top-Mermas-Key", "")).strip()
            if key != WRITE_KEY:
                return jsonify({"ok": False, "error": "Código de equipo incorrecto."}), 401

            incoming = body.get("items") if isinstance(body.get("items"), list) else [body]
            if not incoming or len(incoming) > 50:
                return jsonify({"ok": False, "error": "Lote no válido."}), 400

            with LOCK:
                stored = []
                for raw in incoming:
                    stored.append(upsert_item(raw))

            items, _ = load_all_items()
            return jsonify({
                "ok": True,
                "stored": len(stored),
                "total": len(items),
            })
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception:
            app.logger.exception("Error en histórico compartido seguro v60")
            return jsonify({"ok": False, "error": "No se pudo actualizar el histórico compartido."}), 500

    @app.route("/api/health", methods=["GET"])
    def health_safe():
        try:
            items, _ = load_all_items()
            return jsonify({
                "ok": True,
                "service": "top-mermas-shared-v60-safe",
                "items": len(items),
                "githubConfigured": bool(TOKEN),
                "writeKeyConfigured": bool(WRITE_KEY),
            })
        except Exception as exc:
            return jsonify({
                "ok": False,
                "service": "top-mermas-shared-v60-safe",
                "error": str(exc),
                "githubConfigured": bool(TOKEN),
                "writeKeyConfigured": bool(WRITE_KEY),
            }), 503
