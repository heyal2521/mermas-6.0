from __future__ import annotations
import base64, hashlib, json, os, threading, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timezone
from flask import jsonify, request

FORMAT="TOP_MERMAS_HISTORICO"
REPO=os.getenv("GITHUB_REPO") or os.getenv("GITHUB_REPOSITORY","heyal2521/mermas-6.0")
BRANCH=os.getenv("GITHUB_BRANCH","main")
PATH=os.getenv("TOP_MERMAS_SHARED_V60_PATH","historico/TOP_MERMAS_HISTORICO_SHARED_V60.json")
TOKEN=(os.getenv("GITHUB_HISTORY_TOKEN") or os.getenv("GITHUB_TOKEN") or "").strip()
WRITE_KEY=(os.getenv("TOP_MERMAS_WRITE_KEY") or os.getenv("SHARED_DASHBOARD_WRITE_KEY") or "").strip()
ALLOWED={x.strip() for x in os.getenv("TOP_MERMAS_ALLOWED_ORIGINS","https://heyal2521.github.io,https://mermas-6-0.onrender.com,https://top-mermas.onrender.com").split(",") if x.strip()}
LOCK=threading.Lock()
MAX_BYTES=15*1024*1024

def now():
    return datetime.now(timezone.utc).isoformat()

def empty():
    return {"format":FORMAT,"version":2,"exportedAt":now(),"items":[]}

def gh(method,body=None):
    url=f"https://api.github.com/repos/{REPO}/contents/{urllib.parse.quote(PATH,safe='/')}"
    if method=="GET": url+="?ref="+urllib.parse.quote(BRANCH,safe="")
    headers={"Accept":"application/vnd.github+json","User-Agent":"top-mermas-v60","X-GitHub-Api-Version":"2022-11-28"}
    if TOKEN: headers["Authorization"]=f"Bearer {TOKEN}"
    data=json.dumps(body).encode() if body is not None else None
    req=urllib.request.Request(url,data=data,headers=headers,method=method)
    with urllib.request.urlopen(req,timeout=30) as r:
        return json.loads(r.read().decode())

def load():
    if not TOKEN: raise RuntimeError("GITHUB_TOKEN no configurado.")
    try:
        r=gh("GET")
        p=json.loads(base64.b64decode(r["content"]).decode())
        if p.get("format")!=FORMAT or not isinstance(p.get("items"),list): raise RuntimeError("Formato central no valido.")
        return p,r.get("sha")
    except urllib.error.HTTPError as e:
        if e.code==404: return empty(),None
        raise

def save(p,sha):
    raw=json.dumps(p,ensure_ascii=False,separators=(",",":")).encode()
    body={"message":"Actualizar historico compartido TOP MERMAS v60","content":base64.b64encode(raw).decode(),"branch":BRANCH}
    if sha: body["sha"]=sha
    gh("PUT",body)

def clean(x):
    b64=str(x.get("dataBase64") or "")
    if not b64: raise ValueError("Falta dataBase64.")
    try: data=base64.b64decode(b64,validate=True)
    except Exception as e: raise ValueError("dataBase64 no valido.") from e
    if not data: raise ValueError("Fichero vacio.")
    if len(data)>MAX_BYTES: raise ValueError("Fichero superior a 15 MB.")
    kind=str(x.get("kind") or "original").lower()
    if kind not in ("original","generated"): raise ValueError("Tipo de fichero no valido.")
    return {"name":str(x.get("name") or "FICHERO.xlsx")[:180],"size":len(data),"type":str(x.get("type") or "application/octet-stream")[:150],"kind":kind,"sourceName":str(x.get("sourceName") or "")[:180],"addedAt":str(x.get("addedAt") or now()),"compositionIndex":x.get("compositionIndex"),"dataBase64":base64.b64encode(data).decode(),"sha256":hashlib.sha256(data).hexdigest()}

def merge(p,incoming):
    items=p.setdefault("items",[])
    hashes={i.get("sha256") for i in items if i.get("sha256")}
    added=replaced=0
    for raw in incoming:
        item=clean(raw)
        if item["kind"]=="generated" and item["sourceName"]:
            idx=next((n for n,v in enumerate(items) if v.get("kind")=="generated" and v.get("sourceName")==item["sourceName"]),None)
            if idx is not None:
                items[idx]=item; replaced+=1; hashes.add(item["sha256"]); continue
        if item["sha256"] in hashes: continue
        items.append(item); hashes.add(item["sha256"]); added+=1
    items.sort(key=lambda x:x.get("addedAt",""),reverse=True)
    return added,replaced

def register_shared_history_api(app):
    @app.after_request
    def cors_v60(resp):
        origin=request.headers.get("Origin")
        if origin in ALLOWED or origin=="null":
            resp.headers["Access-Control-Allow-Origin"]=origin
            resp.headers["Vary"]="Origin"
            resp.headers["Access-Control-Allow-Headers"]="Content-Type, X-Top-Mermas-Key"
            resp.headers["Access-Control-Allow-Methods"]="GET, POST, OPTIONS"
            resp.headers["Access-Control-Max-Age"]="600"
        return resp

    @app.route("/api/history",methods=["GET","POST","OPTIONS"])
    @app.route("/api/history/bulk",methods=["POST","OPTIONS"])
    def shared_history_v60():
        if request.method=="OPTIONS":
            resp=app.make_response(("",204))
            origin=request.headers.get("Origin")
            if origin in ALLOWED or origin=="null":
                resp.headers["Access-Control-Allow-Origin"]=origin
                resp.headers["Access-Control-Allow-Headers"]="Content-Type, X-Top-Mermas-Key"
                resp.headers["Access-Control-Allow-Methods"]="GET, POST, OPTIONS"
                resp.headers["Access-Control-Max-Age"]="600"
            return resp
        if request.method=="GET":
            try:
                p,_=load(); return jsonify(p)
            except Exception as e:
                app.logger.exception("v60 history read"); return jsonify({"ok":False,"error":str(e)}),503
        if not WRITE_KEY: return jsonify({"ok":False,"error":"Codigo de escritura no configurado en Render."}),503
        if request.headers.get("X-Top-Mermas-Key","").strip()!=WRITE_KEY: return jsonify({"ok":False,"error":"Codigo de equipo incorrecto."}),401
        body=request.get_json(silent=True) or {}
        incoming=body.get("items") if isinstance(body.get("items"),list) else [body]
        if not incoming or len(incoming)>500: return jsonify({"ok":False,"error":"Lote no valido."}),400
        try:
            with LOCK:
                p,sha=load(); added,replaced=merge(p,incoming); p["exportedAt"]=now(); p["updatedAt"]=now(); save(p,sha)
            return jsonify({"ok":True,"added":added,"replaced":replaced,"total":len(p["items"])})
        except ValueError as e: return jsonify({"ok":False,"error":str(e)}),400
        except Exception as e:
            app.logger.exception("v60 history write"); return jsonify({"ok":False,"error":str(e)}),500

    @app.route("/api/health",methods=["GET"])
    def shared_health_v60():
        try:
            p,_=load()
            return jsonify({"ok":True,"service":"top-mermas-shared-v60","items":len(p.get("items",[])),"githubConfigured":bool(TOKEN),"writeKeyConfigured":bool(WRITE_KEY)})
        except Exception as e:
            return jsonify({"ok":False,"service":"top-mermas-shared-v60","error":str(e),"githubConfigured":bool(TOKEN),"writeKeyConfigured":bool(WRITE_KEY)}),503
