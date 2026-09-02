#!/usr/bin/env python3
"""
descargar_matriz.py — Descarga la Matriz FFIE desde SharePoint usando Microsoft Graph
con autenticación de aplicación (Entra ID, client credentials). Sin usuario, sin enlaces
anónimos: la app registrada por TI tiene permiso de solo lectura sobre el sitio.

Variables de entorno (GitHub Secrets)
  AZURE_TENANT_ID      ID del tenant de Entra ID (GUID o ffie3.onmicrosoft.com)
  AZURE_CLIENT_ID      Application (client) ID del App Registration
  AZURE_CLIENT_SECRET  Client secret del App Registration
  SP_HOSTNAME          ffie3.sharepoint.com                       (opcional, ese es el default)
  SP_SITE_PATH         /sites/SRVINNOVACIN2                       (opcional, ese es el default)
  SP_FILE_PATH         Ruta del archivo dentro de la biblioteca "Documentos" del sitio,
                       p. ej. "00. Compartido/Proyecto Auditoria Digital IE/Matriz FFIE.xlsx"
  SP_FILE_NAME         (alternativa a SP_FILE_PATH) nombre del archivo para buscarlo en el sitio,
                       p. ej. "Matriz FFIE.xlsx". Se usa si SP_FILE_PATH no está definido.

Uso
  python sync/descargar_matriz.py --out /tmp/matriz_ffie.xlsx
Sale con código != 0 si algo falla (el workflow se detiene y el JSON anterior queda intacto).
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

GRAPH = "https://graph.microsoft.com/v1.0"


def die(msg: str) -> None:
    print(f"::error::{msg}", file=sys.stderr)
    sys.exit(1)


def env(name: str, default: str | None = None, required: bool = False) -> str:
    v = os.environ.get(name, "").strip() or (default or "")
    if required and not v:
        die(f"Falta la variable/secreto {name}")
    return v


def http(url: str, token: str | None = None, data: bytes | None = None,
         headers: dict | None = None, raw: bool = False):
    h = {"Accept": "application/json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    if headers:
        h.update(headers)
    req = urllib.request.Request(url, data=data, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            body = r.read()
            return body if raw else json.loads(body or b"{}")
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "replace")[:600]
        die(f"HTTP {e.code} en {url.split('?')[0]} → {detalle}")


def obtener_token(tenant: str, client_id: str, secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }).encode()
    j = http(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"})
    if "access_token" not in j:
        die(f"Entra ID no devolvió token: {j}")
    return j["access_token"]


def resolver_sitio(token: str, hostname: str, site_path: str) -> str:
    j = http(f"{GRAPH}/sites/{hostname}:{site_path}", token)
    if "id" not in j:
        die(f"No se pudo resolver el sitio {hostname}{site_path}: {j}")
    print(f"→ Sitio: {j.get('displayName')} ({j['id'][:40]}…)")
    return j["id"]


def buscar_archivo(token: str, site_id: str, nombre: str) -> str:
    q = urllib.parse.quote(nombre.replace("'", "''"))
    j = http(f"{GRAPH}/sites/{site_id}/drive/root/search(q='{q}')", token)
    hits = [x for x in j.get("value", []) if x.get("name", "").lower() == nombre.lower() and "file" in x]
    if not hits:
        die(f"No se encontró '{nombre}' en el sitio. Resultados: {[x.get('name') for x in j.get('value', [])][:10]}")
    if len(hits) > 1:
        rutas = [x.get("parentReference", {}).get("path", "") for x in hits]
        print(f"⚠ Hay {len(hits)} archivos llamados '{nombre}'; se usa el modificado más recientemente. Rutas: {rutas}")
        hits.sort(key=lambda x: x.get("lastModifiedDateTime", ""), reverse=True)
    h = hits[0]
    print(f"→ Archivo: {h['name']} · modificado {h.get('lastModifiedDateTime')} · {h.get('parentReference', {}).get('path', '')}")
    return h["id"]


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--out", default="/tmp/matriz_ffie.xlsx")
    args = p.parse_args()

    tenant = env("AZURE_TENANT_ID", required=True)
    client_id = env("AZURE_CLIENT_ID", required=True)
    secret = env("AZURE_CLIENT_SECRET", required=True)
    hostname = env("SP_HOSTNAME", "ffie3.sharepoint.com")
    site_path = env("SP_SITE_PATH", "/sites/SRVINNOVACIN2")
    file_path = env("SP_FILE_PATH")
    file_name = env("SP_FILE_NAME")
    if not file_path and not file_name:
        die("Define SP_FILE_PATH (ruta dentro de Documentos) o SP_FILE_NAME (nombre del archivo)")

    token = obtener_token(tenant, client_id, secret)
    print("→ Token de Entra ID obtenido")
    site_id = resolver_sitio(token, hostname, site_path)

    if file_path:
        ruta = urllib.parse.quote(file_path.strip("/"))
        meta = http(f"{GRAPH}/sites/{site_id}/drive/root:/{ruta}", token)
        print(f"→ Archivo: {meta.get('name')} · modificado {meta.get('lastModifiedDateTime')} · {meta.get('size')} bytes")
        url = f"{GRAPH}/sites/{site_id}/drive/root:/{ruta}:/content"
    else:
        item_id = buscar_archivo(token, site_id, file_name)
        url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/content"

    contenido = http(url, token, raw=True)
    if not contenido or contenido[:2] != b"PK":
        die("La descarga no es un .xlsx válido (no empieza con firma ZIP). ¿La ruta apunta al archivo correcto?")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "wb") as f:
        f.write(contenido)
    print(f"✓ Matriz descargada: {args.out} ({len(contenido):,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
