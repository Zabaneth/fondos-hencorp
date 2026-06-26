# -*- coding: utf-8 -*-
"""
sources/atlantida.py  —  Atlántida Capital, S.A.

Fuente: API JSON del sitio (Strapi):
    GET https://atlantidacapital.com.sv/api/fondos/valor-cuota-fondo/{cod}?page={n}&size={s}
    -> { "data": [ {codFondo, fechaProceso, valorCuota, patrimonio, ...}, ... ],
         "meta": { "pagination": { "page", "pageSize", "pageCount", "total" } } }

El sitio está tras un WAF (F5) que bloquea peticiones "de servidor" y además
entrega una cadena TLS incompleta. Para pasarlo se usa `curl_cffi` imitando la
huella TLS de Chrome (impersonate). Si `curl_cffi` no está instalado, se cae a
`requests` (probablemente bloqueado, pero NO rompe el resto del pipeline).

Estrategia: una sesión "calienta" cargando una página del sitio (para obtener la
cookie del WAF) y luego pega a la API JSON, paginando.

Contrato de salida idéntico a sources/sgb.py.
"""
from __future__ import annotations
import time

import pandas as pd

# --- HTTP: curl_cffi (Chrome TLS) con fallback seguro a requests ---
try:
    from curl_cffi import requests as _http
    _IMPERSONATE = "chrome124"
    _CFFI = True
except Exception:  # pragma: no cover
    import requests as _http
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    _IMPERSONATE = None
    _CFFI = False

GESTORA = "Atlántida Capital"
GESTORA_SLUG = "atlantida"
PENDIENTE = False

BASE = "https://atlantidacapital.com.sv"
API = BASE + "/api/fondos/valor-cuota-fondo/{cod}"
WARMUP = BASE + "/fondo-de-inversion-abierto-atlantida-de-liquidez-a-corto-plazo/"

# (codFondo, nombre, tipo). Confirmado: 001 = Liquidez a Corto Plazo.
# Al validar el resto se agregan, p.ej.: ("002","FIA Atlántida de Crecimiento a Mediano Plazo","Abierto"),
FUNDS = [
    ("001", "FIA Atlántida de Liquidez a Corto Plazo", "Abierto"),
]

HEAD = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Referer": BASE + "/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}

SIZE = 500
TIMEOUT = 35
MAXPAGES = 300


class AtlantidaSourceError(RuntimeError):
    """Error al obtener/parsear un fondo de Atlántida (lo captura el orquestador)."""


def _session():
    if _CFFI:
        s = _http.Session(impersonate=_IMPERSONATE)
    else:
        s = _http.Session()
    s.headers.update(HEAD)
    return s


def _get(sess, url):
    return sess.get(url, timeout=TIMEOUT, verify=False)


def _fetch_fondo(sess, cod: str) -> pd.DataFrame:
    rows = []
    page = 1
    while page <= MAXPAGES:
        url = f"{API.format(cod=cod)}?page={page}&size={SIZE}"
        r = _get(sess, url)
        if r.status_code != 200:
            raise AtlantidaSourceError(f"HTTP {r.status_code} en {cod} p{page} (¿WAF/bloqueo?)")
        try:
            js = r.json()
        except Exception:
            raise AtlantidaSourceError(f"respuesta no-JSON en {cod} (probable reto del WAF)")
        data = js.get("data") or []
        if not data:
            break
        for d in data:
            f = d.get("fechaProceso")
            vc = d.get("valorCuota")
            pat = d.get("patrimonio")
            if f and vc is not None:
                rows.append((f, float(vc), None if pat is None else float(pat)))
        meta = (js.get("meta") or {}).get("pagination") or {}
        pc = meta.get("pageCount")
        if pc and page >= pc:
            break
        if not pc and len(data) < SIZE:
            break
        page += 1
        time.sleep(0.3)
    if not rows:
        raise AtlantidaSourceError(f"sin filas para {cod}")
    df = pd.DataFrame(rows, columns=["fecha", "vc", "patrimonio"])
    df["fecha"] = pd.to_datetime(df["fecha"])  # ISO yyyy-mm-dd
    return df.drop_duplicates("fecha").sort_values("fecha").reset_index(drop=True)


def get_fondos() -> list[dict]:
    sess = _session()
    # "calentar" la sesión para obtener la cookie del WAF (best-effort)
    try:
        sess.get(WARMUP, timeout=TIMEOUT, verify=False)
    except Exception:
        pass
    out = []
    for cod, nombre, tipo in FUNDS:
        df = _fetch_fondo(sess, cod)
        lp = df["patrimonio"].dropna()
        out.append({
            "gestora": GESTORA,
            "gestora_slug": GESTORA_SLUG,
            "slug": f"f{cod}",
            "nombre": nombre,
            "tipo": tipo,
            "base": float(df["vc"].iloc[0]),
            "inicio": df["fecha"].min().strftime("%Y-%m-%d"),
            "data_through": df["fecha"].max().strftime("%Y-%m-%d"),
            "n_dias": int(len(df)),
            "last_pat": float(lp.iloc[-1]) if not lp.empty else None,
            "df": df,
        })
    return out


if __name__ == "__main__":
    print("curl_cffi activo:", _CFFI)
    for f in get_fondos():
        print(f"OK {f['nombre']} (cod {f['slug']}): {f['n_dias']} días, "
              f"inicio {f['inicio']}, último {f['data_through']}, "
              f"VC {f['df']['vc'].iloc[-1]:.6f}, patrimonio ${f['last_pat']:,.2f}")
