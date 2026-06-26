# -*- coding: utf-8 -*-
"""
sources/atlantida.py  —  Atlántida Capital, S.A.

Fuente: API JSON del sitio (Strapi):
    GET https://atlantidacapital.com.sv/api/fondos/valor-cuota-fondo/{cod}?page={n}&size={s}
    -> { "data": [ {codFondo, fechaProceso, valorCuota, patrimonio, ...}, ... ],
         "meta": { "pagination": { "page", "pageSize", "pageCount", "total" } } }

NOTA TLS: el certificado del sitio llega con cadena incompleta (falta el intermedio),
así que se intenta verificar normal y, si falla, se reintenta sin verificación
(solo se leen datos públicos de valor cuota; no se envían credenciales).

ESTADO: en validación. Arrancamos con el fondo 001 (Liquidez a Corto Plazo) como
prueba de que GitHub Actions pasa el WAF. Al confirmarse, se agregan los otros 7
a la lista FUNDS (solo hace falta su código 00X y su nombre).

Contrato de salida idéntico a sources/sgb.py.
"""
from __future__ import annotations
import time

import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

GESTORA = "Atlántida Capital"
GESTORA_SLUG = "atlantida"
PENDIENTE = False  # ya implementado (al menos el fondo 001)

API = "https://atlantidacapital.com.sv/api/fondos/valor-cuota-fondo/{cod}"

# (codFondo, nombre, tipo). Confirmado: 001 = Liquidez a Corto Plazo.
# Cuando validemos el resto, se agregan aquí, p.ej.:
#   ("002", "FIA Atlántida de Crecimiento a Mediano Plazo", "Abierto"),
FUNDS = [
    ("001", "FIA Atlántida de Liquidez a Corto Plazo", "Abierto"),
]

HEAD = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Referer": "https://atlantidacapital.com.sv/",
    "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}

SIZE = 500
TIMEOUT = 35
MAXPAGES = 300


class AtlantidaSourceError(RuntimeError):
    """Error al obtener/parsear un fondo de Atlántida (lo captura el orquestador)."""


def _get(url: str) -> requests.Response:
    """GET con verificación TLS; si la cadena está incompleta, reintenta sin verificar."""
    try:
        return requests.get(url, headers=HEAD, timeout=TIMEOUT)
    except requests.exceptions.SSLError:
        return requests.get(url, headers=HEAD, timeout=TIMEOUT, verify=False)


def _fetch_fondo(cod: str) -> pd.DataFrame:
    rows = []
    page = 1
    while page <= MAXPAGES:
        url = f"{API.format(cod=cod)}?page={page}&size={SIZE}"
        r = _get(url)
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
    df["fecha"] = pd.to_datetime(df["fecha"])  # fechaProceso viene ISO yyyy-mm-dd
    return df.drop_duplicates("fecha").sort_values("fecha").reset_index(drop=True)


def get_fondos() -> list[dict]:
    out = []
    for cod, nombre, tipo in FUNDS:
        df = _fetch_fondo(cod)
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
    for f in get_fondos():
        print(f"OK {f['nombre']} (cod {f['slug']}): {f['n_dias']} días, "
              f"inicio {f['inicio']}, último {f['data_through']}, "
              f"VC {f['df']['vc'].iloc[-1]:.6f}, patrimonio ${f['last_pat']:,.2f}")
