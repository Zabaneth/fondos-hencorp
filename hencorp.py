# -*- coding: utf-8 -*-
"""
sources/sgb.py
--------------
Fuente de datos: SGB Fondos de Inversión, S.A., Gestora de Fondos de Inversión.

SGB publica el histórico COMPLETO de cada fondo en un CSV directo y público:
    https://www.sgbfondosdeinversion.com/images/historico-<codigo>.csv

Es la fuente más limpia de las cuatro gestoras (no hay que parsear PDF ni JS).
Verificado contra datos reales:
    - fiarcp: 3,541 días desde 12/10/2016 (base 1.0)
    - fia180: 2,995 días desde 11/04/2018 (base 1.0)

Ambos fondos son ABIERTOS y acumulan limpio (no distribuyen dividendos que
bajen la cuota), así que el rendimiento de la cuota = rendimiento total del
partícipe — análogo a Hencorp Opportunity. No requieren la nota de dividendos
de Renta Fija I.

CONTRATO DE SALIDA (común a todos los módulos de sources/):
    get_fondos() -> list[dict], cada uno:
        {
          "gestora":      "SGB",
          "gestora_slug": "sgb",
          "slug":         "fiarcp",
          "nombre":       "FI Abierto Rentable Corto Plazo",
          "tipo":         "abierto",                 # abierto | cerrado
          "base":         1.0,                        # valor cuota inicial (inferido)
          "inicio":       "2016-10-12",               # primera fecha disponible
          "data_through": "2026-06-22",               # última fecha disponible
          "n_dias":       3541,
          "df":           DataFrame[fecha(datetime64[ns]), vc(float), patrimonio(float)]
        }                                            #   -> orden ASCENDENTE por fecha

El orquestador (scraper.py) toma estos DataFrames diarios y calcula los cierres
de mes / rendimientos con la metodología compartida del proyecto.
"""

from __future__ import annotations
import io
import time
import datetime as dt

import requests
import pandas as pd

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
GESTORA = "SGB"
GESTORA_SLUG = "sgb"
BASE_URL = "https://www.sgbfondosdeinversion.com/images/historico-{code}.csv"

# El servidor responde 404 a peticiones "sospechosas"; con UA de navegador
# + Referer del propio sitio entrega el CSV sin problema (mismo patrón que Hencorp).
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": "https://www.sgbfondosdeinversion.com/",
    "Accept": "text/csv,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9",
}

# Catálogo de fondos de SGB (los 2 que publica). Si SGB lanza otro fondo,
# basta agregar aquí su <codigo> (el de la URL del CSV) y su nombre/tipo.
FONDOS = {
    "fiarcp": {"nombre": "FI Abierto Rentable Corto Plazo", "tipo": "abierto"},
    "fia180": {"nombre": "FI Abierto Plazo 180",           "tipo": "abierto"},
}

TIMEOUT = 30
REINTENTOS = 3


class SGBSourceError(RuntimeError):
    """Error al obtener/parsear datos de SGB (lo captura el orquestador
    para marcar la gestora como 'degradada' sin tumbar toda la corrida)."""


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _limpiar_num(x) -> float | None:
    """'$78,876,405.22' -> 78876405.22 ; '4.146%' -> 4.146 ; ''/'-' -> None."""
    if x is None:
        return None
    s = (
        str(x)
        .strip()
        .replace("$", "")
        .replace(",", "")
        .replace("%", "")
        .replace("\xa0", "")
        .strip()
    )
    if s in ("", "-", "--", "N/A", "n/a", "nan", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _col(cols, *claves) -> str | None:
    """Encuentra una columna por nombre (todas las 'claves' presentes, sin
    importar mayúsculas/acentos del orden). Robusto a cambios de posición."""
    for c in cols:
        cl = str(c).lower()
        if all(k in cl for k in claves):
            return c
    return None


def _descargar_csv(code: str) -> str:
    url = BASE_URL.format(code=code)
    ultimo_err = None
    for intento in range(1, REINTENTOS + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 200 and r.text.strip():
                r.encoding = r.apparent_encoding or "utf-8"
                return r.text
            ultimo_err = f"HTTP {r.status_code} (bytes={len(r.content)})"
        except requests.RequestException as e:
            ultimo_err = repr(e)
        time.sleep(1.5 * intento)
    raise SGBSourceError(f"[SGB:{code}] no se pudo descargar {url}: {ultimo_err}")


def _parsear(code: str, txt: str) -> pd.DataFrame:
    """CSV (newest-first, valores entre comillas, montos con $ y miles) ->
    DataFrame[fecha, vc, patrimonio] ASCENDENTE."""
    delim = ";" if txt.splitlines()[0].count(";") > txt.splitlines()[0].count(",") else ","
    df = pd.read_csv(io.StringIO(txt), sep=delim, engine="python", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    c_fecha = _col(df.columns, "fecha")
    c_vc = _col(df.columns, "valor", "cuota") or _col(df.columns, "cuota", "diari")
    c_pat = _col(df.columns, "patrimonio")
    if not (c_fecha and c_vc and c_pat):
        raise SGBSourceError(
            f"[SGB:{code}] columnas no reconocidas: {list(df.columns)[:8]}"
        )

    out = pd.DataFrame(
        {
            "fecha": pd.to_datetime(df[c_fecha].astype(str).str.strip(),
                                    dayfirst=True, errors="coerce"),
            "vc": df[c_vc].map(_limpiar_num),
            "patrimonio": df[c_pat].map(_limpiar_num),
        }
    )
    out = (
        out.dropna(subset=["fecha", "vc"])
        .drop_duplicates(subset=["fecha"], keep="last")
        .sort_values("fecha")
        .reset_index(drop=True)
    )
    if out.empty:
        raise SGBSourceError(f"[SGB:{code}] CSV sin filas válidas tras limpieza")
    return out


# ---------------------------------------------------------------------------
# API pública del módulo
# ---------------------------------------------------------------------------
def get_fondos() -> list[dict]:
    """Devuelve la lista de fondos de SGB con su serie diaria normalizada.
    Lanza SGBSourceError si algún fondo no se pudo obtener."""
    resultados: list[dict] = []
    for code, meta in FONDOS.items():
        df = _parsear(code, _descargar_csv(code))
        base = float(df["vc"].iloc[0])
        inicio = df["fecha"].iloc[0].date().isoformat()
        through = df["fecha"].iloc[-1].date().isoformat()
        last_pat = df["patrimonio"].dropna()
        resultados.append(
            {
                "gestora": GESTORA,
                "gestora_slug": GESTORA_SLUG,
                "slug": code,
                "nombre": meta["nombre"],
                "tipo": meta["tipo"],
                "base": base,
                "inicio": inicio,
                "data_through": through,
                "n_dias": int(len(df)),
                "last_pat": float(last_pat.iloc[-1]) if not last_pat.empty else None,
                "df": df,
            }
        )
    return resultados


# ---------------------------------------------------------------------------
# Autoprueba:  python sources/sgb.py
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("Probando fuente SGB...\n")
    for f in get_fondos():
        d = f["df"]
        print(f"== {f['nombre']}  ({f['slug']}, {f['tipo']}) ==")
        print(f"   {f['n_dias']} días | inicio {f['inicio']} (base {f['base']:.6f}) "
              f"| último {f['data_through']}")
        print(f"   último patrimonio = ${f['last_pat']:,.2f}")
        # cierres de mes + rendimiento mensual (misma fórmula del proyecto)
        eom = d.assign(ym=d["fecha"].dt.to_period("M")).groupby("ym").tail(1)
        eom = eom.set_index("ym")
        eom["rend_m"] = eom["vc"].pct_change()
        print("   últimos 3 cierres de mes (VC | rend. mensual):")
        for ym, row in eom.tail(3).iterrows():
            rm = "base" if pd.isna(row["rend_m"]) else f"{row['rend_m']*100:+.4f}%"
            print(f"     {ym}  VC={row['vc']:.6f}  {rm}")
        print()
    print("OK — SGB lista para integrarse al orquestador.")
