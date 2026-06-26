# -*- coding: utf-8 -*-
"""
sources/atlantida.py  —  Atlántida Capital, S.A.   (MODO: Excel manual / Plan C)

El sitio de Atlántida está tras un WAF (F5) que bloquea el acceso automático
desde servidores. En vez de pelear con eso, este módulo LEE los archivos Excel
que vos exportás desde el sitio ("Exportar Excel" en cada fondo) y subís al repo.

CÓMO USARLO
-----------
1. En atlantidacapital.com.sv, entrá a cada fondo -> "Histórico de Valor Cuota"
   -> "Exportar Excel". Se descarga un archivo tipo:
       HISTORICO_VALOR_CUOTA_<NOMBRE_DEL_FONDO>-export-AAAA-MM-DD.xlsx
2. Subí ese archivo a la RAÍZ del repositorio (tal cual, sin renombrar).
   - Podés subir varios (uno por fondo).
   - Para refrescar, subí un export más nuevo: el pipeline usa SIEMPRE el más
     reciente por fondo (no hace falta borrar el viejo).
   - (Opcional) Si preferís ordenarlos, también lee los .xlsx que pongas en una
     carpeta llamada  atlantida_data/

El módulo detecta el nombre del fondo desde el nombre del archivo, lee la
columna "Valor cuota($)" y "Patrimonio ($)", y arma el histórico.

Contrato de salida idéntico a sources/sgb.py.
"""
from __future__ import annotations
import os
import re
import glob
import datetime
import unicodedata

import pandas as pd

GESTORA = "Atlántida Capital"
GESTORA_SLUG = "atlantida"
PENDIENTE = False
MOTIVO = ("Subí al repositorio el/los Excel de Atlántida "
          "(HISTORICO_VALOR_CUOTA…-export-AAAA-MM-DD.xlsx) para incorporarlos.")

# Patrones donde buscar los Excel exportados.
PATTERNS = [
    "HISTORICO_VALOR_CUOTA*.xlsx",
    "historico_valor_cuota*.xlsx",
    "atlantida_data/*.xlsx",
    "atlantida_data/*.XLSX",
]

# Opcional: sobreescribir el nombre que sale del archivo por uno "oficial".
# clave = nombre derivado (en minúsculas) ; valor = nombre a mostrar.
OVERRIDES = {
    # "fondo de inversión abierto de crecimiento": "FIA Atlántida de Crecimiento a Mediano Plazo",
}

_MINUS = {"de", "del", "la", "las", "los", "el", "a", "y", "e", "o", "u",
          "en", "con", "por", "para"}


class AtlantidaSourceError(RuntimeError):
    """Error al leer/parsear un Excel de Atlántida (lo captura el orquestador)."""


def _titlecase_es(s: str) -> str:
    out = []
    for i, w in enumerate(s.split()):
        wl = w.lower()
        out.append(wl if (i > 0 and wl in _MINUS) else (w[:1].upper() + w[1:].lower()))
    return " ".join(out)


def _name_from_filename(path: str) -> str:
    base = os.path.basename(path)
    base = re.sub(r"\.xlsx$", "", base, flags=re.I)
    base = re.sub(r"^historico_valor_cuota[_\- ]*", "", base, flags=re.I)
    base = re.sub(r"[_\- ]*export[_\- ]*\d{4}-\d{2}-\d{2}.*$", "", base, flags=re.I)
    base = base.replace("_", " ").strip()
    base = re.sub(r"\s+", " ", base)
    name = _titlecase_es(base) if base else "Fondo Atlántida"
    return OVERRIDES.get(name.lower(), name)


def _export_date(path: str) -> str:
    m = re.search(r"export[_\- ]*(\d{4}-\d{2}-\d{2})", os.path.basename(path), re.I)
    if m:
        return m.group(1)
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d")


def _slug(name: str) -> str:
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower()
    return s or "fondo"


def _find_col(cols, *needles):
    for c in cols:
        cl = re.sub(r"[\s$()]", "", str(c).lower())
        if all(n in cl for n in needles):
            return c
    return None


def _read_excel(path: str, nombre: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0, dtype=str)
    cols = list(raw.columns)
    c_fecha = _find_col(cols, "fecha")
    c_vc = _find_col(cols, "valorcuota")
    c_pat = _find_col(cols, "patrimonio")
    if not c_fecha or not c_vc:
        raise AtlantidaSourceError(
            f"'{os.path.basename(path)}': no encuentro columnas Fecha/Valor cuota "
            f"(columnas: {cols[:6]}…)")
    df = pd.DataFrame({
        "fecha": pd.to_datetime(raw[c_fecha], format="%d/%m/%Y", errors="coerce"),
        "vc": pd.to_numeric(raw[c_vc], errors="coerce"),
        "patrimonio": (pd.to_numeric(raw[c_pat], errors="coerce")
                       if c_pat else pd.Series([None] * len(raw))),
    })
    df = (df.dropna(subset=["fecha", "vc"])
            .drop_duplicates("fecha").sort_values("fecha").reset_index(drop=True))
    if df.empty:
        raise AtlantidaSourceError(f"'{os.path.basename(path)}': sin filas válidas.")
    return df


def _collect_files():
    """Devuelve [(path, nombre)] tomando el export más reciente por fondo."""
    paths = []
    for pat in PATTERNS:
        paths.extend(glob.glob(pat))
    paths = sorted(set(paths))
    groups = {}  # nombre.lower() -> (export_date, path, nombre)
    for p in paths:
        nombre = _name_from_filename(p)
        ed = _export_date(p)
        key = nombre.lower()
        if key not in groups or ed > groups[key][0]:
            groups[key] = (ed, p, nombre)
    return [(p, nombre) for (_, p, nombre) in groups.values()]


def get_fondos() -> list[dict]:
    files = _collect_files()
    if not files:
        return []  # -> el orquestador lo marca "pendiente" con MOTIVO

    out, errores = [], []
    for path, nombre in sorted(files, key=lambda x: x[1]):
        try:
            df = _read_excel(path, nombre)
        except Exception as ex:
            errores.append(str(ex))
            continue
        lp = df["patrimonio"].dropna()
        out.append({
            "gestora": GESTORA,
            "gestora_slug": GESTORA_SLUG,
            "slug": _slug(nombre),
            "nombre": nombre,
            "tipo": "Abierto",
            "base": float(df["vc"].iloc[0]),
            "inicio": df["fecha"].min().strftime("%Y-%m-%d"),
            "data_through": df["fecha"].max().strftime("%Y-%m-%d"),
            "n_dias": int(len(df)),
            "last_pat": float(lp.iloc[-1]) if not lp.empty else None,
            "df": df,
        })

    if not out:
        # había archivos pero ninguno se pudo leer -> degradado (problema real)
        raise AtlantidaSourceError("; ".join(errores) or "no se pudo leer ningún Excel.")
    return out


if __name__ == "__main__":
    fs = get_fondos()
    if not fs:
        print("Sin Excel de Atlántida. ", MOTIVO)
    for f in fs:
        print(f"OK {f['nombre']} ({f['slug']}): {f['n_dias']} días, "
              f"{f['inicio']} → {f['data_through']}, VC {f['df']['vc'].iloc[-1]:.6f}, "
              f"patrimonio ${f['last_pat']:,.2f}")
