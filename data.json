# -*- coding: utf-8 -*-
"""
sources/hencorp.py
------------------
Fuente: Hencorp Gestora de Fondos de Inversión, S.A.

Reutiliza TAL CUAL la lógica del scraper original que ya funciona en producción
(descubrimiento del PDF 'Comportamiento Histórico', regex de fecha+VC y de
patrimonio, parseo con pdfplumber). No se cambió ningún regex; solo se envolvió
en la función get_fondos() que devuelve el contrato común de sources/.

CONTRATO DE SALIDA (idéntico a sources/sgb.py):
    get_fondos() -> list[dict] con: gestora, gestora_slug, slug, nombre, tipo,
    base, inicio, data_through, n_dias, last_pat, df[fecha, vc, patrimonio] (asc).
"""
from __future__ import annotations
import io
import re
import time

import requests
import pandas as pd
import pdfplumber

GESTORA = "Hencorp Gestora de Fondos de Inversión"
GESTORA_SLUG = "hencorp"

HEAD = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Referer": "https://www.hencorpgestora.com/",
}

# (slug, nombre, tipo, url de la página del fondo) — los 6 fondos oficiales.
FUNDS = [
    ("01_opportunity", "FI Abierto Hencorp Opportunity", "Abierto",
     "https://www.hencorpgestora.com/fondo-de-inversion-abierto-a-corto-plazo/"),
    ("02_rentafija1", "FIC Renta Fija I", "Cerrado",
     "https://www.hencorpgestora.com/fondo-de-inversion-cerrado-renta-fija-i/"),
    ("03_growth", "FIC Inmobiliario Hencorp Growth", "Cerrado",
     "https://www.hencorpgestora.com/fondo_de_inversion_cerrado_inmobiliario_hencorp_growth/"),
    ("04_vivienda01", "FIC Desarrollo Inmob. Hencorp Vivienda 01", "Cerrado",
     "https://www.hencorpgestora.com/fondo-de-inversion-inmobiliario-vivienda-01/"),
    ("05_bluewhale", "FIC Inmobiliario Hencorp Blue Whale", "Cerrado",
     "https://www.hencorpgestora.com/fondo-de-inversion-cerrado-inmobiliario-hencorp-blue-whale/"),
    ("06_commercial", "FIC Inmobiliario Hencorp Commercial Properties", "Cerrado",
     "https://www.hencorpgestora.com/fondo-inversion-cerrado-inmobiliario-hencorp-commercial-properties/"),
]

# --- regex ORIGINALES (no tocar): toleran día/mes de 1 dígito y el "$ 2 20,547,315.44" ---
DATE_RE = re.compile(r'(\d{1,2}/\d{1,2}/\d{4})\s+(\d+\.\d+)')
PAT_RE = re.compile(r'\$\s*([\d][\d\s,]*?\.\d{2})(?!\d)')

TIMEOUT_PAGE = 60
TIMEOUT_PDF = 180
REINTENTOS = 3


class HencorpSourceError(RuntimeError):
    """Error al obtener/parsear un fondo de Hencorp (lo captura el orquestador)."""


def find_pdf(page_url: str) -> str | None:
    html = requests.get(page_url, headers=HEAD, timeout=TIMEOUT_PAGE).text
    links = re.findall(
        r'https://www\.hencorpgestora\.com/wp-content/uploads/[^"\' ]+\.pdf', html
    )
    cand = sorted({L for L in links if re.search(r'comportamiento.*histor', L, re.I)})
    return cand[0] if cand else None


def parse_pdf(content: bytes) -> pd.DataFrame:
    rows = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pg in pdf.pages:
            for line in (pg.extract_text() or "").split("\n"):
                m = DATE_RE.match(line.strip())
                if not m:
                    continue
                d = pd.to_datetime(m.group(1), dayfirst=True)
                vc = float(m.group(2))
                rest = line.strip()[m.end():]
                pm = PAT_RE.search(rest)
                patr = float(pm.group(1).replace(" ", "").replace(",", "")) if pm else None
                rows.append((d, vc, patr))
    if not rows:
        raise HencorpSourceError("sin filas parseadas")
    return (
        pd.DataFrame(rows, columns=["fecha", "vc", "patrimonio"])
        .drop_duplicates("fecha")
        .sort_values("fecha")
        .reset_index(drop=True)
    )


def _fund(slug: str, nombre: str, tipo: str, url: str) -> dict:
    ultimo_err = None
    for intento in range(1, REINTENTOS + 1):
        try:
            pdf_url = find_pdf(url)
            if not pdf_url:
                raise HencorpSourceError("no se encontró PDF de Comportamiento Histórico")
            content = requests.get(pdf_url, headers=HEAD, timeout=TIMEOUT_PDF).content
            df = parse_pdf(content)
            last_pat = df["patrimonio"].dropna()
            return {
                "gestora": GESTORA,
                "gestora_slug": GESTORA_SLUG,
                "slug": slug,
                "nombre": nombre,
                "tipo": tipo,
                "base": float(df["vc"].iloc[0]),
                "inicio": df["fecha"].min().strftime("%Y-%m-%d"),
                "data_through": df["fecha"].max().strftime("%Y-%m-%d"),
                "n_dias": int(len(df)),
                "last_pat": float(last_pat.iloc[-1]) if not last_pat.empty else None,
                "pdf": pdf_url,
                "df": df,
            }
        except Exception as e:  # noqa: BLE001
            ultimo_err = e
            time.sleep(2 * intento)
    raise HencorpSourceError(f"[Hencorp:{slug}] {ultimo_err}")


def get_fondos() -> list[dict]:
    """Devuelve los 6 fondos de Hencorp con su serie diaria normalizada.
    Lanza HencorpSourceError si algún fondo falla (el orquestador decide qué hacer)."""
    return [_fund(slug, nombre, tipo, url) for slug, nombre, tipo, url in FUNDS]


if __name__ == "__main__":
    print("Probando fuente Hencorp...\n")
    for f in get_fondos():
        print(f"OK {f['nombre']} ({f['slug']}): {f['n_dias']} días, "
              f"inicio {f['inicio']}, último {f['data_through']}, "
              f"patrimonio ${f['last_pat']:,.2f}")
    print("\nHencorp lista.")
