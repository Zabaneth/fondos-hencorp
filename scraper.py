#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Orquestador multi-gestora.
Corre las fuentes de cada gestora (sources/*.py), calcula rendimientos con la
metodología del proyecto y genera:
  - Un Excel por gestora:        hencorp.xlsx · sgb.xlsx · atlantida.xlsx · banagricola.xlsx
  - Un Excel consolidado:        consolidado.xlsx  (todas las gestoras juntas)
  - data.json:                   agrupado por gestora + lista plana 'funds' (para el dashboard)

Resiliencia POR GESTORA: si una fuente falla o está pendiente, las demás se
actualizan igual; la gestora conserva su Excel anterior y, si hay data.json
previo, sus datos previos en el dashboard.
"""
import json
import sys
import datetime

import pandas as pd

import build_excel
from sources import hencorp, sgb, atlantida, banagricola

ABBR = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}


# ----- Cálculo de rendimientos (reutilizado del scraper original, intacto) -----
def fund_json(df):
    me_vc = df.set_index("fecha")["vc"].resample("ME").last().dropna()
    me_pat = df.set_index("fecha")["patrimonio"].resample("ME").last()
    base = float(df["vc"].iloc[0]); ret = me_vc.pct_change()
    ret.iloc[0] = me_vc.iloc[0] / base - 1
    month_end = [{"ym": d.strftime("%Y-%m"), "label": f"{ABBR[d.month]}-{d.year}",
                  "vc": round(float(v), 8),
                  "pat": (None if pd.isna(me_pat.get(d)) else round(float(me_pat.get(d)), 2)),
                  "ret": (None if pd.isna(r) else round(float(r), 6))}
                 for (d, v), r in zip(me_vc.items(), ret.values)]
    annual = []
    dec = {d.year: v for d, v in me_vc.items() if d.month == 12}
    last = me_vc.index[-1]
    for y in range(me_vc.index[0].year, last.year + 1):
        if y == last.year:
            prev = dec.get(y - 1, base if me_vc.index[0].year == y else None)
            endv = me_vc.iloc[-1]; ytd = True
        else:
            if y not in dec:
                continue
            prev = dec.get(y - 1, base if me_vc.index[0].year == y else None); endv = dec[y]; ytd = False
        if prev:
            annual.append({"year": y, "ret": round(float(endv / prev - 1), 6), "ytd": ytd})
    return dict(base=base, inicio=df["fecha"].min().strftime("%Y-%m-%d"),
                last_date=df["fecha"].max().strftime("%Y-%m-%d"),
                last_vc=round(float(df["vc"].iloc[-1]), 8),
                last_pat=(None if pd.isna(df["patrimonio"].iloc[-1]) else round(float(df["patrimonio"].iloc[-1]), 2)),
                nmonths=len(me_vc), acum=round(float(df["vc"].iloc[-1] / base - 1), 6),
                month_end=month_end, annual=annual)


# --------------------------- Configuración por gestora ---------------------------
GESTORAS = [
    dict(slug="hencorp", nombre="Hencorp Gestora de Fondos de Inversión", mod=hencorp,
         xlsx="hencorp.xlsx", download="Fondos Hencorp.xlsx",
         sub_prefix="Grupo Hencorp — Gestora de Fondos | Fuente: hencorpgestora.com (Comportamiento Histórico)",
         titulo="Resumen — Fondos de Inversión Hencorp",
         fuente_nota="Fuente: Hencorp Gestora de Fondos de Inversión, S.A. — sección 'Informe Valor Cuota' (PDF 'Comportamiento Histórico del Fondo') en hencorpgestora.com. Archivo regenerado automáticamente.",
         notas=[("ADVERTENCIA (fondos cerrados):", 11, "C00000", True),
                ("• Renta Fija I DISTRIBUYE en mayo y noviembre: la cuota baja esos meses, por lo que su rendimiento medido SOLO desde la cuota SUBESTIMA el total del partícipe (lo distribuido se paga en efectivo).", 10, "000000", False),
                ("• Growth y Vivienda 01 registran REVALUACIONES inmobiliarias (saltos grandes) y bajas por distribución/markdown. Commercial Properties tiene historia mínima (inició may-2026).", 10, "000000", False)]),
    dict(slug="sgb", nombre="SGB Fondos de Inversión", mod=sgb,
         xlsx="sgb.xlsx", download="Fondos SGB.xlsx",
         sub_prefix="Grupo SGB (Servicios Generales Bursátiles) | Fuente: sgbfondosdeinversion.com (histórico CSV)",
         titulo="Resumen — Fondos de Inversión SGB",
         fuente_nota="Fuente: SGB Fondos de Inversión, S.A. — histórico público (CSV) por fondo en sgbfondosdeinversion.com. Archivo regenerado automáticamente.",
         notas=[("NOTA (fondos abiertos):", 11, "1F3864", True),
                ("• Ambos fondos son ABIERTOS y acumulan limpio: el rendimiento de la cuota equivale al rendimiento total del partícipe (no hay distribuciones que bajen la cuota).", 10, "000000", False)]),
    dict(slug="atlantida", nombre="Atlántida Capital", mod=atlantida,
         xlsx="atlantida.xlsx", download="Fondos Atlántida.xlsx",
         sub_prefix="Atlántida Capital | Fuente: atlantidacapital.com.sv",
         titulo="Resumen — Fondos de Inversión Atlántida",
         fuente_nota="Fuente: Atlántida Capital, S.A. — atlantidacapital.com.sv.",
         notas=[]),
    dict(slug="banagricola", nombre="Gestora Banagrícola", mod=banagricola,
         xlsx="banagricola.xlsx", download="Fondos Banagrícola.xlsx",
         sub_prefix="Gestora Banagrícola | Fuente: gestorabanagricola.com",
         titulo="Resumen — Fondos de Inversión Banagrícola",
         fuente_nota="Fuente: Gestora de Fondos de Inversión Banagrícola, S.A. — gestorabanagricola.com.",
         notas=[]),
]

CONSOL_XLSX = "consolidado.xlsx"
CONSOL_DOWNLOAD = "Fondos de Inversión El Salvador — Todas las Gestoras.xlsx"


def _key(slug, fund_slug):
    return f"{slug}__{fund_slug}"


def main():
    # data.json previo (para fallback de gestoras degradadas)
    try:
        prev = json.load(open("data.json", encoding="utf-8"))
        prev_g = {g["slug"]: g for g in prev.get("gestoras", [])}
    except Exception:
        prev_g = {}

    today = datetime.date.today().isoformat()
    gestoras_out, flat_funds = [], []
    consol_data, consol_order, consol_disp, consol_tipo = {}, [], {}, {}

    for G in GESTORAS:
        slug = G["slug"]; mod = G["mod"]
        pendiente = getattr(mod, "PENDIENTE", False)
        try:
            fondos = [] if pendiente else mod.get_fondos()
        except Exception as ex:
            fondos = None
            print(f"DEGRADADO {slug}: {ex}", file=sys.stderr)

        if fondos:  # ----- OK: hay datos frescos -----
            data_g, order_g, disp_g, tipo_g = {}, [], {}, {}
            funds_g = []
            for f in fondos:
                k = _key(slug, f["slug"])
                data_g[k] = f["df"]; order_g.append(k)
                disp_g[k] = f["nombre"]; tipo_g[k] = str(f["tipo"]).capitalize()
                fj = fund_json(f["df"])
                fj.update(key=k, name=f["nombre"], tipo=str(f["tipo"]).capitalize(),
                          gestora=G["nombre"], gestora_slug=slug, slug=f["slug"])
                if f.get("pdf"):
                    fj["pdf"] = f["pdf"]
                funds_g.append(fj)
                # acumular al consolidado (con etiqueta de gestora)
                gtag = {"hencorp": "Hencorp", "sgb": "SGB",
                        "atlantida": "Atlántida", "banagricola": "Banagrícola"}[slug]
                consol_data[k] = f["df"]; consol_order.append(k)
                consol_disp[k] = f"{gtag} · {f['nombre']}"; consol_tipo[k] = str(f["tipo"]).capitalize()

            corte = f"datos al {max(x['last_date'] for x in funds_g)} (actualización {today})"
            build_excel.build(data_g, G["xlsx"], corte_label=corte,
                              order=order_g, disp=disp_g, tipo=tipo_g,
                              sub_prefix=G["sub_prefix"], titulo_resumen=G["titulo"],
                              fuente_nota=G["fuente_nota"],
                              notas_fondos=(G["notas"] or None))
            flat_funds.extend(funds_g)
            gestoras_out.append(dict(slug=slug, nombre=G["nombre"], estado="ok",
                                     data_through=max(x["last_date"] for x in funds_g),
                                     n_fondos=len(funds_g), excel=G["xlsx"], download=G["download"]))
            print(f"OK {slug}: {len(funds_g)} fondos -> {G['xlsx']}")

        elif fondos is None:  # ----- DEGRADADO: conservar datos previos -----
            pg = prev_g.get(slug)
            if pg and pg.get("estado") == "ok":
                flat_funds.extend([f for f in prev.get("funds", []) if f.get("gestora_slug") == slug])
                pg2 = dict(pg); pg2["estado"] = "degradado"
                gestoras_out.append(pg2)
            else:
                gestoras_out.append(dict(slug=slug, nombre=G["nombre"], estado="degradado",
                                         n_fondos=0, excel=G["xlsx"], download=G["download"]))

        else:  # ----- PENDIENTE -----
            gestoras_out.append(dict(slug=slug, nombre=G["nombre"], estado="pendiente",
                                     motivo=getattr(mod, "MOTIVO", ""),
                                     n_fondos=0, excel=G["xlsx"], download=G["download"]))
            print(f"PENDIENTE {slug}: {getattr(mod,'MOTIVO','')}")

    # ----- Consolidado (con las gestoras frescas de esta corrida) -----
    if consol_order:
        maxd = max(f["last_date"] for f in flat_funds if "last_date" in f)
        corte = f"datos al {maxd} (actualización {today})"
        build_excel.build(consol_data, CONSOL_XLSX, corte_label=corte,
                          order=consol_order, disp=consol_disp, tipo=consol_tipo,
                          sub_prefix="Fondos de Inversión El Salvador — Todas las gestoras | Fuente: sitios oficiales de cada gestora",
                          titulo_resumen="Resumen — Todas las gestoras",
                          fuente_nota="Fuente: sitios públicos de cada gestora (Hencorp: PDF Comportamiento Histórico; SGB: histórico CSV). Archivo regenerado automáticamente.",
                          notas_fondos=[("ADVERTENCIA — los fondos no son directamente comparables entre sí:", 11, "C00000", True),
                                        ("• Mezclan tipos (abiertos de liquidez, cerrados inmobiliarios, etc.) y bases distintas; compare rendimientos en %, no valores cuota.", 10, "000000", False),
                                        ("• Renta Fija I (Hencorp) distribuye en mayo/noviembre: su rendimiento medido desde la cuota subestima el total del partícipe.", 10, "000000", False)])
        print(f"OK consolidado: {len(consol_order)} fondos -> {CONSOL_XLSX}")

    data_through = max([f["last_date"] for f in flat_funds if "last_date" in f], default=today)
    payload = {
        "updated_at": today,
        "data_through": data_through,
        "source": "Sitios públicos de las gestoras de fondos de inversión de El Salvador",
        "gestoras": gestoras_out,
        "consolidado": {"excel": CONSOL_XLSX, "download": CONSOL_DOWNLOAD,
                        "disponible": bool(consol_order)},
        "funds": flat_funds,
    }
    json.dump(payload, open("data.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\nListo. data_through={data_through}. "
          f"{sum(1 for g in gestoras_out if g['estado']=='ok')} gestora(s) OK, "
          f"{len(flat_funds)} fondos en total.")


if __name__ == "__main__":
    main()
