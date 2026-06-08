#!/usr/bin/env python3
"""Robot Hencorp: descubre el PDF 'Comportamiento Histórico' de cada fondo,
lo parsea, calcula rendimientos y escribe data.json + el Excel. Idempotente."""
import re, io, json, sys, datetime
import requests, pdfplumber, pandas as pd
import build_excel

HEAD={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Referer":"https://www.hencorpgestora.com/"}
FUNDS=[
 ("01_opportunity","FI Abierto Hencorp Opportunity","Abierto","https://www.hencorpgestora.com/fondo-de-inversion-abierto-a-corto-plazo/"),
 ("02_rentafija1","FIC Renta Fija I","Cerrado","https://www.hencorpgestora.com/fondo-de-inversion-cerrado-renta-fija-i/"),
 ("03_growth","FIC Inmobiliario Hencorp Growth","Cerrado","https://www.hencorpgestora.com/fondo_de_inversion_cerrado_inmobiliario_hencorp_growth/"),
 ("04_vivienda01","FIC Desarrollo Inmob. Hencorp Vivienda 01","Cerrado","https://www.hencorpgestora.com/fondo-de-inversion-inmobiliario-vivienda-01/"),
 ("05_bluewhale","FIC Inmobiliario Hencorp Blue Whale","Cerrado","https://www.hencorpgestora.com/fondo-de-inversion-cerrado-inmobiliario-hencorp-blue-whale/"),
 ("06_commercial","FIC Inmobiliario Hencorp Commercial Properties","Cerrado","https://www.hencorpgestora.com/fondo-inversion-cerrado-inmobiliario-hencorp-commercial-properties/"),
]
DATE_RE=re.compile(r'(\d{1,2}/\d{1,2}/\d{4})\s+(\d+\.\d+)')
PAT_RE=re.compile(r'\$\s*([\d][\d\s,]*?\.\d{2})(?!\d)')

def find_pdf(page_url):
    html=requests.get(page_url,headers=HEAD,timeout=60).text
    links=re.findall(r'https://www\.hencorpgestora\.com/wp-content/uploads/[^"\' ]+\.pdf',html)
    cand=sorted({L for L in links if re.search(r'comportamiento.*histor',L,re.I)})
    return cand[0] if cand else None

def parse_pdf(content):
    rows=[]
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for pg in pdf.pages:
            for line in (pg.extract_text() or "").split("\n"):
                m=DATE_RE.match(line.strip())
                if not m: continue
                d=pd.to_datetime(m.group(1),dayfirst=True); vc=float(m.group(2))
                rest=line.strip()[m.end():]; pm=PAT_RE.search(rest)
                patr=float(pm.group(1).replace(" ","").replace(",","")) if pm else None
                rows.append((d,vc,patr))
    if not rows: raise ValueError("sin filas parseadas")
    return pd.DataFrame(rows,columns=["fecha","vc","patrimonio"]).drop_duplicates("fecha").sort_values("fecha").reset_index(drop=True)

ABBR={1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}

def fund_json(df):
    me_vc=df.set_index("fecha")["vc"].resample("ME").last().dropna()
    me_pat=df.set_index("fecha")["patrimonio"].resample("ME").last()
    base=float(df["vc"].iloc[0]); ret=me_vc.pct_change()
    ret.iloc[0]=me_vc.iloc[0]/base-1
    month_end=[{"ym":d.strftime("%Y-%m"),"label":f"{ABBR[d.month]}-{d.year}",
                "vc":round(float(v),8),
                "pat":(None if pd.isna(me_pat.get(d)) else round(float(me_pat.get(d)),2)),
                "ret":(None if pd.isna(r) else round(float(r),6))}
               for (d,v),r in zip(me_vc.items(),ret.values)]
    annual=[]
    dec={d.year:v for d,v in me_vc.items() if d.month==12}
    last=me_vc.index[-1]
    for y in range(me_vc.index[0].year, last.year+1):
        if y==last.year:
            prev=dec.get(y-1, base if me_vc.index[0].year==y else None)
            endv=me_vc.iloc[-1]; ytd=True
        else:
            if y not in dec: continue
            prev=dec.get(y-1, base if me_vc.index[0].year==y else None); endv=dec[y]; ytd=False
        if prev: annual.append({"year":y,"ret":round(float(endv/prev-1),6),"ytd":ytd})
    return dict(base=base, inicio=df["fecha"].min().strftime("%Y-%m-%d"),
               last_date=df["fecha"].max().strftime("%Y-%m-%d"),
               last_vc=round(float(df["vc"].iloc[-1]),8),
               last_pat=(None if pd.isna(df["patrimonio"].iloc[-1]) else round(float(df["patrimonio"].iloc[-1]),2)),
               nmonths=len(me_vc), acum=round(float(df["vc"].iloc[-1]/base-1),6),
               month_end=month_end, annual=annual)

def main():
    data={}; out_funds=[]; errors=[]
    for key,name,tipo,url in FUNDS:
        try:
            pdfL=find_pdf(url)
            if not pdfL: raise ValueError("no se encontró PDF de Comportamiento Histórico")
            content=requests.get(pdfL,headers=HEAD,timeout=180).content
            df=parse_pdf(content); data[key]=df
            fj=fund_json(df); fj.update(key=key,name=name,tipo=tipo,pdf=pdfL)
            out_funds.append(fj)
            print(f"OK {name}: {len(df)} filas, último {fj['last_date']}")
        except Exception as ex:
            errors.append(f"{name}: {ex}"); print(f"ERROR {name}: {ex}", file=sys.stderr)
    if len(out_funds)<6:
        # No sobrescribir si falló algún fondo: conserva el data.json previo
        print(f"Solo {len(out_funds)}/6 OK. No se sobrescribe. Errores: {errors}", file=sys.stderr)
        sys.exit(1)
    today=datetime.date.today().isoformat()
    maxd=max(f["last_date"] for f in out_funds)
    payload={"updated_at":today,"data_through":maxd,"source":"hencorpgestora.com (Comportamiento Histórico)","funds":out_funds}
    json.dump(payload, open("data.json","w"), ensure_ascii=False, indent=1)
    build_excel.build(data, "fondos-hencorp-rendimientos.xlsx",
                      corte_label=f"datos al {maxd} (actualización {today})")
    print(f"\nListo. data_through={maxd}. data.json + Excel generados.")

if __name__=="__main__":
    main()
