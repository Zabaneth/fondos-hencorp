# Genera el workbook de 7 hojas a partir de data = {key: DataFrame(fecha, vc, patrimonio)}
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ORDER=["01_opportunity","02_rentafija1","03_growth","04_vivienda01","05_bluewhale","06_commercial"]
DISP={"01_opportunity":"FI Abierto Hencorp Opportunity","02_rentafija1":"FIC Renta Fija I",
      "03_growth":"FIC Inmobiliario Hencorp Growth","04_vivienda01":"FIC Desarrollo Inmob. Hencorp Vivienda 01",
      "05_bluewhale":"FIC Inmobiliario Hencorp Blue Whale","06_commercial":"FIC Inmobiliario Hencorp Commercial Properties"}
TIPO={"01_opportunity":"Abierto","02_rentafija1":"Cerrado","03_growth":"Cerrado","04_vivienda01":"Cerrado","05_bluewhale":"Cerrado","06_commercial":"Cerrado"}
ABBR={1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}

def build(data, out_path, corte_label="datos actualizados automáticamente"):
    F={}
    for k in ORDER:
        df=data[k].copy()
        me_vc=df.set_index("fecha")["vc"].resample("ME").last().dropna()
        me_pat=df.set_index("fecha")["patrimonio"].resample("ME").last()
        F[k]=dict(vc={d.strftime("%Y-%m"):float(v) for d,v in me_vc.items()},
                  pat={d.strftime("%Y-%m"):(None if pd.isna(v) else float(v)) for d,v in me_pat.items()},
                  base=float(df["vc"].iloc[0]), inc=me_vc.index[0].strftime("%Y-%m"),
                  first_date=df["fecha"].min().to_pydatetime(), last_date=df["fecha"].max().to_pydatetime(),
                  nmonths=len(me_vc))
    start=min(pd.Period(F[k]["inc"],freq="M") for k in ORDER)
    end=max(pd.Period(F[k]["vc"] and max(F[k]["vc"].keys()),freq="M") for k in ORDER)
    allm=pd.period_range(start,end,freq="M")
    months=[(f"{ABBR[p.month]}-{p.year}"+("*" if (p.year==end.year and p.month==end.month) else ""), f"{p.year:04d}-{p.month:02d}") for p in allm]
    ym2col={ym:get_column_letter(2+i) for i,(lab,ym) in enumerate(months)}
    last_ym=f"{end.year:04d}-{end.month:02d}"
    years=list(range(min(int(F[k]['inc'][:4]) for k in ORDER), end.year+1))

    ARIAL=lambda **kw: Font(name="Arial", **kw)
    hdr=PatternFill("solid",fgColor="1F3864"); fund_fill=PatternFill("solid",fgColor="F2F2F2")
    white=Font(name="Arial",bold=True,color="FFFFFF",size=10)
    thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    center=Alignment(horizontal="center",vertical="center"); left=Alignment(horizontal="left",vertical="center")
    BLUE="0000FF"; BLACK="000000"; GREEN="006100"
    SH_VC="Valor cuota (fin de mes)"; SH_PAT="Patrimonio (fin de mes)"; SH_MEN="Rendimiento mensual"
    SH_DIC="Cierre 31-dic (mensual)"; SH_ANU="Rendimiento anual"; SH_RES="Resumen"; SH_NOT="Notas y metodologia"
    PCT="0.00%"; VCFMT="#,##0.000000"; USD="$#,##0"; DT="dd/mm/yyyy"
    SUB=f"Grupo Hencorp — Gestora de Fondos | Fuente: hencorpgestora.com (Comportamiento Historico) | {corte_label}"
    wb=Workbook()

    def title_block(ws,title,sub,ncols):
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols)
        c=ws.cell(1,1,title); c.font=ARIAL(bold=True,size=13,color="1F3864"); c.alignment=left
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncols)
        c=ws.cell(2,1,sub); c.font=ARIAL(italic=True,size=9,color="595959"); c.alignment=left
        ws.row_dimensions[1].height=20

    def set_val(ws,r,c,f,k,ck,kind):
        cell=ws.cell(r,c); col=get_column_letter(c); base_ref=f"'{SH_RES}'!$D${4+f}"
        if kind=="vc":
            v=F[k]["vc"].get(ck)
            if v is not None: cell.value=v; cell.number_format=VCFMT; cell.font=ARIAL(color=BLUE,size=9)
        elif kind=="pat":
            v=F[k]["pat"].get(ck)
            if v is not None: cell.value=v; cell.number_format=USD; cell.font=ARIAL(color=BLUE,size=9)
        elif kind=="ret":
            if ck in F[k]["vc"]:
                vccell=f"'{SH_VC}'!{col}{r}"
                if ck==F[k]["inc"]:
                    cell.value=f'=IF({vccell}="","",{vccell}/{base_ref}-1)'
                else:
                    prev=get_column_letter(c-1); prevcell=f"'{SH_VC}'!{prev}{r}"
                    cell.value=f'=IFERROR(IF(OR({vccell}="",{prevcell}=""),"",{vccell}/{prevcell}-1),"")'
                cell.number_format=PCT; cell.font=ARIAL(color=BLACK,size=9)

    def matrix_sheet(name,col_labels,col_keys,kind):
        ws=wb.create_sheet(name); ncols=1+len(col_labels); title_block(ws,name,SUB,ncols)
        h=ws.cell(3,1,"Fondo"); h.font=white; h.fill=hdr; h.alignment=left; h.border=border
        for j,lab in enumerate(col_labels):
            c=ws.cell(3,2+j,lab); c.font=white; c.fill=hdr; c.alignment=center; c.border=border
        for f,k in enumerate(ORDER):
            r=4+f; cc=ws.cell(r,1,DISP[k]); cc.font=ARIAL(bold=True,size=9); cc.fill=fund_fill; cc.alignment=left; cc.border=border
            for j,ck in enumerate(col_keys):
                ws.cell(r,2+j).border=border; ws.cell(r,2+j).alignment=center; set_val(ws,r,2+j,f,k,ck,kind)
        ws.freeze_panes="B4"; ws.column_dimensions["A"].width=42
        for j in range(len(col_labels)): ws.column_dimensions[get_column_letter(2+j)].width=12 if kind in("vc","ret") else 16

    mlabels=[m[0] for m in months]; mkeys=[m[1] for m in months]
    matrix_sheet(SH_VC,mlabels,mkeys,"vc"); matrix_sheet(SH_PAT,mlabels,mkeys,"pat"); matrix_sheet(SH_MEN,mlabels,mkeys,"ret")

    def simple_sheet(name,sub):
        ws=wb.create_sheet(name); ncols=1+len(years); title_block(ws,name,sub,ncols)
        h=ws.cell(3,1,"Fondo"); h.font=white; h.fill=hdr; h.border=border; h.alignment=left
        for j,y in enumerate(years):
            c=ws.cell(3,2+j,f"{end.year} (YTD)" if y==end.year else str(y)); c.font=white; c.fill=hdr; c.alignment=center; c.border=border
        for f,k in enumerate(ORDER):
            r=4+f; cc=ws.cell(r,1,DISP[k]); cc.font=ARIAL(bold=True,size=9); cc.fill=fund_fill; cc.border=border; cc.alignment=left
            for j in range(len(years)): ws.cell(r,2+j).border=border; ws.cell(r,2+j).alignment=center
        ws.freeze_panes="B4"; ws.column_dimensions["A"].width=42
        for j in range(len(years)): ws.column_dimensions[get_column_letter(2+j)].width=13
        return ws

    wsd=simple_sheet(SH_DIC,"Rendimiento del mes de diciembre de cada año (mismo indicador que la hoja mensual). El año en curso = ultimo mes disponible (parcial)")
    for f,k in enumerate(ORDER):
        r=4+f
        for j,y in enumerate(years):
            tgt=last_ym if y==end.year else f"{y}-12"
            if tgt in F[k]["vc"]:
                c=wsd.cell(r,2+j); c.value=f"='{SH_MEN}'!{ym2col[tgt]}{r}"; c.number_format=PCT; c.font=ARIAL(size=9,color=BLACK)

    wsa=simple_sheet(SH_ANU,"Rendimiento ANUAL (ene–dic) desde la cuota. Año de lanzamiento = desde la base; el año en curso = YTD (parcial)")
    for f,k in enumerate(ORDER):
        r=4+f; base_ref=f"'{SH_RES}'!$D${r}"
        for j,y in enumerate(years):
            end_ym=last_ym if y==end.year else f"{y}-12"
            if end_ym not in F[k]["vc"]: continue
            endcell=f"'{SH_VC}'!{ym2col[end_ym]}{r}"; prior=f"{y-1}-12"
            prevref=f"'{SH_VC}'!{ym2col[prior]}{r}" if prior in F[k]["vc"] else base_ref
            c=wsa.cell(r,2+j); c.value=f'=IFERROR({endcell}/{prevref}-1,"")'; c.number_format=PCT; c.font=ARIAL(size=9,color=BLACK)

    ws=wb.create_sheet(SH_RES)
    cols=["Fondo","Tipo","Inicio","VC inicial (base)","VC actual","Fecha dato","Patrimonio actual (US$)","Meses","Rend. acum. desde inicio","Rend. anualizado (CAGR)"]
    title_block(ws,"Resumen — 6 Fondos de Inversión Hencorp",SUB,len(cols))
    for j,h in enumerate(cols):
        c=ws.cell(3,1+j,h); c.font=white; c.fill=hdr; c.border=border
        c.alignment=Alignment(horizontal=("left" if j==0 else "center"),vertical="center",wrap_text=True)
    ws.row_dimensions[3].height=30
    for f,k in enumerate(ORDER):
        r=4+f
        ws.cell(r,1,DISP[k]).font=ARIAL(bold=True,size=9); ws.cell(r,2,TIPO[k]).font=ARIAL(size=9)
        ws.cell(r,3,F[k]["first_date"]).number_format=DT; ws.cell(r,3).font=ARIAL(size=9)
        b=ws.cell(r,4,F[k]["base"]); b.number_format=("0.0000" if F[k]["base"]<10 else "#,##0.0000"); b.font=ARIAL(color=BLUE,size=9)
        e=ws.cell(r,5,f"='{SH_VC}'!{ym2col[last_ym]}{r}"); e.number_format=VCFMT; e.font=ARIAL(color=GREEN,size=9)
        ws.cell(r,6,F[k]["last_date"]).number_format=DT; ws.cell(r,6).font=ARIAL(size=9)
        g=ws.cell(r,7,f"='{SH_PAT}'!{ym2col[last_ym]}{r}"); g.number_format=USD; g.font=ARIAL(color=GREEN,size=9)
        ws.cell(r,8,F[k]["nmonths"]).font=ARIAL(size=9)
        i9=ws.cell(r,9,f"=E{r}/D{r}-1"); i9.number_format=PCT; i9.font=ARIAL(size=9)
        j10=ws.cell(r,10,f'=IFERROR((E{r}/D{r})^(365/(F{r}-C{r}))-1,"")'); j10.number_format=PCT; j10.font=ARIAL(size=9)
        for cc in range(1,len(cols)+1):
            ws.cell(r,cc).border=border; ws.cell(r,cc).alignment=(left if cc==1 else center)
    ws.column_dimensions["A"].width=42
    for j,w in zip(range(2,11),[10,12,15,14,13,20,8,18,18]): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="B4"

    wsn=wb.create_sheet(SH_NOT)
    notas=[("Notas y metodología",16,"1F3864",True),("",10,"000000",False),
     ("Fuente: Hencorp Gestora de Fondos de Inversión, S.A. — sección 'Informe Valor Cuota' (PDF 'Comportamiento Histórico del Fondo') en hencorpgestora.com. Archivo regenerado automáticamente.",10,"000000",False),
     (f"{corte_label}.",10,"000000",False),("",10,"000000",False),
     ("DEFINICIÓN DE 'RENDIMIENTO MENSUAL':",11,"1F3864",True),
     ("Rendimiento del mes = (Valor cuota al cierre del mes ÷ Valor cuota al cierre del mes anterior) − 1. Es el rendimiento real de la cuota en el mes.",10,"000000",False),
     ("El primer mes de cada fondo se mide desde el valor cuota inicial (base) hasta el cierre de ese mes (puede ser parcial). El año en curso es PARCIAL.",10,"C00000",True),("",10,"000000",False),
     ("ADVERTENCIA (fondos cerrados):",11,"C00000",True),
     ("• Renta Fija I DISTRIBUYE en mayo y noviembre: la cuota baja esos meses, por lo que su rendimiento medido SOLO desde la cuota SUBESTIMA el total del partícipe (lo distribuido se paga en efectivo).",10,"000000",False),
     ("• Growth y Vivienda 01 registran REVALUACIONES inmobiliarias (saltos grandes) y bajas por distribución/markdown. Commercial Properties tiene historia mínima (inició may-2026).",10,"000000",False),("",10,"000000",False),
     ("Códigos de color: azul = dato de origen extraído; verde = enlace a otra hoja; negro = fórmula.",9,"595959",False)]
    for i,(txt,sz,color,bold) in enumerate(notas,start=1):
        c=wsn.cell(i,1,txt); c.font=ARIAL(size=sz,color=color,bold=bold); c.alignment=Alignment(wrap_text=True,vertical="top")
    wsn.column_dimensions["A"].width=140

    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    desired=[SH_RES,SH_MEN,SH_DIC,SH_ANU,SH_VC,SH_PAT,SH_NOT]
    wb._sheets.sort(key=lambda s: desired.index(s.title))
    wb.save(out_path)
