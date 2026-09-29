#!/usr/bin/env python3
"""Faithful PNG previews of the generated report (new black+red design, Indian commas)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from openpyxl import load_workbook

WB = "SEPTEMBER 2026 - ZONE SALES REPORT.xlsx"
INK="#0A0A0A"; INK2="#1A1A1A"; RED="#D02128"; REDT="#EF4444"
GREY="#8A8A8A"; GREYD="#5A5A5A"; INKTX="#141414"; LINE="#E6E6E6"
BAND="#FAFAFA"; FILL2="#F4F4F4"; WHITE="#FFFFFF"


def inr(n):
    n=int(round(n)); s=str(abs(n)); sign="-" if n<0 else ""
    if len(s)>3:
        last3=s[-3:]; rest=s[:-3]; parts=[]
        while len(rest)>2: parts.insert(0,rest[-2:]); rest=rest[:-2]
        if rest: parts.insert(0,rest)
        g=",".join(parts)+","+last3
    else: g=s
    return f"{sign}\u20b9{g}"


def render_zone(sheet, path):
    wb=load_workbook(WB); ws=wb[sheet]
    ncol=ws.max_column
    header=[ws.cell(5,c).value for c in range(1,ncol+1)]
    # gather rows 6.. until row whose col1 == 'SHARE %'
    data=[]; kinds=[]; days=[]
    r=6
    while r<=ws.max_row:
        v1=ws.cell(r,1).value
        if v1 in (None,"") and all(ws.cell(r,c).value in (None,"") for c in range(1,ncol+1)):
            break
        rowvals=[ws.cell(r,c).value for c in range(1,ncol+1)]
        day=ws.cell(r,2).value
        if v1=="TOTAL": kind="total"
        elif v1=="SHARE %": kind="share"
        else: kind="data"
        disp=[]
        for c,val in enumerate(rowvals):
            if kind=="share" and c>=2 and isinstance(val,(int,float)):
                disp.append(f"{val*100:.1f}%")
            elif isinstance(val,(int,float)):
                disp.append(inr(val))
            else:
                disp.append("" if val is None else str(val))
        data.append(disp); kinds.append(kind); days.append(day)
        if kind=="share": break
        r+=1

    nrows=len(data)+1
    colw=[]
    for c in range(ncol):
        mx=max([len(str(header[c]))]+[len(row[c]) for row in data])
        colw.append(mx)
    total_units=sum(colw)
    fig_w=min(0.14*total_units+2, 20)
    fig_h=0.36*nrows+1.6
    fig,ax=plt.subplots(figsize=(fig_w,fig_h)); ax.axis("off")

    # title bar
    ax.add_patch(Rectangle((0,1.0),1,0.10,transform=ax.transAxes,color=INK,clip_on=False,zorder=3))
    ax.add_patch(Rectangle((0,0.985),1,0.015,transform=ax.transAxes,color=RED,clip_on=False,zorder=3))
    zt=sheet.strip()
    ax.text(0.008,1.05,zt[0],transform=ax.transAxes,color=REDT,fontsize=17,fontweight="bold",va="center",zorder=4)
    ax.text(0.008+0.014*fig_w/ (fig_w),1.05,"",transform=ax.transAxes)  # noop
    ax.text(0.028,1.05,f"{zt[1:]}      SEPTEMBER 2026",transform=ax.transAxes,color="white",fontsize=17,fontweight="bold",va="center",zorder=4)

    tbl=ax.table(cellText=data, colLabels=[str(h) for h in header],
                 colWidths=[w/total_units for w in colw],
                 cellLoc="center", loc="center", bbox=[0,0,1,0.92])
    tbl.auto_set_font_size(False); tbl.set_fontsize(8.5)

    for (rr,cc),cell in tbl.get_celld().items():
        cell.set_edgecolor(LINE); cell.set_linewidth(0.6)
        txt=cell.get_text()
        if rr==0:  # header
            cell.set_facecolor(RED if cc==ncol-1 else INK2)
            txt.set_color("white"); txt.set_fontweight("bold"); txt.set_fontsize(8)
            cell.set_height(cell.get_height()*1.5)
            continue
        kind=kinds[rr-1]; day=days[rr-1]
        if kind=="total":
            cell.set_facecolor(RED if cc==ncol-1 else INK)
            txt.set_color("white"); txt.set_fontweight("bold")
        elif kind=="share":
            cell.set_facecolor(FILL2)
            txt.set_color(GREYD); txt.set_fontsize(8)
        else:
            cell.set_facecolor(BAND if (rr-1)%2 else WHITE)
            if cc==ncol-1:  # daily total col
                cell.set_facecolor(FILL2); txt.set_fontweight("bold"); txt.set_color(INKTX)
            elif cc==1:  # day
                if day in ("Saturday","Sunday"): txt.set_color(RED); txt.set_fontweight("bold")
                else: txt.set_color(GREYD)
            elif cc>=2 and data[rr-1][cc]=="\u2013":
                txt.set_color("#C4C4C4")
            else:
                txt.set_color(INKTX)
        # alignment: numbers right, labels center
        if cc>=2 and kind!="share":
            txt.set_ha("right"); cell.PAD=0.04
        if cc>=2 and kind=="share":
            txt.set_ha("right")
    plt.savefig(path,dpi=150,bbox_inches="tight",facecolor="white"); print("wrote",path)


def render_summary(path):
    wb=load_workbook(WB); ws=wb["SUMMARY"]
    ncol=6
    header=[ws.cell(5,c).value for c in range(1,ncol+1)]
    data=[]; kinds=[]
    r=6
    while r<=ws.max_row:
        v1=ws.cell(r,1).value
        if v1 in (None,""): break
        rowvals=[ws.cell(r,c).value for c in range(1,ncol+1)]
        kind="total" if v1=="ALL ZONES" else "data"
        disp=[]
        for c,v in enumerate(rowvals):
            if c in (3,5) and isinstance(v,(int,float)):   # avg + grand total -> currency
                disp.append(inr(v))
            elif isinstance(v,(int,float)):
                disp.append(str(int(v)))                    # active brands count -> plain
            else:
                disp.append("" if v is None else str(v))
        data.append(disp); kinds.append(kind)
        if kind=="total": break
        r+=1
    nrows=len(data)+1
    fig,ax=plt.subplots(figsize=(13,0.5*nrows+2)); ax.axis("off")
    ax.add_patch(Rectangle((0,1.0),1,0.13,transform=ax.transAxes,color=INK,clip_on=False,zorder=3))
    ax.add_patch(Rectangle((0,0.982),1,0.018,transform=ax.transAxes,color=RED,clip_on=False,zorder=3))
    ax.text(0.008,1.06,"S",transform=ax.transAxes,color=REDT,fontsize=18,fontweight="bold",va="center",zorder=4)
    ax.text(0.022,1.06,"EPTEMBER 2026      ALL-ZONE SALES SUMMARY",transform=ax.transAxes,color="white",fontsize=18,fontweight="bold",va="center",zorder=4)
    tbl=ax.table(cellText=data, colLabels=[str(h) for h in header],
                 cellLoc="center", loc="center", bbox=[0,0,1,0.9])
    tbl.auto_set_font_size(False); tbl.set_fontsize(10); tbl.scale(1,1.6)
    for (rr,cc),cell in tbl.get_celld().items():
        cell.set_edgecolor(LINE); cell.set_linewidth(0.6); txt=cell.get_text()
        if rr==0:
            cell.set_facecolor(RED if cc==ncol-1 else INK2); txt.set_color("white"); txt.set_fontweight("bold"); txt.set_fontsize(9)
        elif kinds[rr-1]=="total":
            cell.set_facecolor(RED if cc==ncol-1 else INK); txt.set_color("white"); txt.set_fontweight("bold")
        else:
            cell.set_facecolor(BAND if (rr-1)%2 else WHITE)
            if cc==ncol-1: txt.set_color(RED); txt.set_fontweight("bold")
            elif cc==0: txt.set_color(INKTX); txt.set_fontweight("bold")
            else: txt.set_color(INKTX)
        if cc in (3,5): txt.set_ha("right")
    plt.savefig(path,dpi=150,bbox_inches="tight",facecolor="white"); print("wrote",path)


render_summary("preview_summary.png")
render_zone("ALPHA 1","preview_alpha1.png")
render_zone("BENNETT","preview_bennett.png")
