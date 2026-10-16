import datetime as dt, json
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as md
from matplotlib import font_manager as fm
import os, logging
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
HERE=os.path.dirname(os.path.abspath(__file__))
# 日本語フォント：環境にあるものを順に試す（Linux: IPA / Windows: Yu Gothic, Meiryo / Mac: Hiragino）
_IPA="/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"
if os.path.exists(_IPA): fm.fontManager.addfont(_IPA)
plt.rcParams["font.family"]=["IPAPGothic","Yu Gothic","Meiryo","Hiragino Sans","Noto Sans CJK JP","sans-serif"]
D=dt.date
CAP=560
START=D(2026,10,15)
# name, kg, arrival(None=in stock), color
lots=[("Ribáuè 2026",1326,None,"#2a78d6"),("Ribáuè 2025",749,None,"#eb6834"),("Ghana 0",745,None,"#6250d6"),
("India",8400,D(2026,10,16),"#1baf7a"),("Ghana 1",3000,D(2026,10,21),"#eda100"),
("Ghana 2",12000,D(2026,11,14),"#e87ba4"),("Ghana 3",3700,D(2026,11,13),"#008300")]
L=[]
for n,kg,a,c in lots:
    if a is None: dry_end=None; prep=D(2026,10,20); ready=D(2026,10,21)
    else: dry_end=a+dt.timedelta(6); prep=a+dt.timedelta(7); ready=a+dt.timedelta(8)
    L.append(dict(name=n,kg=kg,arr=a,dry_end=dry_end,prep=prep,ready=ready,color=c,rem=kg,start=None,end=None))
days=[]; d=START; idx=0
while idx<len(L):
    stock={}; sown={}
    cap={6:0,5:CAP//2}.get(d.weekday(),CAP)
    daycap=cap
    while cap>0 and idx<len(L) and L[idx]["ready"]<=d:
        lot=L[idx]; s=min(cap,lot["rem"]); lot["rem"]-=s; cap-=s
        sown[lot["name"]]=s; lot["start"]=lot["start"] or d
        if lot["rem"]==0: lot["end"]=d; idx+=1
    for lot in L:
        if lot["arr"] is None or lot["arr"]<=d: stock[lot["name"]]=lot["rem"]
        else: stock[lot["name"]]=0
    days.append(dict(date=d,stock=stock,sown=sown,used=daycap-cap,cap=daycap))
    d+=dt.timedelta(1)
json.dump(dict(lots=[{k:(str(v) if isinstance(v,D) else v) for k,v in l.items()} for l in L],
 days=[dict(date=str(x["date"]),stock=x["stock"],sown=x["sown"],used=x["used"],cap=x["cap"],wd=x["date"].weekday()) for x in days]),open(os.path.join(HERE,"plan.json"),"w"),indent=1,ensure_ascii=False)
for x in days: print(x["date"],sum(x["stock"].values()),x["used"],x["sown"])
for l in L: print(l["name"],l["arr"],l["dry_end"],l["prep"],l["ready"],l["start"],l["end"])

# chart: stacked bars of end-of-day stock
fig,ax=plt.subplots(figsize=(11,6),dpi=160)
x=[t["date"] for t in days]; base=[0]*len(x)
for l in L:
    v=[t["stock"][l["name"]] for t in days]
    ax.bar(x,v,bottom=base,width=0.8,color=l["color"],label=f'{l["name"]} ({l["kg"]:,} kg)',edgecolor="white",linewidth=0.6)
    base=[b+vv for b,vv in zip(base,v)]
# idle days shading
for t in days:
    if t["used"]<t["cap"] and t["date"]>=D(2026,10,21) and not (t["date"]==days[-1]["date"]):
        frac=1-t["used"]/t["cap"]
        c=md.date2num(t["date"]); ax.axvspan(c-0.5,c+0.5,color="#d03b3b",alpha=0.07+0.13*frac,lw=0,zorder=0)
ax.axvline(md.date2num(D(2026,10,21))-0.5,color="#555",ls=":",lw=1.2); ax.text(md.date2num(D(2026,10,21))-0.3,max(base)*1.04,"播種開始 10/21",fontsize=8,color="#333")
mx=max(base); mi=base.index(mx)
ax.annotate(f'最大 {mx:,} kg（11/14〜11/15）',xy=(x[mi],mx),xytext=(x[mi]+dt.timedelta(9),mx*0.99),fontsize=9,arrowprops=dict(arrowstyle="-",color="#555"),color="#222")
ax.set_ylabel("倉庫内の種子量（kg, 各日の播種後）",color="#333")
ax.xaxis.set_major_locator(md.DayLocator(interval=3)); ax.xaxis.set_major_formatter(md.DateFormatter("%m/%d"))
plt.setp(ax.get_xticklabels(),rotation=60,fontsize=8)
ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v,p:f"{int(v):,}"))
ax.grid(axis="y",color="#e5e5e5"); ax.set_axisbelow(True)
for s in ["top","right"]: ax.spines[s].set_visible(False)
ax.spines["left"].set_color("#999"); ax.spines["bottom"].set_color("#999")
from matplotlib.patches import Patch
h,lab=ax.get_legend_handles_labels()
h.append(Patch(color="#d03b3b",alpha=0.2)); lab.append("播種できる種子が不足（作業日）")
ax.legend(h,lab,fontsize=8,frameon=False,loc="upper center",bbox_to_anchor=(0.5,-0.13),ncol=4)
ax.set_title("種子の倉庫保管量の推移（種類別）",fontsize=12,loc="left")
ax.set_xlim(x[0]-dt.timedelta(1),x[-1]+dt.timedelta(1))
fig.tight_layout(); fig.savefig(os.path.join(HERE,"stock.png"))
