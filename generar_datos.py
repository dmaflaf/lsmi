#!/usr/bin/env python3
"""LSMI - Generador completo v6 FUSION · Temporada 2026"""
import openpyxl, json, sys, os, unicodedata, base64
from datetime import datetime

def norm(s):
    return unicodedata.normalize('NFD', str(s)).encode('ascii','ignore').decode('ascii').upper()

HOJAS_POS = {
    '1era División':'1era División','2da División':'2da División',
    '3era División C1':'3era División 1','3era División C2':'3era División 2',
}
DIV_RES = {v:k for k,v in HOJAS_POS.items()}
DIV_RIESGO   = {'1era División':4,'2da División':2,'3era División C1':4,'3era División C2':4}
DIV_CLASSIFY = {'1era División':8,'2da División':8,'3era División C1':8,'3era División C2':8}

def procesar_stats(ruta):
    wb=openpyxl.load_workbook(ruta,read_only=True,data_only=True)
    ws_g=wb['GOLEADORES']; goles_pj={}; uf=0
    for row in ws_g.iter_rows(min_row=3,values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        e=str(row[1]).strip() if row[1] else ''; j=str(row[2]).strip() if row[2] else ''; g=row[3]
        if e and j and g and isinstance(g,(int,float)):
            key=(e,j); goles_pj[key]=goles_pj.get(key,0)+int(g); uf=max(uf,int(row[0]))
    ws_a=wb['AMONESTACIONES']; tarj_pj={}
    for row in ws_a.iter_rows(min_row=3,values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        e=str(row[1]).strip() if row[1] else ''; j=str(row[2]).strip() if row[2] else ''; t=str(row[3]).strip() if row[3] else ''
        if e and j and t:
            key=(e,j)
            if key not in tarj_pj: tarj_pj[key]={'amarillas':0,'dobles':0,'rojas':0}
            if 'Doble' in t or 'doble' in t: tarj_pj[key]['dobles']+=1
            elif 'Roja' in t or 'roja' in t or '🔴' in t: tarj_pj[key]['rojas']+=1
            else: tarj_pj[key]['amarillas']+=1
    equipos=sorted(set(k[0] for k in goles_pj)|set(k[0] for k in tarj_pj))
    teams_data={}
    for eq in equipos:
        ge={j:g for (e,j),g in goles_pj.items() if e==eq}
        te={j:t for (e,j),t in tarj_pj.items() if e==eq}
        tg=sum(ge.values()); tam=sum(t['amarillas'] for t in te.values())
        tdo=sum(t['dobles'] for t in te.values()); tro=sum(t['rojas'] for t in te.values())
        pg=round(tg/uf,2) if uf else 0; pt=round((tam+tdo+tro)/uf,2) if uf else 0
        ts=sorted(ge.items(),key=lambda x:-x[1])[:8]
        tcr=[{'nombre':j,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':t['amarillas']+t['dobles']*2+t['rojas']*3} for j,t in te.items() if t['amarillas']+t['dobles']*2+t['rojas']*3>0]
        tcr.sort(key=lambda x:-x['_p'])
        teams_data[eq]={'total_goles':tg,'promedio_goles':pg,'total_amarillas':tam,'total_doble':tdo,'total_rojas':tro,'promedio_tarjetas':pt,'top_scorers':[{'nombre':j,'goles':g} for j,g in ts],'top_cards':[{k:v for k,v in c.items() if k!='_p'} for c in tcr[:8]]}
    all_s=[{'nombre':j,'equipo':e,'goles':g} for (e,j),g in goles_pj.items()]; all_s.sort(key=lambda x:-x['goles'])
    all_c=[{'nombre':j,'equipo':e,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':t['amarillas']+t['dobles']*2+t['rojas']*3} for (e,j),t in tarj_pj.items() if t['amarillas']+t['dobles']*2+t['rojas']*3>0]
    all_c.sort(key=lambda x:-x['_p'])
    fp=[{'equipo':eq,'total_tarjetas':sum(t['amarillas']+t['dobles']+t['rojas'] for t in {j:t for (e,j),t in tarj_pj.items() if e==eq}.values())} for eq in equipos]
    fp.sort(key=lambda x:x['total_tarjetas'])
    return {'meta':{'ultima_fecha':uf,'generado':datetime.now().strftime('%d/%m/%Y %H:%M')},'globalData':{'top_10_goleadores':all_s[:10],'top_10_amonestados':[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:10]],'fair_play':fp},'teamsData':teams_data,'_goles_raw':{f"{e}||{j}":g for (e,j),g in goles_pj.items()},'_tarjetas_raw':{f"{e}||{j}":t for (e,j),t in tarj_pj.items()}}

def leer_posiciones(wb_pos):
    pos={}
    for div,hoja in HOJAS_POS.items():
        if hoja not in wb_pos.sheetnames: continue
        ws=wb_pos[hoja]; tabla=[]; ok=False
        for row in ws.iter_rows(values_only=True):
            if row[0]=='POS' and row[1]=='EQUIPO': ok=True; continue
            if not ok: continue
            if row[0] and isinstance(row[0],(int,float)) and row[1] and isinstance(row[1],str) and len(str(row[1]).strip())>1:
                if 'PJ=' in str(row[1]) or 'Puntuación' in str(row[1]): break
                tabla.append({'pos':int(row[0]),'equipo':str(row[1]).strip(),'pj':int(row[2] or 0),'pg':int(row[3] or 0),'pe':int(row[4] or 0),'pp':int(row[5] or 0),'gf':int(row[6] or 0),'gc':int(row[7] or 0),'dg':int(row[8] or 0),'pts':int(row[9] or 0)})
        pos[div]=tabla
    return pos

def leer_resultados(wb_pos):
    ws_r=wb_pos['RESULTADOS']; res={}
    for row in ws_r.iter_rows(min_row=2,values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        if not row[2] or not row[5]: continue
        fecha=int(row[0]); div_raw=str(row[1]).strip() if row[1] else ''
        local=str(row[2]).strip(); gf=int(row[3] or 0); gv=int(row[4] or 0); vis=str(row[5]).strip()
        if not div_raw or not local or not vis: continue
        div_nombre=DIV_RES.get(div_raw,div_raw)
        if div_nombre not in res: res[div_nombre]={}
        if fecha not in res[div_nombre]: res[div_nombre][fecha]=[]
        res[div_nombre][fecha].append({'local':local,'gl':gf,'gv':gv,'visitante':vis})
    return {div:[{'fecha':f,'partidos':ps} for f,ps in sorted(fechas.items(),reverse=True)] for div,fechas in res.items()}

def leer_sancionados(ruta):
    DIV_MAP={norm('1era DIVISIÓN'):'1era División',norm('2da DIVISIÓN'):'2da División',norm('3era DIVISIÓN 1'):'3era División C1',norm('3era DIVISIÓN 2'):'3era División C2'}
    wb=openpyxl.load_workbook(ruta,read_only=True,data_only=True); ws=wb[wb.sheetnames[0]]; sanc={}; div_actual=None
    for row in ws.iter_rows(values_only=True):
        c1,c2,c3,c4,c5=row[1],row[2],row[3],row[4],row[5]
        if c1 is None and c2 and isinstance(c2,str):
            c2n=norm(c2)
            for k,v in DIV_MAP.items():
                if k in c2n: div_actual=v; sanc.setdefault(div_actual,[]); break
            continue
        if div_actual and c2 and isinstance(c2,str) and c3 and isinstance(c3,str) and c4 and isinstance(c4,str):
            if norm(c2) in ('NOMBRES Y APELLIDOS','SANCION'): continue
            numero=str(c1).strip() if c1 is not None else '—'
            s_upper=str(c4).upper()
            if any(x in s_upper for x in ['INDEFINIDO','AÑO','AÑOS','MESES','PROFESIONAL']): tipo='grave'
            elif any(x in s_upper for x in ['2 FECHA','3 FECHA']): tipo='media'
            else: tipo='leve'
            sanc[div_actual].append({'numero':numero,'nombre':str(c2).strip(),'club':str(c3).strip(),'sancion':str(c4).strip(),'fase2':str(c5).strip() if c5 else '','tipo':tipo})
    return sanc

def buscar_excels(carpeta):
    archivos=[f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    stats={}
    claves={'1era División':['1era','primera'],'2da División':['2da','segunda'],'3era División C1':['division_1','3era_division_1','_1_v','g1','c1'],'3era División C2':['division_2','3era_division_2','_2_v','g2','c2']}
    for arch in archivos:
        low=arch.lower()
        for div,pals in claves.items():
            if any(p in low for p in pals) and div not in stats: stats[div]=os.path.join(carpeta,arch); break
    pos_ruta=next((os.path.join(carpeta,f) for f in archivos if any(k in f.lower() for k in ['torneo','posicion','fase'])),None)
    sanc_ruta=next((os.path.join(carpeta,f) for f in archivos if 'sancionado' in f.lower()),None)
    logo_ruta=next((os.path.join(carpeta,f) for f in os.listdir(carpeta) if f.lower().endswith('.png')),None)
    return stats,pos_ruta,sanc_ruta,logo_ruta

def procesar_todo(carpeta):
    stats_rutas,pos_ruta,sanc_ruta,logo_ruta=buscar_excels(carpeta)
    print("📁 Archivos:"); [print(f"   {d}: {os.path.basename(r)}") for d,r in stats_rutas.items()]
    if pos_ruta: print(f"   Posiciones: {os.path.basename(pos_ruta)}")
    if sanc_ruta: print(f"   Sancionados: {os.path.basename(sanc_ruta)}")
    if logo_ruta: print(f"   Logo: {os.path.basename(logo_ruta)}")

    res_stats={}
    for nombre,ruta in stats_rutas.items():
        print(f"\n⚽ {nombre}..."); res_stats[nombre]=procesar_stats(ruta)
        print(f"   {len(res_stats[nombre]['teamsData'])} equipos | F{res_stats[nombre]['meta']['ultima_fecha']}")

    posiciones={}; res_partidos={}
    if pos_ruta:
        wb_pos=openpyxl.load_workbook(pos_ruta,read_only=True,data_only=True)
        posiciones=leer_posiciones(wb_pos); res_partidos=leer_resultados(wb_pos)

    sancionados={}
    if sanc_ruta: sancionados=leer_sancionados(sanc_ruta)

    logo_b64=''
    if logo_ruta:
        with open(logo_ruta,'rb') as f: logo_b64='data:image/png;base64,'+base64.b64encode(f.read()).decode()

    # Ranking equipos más amonestados
    ranking_eq_amon=[]
    for dn,data in res_stats.items():
        for eq,stats in data['teamsData'].items():
            tam=stats['total_amarillas']; tdo=stats['total_doble']; tro=stats['total_rojas']
            pts=tam+tdo*2+tro*3
            ranking_eq_amon.append({'equipo':eq,'division':dn,'amarillas':tam,'dobles':tdo,'rojas':tro,'puntos':pts,'total':tam+tdo+tro})
    ranking_eq_amon.sort(key=lambda x:-x['puntos'])

    total_goles_t=0; total_part_t=0
    for div,fechas in res_partidos.items():
        for fd in fechas:
            for p in fd['partidos']: total_goles_t+=p['gl']+p['gv']; total_part_t+=1
    prom=round(total_goles_t/total_part_t,2) if total_part_t else 0

    all_g,all_c,fp_t,stats_div=[],[],[],[]
    for dn,data in res_stats.items():
        for key,goles in data['_goles_raw'].items():
            e,j=key.split('||'); all_g.append({'nombre':j,'equipo':e,'division':dn,'goles':goles})
        for key,t in data['_tarjetas_raw'].items():
            e,j=key.split('||'); p=t['amarillas']+t['dobles']*2+t['rojas']*3
            if p>0: all_c.append({'nombre':j,'equipo':e,'division':dn,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
        for eq_fp in data['globalData']['fair_play']:
            fp_t.append({'equipo':eq_fp['equipo'],'division':dn,'total_tarjetas':eq_fp['total_tarjetas']})
        tg=sum(d['total_goles'] for d in data['teamsData'].values())
        tt=sum(d['total_amarillas']+d['total_doble']+d['total_rojas'] for d in data['teamsData'].values())
        ne=len(data['teamsData'])
        stats_div.append({'division':dn,'total_goles':tg,'total_tarjetas':tt,'num_equipos':ne,'promedio_goles_equipo':round(tg/ne,1) if ne else 0})
    all_g.sort(key=lambda x:-x['goles']); all_c.sort(key=lambda x:-x['_p']); fp_t.sort(key=lambda x:x['total_tarjetas']); stats_div.sort(key=lambda x:-x['total_goles'])
    top_c=[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:15]]
    total_eq=sum(d['num_equipos'] for d in stats_div)
    global_torneo={'top_15_goleadores':all_g[:15],'top_15_amonestados':top_c,'fair_play':fp_t[:20],'ranking_divisiones':stats_div,'prom_goles_partido':prom,'total_partidos':total_part_t,'ranking_equipos_amonestados':ranking_eq_amon[:15],'total_equipos':total_eq,'total_tarjetas':sum(d['total_tarjetas'] for d in stats_div),'total_goles':sum(d['total_goles'] for d in stats_div)}
    for d in res_stats.values(): d.pop('_goles_raw',None); d.pop('_tarjetas_raw',None)
    return {'generado':datetime.now().strftime('%d/%m/%Y %H:%M'),'divisiones':res_stats,'posiciones':posiciones,'resultados':res_partidos,'sancionados':sancionados,'globalTorneo':global_torneo,'logo':logo_b64}

def generar_html(datos):
    return HTML.replace('__JSON__', json.dumps(datos, ensure_ascii=False))

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>LSMI · Estadísticas Oficiales 2026</title>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#121212;--surf:#1e1e1e;--surf2:#2a2a2a;--surf3:#333;
  --border:rgba(255,255,255,.08);--border2:rgba(255,255,255,.12);
  --red:#c8102e;--red2:#e8192c;--red-dim:rgba(200,16,46,.15);
  --gold:#f5a623;--gold-dim:rgba(245,166,35,.15);
  --green:#22c55e;--green-dim:rgba(34,197,94,.12);
  --orange:#f97316;--orange-dim:rgba(249,115,22,.12);
  --blue:#3b82f6;
  --text:#f0f2f8;--text2:rgba(240,242,248,.7);--text3:rgba(240,242,248,.4);
  --1era:#c8102e;--2da:#3b82f6;--c1:#22c55e;--c2:#f5a623;
  --shadow:0 4px 24px rgba(0,0,0,.4);
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden;padding-bottom:80px;-webkit-font-smoothing:antialiased}

/* ── HEADER ── */
.site-header{position:fixed;top:0;width:100%;z-index:100;height:64px;background:rgba(18,18,18,.9);backdrop-filter:blur(20px);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 20px;gap:14px}
.logo-img{width:40px;height:40px;object-fit:contain;border-radius:8px;flex-shrink:0}
.brand-title{font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:900;letter-spacing:1px;text-transform:uppercase;color:var(--text);line-height:1}
.brand-sub{font-size:9px;font-weight:700;color:var(--text3);letter-spacing:3px;text-transform:uppercase;margin-top:2px}
.header-spacer{flex:1}

/* ── NAV ── */
.site-nav{position:fixed;top:64px;width:100%;z-index:90;height:52px;background:rgba(30,30,30,.95);backdrop-filter:blur(16px);border-bottom:1px solid var(--border);display:flex;align-items:center;overflow:hidden}
.nav-inner{display:flex;overflow-x:auto;white-space:nowrap;padding:0 16px;gap:4px;scrollbar-width:none;height:100%;align-items:center}
.nav-inner::-webkit-scrollbar{display:none}
.nav-btn{display:flex;align-items:center;gap:6px;padding:6px 14px;font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:var(--text3);border:none;background:transparent;cursor:pointer;border-bottom:3px solid transparent;height:52px;transition:all .2s;white-space:nowrap;border-radius:0}
.nav-btn .material-symbols-outlined{font-size:18px}
.nav-btn:hover{color:var(--text2)}
.nav-btn.active{color:var(--red);border-bottom-color:var(--red)}

/* ── BOTTOM NAV (móvil) ── */
.bottom-nav{position:fixed;bottom:0;width:100%;z-index:100;background:rgba(30,30,30,.97);backdrop-filter:blur(20px);border-top:1px solid var(--border);display:flex;justify-content:space-around;padding:8px 0 12px;box-shadow:0 -4px 20px rgba(0,0,0,.4)}
.bnav-item{display:flex;flex-direction:column;align-items:center;gap:3px;cursor:pointer;padding:4px 12px;border-radius:8px;transition:all .2s;color:var(--text3);border:none;background:transparent}
.bnav-item.active{color:var(--red)}
.bnav-item .material-symbols-outlined{font-size:22px}
.bnav-item span:last-child{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase}

/* ── MAIN ── */
.main{margin-top:116px;padding:24px 16px;max-width:1200px;margin-left:auto;margin-right:auto}
.main{margin-top:116px}

/* ── VIEWS ── */
.view{display:none}
.view.active{display:block}

/* ── SECTION ── */
.s-eyebrow{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:3px;text-transform:uppercase;color:var(--red);margin-bottom:6px}
.s-title{font-family:'Barlow Condensed',sans-serif;font-size:clamp(28px,5vw,42px);font-weight:900;text-transform:uppercase;color:var(--text);line-height:1;margin-bottom:24px}
.s-title span{color:var(--red)}

/* ── BENTO KPI ── */
.bento-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin-bottom:28px}
@media(min-width:640px){.bento-grid{grid-template-columns:repeat(3,1fr)}}
@media(min-width:900px){.bento-grid{grid-template-columns:repeat(6,1fr)}}
.kpi-card{background:var(--surf);border:1px solid var(--border);border-radius:14px;padding:20px 16px;position:relative;overflow:hidden;transition:background .2s}
.kpi-card:hover{background:var(--surf2)}
.kpi-card .live-dot{position:absolute;top:12px;right:12px;width:7px;height:7px;border-radius:50%;background:var(--red);box-shadow:0 0 8px var(--red);animation:pulse 2s infinite}
@keyframes pulse{0%{transform:scale(.95);opacity:.7}70%{transform:scale(1.1);opacity:1}100%{transform:scale(.95);opacity:.7}}
.kpi-label{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-bottom:10px;display:block}
.kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:38px;font-weight:900;line-height:1;color:var(--text)}
.kpi-val.red{color:var(--red)}
.kpi-icon{position:absolute;bottom:-8px;right:-4px;opacity:.04;font-size:80px}

/* ── DIV CARDS ── */
.div-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:32px}
@media(min-width:768px){.div-grid{grid-template-columns:repeat(4,1fr)}}
.div-card{background:var(--surf);border:1px solid var(--border);border-radius:14px;overflow:hidden;transition:transform .2s}
.div-card:hover{transform:translateY(-2px)}
.div-card-stripe{height:3px;width:100%}
.div-card-body{padding:16px}
.div-card-name{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;letter-spacing:.5px;text-transform:uppercase}
.div-card-eq{font-size:9px;font-weight:700;letter-spacing:1.5px;color:var(--text3);text-transform:uppercase;margin-top:2px}
.div-card-rank{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:900;color:rgba(255,255,255,.06);line-height:1}
.div-card-stats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-top:12px;padding-top:12px;border-top:1px solid var(--border)}
.div-stat{text-align:center}
.div-stat-val{font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:800;color:var(--text)}
.div-stat-lbl{font-size:9px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--text3);margin-top:2px}

/* ── GRID LAYOUT ── */
.grid-2{display:grid;grid-template-columns:1fr;gap:20px;margin-bottom:28px}
@media(min-width:900px){.grid-2{grid-template-columns:1fr 1fr}}
.grid-3{display:grid;grid-template-columns:1fr;gap:16px;margin-bottom:24px}
@media(min-width:900px){.grid-3{grid-template-columns:1fr 1fr 1fr}}

/* ── PANEL ── */
.panel{background:var(--surf);border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:20px}
.panel-head{display:flex;align-items:center;gap:10px;padding:16px 18px;border-bottom:1px solid var(--border)}
.panel-head .material-symbols-outlined{font-size:20px;color:var(--red)}
.panel-head-title{font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:var(--text)}
.panel-body{padding:8px}

/* ── SCORER ITEMS ── */
.scorer-row{display:flex;align-items:center;gap:14px;padding:12px 10px;border-radius:10px;transition:background .15s;margin-bottom:2px}
.scorer-row:hover{background:var(--surf2)}
.scorer-row.top{background:rgba(200,16,46,.07);border-left:3px solid var(--red)}
.rank-num{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900;color:var(--text3);min-width:24px;text-align:center}
.rank-num.gold{color:var(--gold)}
.rank-num.silver{color:#9ca3af}
.rank-num.bronze{color:#cd7f32}
.scorer-info{flex:1;min-width:0}
.scorer-name{font-size:13px;font-weight:700;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.scorer-meta{font-size:10px;font-weight:600;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.scorer-bar-wrap{height:2px;background:rgba(255,255,255,.06);border-radius:2px;overflow:hidden;margin-top:6px}
.scorer-bar{height:100%;border-radius:2px;transition:width .6s ease}
.scorer-val{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:900;color:var(--text);min-width:36px;text-align:right}

/* ── CARD TARJETAS ── */
.card-t{display:inline-flex;flex-direction:column;align-items:center;justify-content:center;width:34px;height:40px;border-radius:5px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:900;margin-left:4px;border:1px solid rgba(255,255,255,.08)}
.card-t.roja{background:rgba(200,16,46,.25);color:#fca5a5}
.card-t.doble{background:rgba(249,115,22,.25);color:#fdba74}
.card-t.amarilla{background:rgba(234,179,8,.2);color:#fde047}

/* ── FP BARS ── */
.fp-row{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;transition:background .15s;margin-bottom:2px}
.fp-row:hover{background:var(--surf2)}
.fp-bar-wrap{flex:1;height:4px;background:rgba(255,255,255,.06);border-radius:2px;overflow:hidden}
.fp-bar{height:100%;border-radius:2px}
.fp-val{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;color:var(--text);min-width:28px;text-align:right}

/* ── TABLA POSICIONES ── */
.pos-wrap{overflow-x:auto;border-radius:10px;border:1px solid var(--border)}
.pos-table{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}
.pos-table thead{background:rgba(255,255,255,.03)}
.pos-table th{font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:11px 6px;text-align:center;border-bottom:1px solid var(--border);white-space:nowrap}
.pos-table th.eq-h{text-align:left;width:130px;padding-left:10px}
.pos-table th.w36{width:32px}
.pos-table td{padding:10px 6px;text-align:center;border-bottom:1px solid rgba(255,255,255,.04);color:var(--text2);font-weight:500}
.pos-table td.eq-td{text-align:left;font-weight:700;color:var(--text);padding-left:10px;width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:130px}
.pos-table tbody tr:hover td{background:var(--surf2)}
.pos-table tbody tr:last-child td{border-bottom:none}
.pts-td{font-family:'Barlow Condensed',sans-serif;font-size:16px;font-weight:900}
.dg-pos{color:var(--green)}
.dg-neg{color:var(--red)}
.pos-badge{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:900}
.pb-gold{background:var(--gold-dim);color:var(--gold)}
.pb-clf{background:var(--green-dim);color:var(--green)}
.pb-bot{background:var(--red-dim);color:var(--red)}
.pb-norm{background:var(--surf2);color:var(--text3)}
.zone-clf td{background:rgba(34,197,94,.04)}
.zone-clf td.eq-td{border-left:3px solid var(--green)}
.zone-bot td{background:rgba(200,16,46,.04)}
.zone-bot td.eq-td{border-left:3px solid var(--red)}
.zone-top td.eq-td{border-left:3px solid var(--gold)}
.tbl-legend{display:flex;gap:14px;flex-wrap:wrap;padding:10px 14px;border-top:1px solid var(--border);font-size:10px;font-weight:600;color:var(--text3);background:rgba(255,255,255,.02)}
.ldot{width:8px;height:8px;border-radius:2px;display:inline-block;margin-right:5px;vertical-align:middle}

/* ── RESULTADOS ── */
.fecha-block{margin-bottom:20px}
.fecha-lbl{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:5px 12px;background:var(--surf2);border:1px solid var(--border);border-radius:5px;margin-bottom:10px;display:inline-block}
.match-card{display:flex;align-items:center;padding:11px 14px;background:var(--surf);border:1px solid var(--border);border-radius:10px;margin-bottom:6px;gap:10px;transition:border-color .15s}
.match-card:hover{border-color:var(--border2)}
.match-team{font-size:13px;font-weight:600;color:var(--text2);flex:1;line-height:1.3}
.match-team.right{text-align:right}
.match-team.winner{color:var(--text);font-weight:700}
.match-score{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900;color:var(--text);background:var(--surf2);border:1px solid var(--border);padding:5px 14px;border-radius:7px;min-width:62px;text-align:center;letter-spacing:2px}
.match-score.draw{color:var(--text3)}

/* ── SANCIONADOS ── */
.sanc-table{width:100%;border-collapse:collapse;font-size:11px}
.sanc-table th{font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:11px 12px;text-align:left;background:rgba(255,255,255,.03);border-bottom:1px solid var(--border);white-space:nowrap}
.sanc-table td{padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04);color:var(--text2);font-size:12px}
.sanc-table tbody tr:hover td{background:var(--surf2)}
.sanc-table tbody tr:last-child td{border-bottom:none}
.sanc-name{font-weight:700;color:var(--text)}
.s-tag{display:inline-flex;align-items:center;font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;padding:2px 7px;border-radius:3px;letter-spacing:.5px;white-space:nowrap}
.s-tag.grave{background:rgba(200,16,46,.2);color:#fca5a5;border:1px solid rgba(200,16,46,.3)}
.s-tag.media{background:rgba(249,115,22,.2);color:#fdba74;border:1px solid rgba(249,115,22,.3)}
.s-tag.leve{background:rgba(234,179,8,.15);color:#fde047;border:1px solid rgba(234,179,8,.25)}
.s-fase2{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;color:var(--red);background:var(--red-dim);border:1px solid rgba(200,16,46,.3);padding:2px 7px;border-radius:3px}
.s-num{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:900;color:var(--text3);background:var(--surf2);border:1px solid var(--border);width:28px;height:28px;border-radius:4px;display:inline-flex;align-items:center;justify-content:center}

/* ── RANKING EQUIPOS AMON ── */
.eq-amon-row{display:flex;align-items:center;gap:10px;padding:10px 10px;border-radius:10px;background:var(--surf);border:1px solid var(--border);margin-bottom:4px;transition:border-color .15s}
.eq-amon-row:hover{border-color:var(--border2)}

/* ── SUB-TABS ── */
.sub-tabs{display:flex;gap:3px;margin-bottom:20px;background:var(--surf);border:1px solid var(--border);border-radius:8px;padding:3px;width:fit-content}
.stab{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:7px 14px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-radius:6px;transition:all .18s;white-space:nowrap}
.stab:hover{color:var(--text2)}
.stab.active{background:var(--red);color:#fff}
.sub-panel{display:none}
.sub-panel.active{display:block}

/* ── DIV HERO ── */
.div-hero{background:var(--surf);border:1px solid var(--border);border-radius:14px;border-top:3px solid var(--accent,var(--red));padding:20px 22px;margin-bottom:22px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
.div-hero-title{font-family:'Barlow Condensed',sans-serif;font-size:clamp(24px,4vw,34px);font-weight:900;letter-spacing:1px;text-transform:uppercase}
.div-hero-meta{font-size:11px;color:var(--text3);font-weight:500;margin-top:3px}
.div-kpis{display:flex;gap:24px;flex-wrap:wrap}
.div-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:900;color:var(--text)}
.div-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3)}

/* ── TEAM SELECTOR ── */
.team-sel-wrap{background:var(--surf);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:18px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.team-sel-lbl{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);white-space:nowrap}
.team-sel{flex:1;min-width:160px;background:var(--surf2);color:var(--text);border:1.5px solid var(--border2);border-radius:6px;padding:8px 12px;font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700;cursor:pointer;outline:none}

/* ── TEAM KPIS ── */
.team-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:1px;background:var(--border);border-radius:10px;overflow:hidden;margin-bottom:18px}
.team-kpi{background:var(--surf2);padding:16px;text-align:center}
.team-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:34px;font-weight:900;color:var(--text);line-height:1}
.team-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-top:4px}

/* ── FOOTER ── */
.site-footer{background:var(--surf);border-top:1px solid var(--border);padding:40px 20px 100px;text-align:center}
.footer-brand{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:3px;text-transform:uppercase;color:var(--text3);margin-bottom:10px}
.footer-links{display:flex;gap:16px;flex-wrap:wrap;justify-content:center;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-bottom:16px}
.footer-links .nix{color:var(--red)}
.footer-quote{font-size:12px;color:var(--text3);font-style:italic;max-width:480px;margin:0 auto;line-height:1.6}

/* ── EMPTY / MISC ── */
.empty{color:var(--text3);font-size:12px;padding:24px;text-align:center;font-style:italic}
.divider{height:1px;background:var(--border);margin:24px 0}

@media(max-width:600px){
  .site-header{padding:0 12px}
  .brand-title{font-size:14px}
  .main{padding:16px 12px}
  .div-hero{padding:14px 16px}
  .bento-grid{gap:8px}
  .kpi-val{font-size:30px}
}
</style>
</head>
<body>

<!-- HEADER -->
<header class="site-header">
  <img id="logoImg" class="logo-img" alt="LSMI">
  <div>
    <div class="brand-title">Liga San Miguel de Ibarra</div>
    <div class="brand-sub">LSMI · Estadísticas Oficiales</div>
  </div>
  <div class="header-spacer"></div>
</header>

<!-- NAV SUPERIOR -->
<nav class="site-nav">
  <div class="nav-inner">
    <button class="nav-btn active" data-view="global" onclick="switchView(this,'global')">
      <span class="material-symbols-outlined">public</span>GLOBAL
    </button>
    <button class="nav-btn" data-view="1era División" onclick="switchView(this,'1era División')">
      <span class="material-symbols-outlined">military_tech</span>1ERA DIVISIÓN
    </button>
    <button class="nav-btn" data-view="2da División" onclick="switchView(this,'2da División')">
      <span class="material-symbols-outlined">workspace_premium</span>2DA DIVISIÓN
    </button>
    <button class="nav-btn" data-view="3era División C1" onclick="switchView(this,'3era División C1')">
      <span class="material-symbols-outlined">star</span>3ERA C1
    </button>
    <button class="nav-btn" data-view="3era División C2" onclick="switchView(this,'3era División C2')">
      <span class="material-symbols-outlined">star_half</span>3ERA C2
    </button>
  </div>
</nav>

<!-- MAIN -->
<main class="main" style="margin-top:116px;padding:24px 16px;max-width:1200px;margin-left:auto;margin-right:auto">

  <!-- ── VISTA GLOBAL ── -->
  <div id="view-global" class="view active">
    <div style="margin-bottom:20px">
      <div class="s-eyebrow">Temporada 2026</div>
      <div class="s-title">Estadística <span>Global</span> del Torneo</div>
    </div>
    <div class="bento-grid" id="bentoKpis"></div>
    <div class="div-grid" id="divCards"></div>
    <div class="grid-2">
      <div>
        <div class="panel">
          <div class="panel-head"><span class="material-symbols-outlined">leaderboard</span><span class="panel-head-title">Top 15 Goleadores</span></div>
          <div class="panel-body" id="gGol"></div>
        </div>
      </div>
      <div class="panel">
        <div class="panel-head"><span class="material-symbols-outlined">warning</span><span class="panel-head-title">Ranking Amonestaciones</span></div>
        <div class="panel-body" id="gTarj"></div>
        <div style="border-top:1px solid var(--border);margin-top:4px"></div>
        <div class="panel-head" style="border-top:none"><span class="material-symbols-outlined" style="color:var(--green)">volunteer_activism</span><span class="panel-head-title">Fair Play General</span></div>
        <div class="panel-body" id="gFP"></div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">gavel</span><span class="panel-head-title">Ranking Amonestación General — Equipos</span></div>
      <div class="panel-body" style="padding:0">
        <div style="overflow-x:auto">
          <table style="width:100%;border-collapse:collapse;font-size:12px" id="eqAmonTable">
            <thead style="background:rgba(255,255,255,.03)">
              <tr>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:center;border-bottom:1px solid var(--border);width:40px">Pos</th>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:left;border-bottom:1px solid var(--border)">Equipo / División</th>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:center;border-bottom:1px solid var(--border);width:48px">R</th>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:center;border-bottom:1px solid var(--border);width:48px">D</th>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:center;border-bottom:1px solid var(--border);width:48px">A</th>
                <th style="font-family:'Barlow Condensed',sans-serif;font-size:9px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:12px 10px;text-align:right;border-bottom:1px solid var(--border);width:60px">Puntos</th>
              </tr>
            </thead>
            <tbody id="eqAmonBody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- ── VISTAS DIVISIONES ── -->
  <div id="view-1era División" class="view"></div>
  <div id="view-2da División" class="view"></div>
  <div id="view-3era División C1" class="view"></div>
  <div id="view-3era División C2" class="view"></div>

</main>

<!-- FOOTER -->
<footer class="site-footer">
  <div class="footer-brand">Liga San Miguel de Ibarra · Estadísticas Oficiales</div>
  <div class="footer-links">
    <span class="nix">Powered by NIX 26</span>
    <span style="opacity:.3">|</span>
    <span id="footerDate"></span>
  </div>
  <p class="footer-quote">"Fomentando el deporte y la disciplina en la comunidad de Ibarra. Datos proporcionados por la Comisión Técnica LSMI."</p>
</footer>

<!-- BOTTOM NAV (móvil) -->
<nav class="bottom-nav" id="bottomNav">
  <button class="bnav-item active" data-sub="pos" onclick="switchBottomNav(this,'pos')">
    <span class="material-symbols-outlined" style="font-variation-settings:'FILL' 1">format_list_numbered</span>
    <span>Posiciones</span>
  </button>
  <button class="bnav-item" data-sub="res" onclick="switchBottomNav(this,'res')">
    <span class="material-symbols-outlined">calendar_today</span>
    <span>Resultados</span>
  </button>
  <button class="bnav-item" data-sub="stats" onclick="switchBottomNav(this,'stats')">
    <span class="material-symbols-outlined">analytics</span>
    <span>Estadísticas</span>
  </button>
  <button class="bnav-item" data-sub="sanc" onclick="switchBottomNav(this,'sanc')">
    <span class="material-symbols-outlined">gavel</span>
    <span>Sancionados</span>
  </button>
</nav>

<script>
const DATA = __JSON__;
const DIV_COLORS={'1era División':'#c8102e','2da División':'#3b82f6','3era División C1':'#22c55e','3era División C2':'#f5a623'};
const DIV_RIESGO={'1era División':4,'2da División':2,'3era División C1':4,'3era División C2':4};
const DIV_CLASSIFY={'1era División':8,'2da División':8,'3era División C1':8,'3era División C2':8};
let currentDiv='global';

// INIT
if(DATA.logo) document.getElementById('logoImg').src=DATA.logo;
document.getElementById('footerDate').textContent='Actualizado: '+DATA.generado;

// ── NAVEGACIÓN ──
function switchView(btn,id){
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('view-'+id).classList.add('active');
  currentDiv=id;
  const bn=document.getElementById('bottomNav');
  bn.style.display=id==='global'?'none':'flex';
  window.scrollTo({top:0,behavior:'smooth'});
}

function switchBottomNav(btn,sub){
  if(currentDiv==='global') return;
  document.querySelectorAll('.bnav-item').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const panel=document.getElementById('view-'+currentDiv);
  panel.querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));
  panel.querySelectorAll('.sub-panel').forEach(p=>p.classList.remove('active'));
  const stab=panel.querySelector(`.stab[data-sub="${sub}"]`);
  const sp=document.getElementById('sub-'+currentDiv+'-'+sub);
  if(stab) stab.classList.add('active');
  if(sp) sp.classList.add('active');
}

function switchSub(btn,divId,subId){
  const panel=document.getElementById('view-'+divId);
  panel.querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));
  panel.querySelectorAll('.sub-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('sub-'+divId+'-'+subId).classList.add('active');
  // Sync bottom nav
  document.querySelectorAll('.bnav-item').forEach(b=>{
    b.classList.toggle('active', b.dataset.sub===subId);
  });
}

// ── HELPERS ──
function rankNum(i){
  const cls=i===0?'gold':i===1?'silver':i===2?'bronze':'';
  return`<span class="rank-num ${cls}">${i+1}</span>`;
}
function posBadge(pos,n,clf,rsk){
  if(pos===1) return`<span class="pos-badge pb-gold">${pos}</span>`;
  if(pos<=clf) return`<span class="pos-badge pb-clf">${pos}</span>`;
  if(pos>n-rsk) return`<span class="pos-badge pb-bot">${pos}</span>`;
  return`<span class="pos-badge pb-norm">${pos}</span>`;
}
function cardBadges(c,style=''){
  let b='';
  if(c.rojas>0)    b+=`<span class="card-t roja" style="${style}">${c.rojas}<small style="font-size:8px">R</small></span>`;
  if(c.dobles>0)   b+=`<span class="card-t doble" style="${style}">${c.dobles}<small style="font-size:8px">D</small></span>`;
  if(c.amarillas>0)b+=`<span class="card-t amarilla" style="${style}">${c.amarillas}<small style="font-size:8px">A</small></span>`;
  return b;
}

// ── GLOBAL ──
function buildGlobal(){
  const g=DATA.globalTorneo;
  // Bento KPIs
  const bentoData=[
    {lbl:'Goles Totales',val:g.total_goles,cls:'red',icon:'sports_soccer',live:true},
    {lbl:'Equipos',val:g.total_equipos||g.ranking_divisiones.reduce((a,d)=>a+d.num_equipos,0)},
    {lbl:'Divisiones',val:g.ranking_divisiones.length},
    {lbl:'Tarjetas Totales',val:g.total_tarjetas||g.ranking_divisiones.reduce((a,d)=>a+d.total_tarjetas,0)},
    {lbl:'Goles Líder',val:g.top_15_goleadores[0]?.goles||0,cls:'red'},
    {lbl:'Prom. Goles/Part.',val:g.prom_goles_partido},
  ];
  const bento=document.getElementById('bentoKpis');
  bento.innerHTML=bentoData.map(k=>`
    <div class="kpi-card">
      ${k.live?'<div class="live-dot"></div>':''}
      <span class="kpi-label">${k.lbl}</span>
      <div class="kpi-val ${k.cls||''}">${k.val}</div>
      ${k.icon?`<span class="material-symbols-outlined kpi-icon">${k.icon}</span>`:''}
    </div>`).join('');

  // Div cards
  const dc=document.getElementById('divCards');
  dc.innerHTML='';
  g.ranking_divisiones.forEach((d,i)=>{
    const color=DIV_COLORS[d.division]||'#c8102e';
    dc.innerHTML+=`<div class="div-card">
      <div class="div-card-stripe" style="background:${color}"></div>
      <div class="div-card-body">
        <div style="display:flex;justify-content:space-between;align-items:flex-start">
          <div>
            <div class="div-card-name" style="color:${color}">${d.division}</div>
            <div class="div-card-eq">${d.num_equipos} Equipos</div>
          </div>
          <div class="div-card-rank">#${i+1}</div>
        </div>
        <div class="div-card-stats">
          <div class="div-stat"><div class="div-stat-val" style="color:${color}">${d.total_goles}</div><div class="div-stat-lbl">Goles</div></div>
          <div class="div-stat"><div class="div-stat-val">${d.promedio_goles_equipo}</div><div class="div-stat-lbl">Prom/Eq</div></div>
          <div class="div-stat"><div class="div-stat-val">${d.total_tarjetas}</div><div class="div-stat-lbl">Tarjetas</div></div>
        </div>
      </div>
    </div>`;
  });

  // Goleadores
  const maxG=g.top_15_goleadores[0]?.goles||1;
  const gg=document.getElementById('gGol');
  gg.innerHTML='';
  g.top_15_goleadores.forEach((p,i)=>{
    const color=DIV_COLORS[p.division]||'#c8102e';
    const pct=Math.round(p.goles/maxG*100);
    const isTop=i<3;
    gg.innerHTML+=`<div class="scorer-row ${isTop?'top':''}">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name">${p.nombre}</div>
        <div class="scorer-meta" style="color:${color}">${p.equipo} · ${p.division}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="scorer-val">${p.goles}</div>
    </div>`;
  });

  // Amonestaciones
  const gt=document.getElementById('gTarj');
  gt.innerHTML='';
  g.top_15_amonestados.forEach((c,i)=>{
    const color=DIV_COLORS[c.division]||'#c8102e';
    gt.innerHTML+=`<div class="scorer-row">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name">${c.nombre}</div>
        <div class="scorer-meta" style="color:${color}">${c.equipo} · ${c.division}</div>
      </div>
      <div style="display:flex;gap:2px;flex-wrap:wrap;justify-content:flex-end">${cardBadges(c)}</div>
    </div>`;
  });

  // Fair play
  const maxFP=Math.max(...g.fair_play.map(f=>f.total_tarjetas))||1;
  const gfp=document.getElementById('gFP');
  gfp.innerHTML='';
  g.fair_play.slice(0,14).forEach((eq,i)=>{
    const color=DIV_COLORS[eq.division]||'#c8102e';
    const pct=Math.round(eq.total_tarjetas/maxFP*100);
    const bc=i<3?'#22c55e':i<8?'#f5a623':'#c8102e';
    gfp.innerHTML+=`<div class="fp-row">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name" style="font-size:11px">${eq.equipo}</div>
        <div class="scorer-meta" style="color:${color};font-size:9px">${eq.division}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div>
    </div>`;
  });

  // Ranking equipos amonestados — tabla
  const tbody=document.getElementById('eqAmonBody');
  tbody.innerHTML='';
  g.ranking_equipos_amonestados.forEach((eq,i)=>{
    const color=DIV_COLORS[eq.division]||'#c8102e';
    const numStyle=i===0?`color:var(--gold);font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900`:`color:var(--text3);font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:800;opacity:.4`;
    tbody.innerHTML+=`<tr style="transition:background .15s" onmouseover="this.style.background='var(--surf2)'" onmouseout="this.style.background=''">
      <td style="text-align:center;padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04)"><span style="${numStyle}">${i+1}</span></td>
      <td style="padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04)">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:800;color:var(--text)">${eq.equipo}</div>
        <div style="font-size:9px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${color};margin-top:2px">${eq.division}</div>
      </td>
      <td style="text-align:center;padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04)"><span style="background:rgba(200,16,46,.2);color:#fca5a5;border:1px solid rgba(200,16,46,.3);font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;padding:2px 7px;border-radius:3px">${eq.rojas}R</span></td>
      <td style="text-align:center;padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04)"><span style="background:rgba(249,115,22,.2);color:#fdba74;border:1px solid rgba(249,115,22,.3);font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;padding:2px 7px;border-radius:3px">${eq.dobles}D</span></td>
      <td style="text-align:center;padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04)"><span style="background:rgba(234,179,8,.15);color:#fde047;border:1px solid rgba(234,179,8,.25);font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;padding:2px 7px;border-radius:3px">${eq.amarillas}A</span></td>
      <td style="text-align:right;padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.04);font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900;color:${i===0?'var(--red)':'var(--text)'}">${eq.puntos}</td>
    </tr>`;
  });
}

// ── TABLA POSICIONES ──
function buildTablaPos(divNombre){
  const tabla=DATA.posiciones[divNombre]||[];
  const color=DIV_COLORS[divNombre]||'#c8102e';
  const clf=DIV_CLASSIFY[divNombre]||8;
  const rsk=DIV_RIESGO[divNombre]||3;
  if(!tabla.length) return`<div class="empty">Sin tabla disponible</div>`;
  const n=tabla.length;
  let rows='';
  tabla.forEach(eq=>{
    const pos=eq.pos;
    const isTop=pos===1,isClf=pos>1&&pos<=clf,isBot=pos>n-rsk;
    const rowCls=isTop?'zone-top':isClf?'zone-clf':isBot?'zone-bot':'';
    const dgCls=eq.dg>0?'dg-pos':eq.dg<0?'dg-neg':'';
    rows+=`<tr class="${rowCls}">
      <td class="w36">${posBadge(pos,n,clf,rsk)}</td>
      <td class="eq-td">${eq.equipo}</td>
      <td class="w36">${eq.pj}</td><td class="w36">${eq.pg}</td>
      <td class="w36">${eq.pe}</td><td class="w36">${eq.pp}</td>
      <td class="w36">${eq.gf}</td><td class="w36">${eq.gc}</td>
      <td class="w36 ${dgCls}">${eq.dg>0?'+':''}${eq.dg}</td>
      <td class="w36 pts-td" style="color:${color}">${eq.pts}</td>
    </tr>`;
  });
  return`<div class="pos-wrap">
    <table class="pos-table">
      <thead><tr>
        <th class="w36">#</th><th class="eq-h">Equipo</th>
        <th class="w36">PJ</th><th class="w36">G</th><th class="w36">E</th><th class="w36">P</th>
        <th class="w36">GF</th><th class="w36">GC</th><th class="w36">DG</th>
        <th class="w36" style="color:${color}">PTS</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="tbl-legend">
      <span><span class="ldot" style="background:var(--green)"></span>Clasificación Fase 3 (top ${clf})</span>
      <span><span class="ldot" style="background:var(--red)"></span>Zona de riesgo (últimos ${rsk})</span>
    </div>
  </div>`;
}

// ── RESULTADOS ──
function buildResultados(divNombre){
  const fechas=DATA.resultados[divNombre]||[];
  if(!fechas.length) return`<div class="empty">Sin resultados registrados</div>`;
  let html='';
  fechas.forEach(f=>{
    html+=`<div class="fecha-block"><div class="fecha-lbl">Fecha ${f.fecha} — Fase 2</div>`;
    f.partidos.forEach(p=>{
      const draw=p.gl===p.gv,lW=p.gl>p.gv,vW=p.gv>p.gl;
      html+=`<div class="match-card">
        <div class="match-team ${lW?'winner':''}" style="opacity:${vW?.55:1}">${p.local}</div>
        <div class="match-score ${draw?'draw':''}">${p.gl} — ${p.gv}</div>
        <div class="match-team right ${vW?'winner':''}" style="opacity:${lW?.55:1}">${p.visitante}</div>
      </div>`;
    });
    html+='</div>';
  });
  return html;
}

// ── SANCIONADOS ──
function buildSancionados(divNombre){
  const lista=DATA.sancionados[divNombre]||[];
  if(!lista.length) return`<div class="empty">Sin sancionados registrados</div>`;
  let rows='';
  lista.forEach(s=>{
    const numD=isNaN(s.numero)?`<span style="font-size:9px;font-weight:700;color:var(--text3)">${s.numero}</span>`:`<span class="s-num">${s.numero}</span>`;
    rows+=`<tr onmouseover="this.querySelectorAll('td').forEach(td=>td.style.background='var(--surf2)')" onmouseout="this.querySelectorAll('td').forEach(td=>td.style.background='')">
      <td style="text-align:center;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04)">${numD}</td>
      <td style="padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04)"><span class="sanc-name">${s.nombre}</span></td>
      <td style="padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04);color:var(--text2);font-weight:600">${s.club}</td>
      <td style="padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04)"><span class="s-tag ${s.tipo}">${s.sancion}</span></td>
      <td style="text-align:center;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.04)">${s.fase2?`<span class="s-fase2">${s.fase2}</span>`:'-'}</td>
    </tr>`;
  });
  return`<div style="overflow-x:auto;border-radius:10px;border:1px solid var(--border)">
    <table class="sanc-table">
      <thead><tr>
        <th style="text-align:center;width:44px">#</th>
        <th>Jugador / Dirigente</th><th>Club</th>
        <th>Sanción</th>
        <th style="text-align:center">Fase 2</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>
  <div style="display:flex;gap:14px;flex-wrap:wrap;padding:10px 4px;font-size:10px;font-weight:600;color:var(--text3)">
    <span><span class="s-tag grave" style="font-size:9px">Grave</span> Indefinido / Años / Meses</span>
    <span><span class="s-tag media" style="font-size:9px">Media</span> 2-3 Fechas</span>
    <span><span class="s-tag leve" style="font-size:9px">Leve</span> 1 Fecha</span>
  </div>`;
}

// ── PANEL DIVISIÓN ──
function buildDivPanel(divNombre){
  const divData=DATA.divisiones[divNombre];
  const color=DIV_COLORS[divNombre]||'#c8102e';
  const meta=divData.meta;
  const gd=divData.globalData;
  const totalG=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_goles,0);
  const totalA=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_amarillas,0);
  const totalR=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_rojas+t.total_doble,0);
  const ne=Object.keys(divData.teamsData).length;
  const p=document.getElementById('view-'+divNombre);

  p.innerHTML=`
  <div class="div-hero" style="--accent:${color}">
    <div>
      <div class="div-hero-title" style="color:${color}">${divNombre}</div>
      <div class="div-hero-meta">Fecha ${meta.ultima_fecha} · ${ne} equipos</div>
    </div>
    <div class="div-kpis">
      <div style="text-align:center"><div class="div-kpi-val" style="color:${color}">${totalG}</div><div class="div-kpi-lbl">Goles</div></div>
      <div style="text-align:center"><div class="div-kpi-val">${totalA}</div><div class="div-kpi-lbl">Amarillas</div></div>
      <div style="text-align:center"><div class="div-kpi-val">${totalR}</div><div class="div-kpi-lbl">Rojas/Dbl</div></div>
    </div>
  </div>

  <!-- SUB-TABS escritorio -->
  <div class="sub-tabs">
    <button class="stab active" data-sub="pos" onclick="switchSub(this,'${divNombre}','pos')">POS</button>
    <button class="stab" data-sub="res" onclick="switchSub(this,'${divNombre}','res')">RES</button>
    <button class="stab" data-sub="stats" onclick="switchSub(this,'${divNombre}','stats')">EST</button>
    <button class="stab" data-sub="sanc" onclick="switchSub(this,'${divNombre}','sanc')">SAN</button>
  </div>

  <!-- POSICIONES -->
  <div id="sub-${divNombre}-pos" class="sub-panel active">
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">format_list_numbered</span><span class="panel-head-title">Tabla de Posiciones · ${divNombre}</span></div>
      <div class="panel-body" style="padding:0">${buildTablaPos(divNombre)}</div>
    </div>
  </div>

  <!-- RESULTADOS -->
  <div id="sub-${divNombre}-res" class="sub-panel">
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">calendar_today</span><span class="panel-head-title">Resultados · ${divNombre}</span></div>
      <div class="panel-body">${buildResultados(divNombre)}</div>
    </div>
  </div>

  <!-- ESTADÍSTICAS -->
  <div id="sub-${divNombre}-stats" class="sub-panel">
    <div class="team-sel-wrap">
      <span class="team-sel-lbl">Ver equipo</span>
      <select class="team-sel" id="sel-${divNombre}" onchange="renderTeam('${divNombre}',this.value)" style="border-color:${color}">
        <option value="__global__">📊 Estadística de la División</option>
      </select>
    </div>
    <div id="dv-global-${divNombre}">
      <div class="grid-3">
        <div class="panel">
          <div class="panel-head"><span class="material-symbols-outlined">leaderboard</span><span class="panel-head-title">Top Goleadores</span></div>
          <div class="panel-body"><div id="dg-${divNombre}"></div></div>
        </div>
        <div class="panel">
          <div class="panel-head"><span class="material-symbols-outlined">warning</span><span class="panel-head-title">Ranking Amonestaciones</span></div>
          <div class="panel-body"><div id="dt-${divNombre}"></div></div>
        </div>
        <div class="panel">
          <div class="panel-head"><span class="material-symbols-outlined" style="color:var(--green)">volunteer_activism</span><span class="panel-head-title">Fair Play</span></div>
          <div class="panel-body" id="df-${divNombre}"></div>
        </div>
      </div>
    </div>
    <div id="dv-team-${divNombre}" style="display:none">
      <div class="team-kpi-row" id="tkpi-${divNombre}"></div>
      <div class="grid-2" id="tgrid-${divNombre}"></div>
    </div>
  </div>

  <!-- SANCIONADOS -->
  <div id="sub-${divNombre}-sanc" class="sub-panel">
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">gavel</span><span class="panel-head-title">Sancionados · ${divNombre}</span></div>
      <div class="panel-body">${buildSancionados(divNombre)}</div>
    </div>
  </div>`;

  // Selector equipos
  const sel=document.getElementById('sel-'+divNombre);
  Object.keys(divData.teamsData).sort().forEach(t=>{
    const o=document.createElement('option');o.value=t;o.textContent='⚽ '+t;sel.appendChild(o);
  });

  // Listas globales división
  const maxG2=gd.top_10_goleadores[0]?.goles||1;
  const dgEl=document.getElementById('dg-'+divNombre);
  gd.top_10_goleadores.forEach((pl,i)=>{
    const pct=Math.round(pl.goles/maxG2*100);
    dgEl.innerHTML+=`<div class="scorer-row ${i<3?'top':''}">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name">${pl.nombre}</div>
        <div class="scorer-meta" style="color:${color}">${pl.equipo}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="scorer-val">${pl.goles}</div>
    </div>`;
  });

  const dtEl=document.getElementById('dt-'+divNombre);
  gd.top_10_amonestados.forEach((c,i)=>{
    dtEl.innerHTML+=`<div class="scorer-row">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name">${c.nombre}</div>
        <div class="scorer-meta" style="color:${color}">${c.equipo}</div>
      </div>
      <div style="display:flex;gap:2px;flex-wrap:wrap">${cardBadges(c)}</div>
    </div>`;
  });

  const maxFP2=Math.max(...gd.fair_play.map(f=>f.total_tarjetas))||1;
  const dfEl=document.getElementById('df-'+divNombre);
  gd.fair_play.forEach((eq,i)=>{
    const pct=Math.round(eq.total_tarjetas/maxFP2*100);
    const bc=i<3?'#22c55e':i<7?'#f5a623':'#c8102e';
    dfEl.innerHTML+=`<div class="fp-row">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name" style="font-size:11px">${eq.equipo}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div>
    </div>`;
  });
}

function renderTeam(divNombre,equipo){
  const gl=document.getElementById('dv-global-'+divNombre);
  const tm=document.getElementById('dv-team-'+divNombre);
  if(equipo==='__global__'){tm.style.display='none';gl.style.display='block';return;}
  gl.style.display='none';tm.style.display='block';
  const d=DATA.divisiones[divNombre].teamsData[equipo];
  const color=DIV_COLORS[divNombre]||'#c8102e';
  document.getElementById('tkpi-'+divNombre).innerHTML=`
    <div class="team-kpi"><div class="team-kpi-val" style="color:${color}">${d.total_goles}</div><div class="team-kpi-lbl">Goles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_goles}</div><div class="team-kpi-lbl">Prom. Goles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_amarillas}</div><div class="team-kpi-lbl">Amarillas</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_rojas+d.total_doble}</div><div class="team-kpi-lbl">Rojas/Dbl</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_tarjetas}</div><div class="team-kpi-lbl">Prom. Tarj.</div></div>`;
  const maxG3=d.top_scorers[0]?.goles||1;
  let golesH=d.top_scorers.length?'':'<div class="empty">Sin goles</div>';
  d.top_scorers.forEach((g,i)=>{
    const pct=Math.round(g.goles/maxG3*100);
    golesH+=`<div class="scorer-row ${i<3?'top':''}">
      ${rankNum(i)}
      <div class="scorer-info">
        <div class="scorer-name">${g.nombre}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="scorer-val">${g.goles}</div>
    </div>`;
  });
  let tarjH=d.top_cards.length?'':'<div class="empty">Sin tarjetas</div>';
  d.top_cards.forEach((c,i)=>{
    tarjH+=`<div class="scorer-row">
      ${rankNum(i)}
      <div class="scorer-info"><div class="scorer-name">${c.nombre}</div></div>
      <div style="display:flex;gap:2px;flex-wrap:wrap">${cardBadges(c,'font-size:12px;width:38px;height:44px')}</div>
    </div>`;
  });
  document.getElementById('tgrid-'+divNombre).innerHTML=`
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">leaderboard</span><span class="panel-head-title">${equipo} · Goleadores</span></div>
      <div class="panel-body">${golesH}</div>
    </div>
    <div class="panel">
      <div class="panel-head"><span class="material-symbols-outlined">warning</span><span class="panel-head-title">${equipo} · Disciplina</span></div>
      <div class="panel-body">${tarjH}</div>
    </div>`;
}

// INIT
buildGlobal();
['1era División','2da División','3era División C1','3era División C2'].forEach(div=>buildDivPanel(div));
// Bottom nav oculto en global
document.getElementById('bottomNav').style.display='none';
</script>
</body>
</html>"""

if __name__=='__main__':
    carpeta=os.path.dirname(os.path.abspath(__file__))
    datos=procesar_todo(carpeta)
    html=generar_html(datos)
    salida=os.path.join(carpeta,'index.html')
    with open(salida,'w',encoding='utf-8') as f: f.write(html)
    print(f"\n✅ ¡Listo! → {salida}")
    print(f"   Logo: {'✓' if datos['logo'] else '✗'}")
    print(f"   Sancionados: {sum(len(v) for v in datos['sancionados'].values())}")
    print(f"   Prom goles/partido: {datos['globalTorneo']['prom_goles_partido']}")
