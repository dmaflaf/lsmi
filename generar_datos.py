#!/usr/bin/env python3
"""
LSMI - Generador completo · Temporada 2026
Ejecutar: python generar_datos.py
Archivos necesarios en la misma carpeta:
  - LSMI_1era_Division_Vxx.xlsx
  - LSMI_2da_Division_Vxx.xlsx
  - LSMI_3era_Division_1_Vxx.xlsx
  - LSMI_3era_Division_2_Vxx.xlsx
  - LSMI_Torneo_vx_Fase_x.xlsx  (posiciones y resultados)
  - sancionados_san_miguel_ibarra.xlsx
"""
import openpyxl, json, sys, os, unicodedata
from datetime import datetime

def norm(s):
    return unicodedata.normalize('NFD', str(s)).encode('ascii','ignore').decode('ascii').upper()

HOJAS_POS = {
    '1era División':'1era División','2da División':'2da División',
    '3era División C1':'3era División 1','3era División C2':'3era División 2',
}
DIV_RES = {
    '1era División':'1era División','2da División':'2da División',
    '3era División C1':'3era División 1','3era División C2':'3era División 2',
}

# ── ESTADÍSTICAS ──────────────────────────────────────────────────────────────
def procesar_stats(ruta):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws_g = wb['GOLEADORES']
    goles_pj={}; uf=0
    for row in ws_g.iter_rows(min_row=3, values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        e=str(row[1]).strip() if row[1] else ''; j=str(row[2]).strip() if row[2] else ''; g=row[3]
        if e and j and g and isinstance(g,(int,float)):
            key=(e,j); goles_pj[key]=goles_pj.get(key,0)+int(g); uf=max(uf,int(row[0]))
    ws_a=wb['AMONESTACIONES']; tarj_pj={}
    for row in ws_a.iter_rows(min_row=3, values_only=True):
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
        tcr=[{'nombre':j,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':t['amarillas']+t['dobles']*2+t['rojas']*3}
             for j,t in te.items() if t['amarillas']+t['dobles']*2+t['rojas']*3>0]
        tcr.sort(key=lambda x:-x['_p'])
        teams_data[eq]={'total_goles':tg,'promedio_goles':pg,'total_amarillas':tam,'total_doble':tdo,
            'total_rojas':tro,'promedio_tarjetas':pt,'top_scorers':[{'nombre':j,'goles':g} for j,g in ts],
            'top_cards':[{k:v for k,v in c.items() if k!='_p'} for c in tcr[:8]]}
    all_s=[{'nombre':j,'equipo':e,'goles':g} for (e,j),g in goles_pj.items()]; all_s.sort(key=lambda x:-x['goles'])
    all_c=[{'nombre':j,'equipo':e,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':t['amarillas']+t['dobles']*2+t['rojas']*3}
           for (e,j),t in tarj_pj.items() if t['amarillas']+t['dobles']*2+t['rojas']*3>0]
    all_c.sort(key=lambda x:-x['_p'])
    fp=[{'equipo':eq,'total_tarjetas':sum(t['amarillas']+t['dobles']+t['rojas'] for t in {j:t for (e,j),t in tarj_pj.items() if e==eq}.values())} for eq in equipos]
    fp.sort(key=lambda x:x['total_tarjetas'])
    return {'meta':{'ultima_fecha':uf,'generado':datetime.now().strftime('%d/%m/%Y %H:%M')},
        'globalData':{'top_10_goleadores':all_s[:10],'top_10_amonestados':[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:10]],'fair_play':fp},
        'teamsData':teams_data,
        '_goles_raw':{f"{e}||{j}":g for (e,j),g in goles_pj.items()},
        '_tarjetas_raw':{f"{e}||{j}":t for (e,j),t in tarj_pj.items()}}

# ── POSICIONES ────────────────────────────────────────────────────────────────
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
                tabla.append({'pos':int(row[0]),'equipo':str(row[1]).strip(),
                    'pj':int(row[2] or 0),'pg':int(row[3] or 0),'pe':int(row[4] or 0),'pp':int(row[5] or 0),
                    'gf':int(row[6] or 0),'gc':int(row[7] or 0),'dg':int(row[8] or 0),'pts':int(row[9] or 0)})
        pos[div]=tabla
    return pos

# ── RESULTADOS ────────────────────────────────────────────────────────────────
def leer_resultados(wb_pos):
    ws_r=wb_pos['RESULTADOS']; res={}
    for row in ws_r.iter_rows(min_row=2, values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        if not row[2] or not row[5]: continue
        fecha=int(row[0]); div_raw=str(row[1]).strip() if row[1] else ''
        local=str(row[2]).strip(); gf=int(row[3] or 0); gv=int(row[4] or 0); vis=str(row[5]).strip()
        if not div_raw or not local or not vis: continue
        div_nombre=None
        for k,v in DIV_RES.items():
            if v==div_raw: div_nombre=k; break
        if not div_nombre: div_nombre=div_raw
        if div_nombre not in res: res[div_nombre]={}
        if fecha not in res[div_nombre]: res[div_nombre][fecha]=[]
        res[div_nombre][fecha].append({'local':local,'gl':gf,'gv':gv,'visitante':vis})
    return {div:[{'fecha':f,'partidos':ps} for f,ps in sorted(fechas.items(),reverse=True)] for div,fechas in res.items()}

# ── SANCIONADOS ───────────────────────────────────────────────────────────────
def leer_sancionados(ruta):
    DIV_MAP={norm('1era DIVISIÓN'):'1era División',norm('2da DIVISIÓN'):'2da División',
        norm('3era DIVISIÓN 1'):'3era División C1',norm('3era DIVISIÓN 2'):'3era División C2'}
    wb=openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws=wb[wb.sheetnames[0]]; sanc={}; div_actual=None
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
            sanc[div_actual].append({'numero':numero,'nombre':str(c2).strip(),'club':str(c3).strip(),
                'sancion':str(c4).strip(),'fase2':str(c5).strip() if c5 else '','tipo':tipo})
    return sanc

# ── BUSCAR ARCHIVOS ───────────────────────────────────────────────────────────
def buscar_excels(carpeta):
    archivos=[f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    stats={}
    claves={'1era División':['1era','primera'],'2da División':['2da','segunda'],
        '3era División C1':['division_1','3era_division_1','_1_v','g1','c1'],
        '3era División C2':['division_2','3era_division_2','_2_v','g2','c2']}
    for arch in archivos:
        low=arch.lower()
        for div,pals in claves.items():
            if any(p in low for p in pals) and div not in stats:
                stats[div]=os.path.join(carpeta,arch); break
    pos_ruta=next((os.path.join(carpeta,f) for f in archivos if any(k in f.lower() for k in ['torneo','posicion','fase'])),None)
    sanc_ruta=next((os.path.join(carpeta,f) for f in archivos if 'sancionado' in f.lower()),None)
    return stats,pos_ruta,sanc_ruta

# ── MAIN ─────────────────────────────────────────────────────────────────────
def procesar_todo(carpeta):
    stats_rutas,pos_ruta,sanc_ruta=buscar_excels(carpeta)
    print("📁 Estadísticas:"); [print(f"   {d}: {os.path.basename(r)}") for d,r in stats_rutas.items()]
    if pos_ruta:  print(f"📁 Posiciones: {os.path.basename(pos_ruta)}")
    if sanc_ruta: print(f"📁 Sancionados: {os.path.basename(sanc_ruta)}")

    res_stats={}
    for nombre,ruta in stats_rutas.items():
        print(f"\n⚽ {nombre}..."); res_stats[nombre]=procesar_stats(ruta)
        print(f"   {len(res_stats[nombre]['teamsData'])} equipos | F{res_stats[nombre]['meta']['ultima_fecha']}")

    posiciones={}; res_partidos={}
    if pos_ruta:
        wb_pos=openpyxl.load_workbook(pos_ruta,read_only=True,data_only=True)
        posiciones=leer_posiciones(wb_pos); res_partidos=leer_resultados(wb_pos)
        print(f"\n📊 Posiciones y resultados cargados")

    sancionados={}
    if sanc_ruta:
        sancionados=leer_sancionados(sanc_ruta)
        total_sanc=sum(len(v) for v in sancionados.values())
        print(f"⚠️  Sancionados: {total_sanc} jugadores")

    # Calcular promedio goles por partido (todos los partidos del torneo)
    total_goles_torneo=0; total_partidos_torneo=0
    for div,fechas in res_partidos.items():
        for fecha_data in fechas:
            for p in fecha_data['partidos']:
                total_goles_torneo+=p['gl']+p['gv']; total_partidos_torneo+=1
    prom_goles_partido=round(total_goles_torneo/total_partidos_torneo,2) if total_partidos_torneo else 0

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
        stats_div.append({'division':dn,'total_goles':tg,'total_tarjetas':tt,'num_equipos':ne,
            'promedio_goles_equipo':round(tg/ne,1) if ne else 0,'promedio_tarjetas_equipo':round(tt/ne,1) if ne else 0})
    all_g.sort(key=lambda x:-x['goles']); all_c.sort(key=lambda x:-x['_p']); fp_t.sort(key=lambda x:x['total_tarjetas']); stats_div.sort(key=lambda x:-x['total_goles'])
    top_c=[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:15]]
    global_torneo={'top_15_goleadores':all_g[:15],'top_15_amonestados':top_c,'fair_play':fp_t[:20],
        'ranking_divisiones':stats_div,'prom_goles_partido':prom_goles_partido,
        'total_partidos':total_partidos_torneo,'total_sancionados':sum(len(v) for v in sancionados.values())}
    for d in res_stats.values(): d.pop('_goles_raw',None); d.pop('_tarjetas_raw',None)
    return {'generado':datetime.now().strftime('%d/%m/%Y %H:%M'),'divisiones':res_stats,
        'posiciones':posiciones,'resultados':res_partidos,'sancionados':sancionados,'globalTorneo':global_torneo}

def generar_html(datos):
    return HTML.replace('__JSON__', json.dumps(datos, ensure_ascii=False))

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>LSMI · Liga San Miguel de Ibarra</title>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f5f6f8;--surface:#ffffff;--surface2:#f0f2f5;--surface3:#e8ebf0;
  --border:#dde1ea;--border2:#c8cdd8;
  --red:#c8102e;--red-light:#fdf0f2;--red-mid:#f5c6cc;
  --gray:#1a1f2e;--gray2:#3d4460;--gray3:#6b7394;--gray4:#9ba3bf;
  --gold:#b8860b;--green:#1a7a3c;--green-light:#f0faf4;
  --orange:#c45c00;--orange-light:#fff4ec;
  --blue:#1a3a6e;
  --text:#1a1f2e;--text2:#3d4460;--text3:#6b7394;--text4:#9ba3bf;
  --shadow:0 1px 3px rgba(0,0,0,.08),0 4px 12px rgba(0,0,0,.05);
  --shadow-md:0 4px 16px rgba(0,0,0,.10),0 1px 4px rgba(0,0,0,.06);
  --1era:#c8102e;--2da:#1a3a6e;--c1:#1a7a3c;--c2:#b8860b;
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

/* ── HEADER ── */
.site-header{background:var(--gray);border-bottom:3px solid var(--red);position:sticky;top:0;z-index:100;box-shadow:var(--shadow-md)}
.header-inner{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;padding:14px 24px;gap:16px}
.logo-block{display:flex;align-items:center;gap:14px}
.logo-emblem{width:48px;height:48px;background:var(--red);border-radius:6px;display:flex;align-items:center;justify-content:center;font-family:'Barlow Condensed',sans-serif;font-size:24px;font-weight:900;color:white;letter-spacing:-1px;flex-shrink:0}
.logo-title{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900;letter-spacing:1.5px;color:#fff;text-transform:uppercase;line-height:1}
.logo-sub{font-size:10px;font-weight:600;color:rgba(255,255,255,.45);letter-spacing:2.5px;text-transform:uppercase;margin-top:3px}
.header-season{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;color:var(--red);letter-spacing:2px;text-transform:uppercase;background:rgba(200,16,46,.12);border:1px solid rgba(200,16,46,.25);padding:5px 12px;border-radius:4px}

/* ── NAV TABS ── */
.div-tabs{background:var(--surface);border-bottom:2px solid var(--border);position:sticky;top:79px;z-index:90;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.div-tabs-inner{max-width:1400px;margin:0 auto;display:flex;overflow-x:auto;scrollbar-width:none;padding:0 20px}
.div-tabs-inner::-webkit-scrollbar{display:none}
.dtab{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;padding:14px 22px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .18s;white-space:nowrap}
.dtab:hover{color:var(--text2);background:var(--surface2)}
.dtab.active{color:var(--red);border-bottom-color:var(--red);background:var(--red-light)}

/* ── MAIN ── */
.main{max-width:1400px;margin:0 auto;padding:28px 20px}
.tab-panel{display:none}
.tab-panel.active{display:block}

/* ── SUB-TABS ── */
.sub-tabs{display:flex;gap:4px;margin-bottom:24px;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:4px;width:fit-content}
.stab{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:8px 18px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-radius:6px;transition:all .18s}
.stab:hover{color:var(--text2);background:var(--surface2)}
.stab.active{background:var(--gray);color:#fff}
.sub-panel{display:none}
.sub-panel.active{display:block}

/* ── SECTION HEADER ── */
.section-eyebrow{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;letter-spacing:3px;text-transform:uppercase;color:var(--red);margin-bottom:4px}
.section-title{font-family:'Barlow Condensed',sans-serif;font-size:30px;font-weight:900;letter-spacing:.5px;text-transform:uppercase;color:var(--gray);line-height:1}
.section-title span{color:var(--red)}

/* ── HERO STATS ── */
.hero-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:0;background:var(--border);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin:20px 0;box-shadow:var(--shadow)}
.hero-stat{background:var(--surface);padding:22px 20px;text-align:center;border-right:1px solid var(--border)}
.hero-stat:last-child{border-right:none}
.hero-stat-val{font-family:'Barlow Condensed',sans-serif;font-size:40px;font-weight:900;line-height:1;color:var(--gray)}
.hero-stat-val.red{color:var(--red)}
.hero-stat-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text4);margin-top:4px}

/* ── DIV CARDS ── */
.div-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px;margin:20px 0}
.div-card{background:var(--surface);border-radius:10px;border:1px solid var(--border);overflow:hidden;box-shadow:var(--shadow);transition:box-shadow .2s}
.div-card:hover{box-shadow:var(--shadow-md)}
.div-card-header{padding:14px 18px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border)}
.div-card-title{font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:800;letter-spacing:1px;text-transform:uppercase}
.div-card-rank{font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:900;color:var(--border2);line-height:1}
.div-card-body{padding:14px 18px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;background:var(--surface2)}
.div-stat-mini{text-align:center}
.div-stat-mini-val{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--gray)}
.div-stat-mini-lbl{font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--text4);margin-top:2px}

/* ── GRID ── */
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin:20px 0}
.grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:16px 0}
@media(max-width:960px){.grid-3{grid-template-columns:1fr}.grid-2{grid-template-columns:1fr}}

/* ── PANEL ── */
.panel{background:var(--surface);border-radius:10px;border:1px solid var(--border);overflow:hidden;box-shadow:var(--shadow)}
.panel-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;background:var(--surface)}
.panel-head-title{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:var(--gray)}
.panel-body{padding:12px}

/* ── RANK ITEMS ── */
.rank-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;border:1px solid transparent;transition:all .15s;margin-bottom:3px}
.rank-item:hover{background:var(--surface2);border-color:var(--border)}
.rank-num{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;color:var(--text4);min-width:18px;text-align:center}
.rank-num.gold{color:var(--gold)}.rank-num.silver{color:#8a939a}.rank-num.bronze{color:#a0784a}
.rank-info{flex:1;overflow:hidden}
.rank-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-sub{font-size:10px;font-weight:600;color:var(--text3);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-val{font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:800;color:var(--gray);min-width:32px;text-align:right}
.rank-val small{font-size:10px;color:var(--text4)}

/* ── BADGES ── */
.badge{display:inline-flex;align-items:center;font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;padding:2px 6px;border-radius:3px;margin-left:3px;letter-spacing:.5px}
.badge-yellow{background:#fef9c3;color:#854d0e;border:1px solid #fde68a}
.badge-red{background:#fef2f2;color:#991b1b;border:1px solid #fecaca}
.badge-orange{background:#fff7ed;color:#9a3412;border:1px solid #fed7aa}

/* ── SCORER BARS ── */
.scorer-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;border:1px solid transparent;transition:all .15s;margin-bottom:3px}
.scorer-item:hover{background:var(--surface2);border-color:var(--border)}
.scorer-bar-wrap{flex:1;height:2px;background:var(--surface3);border-radius:2px;overflow:hidden;margin-top:4px}
.scorer-bar{height:100%;border-radius:2px}

/* ── FAIR PLAY ── */
.fp-item{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:6px;border:1px solid transparent;margin-bottom:3px;transition:all .15s}
.fp-item:hover{background:var(--surface2);border-color:var(--border)}
.fp-bar-wrap{flex:1;height:3px;background:var(--surface3);border-radius:2px;overflow:hidden}
.fp-bar{height:100%;border-radius:2px}
.fp-val{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;color:var(--gray);min-width:28px;text-align:right}

/* ── TABLA POSICIONES ── */
.pos-table-wrap{overflow-x:auto;border-radius:8px;border:1px solid var(--border)}
.pos-table{width:100%;border-collapse:collapse;font-size:12px}
.pos-table thead{background:var(--gray)}
.pos-table th{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:rgba(255,255,255,.6);padding:10px 10px;text-align:center;white-space:nowrap}
.pos-table th:nth-child(2){text-align:left;min-width:140px}
.pos-table td{padding:10px 10px;text-align:center;border-bottom:1px solid var(--border);color:var(--text2);font-weight:500;font-size:12px}
.pos-table td:nth-child(2){text-align:left;font-weight:700;color:var(--text)}
.pos-table tbody tr:hover td{background:var(--surface2)}
.pos-table tbody tr:last-child td{border-bottom:none}
.pts-col{font-family:'Barlow Condensed',sans-serif;font-size:17px;font-weight:900;color:var(--gray) !important}
.dg-pos{color:var(--green) !important;font-weight:700}
.dg-neg{color:var(--red) !important;font-weight:700}
/* Zonas clasificación */
.zone-clf td{background:var(--green-light)}
.zone-clf td:nth-child(2){border-left:3px solid var(--green)}
.zone-bot td{background:var(--orange-light)}
.zone-bot td:nth-child(2){border-left:3px solid var(--orange)}
.zone-top td:nth-child(2){border-left:3px solid var(--gold)}
.pos-badge{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:900}
.pos-badge.gold{background:#fef9c3;color:var(--gold)}
.pos-badge.clf{background:var(--green-light);color:var(--green)}
.pos-badge.bot{background:var(--orange-light);color:var(--orange)}
.pos-badge.norm{background:var(--surface2);color:var(--text3)}
.table-legend{display:flex;gap:16px;flex-wrap:wrap;padding:10px 14px;border-top:1px solid var(--border);font-size:10px;font-weight:600;color:var(--text4);background:var(--surface2)}
.legend-dot{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px}

/* ── RESULTADOS ── */
.fecha-block{margin-bottom:20px}
.fecha-label{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:6px 12px;background:var(--surface2);border:1px solid var(--border);border-radius:5px;margin-bottom:10px;display:inline-block}
.match-card{display:flex;align-items:center;padding:10px 16px;background:var(--surface);border:1px solid var(--border);border-radius:8px;margin-bottom:6px;gap:8px;transition:box-shadow .15s}
.match-card:hover{box-shadow:var(--shadow-md)}
.match-team{font-size:13px;font-weight:600;color:var(--text2);flex:1;line-height:1.3}
.match-team.right{text-align:right}
.match-team.winner{color:var(--text);font-weight:700}
.match-score{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:900;color:var(--gray);background:var(--surface2);border:1px solid var(--border);padding:5px 14px;border-radius:6px;min-width:64px;text-align:center;letter-spacing:2px}
.match-score.draw{color:var(--text3)}

/* ── SANCIONADOS ── */
.sanc-table{width:100%;border-collapse:collapse;font-size:12px}
.sanc-table th{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:rgba(255,255,255,.6);padding:10px 12px;text-align:left;background:var(--gray);white-space:nowrap}
.sanc-table td{padding:9px 12px;border-bottom:1px solid var(--border);color:var(--text2);font-size:12px;font-weight:500;vertical-align:middle}
.sanc-table tbody tr:hover td{background:var(--surface2)}
.sanc-table tbody tr:last-child td{border-bottom:none}
.sanc-name{font-weight:700;color:var(--text)}
.sanc-tag{display:inline-flex;align-items:center;font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;padding:2px 8px;border-radius:3px;letter-spacing:.5px;white-space:nowrap}
.sanc-tag.grave{background:#fef2f2;color:#991b1b;border:1px solid #fecaca}
.sanc-tag.media{background:#fff7ed;color:#9a3412;border:1px solid #fed7aa}
.sanc-tag.leve{background:#fefce8;color:#854d0e;border:1px solid #fde68a}
.sanc-fase2{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;color:var(--red);background:var(--red-light);border:1px solid var(--red-mid);padding:2px 8px;border-radius:3px;white-space:nowrap}
.sanc-numero{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:900;color:var(--text4);background:var(--surface2);border:1px solid var(--border);width:30px;height:30px;border-radius:4px;display:inline-flex;align-items:center;justify-content:center}

/* ── DIV HERO ── */
.div-hero{background:var(--surface);border-radius:10px;border:1px solid var(--border);border-top:3px solid var(--accent,var(--red));padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;box-shadow:var(--shadow)}
.div-hero-title{font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:900;letter-spacing:1px;text-transform:uppercase}
.div-hero-meta{font-size:11px;color:var(--text3);font-weight:500;margin-top:3px}
.div-kpis{display:flex;gap:28px;flex-wrap:wrap}
.div-kpi{text-align:center}
.div-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:900;color:var(--gray)}
.div-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text4)}

/* ── TEAM SELECTOR ── */
.team-selector-wrap{background:var(--surface);border-radius:10px;border:1px solid var(--border);padding:14px 18px;margin-bottom:18px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;box-shadow:var(--shadow)}
.team-selector-lbl{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);white-space:nowrap}
.team-select{flex:1;min-width:180px;background:var(--surface);color:var(--text);border:1.5px solid var(--border2);border-radius:6px;padding:8px 12px;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;cursor:pointer;outline:none;transition:border-color .15s}
.team-select:focus{border-color:var(--red)}

/* ── TEAM KPIS ── */
.team-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:1px;background:var(--border);border-radius:10px;overflow:hidden;margin-bottom:18px;box-shadow:var(--shadow)}
.team-kpi{background:var(--surface);padding:16px;text-align:center}
.team-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:900;color:var(--gray);line-height:1}
.team-kpi-val.accent{color:var(--accent,var(--red))}
.team-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text4);margin-top:4px}

/* ── EMPTY ── */
.empty{color:var(--text4);font-size:12px;padding:24px;text-align:center;font-style:italic}

/* ── FOOTER ── */
.site-footer{background:var(--gray);color:rgba(255,255,255,.4);text-align:center;padding:20px 16px;margin-top:40px;font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;line-height:2}
.site-footer .powered{color:var(--red);font-weight:900;letter-spacing:2px}
.site-footer .update-time{color:rgba(255,255,255,.6);font-size:10px}

/* ── DIVIDER ── */
.divider{height:1px;background:var(--border);margin:20px 0}

@media(max-width:600px){
  .header-inner{padding:10px 14px}
  .logo-title{font-size:15px}
  .main{padding:16px 12px}
  .div-hero{padding:14px 16px}
  .div-hero-title{font-size:24px}
  .pos-table th,.pos-table td{padding:8px 6px;font-size:11px}
  .match-card{padding:8px 10px}
  .match-score{padding:4px 10px;min-width:52px;font-size:17px}
}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-inner">
    <div class="logo-block">
      <div class="logo-emblem">L</div>
      <div>
        <div class="logo-title">Liga San Miguel de Ibarra</div>
        <div class="logo-sub">LSMI · Estadísticas Oficiales</div>
      </div>
    </div>
    <div class="header-season">Temporada 2026</div>
  </div>
</header>

<nav class="div-tabs">
  <div class="div-tabs-inner">
    <button class="dtab active" data-tab="global" onclick="switchTab(this,'global')">🏆 Global</button>
    <button class="dtab" data-tab="1era División" onclick="switchTab(this,'1era División')">1era División</button>
    <button class="dtab" data-tab="2da División" onclick="switchTab(this,'2da División')">2da División</button>
    <button class="dtab" data-tab="3era División C1" onclick="switchTab(this,'3era División C1')">3era Div. C1</button>
    <button class="dtab" data-tab="3era División C2" onclick="switchTab(this,'3era División C2')">3era Div. C2</button>
  </div>
</nav>

<main class="main">
  <div id="panel-global" class="tab-panel active"></div>
  <div id="panel-1era División" class="tab-panel"></div>
  <div id="panel-2da División" class="tab-panel"></div>
  <div id="panel-3era División C1" class="tab-panel"></div>
  <div id="panel-3era División C2" class="tab-panel"></div>
</main>

<footer class="site-footer">
  <div>LIGA SAN MIGUEL DE IBARRA · ESTADÍSTICAS OFICIALES · <span class="powered">POWERED BY NIX 26</span></div>
  <div class="update-time" id="footerDate"></div>
</footer>

<script>
const DATA = __JSON__;
const DIV_COLORS={'1era División':'#c8102e','2da División':'#1a3a6e','3era División C1':'#1a7a3c','3era División C2':'#b8860b'};
const DIV_CLASSIFY={
  '1era División':8,'2da División':8,'3era División C1':8,'3era División C2':8
};

document.getElementById('footerDate').textContent='Actualizado: '+DATA.generado;

// ── TABS ──
function switchTab(btn,id){
  document.querySelectorAll('.dtab').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('panel-'+id).classList.add('active');
  window.scrollTo({top:0,behavior:'smooth'});
}
function switchSub(btn,divId,subId){
  const panel=document.getElementById('panel-'+divId);
  panel.querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));
  panel.querySelectorAll('.sub-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('sub-'+divId+'-'+subId).classList.add('active');
}

// ── HELPERS ──
function badge(c,big=false){
  const sz=big?'font-size:12px;padding:3px 9px':'';
  let b='';
  if(c.rojas>0)   b+=`<span class="badge badge-red" style="${sz}">${c.rojas} ROJA${c.rojas>1?'S':''}</span>`;
  if(c.dobles>0)  b+=`<span class="badge badge-orange" style="${sz}">${c.dobles} DBL</span>`;
  if(c.amarillas>0)b+=`<span class="badge badge-yellow" style="${sz}">${c.amarillas} AMAR.</span>`;
  return b;
}
function rankNum(i){
  if(i===0)return'<span class="rank-num gold">1</span>';
  if(i===1)return'<span class="rank-num silver">2</span>';
  if(i===2)return'<span class="rank-num bronze">3</span>';
  return`<span class="rank-num">${i+1}</span>`;
}
function posBadge(pos,n,classify){
  if(pos===1) return`<span class="pos-badge gold">${pos}</span>`;
  if(pos<=classify) return`<span class="pos-badge clf">${pos}</span>`;
  if(pos>n-3) return`<span class="pos-badge bot">${pos}</span>`;
  return`<span class="pos-badge norm">${pos}</span>`;
}

// ── GLOBAL PANEL ──
function buildGlobal(){
  const g=DATA.globalTorneo;
  const totalG=g.ranking_divisiones.reduce((a,d)=>a+d.total_goles,0);
  const totalT=g.ranking_divisiones.reduce((a,d)=>a+d.total_tarjetas,0);
  const totalE=g.ranking_divisiones.reduce((a,d)=>a+d.num_equipos,0);
  const p=document.getElementById('panel-global');
  p.innerHTML=`
  <div style="margin-bottom:24px">
    <div class="section-eyebrow">Temporada 2026</div>
    <div class="section-title">Estadística <span>Global</span> del Torneo</div>
  </div>
  <div class="hero-stats">
    <div class="hero-stat"><div class="hero-stat-val red">${totalG}</div><div class="hero-stat-lbl">Goles en el Torneo</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalE}</div><div class="hero-stat-lbl">Equipos</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${g.ranking_divisiones.length}</div><div class="hero-stat-lbl">Divisiones</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalT}</div><div class="hero-stat-lbl">Tarjetas Totales</div></div>
    <div class="hero-stat"><div class="hero-stat-val red">${g.top_15_goleadores[0]?.goles||0}</div><div class="hero-stat-lbl">Goles Líder</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${g.prom_goles_partido}</div><div class="hero-stat-lbl">Promedio Goles/Partido</div></div>
  </div>
  <div class="div-cards" id="gDivCards"></div>
  <div class="grid-3">
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">⚽</span><span class="panel-head-title">Top 15 Goleadores</span></div>
      <div class="panel-body"><div id="gGol"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🟨</span><span class="panel-head-title">Top 15 Disciplina</span></div>
      <div class="panel-body"><div id="gTarj"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🤝</span><span class="panel-head-title">Fair Play General</span></div>
      <div class="panel-body" id="gFP"></div>
    </div>
  </div>`;

  // Ranking divisiones
  const rd=document.getElementById('gDivCards');
  g.ranking_divisiones.forEach((d,i)=>{
    const color=DIV_COLORS[d.division]||'#c8102e';
    rd.innerHTML+=`<div class="div-card">
      <div class="div-card-header">
        <div>
          <div class="div-card-title" style="color:${color}">${d.division}</div>
          <div style="font-size:9px;color:var(--text4);font-weight:700;letter-spacing:1px;margin-top:2px">${d.num_equipos} EQUIPOS</div>
        </div>
        <div class="div-card-rank">#${i+1}</div>
      </div>
      <div class="div-card-body">
        <div class="div-stat-mini"><div class="div-stat-mini-val" style="color:${color}">${d.total_goles}</div><div class="div-stat-mini-lbl">Goles</div></div>
        <div class="div-stat-mini"><div class="div-stat-mini-val">${d.promedio_goles_equipo}</div><div class="div-stat-mini-lbl">Prom/Eq</div></div>
        <div class="div-stat-mini"><div class="div-stat-mini-val">${d.total_tarjetas}</div><div class="div-stat-mini-lbl">Tarjetas</div></div>
      </div>
    </div>`;
  });

  // Goleadores
  const maxG=g.top_15_goleadores[0]?.goles||1;
  const gg=document.getElementById('gGol');
  g.top_15_goleadores.forEach((pl,i)=>{
    const color=DIV_COLORS[pl.division]||'#c8102e';
    const pct=Math.round(pl.goles/maxG*100);
    gg.innerHTML+=`<div class="scorer-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${pl.nombre}</div>
        <div class="rank-sub" style="color:${color}">${pl.equipo} · ${pl.division}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${pl.goles}<small> ⚽</small></div></div>`;
  });

  // Tarjetas
  const gt=document.getElementById('gTarj');
  g.top_15_amonestados.forEach((c,i)=>{
    const color=DIV_COLORS[c.division]||'#c8102e';
    gt.innerHTML+=`<div class="rank-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${c.nombre}</div>
        <div class="rank-sub" style="color:${color}">${c.equipo} · ${c.division}</div>
      </div>
      <div style="display:flex;gap:3px;flex-wrap:wrap">${badge(c)}</div></div>`;
  });

  // Fair play
  const maxFP=Math.max(...g.fair_play.map(f=>f.total_tarjetas))||1;
  const gfp=document.getElementById('gFP');
  g.fair_play.slice(0,16).forEach((eq,i)=>{
    const color=DIV_COLORS[eq.division]||'#c8102e';
    const pct=Math.round(eq.total_tarjetas/maxFP*100);
    const bc=i<3?'#1a7a3c':i<8?'#b8860b':'#c8102e';
    gfp.innerHTML+=`<div class="fp-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${eq.equipo}</div>
        <div class="rank-sub" style="color:${color}">${eq.division}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div></div>`;
  });
}

// ── TABLA POSICIONES ──
function buildTablaPos(divNombre){
  const tabla=DATA.posiciones[divNombre]||[];
  const color=DIV_COLORS[divNombre]||'#c8102e';
  const classify=DIV_CLASSIFY[divNombre]||8;
  if(!tabla.length) return`<div class="empty">Sin tabla de posiciones disponible</div>`;
  const n=tabla.length;
  let rows='';
  tabla.forEach((eq,i)=>{
    const pos=eq.pos;
    const isTop=pos===1; const isClf=pos>1&&pos<=classify; const isBot=pos>n-3;
    const rowCls=isTop?'zone-top':isClf?'zone-clf':isBot?'zone-bot':'';
    const dgCls=eq.dg>0?'dg-pos':eq.dg<0?'dg-neg':'';
    rows+=`<tr class="${rowCls}">
      <td>${posBadge(pos,n,classify)}</td>
      <td>${eq.equipo}</td>
      <td>${eq.pj}</td><td>${eq.pg}</td><td>${eq.pe}</td><td>${eq.pp}</td>
      <td>${eq.gf}</td><td>${eq.gc}</td>
      <td class="${dgCls}">${eq.dg>0?'+':''}${eq.dg}</td>
      <td class="pts-col" style="color:${color}">${eq.pts}</td>
    </tr>`;
  });
  return`<div class="pos-table-wrap">
    <table class="pos-table">
      <thead><tr>
        <th>#</th><th>Equipo</th>
        <th title="Partidos Jugados">PJ</th><th title="Ganados">G</th>
        <th title="Empatados">E</th><th title="Perdidos">P</th>
        <th title="Goles a Favor">GF</th><th title="Goles en Contra">GC</th>
        <th title="Diferencia">DG</th><th style="color:${color}">PTS</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="table-legend">
      <span><span class="legend-dot" style="background:#1a7a3c"></span>Clasificación Fase 3 (top ${classify})</span>
      <span><span class="legend-dot" style="background:#c45c00"></span>Zona de riesgo</span>
      <span>PJ · G · E · P · GF · GC · DG · PTS</span>
    </div>
  </div>`;
}

// ── RESULTADOS ──
function buildResultados(divNombre){
  const fechas=DATA.resultados[divNombre]||[];
  if(!fechas.length) return`<div class="empty">Sin resultados registrados</div>`;
  let html='';
  fechas.forEach((f,fi)=>{
    html+=`<div class="fecha-block"><div class="fecha-label">Fecha ${f.fecha} — Fase 2</div>`;
    f.partidos.forEach(p=>{
      const draw=p.gl===p.gv, lW=p.gl>p.gv, vW=p.gv>p.gl;
      html+=`<div class="match-card">
        <div class="match-team ${lW?'winner':''}" style="opacity:${vW?.6:1}">${p.local}</div>
        <div class="match-score ${draw?'draw':''}">${p.gl} — ${p.gv}</div>
        <div class="match-team right ${vW?'winner':''}" style="opacity:${lW?.6:1}">${p.visitante}</div>
      </div>`;
    });
    html+='</div>';
  });
  return html;
}

// ── SANCIONADOS ──
function buildSancionados(divNombre){
  const lista=DATA.sancionados[divNombre]||[];
  if(!lista.length) return`<div class="empty">Sin sancionados registrados para esta división</div>`;
  let rows='';
  lista.forEach(s=>{
    const numDisplay=isNaN(s.numero)?`<span style="font-size:10px;font-weight:700;color:var(--text3)">${s.numero}</span>`:`<span class="sanc-numero">${s.numero}</span>`;
    rows+=`<tr>
      <td style="text-align:center">${numDisplay}</td>
      <td><span class="sanc-name">${s.nombre}</span></td>
      <td style="color:var(--text2);font-weight:600">${s.club}</td>
      <td><span class="sanc-tag ${s.tipo}">${s.sancion}</span></td>
      <td style="text-align:center">${s.fase2?`<span class="sanc-fase2">${s.fase2}</span>`:'-'}</td>
    </tr>`;
  });
  return`<div style="overflow-x:auto;border-radius:8px;border:1px solid var(--border)">
    <table class="sanc-table">
      <thead><tr>
        <th style="text-align:center;width:48px">#</th>
        <th>Jugador / Dirigente</th>
        <th>Club</th>
        <th>Sanción</th>
        <th style="text-align:center">Suspensión Fase 2</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>
  <div style="font-size:10px;color:var(--text4);padding:8px 4px;display:flex;gap:14px;flex-wrap:wrap;margin-top:8px">
    <span><span class="sanc-tag grave" style="font-size:9px">Grave</span> Indefinido / Años / Meses</span>
    <span><span class="sanc-tag media" style="font-size:9px">Media</span> 2-3 Fechas</span>
    <span><span class="sanc-tag leve" style="font-size:9px">Leve</span> 1 Fecha</span>
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
  const p=document.getElementById('panel-'+divNombre);

  p.innerHTML=`
  <div class="div-hero" style="--accent:${color}">
    <div>
      <div class="div-hero-title" style="color:${color}">${divNombre}</div>
      <div class="div-hero-meta">Fecha ${meta.ultima_fecha} · ${ne} equipos</div>
    </div>
    <div class="div-kpis">
      <div class="div-kpi"><div class="div-kpi-val" style="color:${color}">${totalG}</div><div class="div-kpi-lbl">Goles</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalA}</div><div class="div-kpi-lbl">Amarillas</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalR}</div><div class="div-kpi-lbl">Rojas/Dobles</div></div>
    </div>
  </div>

  <div class="sub-tabs">
    <button class="stab active" onclick="switchSub(this,'${divNombre}','pos')">📊 Posiciones</button>
    <button class="stab" onclick="switchSub(this,'${divNombre}','res')">📅 Resultados</button>
    <button class="stab" onclick="switchSub(this,'${divNombre}','stats')">⚽ Estadísticas</button>
    <button class="stab" onclick="switchSub(this,'${divNombre}','sanc')">🚫 Sancionados</button>
  </div>

  <div id="sub-${divNombre}-pos" class="sub-panel active">
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🏆</span><span class="panel-head-title">Tabla de Posiciones · ${divNombre}</span></div>
      <div class="panel-body">${buildTablaPos(divNombre)}</div>
    </div>
  </div>

  <div id="sub-${divNombre}-res" class="sub-panel">
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">📅</span><span class="panel-head-title">Resultados · ${divNombre}</span></div>
      <div class="panel-body">${buildResultados(divNombre)}</div>
    </div>
  </div>

  <div id="sub-${divNombre}-stats" class="sub-panel">
    <div class="team-selector-wrap">
      <span class="team-selector-lbl">Ver equipo</span>
      <select class="team-select" id="sel-${divNombre}" onchange="renderTeam('${divNombre}',this.value)" style="border-color:${color}">
        <option value="__global__">📊 Estadística de la División</option>
      </select>
    </div>
    <div id="divglobal-${divNombre}">
      <div class="grid-3">
        <div class="panel">
          <div class="panel-head"><span style="font-size:16px">⚽</span><span class="panel-head-title">Top Goleadores</span></div>
          <div class="panel-body"><div id="dg-${divNombre}"></div></div>
        </div>
        <div class="panel">
          <div class="panel-head"><span style="font-size:16px">🟨</span><span class="panel-head-title">Más Amonestados</span></div>
          <div class="panel-body"><div id="dt-${divNombre}"></div></div>
        </div>
        <div class="panel">
          <div class="panel-head"><span style="font-size:16px">🤝</span><span class="panel-head-title">Fair Play</span></div>
          <div class="panel-body" id="df-${divNombre}"></div>
        </div>
      </div>
    </div>
    <div id="divteam-${divNombre}" style="display:none">
      <div class="team-kpi-row" id="tkpi-${divNombre}"></div>
      <div class="grid-2" id="tgrid-${divNombre}"></div>
    </div>
  </div>

  <div id="sub-${divNombre}-sanc" class="sub-panel">
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🚫</span><span class="panel-head-title">Sancionados · ${divNombre}</span></div>
      <div class="panel-body">${buildSancionados(divNombre)}</div>
    </div>
  </div>`;

  // Selector equipos
  const sel=document.getElementById('sel-'+divNombre);
  Object.keys(divData.teamsData).sort().forEach(t=>{
    const o=document.createElement('option');o.value=t;o.textContent='⚽ '+t;sel.appendChild(o);
  });

  // Listas stats globales de la división
  const maxG2=gd.top_10_goleadores[0]?.goles||1;
  const dgEl=document.getElementById('dg-'+divNombre);
  gd.top_10_goleadores.forEach((pl,i)=>{
    const pct=Math.round(pl.goles/maxG2*100);
    dgEl.innerHTML+=`<div class="scorer-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${pl.nombre}</div>
        <div class="rank-sub" style="color:${color}">${pl.equipo}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${pl.goles}<small> ⚽</small></div></div>`;
  });
  const dtEl=document.getElementById('dt-'+divNombre);
  gd.top_10_amonestados.forEach((c,i)=>{
    dtEl.innerHTML+=`<div class="rank-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${c.nombre}</div>
        <div class="rank-sub" style="color:${color}">${c.equipo}</div>
      </div>
      <div style="display:flex;gap:3px;flex-wrap:wrap">${badge(c)}</div></div>`;
  });
  const maxFP2=Math.max(...gd.fair_play.map(f=>f.total_tarjetas))||1;
  const dfEl=document.getElementById('df-'+divNombre);
  gd.fair_play.forEach((eq,i)=>{
    const pct=Math.round(eq.total_tarjetas/maxFP2*100);
    const bc=i<3?'#1a7a3c':i<7?'#b8860b':'#c8102e';
    dfEl.innerHTML+=`<div class="fp-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${eq.equipo}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div></div>`;
  });
}

function renderTeam(divNombre,equipo){
  const divGl=document.getElementById('divglobal-'+divNombre);
  const divTm=document.getElementById('divteam-'+divNombre);
  if(equipo==='__global__'){divTm.style.display='none';divGl.style.display='block';return;}
  divGl.style.display='none';divTm.style.display='block';
  const d=DATA.divisiones[divNombre].teamsData[equipo];
  const color=DIV_COLORS[divNombre]||'#c8102e';
  document.getElementById('tkpi-'+divNombre).innerHTML=`
    <div class="team-kpi"><div class="team-kpi-val accent" style="color:${color}">${d.total_goles}</div><div class="team-kpi-lbl">Goles Totales</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_goles}</div><div class="team-kpi-lbl">Promedio Goles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_amarillas}</div><div class="team-kpi-lbl">Amarillas</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_rojas+d.total_doble}</div><div class="team-kpi-lbl">Rojas/Dobles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_tarjetas}</div><div class="team-kpi-lbl">Prom. Tarjetas</div></div>`;
  const maxG3=d.top_scorers[0]?.goles||1;
  let golesH=d.top_scorers.length?'':'<div class="empty">Sin goles registrados</div>';
  d.top_scorers.forEach((g,i)=>{
    const pct=Math.round(g.goles/maxG3*100);
    golesH+=`<div class="scorer-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${g.nombre}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${g.goles}<small> ⚽</small></div></div>`;
  });
  let tarjH=d.top_cards.length?'':'<div class="empty">Sin tarjetas</div>';
  d.top_cards.forEach((c,i)=>{
    tarjH+=`<div class="rank-item">${rankNum(i)}
      <div class="rank-info"><div class="rank-name">${c.nombre}</div></div>
      <div style="display:flex;gap:3px;flex-wrap:wrap">${badge(c,true)}</div></div>`;
  });
  document.getElementById('tgrid-'+divNombre).innerHTML=`
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">⚽</span><span class="panel-head-title">${equipo} · Goleadores</span></div>
      <div class="panel-body">${golesH}</div>
    </div>
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🟨</span><span class="panel-head-title">${equipo} · Disciplina</span></div>
      <div class="panel-body">${tarjH}</div>
    </div>`;
}

// INICIALIZAR
buildGlobal();
['1era División','2da División','3era División C1','3era División C2'].forEach(div=>buildDivPanel(div));
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
    print(f"   Sancionados incluidos: {sum(len(v) for v in datos['sancionados'].values())}")
    print(f"   Prom. goles/partido: {datos['globalTorneo']['prom_goles_partido']}")
    print("👉 Abre index.html con doble clic.")
