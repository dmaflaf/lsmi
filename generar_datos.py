#!/usr/bin/env python3
"""
LSMI - Generador de página web completa (4 divisiones)
Incluye: Estadísticas, Tabla de Posiciones, Resultados
Ejecutar: python generar_datos.py
"""
import openpyxl, json, sys, os
from datetime import datetime

# ─── NOMBRES DE HOJAS EN EL EXCEL DE POSICIONES ───────────────────────────────
HOJAS_POS = {
    '1era División':    '1era División',
    '2da División':     '2da División',
    '3era División C1': '3era División 1',
    '3era División C2': '3era División 2',
}
DIV_RESULTADOS = {
    '1era División':    '1era División',
    '2da División':     '2da División',
    '3era División C1': '3era División 1',
    '3era División C2': '3era División 2',
}

# ─── ESTADÍSTICAS (goles / tarjetas) ──────────────────────────────────────────
def procesar_division_stats(ruta):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws_g = wb['GOLEADORES']
    goles_pj = {}; ultima_fecha = 0
    for row in ws_g.iter_rows(min_row=3, values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        e=str(row[1]).strip() if row[1] else ''; j=str(row[2]).strip() if row[2] else ''; g=row[3]
        if e and j and g and isinstance(g,(int,float)):
            key=(e,j); goles_pj[key]=goles_pj.get(key,0)+int(g); ultima_fecha=max(ultima_fecha,int(row[0]))
    ws_a = wb['AMONESTACIONES']
    tarj_pj = {}
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
    for equipo in equipos:
        ge={j:g for (e,j),g in goles_pj.items() if e==equipo}
        te={j:t for (e,j),t in tarj_pj.items() if e==equipo}
        tg=sum(ge.values()); tam=sum(t['amarillas'] for t in te.values())
        tdo=sum(t['dobles'] for t in te.values()); tro=sum(t['rojas'] for t in te.values())
        pg=round(tg/ultima_fecha,2) if ultima_fecha else 0
        pt=round((tam+tdo+tro)/ultima_fecha,2) if ultima_fecha else 0
        ts=sorted(ge.items(),key=lambda x:-x[1])[:8]
        tcr=[]
        for j,t in te.items():
            p=t['amarillas']+t['dobles']*2+t['rojas']*3
            if p>0: tcr.append({'nombre':j,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
        tcr.sort(key=lambda x:-x['_p'])
        teams_data[equipo]={'total_goles':tg,'promedio_goles':pg,'total_amarillas':tam,'total_doble':tdo,
            'total_rojas':tro,'promedio_tarjetas':pt,'top_scorers':[{'nombre':j,'goles':g} for j,g in ts],
            'top_cards':[{k:v for k,v in c.items() if k!='_p'} for c in tcr[:8]]}
    all_s=[{'nombre':j,'equipo':e,'goles':g} for (e,j),g in goles_pj.items()]; all_s.sort(key=lambda x:-x['goles'])
    all_c=[]
    for (e,j),t in tarj_pj.items():
        p=t['amarillas']+t['dobles']*2+t['rojas']*3
        if p>0: all_c.append({'nombre':j,'equipo':e,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
    all_c.sort(key=lambda x:-x['_p'])
    fp=[]; 
    for eq in equipos:
        te={j:t for (e,j),t in tarj_pj.items() if e==eq}
        fp.append({'equipo':eq,'total_tarjetas':sum(t['amarillas']+t['dobles']+t['rojas'] for t in te.values())})
    fp.sort(key=lambda x:x['total_tarjetas'])
    return {'meta':{'ultima_fecha':ultima_fecha,'generado':datetime.now().strftime('%d/%m/%Y %H:%M')},
        'globalData':{'top_10_goleadores':all_s[:10],'top_10_amonestados':[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:10]],'fair_play':fp},
        'teamsData':teams_data,
        '_goles_raw':{f"{e}||{j}":g for (e,j),g in goles_pj.items()},
        '_tarjetas_raw':{f"{e}||{j}":t for (e,j),t in tarj_pj.items()}}

# ─── TABLA DE POSICIONES ──────────────────────────────────────────────────────
def leer_posiciones(wb_pos):
    posiciones = {}
    for div_nombre, hoja in HOJAS_POS.items():
        if hoja not in wb_pos.sheetnames: continue
        ws = wb_pos[hoja]; tabla = []; header_ok = False
        for row in ws.iter_rows(values_only=True):
            if row[0]=='POS' and row[1]=='EQUIPO': header_ok=True; continue
            if not header_ok: continue
            if row[0] and isinstance(row[0],(int,float)) and row[1] and isinstance(row[1],str) and len(str(row[1]).strip())>1:
                if 'PJ=' in str(row[1]) or 'Puntuación' in str(row[1]): break
                tabla.append({'pos':int(row[0]),'equipo':str(row[1]).strip(),
                    'pj':int(row[2] or 0),'pg':int(row[3] or 0),'pe':int(row[4] or 0),'pp':int(row[5] or 0),
                    'gf':int(row[6] or 0),'gc':int(row[7] or 0),'dg':int(row[8] or 0),'pts':int(row[9] or 0)})
        posiciones[div_nombre] = tabla
    return posiciones

# ─── RESULTADOS ───────────────────────────────────────────────────────────────
def leer_resultados(wb_pos):
    ws_r = wb_pos['RESULTADOS']
    resultados = {}
    for row in ws_r.iter_rows(min_row=2, values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        if not row[2] or not row[5]: continue
        fecha=int(row[0]); div_raw=str(row[1]).strip() if row[1] else ''
        local=str(row[2]).strip(); gf=int(row[3] or 0); gv=int(row[4] or 0); vis=str(row[5]).strip()
        if not div_raw or not local or not vis: continue
        # Mapear nombre de división al interno
        div_nombre = None
        for k,v in DIV_RESULTADOS.items():
            if v == div_raw: div_nombre=k; break
        if not div_nombre: div_nombre = div_raw
        if div_nombre not in resultados: resultados[div_nombre]={}
        if fecha not in resultados[div_nombre]: resultados[div_nombre][fecha]=[]
        resultados[div_nombre][fecha].append({'local':local,'gl':gf,'gv':gv,'visitante':vis})
    # Convertir a lista ordenada
    out = {}
    for div, fechas in resultados.items():
        out[div] = [{'fecha':f,'partidos':ps} for f,ps in sorted(fechas.items(),reverse=True)]
    return out

# ─── BUSCAR ARCHIVOS ─────────────────────────────────────────────────────────
def buscar_excels(carpeta):
    archivos=[f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    stats={}
    claves_stats={'1era División':['1era','primera'],'2da División':['2da','segunda'],
        '3era División C1':['division_1','3era_division_1','_1_v','grupo_1','c1','g1'],
        '3era División C2':['division_2','3era_division_2','_2_v','grupo_2','c2','g2']}
    for archivo in archivos:
        low=archivo.lower()
        for div,pals in claves_stats.items():
            if any(p in low for p in pals) and div not in stats:
                stats[div]=os.path.join(carpeta,archivo); break
    # Archivo de posiciones/torneo (contiene 'torneo' o tiene múltiples divisiones)
    posiciones_ruta=None
    for archivo in archivos:
        low=archivo.lower()
        if 'torneo' in low or 'posicion' in low or 'tabla' in low or 'fase' in low:
            posiciones_ruta=os.path.join(carpeta,archivo); break
    return stats, posiciones_ruta

# ─── PROCESAMIENTO PRINCIPAL ─────────────────────────────────────────────────
def procesar_todo(carpeta):
    stats_rutas, pos_ruta = buscar_excels(carpeta)
    print("📁 Excel de estadísticas:")
    for div,r in stats_rutas.items(): print(f"   {div}: {os.path.basename(r)}")
    if pos_ruta: print(f"📁 Excel de posiciones/resultados: {os.path.basename(pos_ruta)}")
    else: print("⚠️  No se encontró Excel de posiciones (busca uno con 'torneo' o 'fase' en el nombre)")

    # Estadísticas
    resultados_stats={}
    for nombre,ruta in stats_rutas.items():
        print(f"\n⚽ Procesando estadísticas {nombre}...")
        resultados_stats[nombre]=procesar_division_stats(ruta)
        print(f"   Equipos: {len(resultados_stats[nombre]['teamsData'])} | Fecha: {resultados_stats[nombre]['meta']['ultima_fecha']}")

    # Posiciones y resultados
    posiciones={}; resultados_partidos={}
    if pos_ruta:
        print(f"\n📊 Leyendo tablas de posiciones y resultados...")
        wb_pos=openpyxl.load_workbook(pos_ruta, read_only=True, data_only=True)
        posiciones=leer_posiciones(wb_pos)
        resultados_partidos=leer_resultados(wb_pos)
        for div,tabla in posiciones.items(): print(f"   {div}: {len(tabla)} equipos en tabla")
        for div,fechas in resultados_partidos.items(): print(f"   {div}: {sum(len(f['partidos']) for f in fechas)} partidos")

    # Global torneo
    all_g,all_c,fp_torneo,stats_div=[],[],[],[]
    for div_nombre,data in resultados_stats.items():
        for key,goles in data['_goles_raw'].items():
            e,j=key.split('||'); all_g.append({'nombre':j,'equipo':e,'division':div_nombre,'goles':goles})
        for key,t in data['_tarjetas_raw'].items():
            e,j=key.split('||'); p=t['amarillas']+t['dobles']*2+t['rojas']*3
            if p>0: all_c.append({'nombre':j,'equipo':e,'division':div_nombre,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
        for eq_fp in data['globalData']['fair_play']:
            fp_torneo.append({'equipo':eq_fp['equipo'],'division':div_nombre,'total_tarjetas':eq_fp['total_tarjetas']})
        tg=sum(d['total_goles'] for d in data['teamsData'].values())
        tt=sum(d['total_amarillas']+d['total_doble']+d['total_rojas'] for d in data['teamsData'].values())
        ne=len(data['teamsData'])
        stats_div.append({'division':div_nombre,'total_goles':tg,'total_tarjetas':tt,'num_equipos':ne,
            'promedio_goles_equipo':round(tg/ne,1) if ne else 0,'promedio_tarjetas_equipo':round(tt/ne,1) if ne else 0})
    all_g.sort(key=lambda x:-x['goles']); all_c.sort(key=lambda x:-x['_p']); fp_torneo.sort(key=lambda x:x['total_tarjetas']); stats_div.sort(key=lambda x:-x['total_goles'])
    top_c=[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:15]]
    global_torneo={'top_15_goleadores':all_g[:15],'top_15_amonestados':top_c,'fair_play':fp_torneo[:20],'ranking_divisiones':stats_div}
    for d in resultados_stats.values(): d.pop('_goles_raw',None); d.pop('_tarjetas_raw',None)
    return {'generado':datetime.now().strftime('%d/%m/%Y %H:%M'),'divisiones':resultados_stats,
        'posiciones':posiciones,'resultados':resultados_partidos,'globalTorneo':global_torneo}

# ─── HTML ─────────────────────────────────────────────────────────────────────
def generar_html(datos):
    json_str = json.dumps(datos, ensure_ascii=False)
    return HTML_TEMPLATE.replace('__JSON_DATA__', json_str)

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>LSMI · Liga San Miguel de Ibarra</title>
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d0f14;--surface:#161a23;--surface2:#1e2330;--surface3:#252b3b;
  --border:#2a3045;--border2:#323a50;
  --red:#e8192c;--gold:#f5a623;--green:#00c853;--blue:#2979ff;
  --text:#f0f2f8;--text2:#8b95b0;--text3:#5a6480;
  --1era:#e8192c;--2da:#2979ff;--c1:#00c853;--c2:#f5a623;
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}
.site-header{background:linear-gradient(135deg,#0d0f14 0%,#1a0508 50%,#0d0f14 100%);border-bottom:2px solid var(--red);position:sticky;top:0;z-index:100;box-shadow:0 4px 30px rgba(232,25,44,0.2)}
.header-inner{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;padding:12px 24px;gap:16px}
.logo-block{display:flex;align-items:center;gap:14px}
.logo-emblem{width:52px;height:52px;background:var(--red);border-radius:8px;display:flex;align-items:center;justify-content:center;font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:900;color:white;box-shadow:0 0 20px rgba(232,25,44,0.4)}
.logo-title{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:900;letter-spacing:2px;text-transform:uppercase;line-height:1}
.logo-sub{font-size:10px;font-weight:600;color:var(--text2);letter-spacing:3px;text-transform:uppercase;margin-top:2px}
.header-meta{font-size:11px;color:var(--text3);font-weight:500;text-align:right;line-height:1.6}
.header-meta span{color:var(--red);font-weight:700}
/* TABS */
.div-tabs{background:var(--surface);border-bottom:1px solid var(--border);position:sticky;top:78px;z-index:90}
.div-tabs-inner{max-width:1400px;margin:0 auto;display:flex;overflow-x:auto;scrollbar-width:none;padding:0 16px}
.div-tabs-inner::-webkit-scrollbar{display:none}
.dtab{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;padding:14px 20px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap}
.dtab:hover{color:var(--text);background:var(--surface2)}
.dtab.active{color:var(--text);background:var(--surface2)}
/* SUB-TABS */
.sub-tabs{display:flex;gap:4px;margin-bottom:20px;background:var(--surface);border-radius:8px;padding:4px}
.stab{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:8px 16px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-radius:6px;transition:all .2s}
.stab:hover{color:var(--text)}
.stab.active{background:var(--surface3);color:var(--text)}
.sub-panel{display:none}
.sub-panel.active{display:block}
/* LAYOUT */
.main{max-width:1400px;margin:0 auto;padding:24px 16px}
.tab-panel{display:none}
.tab-panel.active{display:block}
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin:20px 0}
.grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:20px 0}
@media(max-width:900px){.grid-3{grid-template-columns:1fr}.grid-2{grid-template-columns:1fr}}
/* PANELS */
.panel{background:var(--surface);border-radius:10px;border:1px solid var(--border);overflow:hidden}
.panel-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.panel-head-title{font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:var(--text)}
.panel-body{padding:12px}
/* HERO STATS */
.hero-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1px;background:var(--border);border-radius:10px;overflow:hidden;margin:20px 0}
.hero-stat{background:var(--surface);padding:20px 24px;text-align:center}
.hero-stat-val{font-family:'Barlow Condensed',sans-serif;font-size:42px;font-weight:900;line-height:1;color:var(--text)}
.hero-stat-val.red{color:var(--red)}
.hero-stat-lbl{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-top:4px}
/* DIV CARDS */
.div-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin:20px 0}
.div-card{background:var(--surface);border-radius:10px;overflow:hidden;border:1px solid var(--border)}
.div-card-header{padding:14px 18px;display:flex;align-items:center;justify-content:space-between}
.div-card-title{font-family:'Barlow Condensed',sans-serif;font-size:16px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:white}
.div-card-rank{font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:900;color:rgba(255,255,255,0.2)}
.div-card-body{padding:14px 18px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;background:var(--surface2)}
.div-stat-mini{text-align:center}
.div-stat-mini-val{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--text)}
.div-stat-mini-lbl{font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--text3);margin-top:2px}
/* RANKING */
.rank-list{display:flex;flex-direction:column;gap:4px}
.rank-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;background:var(--surface2);border:1px solid transparent;transition:border-color .15s}
.rank-item:hover{border-color:var(--border2)}
.rank-num{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;color:var(--text3);min-width:18px;text-align:center}
.rank-num.gold{color:#f5a623}.rank-num.silver{color:#9aa0b0}.rank-num.bronze{color:#cd7f32}
.rank-info{flex:1;overflow:hidden}
.rank-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-sub{font-size:10px;font-weight:600;color:var(--text3);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-val{font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:800;color:var(--text);min-width:32px;text-align:right}
.rank-val small{font-size:10px;font-weight:600;color:var(--text3)}
/* BADGES */
.badge{display:inline-flex;align-items:center;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;padding:2px 6px;border-radius:4px;margin-left:3px}
.badge-yellow{background:#7c6300;color:#ffd600}
.badge-red{background:#5c0010;color:#ff5252}
.badge-orange{background:#5c2c00;color:#ff9100}
/* SCORER BAR */
.scorer-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;background:var(--surface2);margin-bottom:4px;border:1px solid transparent;transition:border-color .15s}
.scorer-item:hover{border-color:var(--border2)}
.scorer-bar-wrap{flex:1;height:3px;background:var(--surface3);border-radius:2px;overflow:hidden}
.scorer-bar{height:100%;border-radius:2px}
/* FAIR PLAY */
.fp-item{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:6px;background:var(--surface2);margin-bottom:4px}
.fp-bar-wrap{flex:1;height:4px;background:var(--surface3);border-radius:2px;overflow:hidden}
.fp-bar{height:100%;border-radius:2px}
.fp-val{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;color:var(--text);min-width:28px;text-align:right}
/* TABLA DE POSICIONES */
.pos-table{width:100%;border-collapse:collapse;font-size:12px}
.pos-table th{font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:8px 10px;text-align:center;border-bottom:2px solid var(--border2)}
.pos-table th:nth-child(2){text-align:left}
.pos-table td{padding:9px 10px;text-align:center;border-bottom:1px solid var(--border);color:var(--text2);font-weight:500}
.pos-table td:nth-child(2){text-align:left;font-weight:700;color:var(--text)}
.pos-table tr:hover td{background:var(--surface2)}
.pos-table .pos-1{background:linear-gradient(90deg,rgba(245,166,35,0.08) 0%,transparent 100%)}
.pos-table .pos-1 td:nth-child(2){color:var(--gold)}
.pos-table .pts-col{font-family:'Barlow Condensed',sans-serif;font-size:16px;font-weight:900;color:var(--text)}
.pos-table .dg-pos{color:var(--green)}
.pos-table .dg-neg{color:var(--red)}
.zone-top{border-left:3px solid var(--gold)}
.zone-mid{border-left:3px solid var(--green)}
.zone-bot{border-left:3px solid var(--red)}
/* RESULTADOS */
.fecha-block{margin-bottom:16px}
.fecha-label{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);padding:8px 12px;background:var(--surface3);border-radius:6px;margin-bottom:8px;display:inline-block}
.match-card{display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:var(--surface2);border-radius:8px;border:1px solid var(--border);margin-bottom:6px;gap:8px}
.match-team{font-size:13px;font-weight:600;color:var(--text);flex:1}
.match-team.right{text-align:right}
.match-team.winner{color:var(--text);font-weight:700}
.match-score{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:900;color:var(--text);background:var(--surface3);padding:4px 16px;border-radius:6px;min-width:70px;text-align:center;letter-spacing:2px}
.match-score.draw{color:var(--text2)}
/* DIV HERO */
.div-hero{background:var(--surface);border-radius:10px;border:1px solid var(--border);padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.div-hero-title{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:900;letter-spacing:1px;text-transform:uppercase}
.div-hero-meta{font-size:12px;color:var(--text2);font-weight:500;margin-top:4px}
.div-kpis{display:flex;gap:24px;flex-wrap:wrap}
.div-kpi{text-align:center}
.div-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:900;color:var(--text)}
.div-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3)}
/* TEAM SELECTOR */
.team-selector-wrap{background:var(--surface);border-radius:10px;border:1px solid var(--border);padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.team-selector-lbl{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:var(--text3);white-space:nowrap}
.team-select{flex:1;min-width:180px;background:var(--surface2);color:var(--text);border:1px solid var(--border2);border-radius:6px;padding:8px 12px;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;cursor:pointer;outline:none}
/* TEAM KPIS */
.team-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:1px;background:var(--border);border-radius:8px;overflow:hidden;margin-bottom:16px}
.team-kpi{background:var(--surface2);padding:16px;text-align:center}
.team-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:900;color:var(--text);line-height:1}
.team-kpi-val.accent{color:var(--accent,var(--red))}
.team-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-top:4px}
/* EMPTY */
.empty{color:var(--text3);font-size:12px;padding:20px;text-align:center;font-style:italic}
@media(max-width:600px){
  .header-inner{padding:10px 12px}
  .logo-title{font-size:16px}
  .main{padding:16px 10px}
  .match-card{padding:8px 10px}
  .pos-table{font-size:11px}
  .pos-table th,.pos-table td{padding:6px 6px}
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
    <div class="header-meta" id="headerMeta"></div>
  </div>
</header>
<nav class="div-tabs">
  <div class="div-tabs-inner">
    <button class="dtab active" data-tab="global" data-color="#e8192c" onclick="switchTab(this,'global')">🏆 Global</button>
    <button class="dtab" data-tab="1era División" data-color="#e8192c" onclick="switchTab(this,'1era División')">1era División</button>
    <button class="dtab" data-tab="2da División" data-color="#2979ff" onclick="switchTab(this,'2da División')">2da División</button>
    <button class="dtab" data-tab="3era División C1" data-color="#00c853" onclick="switchTab(this,'3era División C1')">3era Div. C1</button>
    <button class="dtab" data-tab="3era División C2" data-color="#f5a623" onclick="switchTab(this,'3era División C2')">3era Div. C2</button>
  </div>
</nav>
<main class="main">
  <div id="panel-global" class="tab-panel active"></div>
  <div id="panel-1era División" class="tab-panel"></div>
  <div id="panel-2da División" class="tab-panel"></div>
  <div id="panel-3era División C1" class="tab-panel"></div>
  <div id="panel-3era División C2" class="tab-panel"></div>
</main>
<footer style="text-align:center;padding:24px 16px;color:var(--text3);font-size:11px;font-weight:600;letter-spacing:1px;border-top:1px solid var(--border);margin-top:20px">
  LIGA SAN MIGUEL DE IBARRA · ESTADÍSTICAS OFICIALES · <span id="footerDate"></span>
</footer>
<script>
const DATA = __JSON_DATA__;
const DIV_COLORS={'1era División':'#e8192c','2da División':'#2979ff','3era División C1':'#00c853','3era División C2':'#f5a623'};
const DIV_BG={'1era División':'#3d0a10','2da División':'#0a1a40','3era División C1':'#003320','3era División C2':'#3d2800'};
document.getElementById('headerMeta').innerHTML=`Actualizado: <span>${DATA.generado}</span>`;
document.getElementById('footerDate').textContent=DATA.generado;

function switchTab(btn,id){
  document.querySelectorAll('.dtab').forEach(b=>{b.classList.remove('active');b.style.borderBottomColor='transparent'});
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  btn.style.borderBottomColor=btn.dataset.color||'#e8192c';
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
function badge(c,big=false){
  const sz=big?'font-size:13px;padding:3px 8px':'';
  let b='';
  if(c.rojas>0) b+=`<span class="badge badge-red" style="${sz}">${c.rojas}🔴</span>`;
  if(c.dobles>0) b+=`<span class="badge badge-orange" style="${sz}">${c.dobles}🟠</span>`;
  if(c.amarillas>0) b+=`<span class="badge badge-yellow" style="${sz}">${c.amarillas}🟡</span>`;
  return b;
}
function rankNum(i){
  if(i===0)return'<span class="rank-num gold">1</span>';
  if(i===1)return'<span class="rank-num silver">2</span>';
  if(i===2)return'<span class="rank-num bronze">3</span>';
  return`<span class="rank-num">${i+1}</span>`;
}

// ── PANEL GLOBAL ──────────────────────────────────────────────────────────────
function buildGlobal(){
  const g=DATA.globalTorneo;
  const totalG=g.ranking_divisiones.reduce((a,d)=>a+d.total_goles,0);
  const totalT=g.ranking_divisiones.reduce((a,d)=>a+d.total_tarjetas,0);
  const totalE=g.ranking_divisiones.reduce((a,d)=>a+d.num_equipos,0);
  const p=document.getElementById('panel-global');
  p.innerHTML=`
  <div style="margin-bottom:20px">
    <div style="font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--text3);margin-bottom:6px">Temporada 2025</div>
    <div style="font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:900;text-transform:uppercase">Estadística <span style="color:var(--red)">Global</span> del Torneo</div>
  </div>
  <div class="hero-stats">
    <div class="hero-stat"><div class="hero-stat-val red">${totalG}</div><div class="hero-stat-lbl">Goles en el Torneo</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalE}</div><div class="hero-stat-lbl">Equipos</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${g.ranking_divisiones.length}</div><div class="hero-stat-lbl">Divisiones</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalT}</div><div class="hero-stat-lbl">Tarjetas Totales</div></div>
    <div class="hero-stat"><div class="hero-stat-val red">${g.top_15_goleadores[0]?.goles||0}</div><div class="hero-stat-lbl">Goles Líder</div></div>
  </div>
  <div class="div-cards" id="rankDivCards"></div>
  <div class="grid-3">
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">⚽</span><span class="panel-head-title">Top 15 Goleadores</span></div>
      <div class="panel-body"><div class="rank-list" id="gGol"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🟨</span><span class="panel-head-title">Top 15 Disciplina</span></div>
      <div class="panel-body"><div class="rank-list" id="gTarj"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span style="font-size:16px">🤝</span><span class="panel-head-title">Fair Play General</span></div>
      <div class="panel-body" id="gFP"></div>
    </div>
  </div>`;

  // ranking divisiones
  const rd=document.getElementById('rankDivCards');
  g.ranking_divisiones.forEach((d,i)=>{
    const color=DIV_COLORS[d.division]||'#e8192c', bg=DIV_BG[d.division]||'#3d0a10';
    rd.innerHTML+=`<div class="div-card">
      <div class="div-card-header" style="background:${bg};border-bottom:2px solid ${color}">
        <div><div class="div-card-title" style="color:${color}">${d.division}</div>
        <div style="font-size:10px;color:rgba(255,255,255,0.4);font-weight:600;margin-top:2px">${d.num_equipos} EQUIPOS</div></div>
        <div class="div-card-rank">#${i+1}</div>
      </div>
      <div class="div-card-body">
        <div class="div-stat-mini"><div class="div-stat-mini-val" style="color:${color}">${d.total_goles}</div><div class="div-stat-mini-lbl">Goles</div></div>
        <div class="div-stat-mini"><div class="div-stat-mini-val">${d.promedio_goles_equipo}</div><div class="div-stat-mini-lbl">Prom/Eq</div></div>
        <div class="div-stat-mini"><div class="div-stat-mini-val">${d.total_tarjetas}</div><div class="div-stat-mini-lbl">Tarjetas</div></div>
      </div>
    </div>`;
  });

  const maxG2=g.top_15_goleadores[0]?.goles||1;
  const gg=document.getElementById('gGol');
  g.top_15_goleadores.forEach((pl,i)=>{
    const color=DIV_COLORS[pl.division]||'#e8192c', pct=Math.round(pl.goles/maxG2*100);
    gg.innerHTML+=`<div class="scorer-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${pl.nombre}</div>
        <div class="rank-sub" style="color:${color}">${pl.equipo} · ${pl.division}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${pl.goles}<small> ⚽</small></div></div>`;
  });

  const gt=document.getElementById('gTarj');
  g.top_15_amonestados.forEach((c,i)=>{
    const color=DIV_COLORS[c.division]||'#e8192c';
    gt.innerHTML+=`<div class="rank-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${c.nombre}</div>
        <div class="rank-sub" style="color:${color}">${c.equipo} · ${c.division}</div>
      </div>
      <div style="display:flex;gap:3px">${badge(c)}</div></div>`;
  });

  const maxFP=Math.max(...g.fair_play.map(f=>f.total_tarjetas))||1;
  const gfp=document.getElementById('gFP');
  g.fair_play.slice(0,16).forEach((eq,i)=>{
    const color=DIV_COLORS[eq.division]||'#e8192c', pct=Math.round(eq.total_tarjetas/maxFP*100);
    const bc=i<3?'#00c853':i<8?'#f5a623':'#e8192c';
    gfp.innerHTML+=`<div class="fp-item">${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${eq.equipo}</div>
        <div class="rank-sub" style="color:${color}">${eq.division}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div></div>`;
  });
}

// ── TABLA DE POSICIONES ───────────────────────────────────────────────────────
function buildTablaPos(divNombre){
  const tabla=DATA.posiciones[divNombre]||[];
  const color=DIV_COLORS[divNombre]||'#e8192c';
  if(!tabla.length) return`<div class="empty">Sin tabla de posiciones disponible</div>`;
  const n=tabla.length;
  let html=`<div style="overflow-x:auto"><table class="pos-table">
    <thead><tr>
      <th>#</th><th style="text-align:left">Equipo</th>
      <th title="Partidos Jugados">PJ</th><th title="Ganados">G</th>
      <th title="Empatados">E</th><th title="Perdidos">P</th>
      <th title="Goles a Favor">GF</th><th title="Goles en Contra">GC</th>
      <th title="Diferencia de Goles">DG</th>
      <th title="Puntos" style="color:${color}">PTS</th>
    </tr></thead><tbody>`;
  tabla.forEach((eq,i)=>{
    const zone = i===0?'zone-top':i<3?'zone-mid':i>=n-3?'zone-bot':'';
    const dgClass=eq.dg>0?'dg-pos':eq.dg<0?'dg-neg':'';
    const rowClass=i===0?'pos-1':'';
    html+=`<tr class="${rowClass} ${zone}">
      <td><strong style="color:${i===0?'var(--gold)':i<3?color:'var(--text2)'}">${eq.pos}</strong></td>
      <td>${eq.equipo}</td>
      <td>${eq.pj}</td><td>${eq.pg}</td><td>${eq.pe}</td><td>${eq.pp}</td>
      <td>${eq.gf}</td><td>${eq.gc}</td>
      <td class="${dgClass}">${eq.dg>0?'+':''}${eq.dg}</td>
      <td class="pts-col" style="color:${color}">${eq.pts}</td>
    </tr>`;
  });
  html+=`</tbody></table></div>
  <div style="font-size:10px;color:var(--text3);padding:10px 4px;display:flex;gap:16px;flex-wrap:wrap">
    <span><span style="color:var(--gold)">■</span> Líder</span>
    <span><span style="color:${color}">■</span> Zona ascenso</span>
    <span><span style="color:var(--red)">■</span> Zona de riesgo</span>
    <span>PJ=Partidos Jugados · G/E/P · GF=Goles Favor · GC=Goles Contra · DG=Diferencia · PTS=Puntos</span>
  </div>`;
  return html;
}

// ── RESULTADOS ────────────────────────────────────────────────────────────────
function buildResultados(divNombre){
  const fechas=DATA.resultados[divNombre]||[];
  if(!fechas.length) return`<div class="empty">Sin resultados registrados</div>`;
  let html='';
  fechas.forEach(f=>{
    html+=`<div class="fecha-block"><div class="fecha-label">⚽ Fecha ${f.fecha}</div>`;
    f.partidos.forEach(p=>{
      const draw=p.gl===p.gv;
      const localW=p.gl>p.gv, visW=p.gv>p.gl;
      html+=`<div class="match-card">
        <div class="match-team ${localW?'winner':''}" style="opacity:${visW?0.5:1}">${p.local}</div>
        <div class="match-score ${draw?'draw':''}">${p.gl} - ${p.gv}</div>
        <div class="match-team right ${visW?'winner':''}" style="opacity:${localW?0.5:1}">${p.visitante}</div>
      </div>`;
    });
    html+='</div>';
  });
  return html;
}

// ── PANEL DIVISIÓN ────────────────────────────────────────────────────────────
function buildDivPanel(divNombre){
  const divData=DATA.divisiones[divNombre];
  const color=DIV_COLORS[divNombre]||'#e8192c';
  const bg=DIV_BG[divNombre]||'#3d0a10';
  const meta=divData.meta;
  const gd=divData.globalData;
  const totalG=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_goles,0);
  const totalA=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_amarillas,0);
  const totalR=Object.values(divData.teamsData).reduce((a,t)=>a+t.total_rojas+t.total_doble,0);
  const ne=Object.keys(divData.teamsData).length;
  const p=document.getElementById('panel-'+divNombre);

  p.innerHTML=`
  <div class="div-hero" style="border-top:3px solid ${color}">
    <div>
      <div class="div-hero-title" style="color:${color}">${divNombre}</div>
      <div class="div-hero-meta">Fecha ${meta.ultima_fecha} · Actualizado ${meta.generado} · ${ne} equipos</div>
    </div>
    <div class="div-kpis">
      <div class="div-kpi"><div class="div-kpi-val" style="color:${color}">${totalG}</div><div class="div-kpi-lbl">Goles</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalA}</div><div class="div-kpi-lbl">Amarillas</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalR}</div><div class="div-kpi-lbl">Rojas/Dobles</div></div>
    </div>
  </div>

  <!-- SUB-TABS -->
  <div class="sub-tabs">
    <button class="stab active" onclick="switchSub(this,'${divNombre}','pos')">📊 Tabla de Posiciones</button>
    <button class="stab" onclick="switchSub(this,'${divNombre}','res')">📅 Resultados</button>
    <button class="stab" onclick="switchSub(this,'${divNombre}','stats')">⚽ Estadísticas</button>
  </div>

  <!-- POSICIONES -->
  <div id="sub-${divNombre}-pos" class="sub-panel active">
    <div class="panel">
      <div class="panel-head" style="border-bottom-color:${color}40">
        <span style="font-size:16px">🏆</span>
        <span class="panel-head-title">Tabla de Posiciones · ${divNombre}</span>
      </div>
      <div class="panel-body">${buildTablaPos(divNombre)}</div>
    </div>
  </div>

  <!-- RESULTADOS -->
  <div id="sub-${divNombre}-res" class="sub-panel">
    <div class="panel">
      <div class="panel-head">
        <span style="font-size:16px">📅</span>
        <span class="panel-head-title">Resultados · ${divNombre}</span>
      </div>
      <div class="panel-body">${buildResultados(divNombre)}</div>
    </div>
  </div>

  <!-- ESTADÍSTICAS -->
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
          <div class="panel-body"><div class="rank-list" id="dg-${divNombre}"></div></div>
        </div>
        <div class="panel">
          <div class="panel-head"><span style="font-size:16px">🟨</span><span class="panel-head-title">Más Amonestados</span></div>
          <div class="panel-body"><div class="rank-list" id="dt-${divNombre}"></div></div>
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
  </div>`;

  // Selector equipos
  const sel=document.getElementById('sel-'+divNombre);
  Object.keys(divData.teamsData).sort().forEach(t=>{
    const o=document.createElement('option'); o.value=t; o.textContent='⚽ '+t; sel.appendChild(o);
  });

  // Listas globales estadísticas
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
      <div style="display:flex;gap:3px">${badge(c)}</div></div>`;
  });
  const maxFP2=Math.max(...gd.fair_play.map(f=>f.total_tarjetas))||1;
  const dfEl=document.getElementById('df-'+divNombre);
  gd.fair_play.forEach((eq,i)=>{
    const pct=Math.round(eq.total_tarjetas/maxFP2*100);
    const bc=i<3?'#00c853':i<7?'#f5a623':'#e8192c';
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
  divGl.style.display='none'; divTm.style.display='block';
  const d=DATA.divisiones[divNombre].teamsData[equipo];
  const color=DIV_COLORS[divNombre]||'#e8192c';
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
      <div style="display:flex;gap:3px">${badge(c,true)}</div></div>`;
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
['1era División','2da División','3era División C1','3era División C2'].forEach(div=>buildDivPanel(div,DATA.divisiones[div]));
// Activar borde del tab inicial
document.querySelector('.dtab.active').style.borderBottomColor='#e8192c';
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
    print("👉 Abre index.html con doble clic.")
