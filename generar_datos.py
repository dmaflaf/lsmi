#!/usr/bin/env python3
"""
LSMI - Generador de página web completa (4 divisiones)
Ejecutar: python generar_datos.py
"""
import openpyxl, json, sys, os
from datetime import datetime

def procesar_division(ruta):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws_g = wb['GOLEADORES']
    goles_por_jugador = {}
    ultima_fecha = 0
    for row in ws_g.iter_rows(min_row=3, values_only=True):
        if not row[0] or not isinstance(row[0], (int, float)): continue
        fecha=int(row[0]); equipo=str(row[1]).strip() if row[1] else ''; jugador=str(row[2]).strip() if row[2] else ''; goles=row[3]
        if equipo and jugador and goles and isinstance(goles,(int,float)):
            key=(equipo,jugador); goles_por_jugador[key]=goles_por_jugador.get(key,0)+int(goles); ultima_fecha=max(ultima_fecha,fecha)
    ws_a = wb['AMONESTACIONES']
    tarjetas_por_jugador = {}
    for row in ws_a.iter_rows(min_row=3, values_only=True):
        if not row[0] or not isinstance(row[0],(int,float)): continue
        equipo=str(row[1]).strip() if row[1] else ''; jugador=str(row[2]).strip() if row[2] else ''; tarjeta=str(row[3]).strip() if row[3] else ''
        if equipo and jugador and tarjeta:
            key=(equipo,jugador)
            if key not in tarjetas_por_jugador: tarjetas_por_jugador[key]={'amarillas':0,'dobles':0,'rojas':0}
            if 'Doble' in tarjeta or 'doble' in tarjeta: tarjetas_por_jugador[key]['dobles']+=1
            elif 'Roja' in tarjeta or 'roja' in tarjeta or '🔴' in tarjeta: tarjetas_por_jugador[key]['rojas']+=1
            else: tarjetas_por_jugador[key]['amarillas']+=1
    equipos=sorted(set(k[0] for k in goles_por_jugador)|set(k[0] for k in tarjetas_por_jugador))
    teams_data={}
    for equipo in equipos:
        goles_eq={j:g for (e,j),g in goles_por_jugador.items() if e==equipo}
        tarj_eq={j:t for (e,j),t in tarjetas_por_jugador.items() if e==equipo}
        total_goles=sum(goles_eq.values()); total_amarillas=sum(t['amarillas'] for t in tarj_eq.values())
        total_dobles=sum(t['dobles'] for t in tarj_eq.values()); total_rojas=sum(t['rojas'] for t in tarj_eq.values())
        promedio_goles=round(total_goles/ultima_fecha,2) if ultima_fecha else 0
        promedio_tarjetas=round((total_amarillas+total_dobles+total_rojas)/ultima_fecha,2) if ultima_fecha else 0
        top_scorers=sorted(goles_eq.items(),key=lambda x:-x[1])[:8]
        top_cards_raw=[]
        for j,t in tarj_eq.items():
            p=t['amarillas']+t['dobles']*2+t['rojas']*3
            if p>0: top_cards_raw.append({'nombre':j,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
        top_cards_raw.sort(key=lambda x:-x['_p'])
        top_cards=[{k:v for k,v in c.items() if k!='_p'} for c in top_cards_raw[:8]]
        teams_data[equipo]={'total_goles':total_goles,'promedio_goles':promedio_goles,'total_amarillas':total_amarillas,
            'total_doble':total_dobles,'total_rojas':total_rojas,'promedio_tarjetas':promedio_tarjetas,
            'top_scorers':[{'nombre':j,'goles':g} for j,g in top_scorers],'top_cards':top_cards}
    all_scorers=[{'nombre':j,'equipo':e,'goles':g} for (e,j),g in goles_por_jugador.items()]
    all_scorers.sort(key=lambda x:-x['goles'])
    all_cards=[]
    for (e,j),t in tarjetas_por_jugador.items():
        p=t['amarillas']+t['dobles']*2+t['rojas']*3
        if p>0: all_cards.append({'nombre':j,'equipo':e,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
    all_cards.sort(key=lambda x:-x['_p'])
    top10_cards=[{k:v for k,v in c.items() if k!='_p'} for c in all_cards[:10]]
    fair_play=[]
    for equipo in equipos:
        tarj_e={j:t for (e,j),t in tarjetas_por_jugador.items() if e==equipo}
        total=sum(t['amarillas']+t['dobles']+t['rojas'] for t in tarj_e.values())
        fair_play.append({'equipo':equipo,'total_tarjetas':total})
    fair_play.sort(key=lambda x:x['total_tarjetas'])
    return {'meta':{'ultima_fecha':ultima_fecha,'generado':datetime.now().strftime('%d/%m/%Y %H:%M')},
        'globalData':{'top_10_goleadores':all_scorers[:10],'top_10_amonestados':top10_cards,'fair_play':fair_play},
        'teamsData':teams_data,
        '_goles_raw':{f"{e}||{j}":g for (e,j),g in goles_por_jugador.items()},
        '_tarjetas_raw':{f"{e}||{j}":t for (e,j),t in tarjetas_por_jugador.items()}}

def buscar_excel(carpeta):
    archivos=[f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    mapa={}
    claves={'1era División':['1era','primera'],'2da División':['2da','segunda'],'3era División C1':['division_1','3era_division_1','grupo_1','c1','g1','_1_v'],'3era División C2':['division_2','3era_division_2','grupo_2','c2','g2','_2_v']}
    for archivo in archivos:
        low=archivo.lower()
        for div,palabras in claves.items():
            if any(p in low for p in palabras) and div not in mapa:
                mapa[div]=os.path.join(carpeta,archivo); break
    return mapa

def procesar_todo(carpeta):
    mapa=buscar_excel(carpeta)
    print("📁 Archivos encontrados:")
    for div,ruta in mapa.items(): print(f"   {div}: {os.path.basename(ruta)}")
    if len(mapa)<4:
        faltantes=[d for d in ['1era División','2da División','3era División C1','3era División C2'] if d not in mapa]
        print(f"⚠️  Faltan: {faltantes}"); sys.exit(1)
    resultados={}
    for nombre,ruta in mapa.items():
        print(f"\n⚽ Procesando {nombre}...")
        resultados[nombre]=procesar_division(ruta)
        print(f"   Equipos: {len(resultados[nombre]['teamsData'])} | Fecha: {resultados[nombre]['meta']['ultima_fecha']}")
    all_g,all_c,fp_torneo,stats_div=[],[],[],[]
    for div_nombre,data in resultados.items():
        for key,goles in data['_goles_raw'].items():
            e,j=key.split('||'); all_g.append({'nombre':j,'equipo':e,'division':div_nombre,'goles':goles})
        for key,t in data['_tarjetas_raw'].items():
            e,j=key.split('||'); p=t['amarillas']+t['dobles']*2+t['rojas']*3
            if p>0: all_c.append({'nombre':j,'equipo':e,'division':div_nombre,'amarillas':t['amarillas'],'dobles':t['dobles'],'rojas':t['rojas'],'_p':p})
        for eq_fp in data['globalData']['fair_play']:
            fp_torneo.append({'equipo':eq_fp['equipo'],'division':div_nombre,'total_tarjetas':eq_fp['total_tarjetas']})
        total_g=sum(d['total_goles'] for d in data['teamsData'].values())
        total_t=sum(d['total_amarillas']+d['total_doble']+d['total_rojas'] for d in data['teamsData'].values())
        n_eq=len(data['teamsData'])
        stats_div.append({'division':div_nombre,'total_goles':total_g,'total_tarjetas':total_t,'num_equipos':n_eq,
            'promedio_goles_equipo':round(total_g/n_eq,1) if n_eq else 0,'promedio_tarjetas_equipo':round(total_t/n_eq,1) if n_eq else 0})
    all_g.sort(key=lambda x:-x['goles']); all_c.sort(key=lambda x:-x['_p']); fp_torneo.sort(key=lambda x:x['total_tarjetas']); stats_div.sort(key=lambda x:-x['total_goles'])
    top_c=[{k:v for k,v in c.items() if k!='_p'} for c in all_c[:15]]
    global_torneo={'top_15_goleadores':all_g[:15],'top_15_amonestados':top_c,'fair_play':fp_torneo[:20],'ranking_divisiones':stats_div}
    for d in resultados.values(): d.pop('_goles_raw',None); d.pop('_tarjetas_raw',None)
    return {'generado':datetime.now().strftime('%d/%m/%Y %H:%M'),'divisiones':resultados,'globalTorneo':global_torneo}

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
  --red:#e8192c;--red2:#ff3347;--red-dim:#3d0a10;
  --gold:#f5a623;--green:#00c853;--blue:#2979ff;--orange:#ff6d00;
  --text:#f0f2f8;--text2:#8b95b0;--text3:#5a6480;
  --1era:#e8192c;--2da:#2979ff;--c1:#00c853;--c2:#f5a623;
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

/* HEADER */
.site-header{background:linear-gradient(135deg,#0d0f14 0%,#1a0508 50%,#0d0f14 100%);border-bottom:2px solid var(--red);padding:0;position:sticky;top:0;z-index:100;box-shadow:0 4px 30px rgba(232,25,44,0.2)}
.header-inner{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;padding:12px 24px;gap:16px}
.logo-block{display:flex;align-items:center;gap:14px}
.logo-emblem{width:52px;height:52px;background:var(--red);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:26px;font-family:'Barlow Condensed',sans-serif;font-weight:900;color:white;letter-spacing:-1px;box-shadow:0 0 20px rgba(232,25,44,0.4)}
.logo-text{display:flex;flex-direction:column}
.logo-title{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:900;letter-spacing:2px;color:var(--text);text-transform:uppercase;line-height:1}
.logo-sub{font-size:10px;font-weight:600;color:var(--text2);letter-spacing:3px;text-transform:uppercase;margin-top:2px}
.header-meta{font-size:11px;color:var(--text3);font-weight:500;text-align:right;line-height:1.6}
.header-meta span{color:var(--red);font-weight:700}

/* DIVISION TABS */
.div-tabs{background:var(--surface);border-bottom:1px solid var(--border);position:sticky;top:78px;z-index:90}
.div-tabs-inner{max-width:1400px;margin:0 auto;display:flex;overflow-x:auto;scrollbar-width:none;gap:0;padding:0 16px}
.div-tabs-inner::-webkit-scrollbar{display:none}
.dtab{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;padding:14px 20px;border:none;background:transparent;color:var(--text3);cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap}
.dtab:hover{color:var(--text);background:var(--surface2)}
.dtab.active{color:var(--text);border-bottom-color:var(--accent-color, var(--red));background:var(--surface2)}

/* MAIN CONTENT */
.main{max-width:1400px;margin:0 auto;padding:24px 16px}
.tab-panel{display:none}
.tab-panel.active{display:block}

/* SECTION HEADER */
.section-eyebrow{font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--text3);margin-bottom:6px}
.section-title{font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:900;letter-spacing:1px;text-transform:uppercase;color:var(--text);line-height:1}
.section-title span{color:var(--accent, var(--red))}

/* HERO STATS (global) */
.hero-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1px;background:var(--border);border-radius:10px;overflow:hidden;margin:20px 0}
.hero-stat{background:var(--surface);padding:20px 24px;text-align:center}
.hero-stat-val{font-family:'Barlow Condensed',sans-serif;font-size:42px;font-weight:900;line-height:1;color:var(--text)}
.hero-stat-val.red{color:var(--red)}
.hero-stat-lbl{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-top:4px}

/* DIV CARDS (ranking divisiones) */
.div-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px;margin:20px 0}
.div-card{background:var(--surface);border-radius:10px;overflow:hidden;border:1px solid var(--border)}
.div-card-header{padding:14px 18px;display:flex;align-items:center;justify-content:space-between}
.div-card-title{font-family:'Barlow Condensed',sans-serif;font-size:16px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:white}
.div-card-rank{font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:900;color:rgba(255,255,255,0.2)}
.div-card-body{padding:14px 18px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;background:var(--surface2)}
.div-stat-mini{text-align:center}
.div-stat-mini-val{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--text)}
.div-stat-mini-lbl{font-size:9px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--text3);margin-top:2px}

/* GRID LAYOUT */
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin:20px 0}
.grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin:20px 0}
@media(max-width:900px){.grid-3{grid-template-columns:1fr}.grid-2{grid-template-columns:1fr}}

/* PANEL */
.panel{background:var(--surface);border-radius:10px;border:1px solid var(--border);overflow:hidden}
.panel-head{padding:14px 18px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.panel-head-icon{font-size:16px}
.panel-head-title{font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:var(--text)}
.panel-body{padding:12px}

/* RANKING LIST */
.rank-list{display:flex;flex-direction:column;gap:4px}
.rank-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;background:var(--surface2);border:1px solid transparent;transition:border-color .15s}
.rank-item:hover{border-color:var(--border2)}
.rank-num{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:800;color:var(--text3);min-width:18px;text-align:center}
.rank-num.gold{color:var(--gold)}
.rank-num.silver{color:#9aa0b0}
.rank-num.bronze{color:#cd7f32}
.rank-info{flex:1;overflow:hidden}
.rank-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-sub{font-size:10px;font-weight:600;color:var(--text3);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-val{font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:800;color:var(--text);min-width:32px;text-align:right}
.rank-val small{font-size:10px;font-weight:600;color:var(--text3)}

/* BADGES */
.badge{display:inline-flex;align-items:center;justify-content:center;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:800;padding:2px 6px;border-radius:4px;margin-left:3px}
.badge-yellow{background:#7c6300;color:#ffd600}
.badge-red{background:#5c0010;color:#ff5252}
.badge-orange{background:#5c2c00;color:#ff9100}

/* FAIR PLAY BAR */
.fp-item{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:6px;background:var(--surface2);margin-bottom:4px}
.fp-bar-wrap{flex:1;height:4px;background:var(--surface3);border-radius:2px;overflow:hidden}
.fp-bar{height:100%;border-radius:2px;background:var(--green);transition:width .6s}
.fp-val{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:800;color:var(--text);min-width:28px;text-align:right}

/* DIVISION DETAIL */
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
.team-select:focus{border-color:var(--accent,var(--red))}

/* TEAM DETAIL CARDS */
.team-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:1px;background:var(--border);border-radius:8px;overflow:hidden;margin-bottom:16px}
.team-kpi{background:var(--surface2);padding:16px;text-align:center}
.team-kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:900;color:var(--text);line-height:1}
.team-kpi-val.accent{color:var(--accent,var(--red))}
.team-kpi-lbl{font-size:9px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--text3);margin-top:4px}

/* SCORE BAR (top scorer visual) */
.scorer-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:6px;background:var(--surface2);margin-bottom:4px;border:1px solid transparent;transition:border-color .15s}
.scorer-item:hover{border-color:var(--border2)}
.scorer-bar-wrap{flex:1;height:3px;background:var(--surface3);border-radius:2px;overflow:hidden}
.scorer-bar{height:100%;border-radius:2px}

/* SANCIONADOS */
.sanc-item{display:flex;align-items:center;justify-content:space-between;padding:8px 10px;border-radius:6px;background:var(--surface2);margin-bottom:4px;border-left:3px solid var(--red)}
.sanc-info{flex:1;overflow:hidden}
.sanc-name{font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sanc-team{font-size:10px;color:var(--text3);margin-top:1px}
.sanc-badges{display:flex;gap:3px;flex-wrap:wrap;justify-content:flex-end}

/* RESPONSIVE */
@media(max-width:600px){
  .header-inner{padding:10px 12px}
  .logo-title{font-size:16px}
  .main{padding:16px 10px}
  .div-hero{padding:14px}
  .div-hero-title{font-size:24px}
  .hero-stat-val{font-size:30px}
}
</style>
</head>
<body>

<!-- HEADER -->
<header class="site-header">
  <div class="header-inner">
    <div class="logo-block">
      <div class="logo-emblem">L</div>
      <div class="logo-text">
        <div class="logo-title">Liga San Miguel</div>
        <div class="logo-sub">de Ibarra · LSMI</div>
      </div>
    </div>
    <div class="header-meta" id="headerMeta"></div>
  </div>
</header>

<!-- DIVISION TABS -->
<nav class="div-tabs">
  <div class="div-tabs-inner">
    <button class="dtab active" data-tab="global" data-color="var(--red)" onclick="switchTab(this,'global')">🏆 Global</button>
    <button class="dtab" data-tab="1era División" data-color="var(--1era)" onclick="switchTab(this,'1era División')">1era División</button>
    <button class="dtab" data-tab="2da División" data-color="var(--2da)" onclick="switchTab(this,'2da División')">2da División</button>
    <button class="dtab" data-tab="3era División C1" data-color="var(--c1)" onclick="switchTab(this,'3era División C1')">3era Div. C1</button>
    <button class="dtab" data-tab="3era División C2" data-color="var(--c2)" onclick="switchTab(this,'3era División C2')">3era Div. C2</button>
  </div>
</nav>

<main class="main">

<!-- PANEL GLOBAL -->
<div id="panel-global" class="tab-panel active">
  <div style="margin-bottom:20px">
    <div class="section-eyebrow">Temporada 2025</div>
    <div class="section-title">Estadística <span>Global</span> del Torneo</div>
  </div>
  <div class="hero-stats" id="heroStats"></div>
  <div class="div-cards" id="rankingDivisiones"></div>
  <div class="grid-3">
    <div class="panel">
      <div class="panel-head"><span class="panel-head-icon">⚽</span><span class="panel-head-title">Top 15 Goleadores</span></div>
      <div class="panel-body"><div class="rank-list" id="globalGoleadores"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span class="panel-head-icon">🟨</span><span class="panel-head-title">Top 15 Disciplina</span></div>
      <div class="panel-body"><div class="rank-list" id="globalTarjetas"></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><span class="panel-head-icon">🤝</span><span class="panel-head-title">Fair Play General</span></div>
      <div class="panel-body" id="globalFairPlay"></div>
    </div>
  </div>
</div>

<!-- PANELES DIVISIONES (generados dinámicamente) -->
<div id="panel-1era División" class="tab-panel"></div>
<div id="panel-2da División" class="tab-panel"></div>
<div id="panel-3era División C1" class="tab-panel"></div>
<div id="panel-3era División C2" class="tab-panel"></div>

</main>

<footer style="text-align:center;padding:24px 16px;color:var(--text3);font-size:11px;font-weight:600;letter-spacing:1px;border-top:1px solid var(--border);margin-top:20px">
  LIGA SAN MIGUEL DE IBARRA · ESTADÍSTICAS OFICIALES · <span id="footerDate"></span>
</footer>

<script>
// ============================================================
// DATOS INCRUSTADOS — NO editar. Ejecutar: python generar_datos.py
// ============================================================
const DATA = __JSON_DATA__;

const DIV_COLORS = {'1era División':'#e8192c','2da División':'#2979ff','3era División C1':'#00c853','3era División C2':'#f5a623'};
const DIV_BG     = {'1era División':'#3d0a10','2da División':'#0a1a40','3era División C1':'#003320','3era División C2':'#3d2800'};

// META
document.getElementById('headerMeta').innerHTML = `Fecha <span>${DATA.divisiones['1era División'].meta.ultima_fecha}</span> &nbsp;·&nbsp; Actualizado: ${DATA.generado}`;
document.getElementById('footerDate').textContent = DATA.generado;

// TABS
function switchTab(btn, id) {
  document.querySelectorAll('.dtab').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('panel-' + id).classList.add('active');
  const color = btn.dataset.color;
  document.querySelectorAll('.dtab.active').forEach(b => b.style.borderBottomColor = color);
  window.scrollTo({top:0,behavior:'smooth'});
}

// BADGES
function badge(c, big=false) {
  const sz = big ? 'font-size:13px;padding:3px 8px' : '';
  let b = '';
  if(c.rojas   > 0) b += `<span class="badge badge-red" style="${sz}">${c.rojas}🔴</span>`;
  if(c.dobles  > 0) b += `<span class="badge badge-orange" style="${sz}">${c.dobles}🟠</span>`;
  if(c.amarillas>0) b += `<span class="badge badge-yellow" style="${sz}">${c.amarillas}🟡</span>`;
  return b;
}

function rankNum(i) {
  if(i===0) return '<span class="rank-num gold">1</span>';
  if(i===1) return '<span class="rank-num silver">2</span>';
  if(i===2) return '<span class="rank-num bronze">3</span>';
  return `<span class="rank-num">${i+1}</span>`;
}

// ---- GLOBAL TORNEO ----
function renderGlobal() {
  const g = DATA.globalTorneo;
  const divs = DATA.divisiones;

  // Hero stats
  const totalGoles = g.ranking_divisiones.reduce((a,d)=>a+d.total_goles,0);
  const totalTarj  = g.ranking_divisiones.reduce((a,d)=>a+d.total_tarjetas,0);
  const totalEq    = g.ranking_divisiones.reduce((a,d)=>a+d.num_equipos,0);
  const totalDiv   = g.ranking_divisiones.length;
  document.getElementById('heroStats').innerHTML = `
    <div class="hero-stat"><div class="hero-stat-val red">${totalGoles}</div><div class="hero-stat-lbl">Goles en el Torneo</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalEq}</div><div class="hero-stat-lbl">Equipos</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalDiv}</div><div class="hero-stat-lbl">Divisiones</div></div>
    <div class="hero-stat"><div class="hero-stat-val">${totalTarj}</div><div class="hero-stat-lbl">Tarjetas Totales</div></div>
    <div class="hero-stat"><div class="hero-stat-val red">${g.top_15_goleadores[0].goles}</div><div class="hero-stat-lbl">Goles Líder</div></div>`;

  // Ranking divisiones
  const rd = document.getElementById('rankingDivisiones');
  rd.innerHTML = '';
  g.ranking_divisiones.forEach((d,i) => {
    const color = DIV_COLORS[d.division]||'#e8192c';
    const bg    = DIV_BG[d.division]||'#3d0a10';
    rd.innerHTML += `<div class="div-card">
      <div class="div-card-header" style="background:${bg};border-bottom:2px solid ${color}">
        <div>
          <div class="div-card-title" style="color:${color}">${d.division}</div>
          <div style="font-size:10px;color:rgba(255,255,255,0.4);font-weight:600;letter-spacing:1px;margin-top:2px">${d.num_equipos} EQUIPOS</div>
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

  // Top goleadores
  const maxG = g.top_15_goleadores[0]?.goles || 1;
  const lg = document.getElementById('globalGoleadores');
  lg.innerHTML = '';
  g.top_15_goleadores.forEach((p,i) => {
    const color = DIV_COLORS[p.division]||'#e8192c';
    const pct = Math.round(p.goles/maxG*100);
    lg.innerHTML += `<div class="scorer-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${p.nombre}</div>
        <div class="rank-sub" style="color:${color}">${p.equipo} · ${p.division}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${p.goles}<small> ⚽</small></div>
    </div>`;
  });

  // Top indisciplina
  const lt = document.getElementById('globalTarjetas');
  lt.innerHTML = '';
  g.top_15_amonestados.forEach((c,i) => {
    const color = DIV_COLORS[c.division]||'#e8192c';
    lt.innerHTML += `<div class="rank-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${c.nombre}</div>
        <div class="rank-sub" style="color:${color}">${c.equipo} · ${c.division}</div>
      </div>
      <div style="display:flex;gap:3px">${badge(c)}</div>
    </div>`;
  });

  // Fair play
  const maxFP = Math.max(...g.fair_play.map(f=>f.total_tarjetas))||1;
  const fp = document.getElementById('globalFairPlay');
  fp.innerHTML = '';
  g.fair_play.slice(0,16).forEach((eq,i) => {
    const color = DIV_COLORS[eq.division]||'#e8192c';
    const pct = Math.round(eq.total_tarjetas/maxFP*100);
    const barColor = i<3?'#00c853':i<8?'#f5a623':'#e8192c';
    fp.innerHTML += `<div class="fp-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${eq.equipo}</div>
        <div class="rank-sub" style="color:${color}">${eq.division}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${barColor}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div>
    </div>`;
  });
}

// ---- DIVISION PANEL ----
function buildDivPanel(divNombre, divData) {
  const container = document.getElementById('panel-' + divNombre);
  const color = DIV_COLORS[divNombre]||'#e8192c';
  const bg    = DIV_BG[divNombre]||'#3d0a10';
  const meta  = divData.meta;
  const gd    = divData.globalData;

  // KPIs globales de la división
  const totalGoles = Object.values(divData.teamsData).reduce((a,t)=>a+t.total_goles,0);
  const totalAmar  = Object.values(divData.teamsData).reduce((a,t)=>a+t.total_amarillas,0);
  const totalRojas = Object.values(divData.teamsData).reduce((a,t)=>a+t.total_rojas+t.total_doble,0);
  const numEq      = Object.keys(divData.teamsData).length;

  container.innerHTML = `
  <div class="div-hero" style="border-color:${color};border-top:3px solid ${color}">
    <div>
      <div class="div-hero-title" style="color:${color}">${divNombre}</div>
      <div class="div-hero-meta">Fecha ${meta.ultima_fecha} · Actualizado ${meta.generado} · ${numEq} equipos</div>
    </div>
    <div class="div-kpis">
      <div class="div-kpi"><div class="div-kpi-val" style="color:${color}">${totalGoles}</div><div class="div-kpi-lbl">Goles</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalAmar}</div><div class="div-kpi-lbl">Amarillas</div></div>
      <div class="div-kpi"><div class="div-kpi-val">${totalRojas}</div><div class="div-kpi-lbl">Rojas/Dobles</div></div>
    </div>
  </div>

  <div class="team-selector-wrap">
    <span class="team-selector-lbl">Ver equipo</span>
    <select class="team-select" id="sel-${divNombre}" onchange="renderTeam('${divNombre}',this.value)" style="border-color:${color}">
      <option value="__global__">📊 Estadística de la División</option>
    </select>
  </div>

  <!-- GLOBAL DIV -->
  <div id="divglobal-${divNombre}">
    <div class="grid-3">
      <div class="panel">
        <div class="panel-head" style="border-bottom-color:${color}40"><span class="panel-head-icon">⚽</span><span class="panel-head-title">Top Goleadores</span></div>
        <div class="panel-body"><div class="rank-list" id="dg-${divNombre}"></div></div>
      </div>
      <div class="panel">
        <div class="panel-head"><span class="panel-head-icon">🟨</span><span class="panel-head-title">Más Amonestados</span></div>
        <div class="panel-body"><div class="rank-list" id="dt-${divNombre}"></div></div>
      </div>
      <div class="panel">
        <div class="panel-head"><span class="panel-head-icon">🤝</span><span class="panel-head-title">Fair Play</span></div>
        <div class="panel-body" id="df-${divNombre}"></div>
      </div>
    </div>
  </div>

  <!-- EQUIPO DETAIL -->
  <div id="divteam-${divNombre}" style="display:none">
    <div class="team-kpi-row" id="tkpi-${divNombre}"></div>
    <div class="grid-2" id="tgrid-${divNombre}"></div>
  </div>`;

  // Poblar selector
  const sel = document.getElementById(`sel-${divNombre}`);
  Object.keys(divData.teamsData).sort().forEach(t => {
    const o=document.createElement('option'); o.value=t; o.textContent='⚽ '+t; sel.appendChild(o);
  });

  // Render listas globales de la división
  const maxG = gd.top_10_goleadores[0]?.goles||1;
  const dg = document.getElementById(`dg-${divNombre}`);
  gd.top_10_goleadores.forEach((p,i) => {
    const pct=Math.round(p.goles/maxG*100);
    dg.innerHTML += `<div class="scorer-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${p.nombre}</div>
        <div class="rank-sub" style="color:${color}">${p.equipo}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${p.goles}<small> ⚽</small></div>
    </div>`;
  });

  const dt = document.getElementById(`dt-${divNombre}`);
  gd.top_10_amonestados.forEach((c,i) => {
    dt.innerHTML += `<div class="rank-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${c.nombre}</div>
        <div class="rank-sub" style="color:${color}">${c.equipo}</div>
      </div>
      <div style="display:flex;gap:3px">${badge(c)}</div>
    </div>`;
  });

  const maxFP = Math.max(...gd.fair_play.map(f=>f.total_tarjetas))||1;
  const df = document.getElementById(`df-${divNombre}`);
  gd.fair_play.forEach((eq,i) => {
    const pct=Math.round(eq.total_tarjetas/maxFP*100);
    const bc=i<3?'#00c853':i<7?'#f5a623':'#e8192c';
    df.innerHTML += `<div class="fp-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${eq.equipo}</div>
        <div class="fp-bar-wrap"><div class="fp-bar" style="width:${pct}%;background:${bc}"></div></div>
      </div>
      <div class="fp-val">${eq.total_tarjetas}</div>
    </div>`;
  });
}

function renderTeam(divNombre, equipo) {
  const divGl = document.getElementById(`divglobal-${divNombre}`);
  const divTm = document.getElementById(`divteam-${divNombre}`);
  if(equipo==='__global__') { divTm.style.display='none'; divGl.style.display='block'; return; }
  divGl.style.display='none'; divTm.style.display='block';

  const d = DATA.divisiones[divNombre].teamsData[equipo];
  const color = DIV_COLORS[divNombre]||'#e8192c';

  document.getElementById(`tkpi-${divNombre}`).innerHTML = `
    <div class="team-kpi"><div class="team-kpi-val accent">${d.total_goles}</div><div class="team-kpi-lbl">Goles Totales</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_goles}</div><div class="team-kpi-lbl">Promedio Goles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_amarillas}</div><div class="team-kpi-lbl">Amarillas</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.total_rojas + d.total_doble}</div><div class="team-kpi-lbl">Rojas/Dobles</div></div>
    <div class="team-kpi"><div class="team-kpi-val">${d.promedio_tarjetas}</div><div class="team-kpi-lbl">Prom. Tarjetas</div></div>`;

  const maxG = d.top_scorers[0]?.goles||1;
  let golesHTML = d.top_scorers.length ? '' : '<div style="color:var(--text3);font-size:12px;padding:10px;text-align:center">Sin goles registrados</div>';
  d.top_scorers.forEach((g,i) => {
    const pct=Math.round(g.goles/maxG*100);
    golesHTML += `<div class="scorer-item">
      ${rankNum(i)}
      <div class="rank-info">
        <div class="rank-name">${g.nombre}</div>
        <div class="scorer-bar-wrap"><div class="scorer-bar" style="width:${pct}%;background:${color}"></div></div>
      </div>
      <div class="rank-val">${g.goles}<small> ⚽</small></div>
    </div>`;
  });

  let tarjHTML = d.top_cards.length ? '' : '<div style="color:var(--text3);font-size:12px;padding:10px;text-align:center">Sin tarjetas registradas</div>';
  d.top_cards.forEach((c,i) => {
    tarjHTML += `<div class="rank-item">
      ${rankNum(i)}
      <div class="rank-info"><div class="rank-name">${c.nombre}</div></div>
      <div style="display:flex;gap:3px">${badge(c,true)}</div>
    </div>`;
  });

  document.getElementById(`tgrid-${divNombre}`).innerHTML = `
    <div class="panel">
      <div class="panel-head" style="border-bottom-color:${color}40"><span class="panel-head-icon">⚽</span><span class="panel-head-title">${equipo} · Goleadores</span></div>
      <div class="panel-body">${golesHTML}</div>
    </div>
    <div class="panel">
      <div class="panel-head"><span class="panel-head-icon">🟨</span><span class="panel-head-title">${equipo} · Disciplina</span></div>
      <div class="panel-body">${tarjHTML}</div>
    </div>`;
}

// INICIALIZAR
renderGlobal();
['1era División','2da División','3era División C1','3era División C2'].forEach(div => buildDivPanel(div, DATA.divisiones[div]));
</script>
</body>
</html>"""

def generar_html(datos):
    json_str = json.dumps(datos, ensure_ascii=False)
    return HTML_TEMPLATE.replace('__JSON_DATA__', json_str)

if __name__ == '__main__':
    carpeta = os.path.dirname(os.path.abspath(__file__))
    datos   = procesar_todo(carpeta)
    html    = generar_html(datos)
    salida  = os.path.join(carpeta, 'index.html')
    with open(salida, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n✅ ¡Listo! → {salida}")
    print(f"👉 Abre index.html con doble clic — funciona directo.")
