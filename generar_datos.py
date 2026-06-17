#!/usr/bin/env python3
"""
LSMI - Generador de página web completa (4 divisiones)
========================================================
Uso: python generar_datos.py

Busca automáticamente los 4 Excel en la misma carpeta y genera index.html
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
        fecha = int(row[0])
        equipo  = str(row[1]).strip() if row[1] else ''
        jugador = str(row[2]).strip() if row[2] else ''
        goles   = row[3]
        if equipo and jugador and goles and isinstance(goles, (int, float)):
            key = (equipo, jugador)
            goles_por_jugador[key] = goles_por_jugador.get(key, 0) + int(goles)
            ultima_fecha = max(ultima_fecha, fecha)

    ws_a = wb['AMONESTACIONES']
    tarjetas_por_jugador = {}
    for row in ws_a.iter_rows(min_row=3, values_only=True):
        if not row[0] or not isinstance(row[0], (int, float)): continue
        equipo  = str(row[1]).strip() if row[1] else ''
        jugador = str(row[2]).strip() if row[2] else ''
        tarjeta = str(row[3]).strip() if row[3] else ''
        if equipo and jugador and tarjeta:
            key = (equipo, jugador)
            if key not in tarjetas_por_jugador:
                tarjetas_por_jugador[key] = {'amarillas': 0, 'dobles': 0, 'rojas': 0}
            if 'Doble' in tarjeta or 'doble' in tarjeta:
                tarjetas_por_jugador[key]['dobles'] += 1
            elif 'Roja' in tarjeta or 'roja' in tarjeta or '🔴' in tarjeta:
                tarjetas_por_jugador[key]['rojas'] += 1
            else:
                tarjetas_por_jugador[key]['amarillas'] += 1

    equipos = sorted(set(k[0] for k in goles_por_jugador) | set(k[0] for k in tarjetas_por_jugador))
    teams_data = {}
    for equipo in equipos:
        goles_eq = {j: g for (e,j),g in goles_por_jugador.items() if e == equipo}
        tarj_eq  = {j: t for (e,j),t in tarjetas_por_jugador.items() if e == equipo}
        total_goles     = sum(goles_eq.values())
        total_amarillas = sum(t['amarillas'] for t in tarj_eq.values())
        total_dobles    = sum(t['dobles']    for t in tarj_eq.values())
        total_rojas     = sum(t['rojas']     for t in tarj_eq.values())
        promedio_goles    = round(total_goles / ultima_fecha, 2) if ultima_fecha else 0
        promedio_tarjetas = round((total_amarillas+total_dobles+total_rojas) / ultima_fecha, 2) if ultima_fecha else 0
        top_scorers = sorted(goles_eq.items(), key=lambda x: -x[1])[:8]
        top_cards_raw = []
        for j, t in tarj_eq.items():
            p = t['amarillas'] + t['dobles']*2 + t['rojas']*3
            if p > 0:
                top_cards_raw.append({'nombre': j, 'amarillas': t['amarillas'],
                                      'dobles': t['dobles'], 'rojas': t['rojas'], '_p': p})
        top_cards_raw.sort(key=lambda x: -x['_p'])
        top_cards = [{k: v for k,v in c.items() if k != '_p'} for c in top_cards_raw[:8]]
        teams_data[equipo] = {
            'total_goles': total_goles, 'promedio_goles': promedio_goles,
            'total_amarillas': total_amarillas, 'total_doble': total_dobles,
            'total_rojas': total_rojas, 'promedio_tarjetas': promedio_tarjetas,
            'top_scorers': [{'nombre': j, 'goles': g} for j, g in top_scorers],
            'top_cards': top_cards,
        }

    all_scorers = [{'nombre': j, 'equipo': e, 'goles': g} for (e,j),g in goles_por_jugador.items()]
    all_scorers.sort(key=lambda x: -x['goles'])
    all_cards = []
    for (e,j),t in tarjetas_por_jugador.items():
        p = t['amarillas'] + t['dobles']*2 + t['rojas']*3
        if p > 0:
            all_cards.append({'nombre': j, 'equipo': e, 'amarillas': t['amarillas'],
                              'dobles': t['dobles'], 'rojas': t['rojas'], '_p': p})
    all_cards.sort(key=lambda x: -x['_p'])
    top10_cards = [{k: v for k,v in c.items() if k != '_p'} for c in all_cards[:10]]
    fair_play = []
    for equipo in equipos:
        tarj_e = {j: t for (e,j),t in tarjetas_por_jugador.items() if e == equipo}
        total = sum(t['amarillas']+t['dobles']+t['rojas'] for t in tarj_e.values())
        fair_play.append({'equipo': equipo, 'total_tarjetas': total})
    fair_play.sort(key=lambda x: x['total_tarjetas'])

    return {
        'meta': {'ultima_fecha': ultima_fecha, 'generado': datetime.now().strftime('%d/%m/%Y %H:%M')},
        'globalData': {'top_10_goleadores': all_scorers[:10], 'top_10_amonestados': top10_cards, 'fair_play': fair_play},
        'teamsData': teams_data,
        '_goles_raw': {f"{e}||{j}": g for (e,j),g in goles_por_jugador.items()},
        '_tarjetas_raw': {f"{e}||{j}": t for (e,j),t in tarjetas_por_jugador.items()},
    }


def buscar_excel(carpeta):
    """Busca los 4 Excel por palabras clave en el nombre."""
    archivos = [f for f in os.listdir(carpeta) if f.endswith('.xlsx')]
    mapa = {}
    claves = {
        '1era División':    ['1era','1era_div','primera'],
        '2da División':     ['2da','2da_div','segunda'],
        '3era División C1': ['3era_div_1','3era_division_1','division_1_v','grupo_1','c1','g1'],
        '3era División C2': ['3era_div_2','3era_division_2','division_2_v','grupo_2','c2','g2'],
    }
    for archivo in archivos:
        low = archivo.lower()
        for div, palabras in claves.items():
            if any(p in low for p in palabras):
                mapa[div] = os.path.join(carpeta, archivo)
                break
    return mapa


def procesar_todo(carpeta):
    mapa = buscar_excel(carpeta)
    print(f"📁 Archivos encontrados:")
    for div, ruta in mapa.items():
        print(f"   {div}: {os.path.basename(ruta)}")

    if len(mapa) < 4:
        faltantes = [d for d in ['1era División','2da División','3era División C1','3era División C2'] if d not in mapa]
        print(f"⚠️  No se encontraron: {faltantes}")
        print("   Asegúrate que los 4 Excel estén en la misma carpeta que este script.")
        sys.exit(1)

    resultados = {}
    for nombre, ruta in mapa.items():
        print(f"\n⚽ Procesando {nombre}...")
        resultados[nombre] = procesar_division(ruta)
        print(f"   Equipos: {len(resultados[nombre]['teamsData'])} | Fecha: {resultados[nombre]['meta']['ultima_fecha']}")

    # GLOBAL TORNEO
    all_g, all_c, fp_torneo, stats_div = [], [], [], []
    for div_nombre, data in resultados.items():
        for key, goles in data['_goles_raw'].items():
            e, j = key.split('||')
            all_g.append({'nombre': j, 'equipo': e, 'division': div_nombre, 'goles': goles})
        for key, t in data['_tarjetas_raw'].items():
            e, j = key.split('||')
            p = t['amarillas'] + t['dobles']*2 + t['rojas']*3
            if p > 0:
                all_c.append({'nombre': j, 'equipo': e, 'division': div_nombre,
                               'amarillas': t['amarillas'], 'dobles': t['dobles'],
                               'rojas': t['rojas'], '_p': p})
        for eq_fp in data['globalData']['fair_play']:
            fp_torneo.append({'equipo': eq_fp['equipo'], 'division': div_nombre,
                              'total_tarjetas': eq_fp['total_tarjetas']})
        total_g = sum(d['total_goles'] for d in data['teamsData'].values())
        total_t = sum(d['total_amarillas']+d['total_doble']+d['total_rojas'] for d in data['teamsData'].values())
        n_eq    = len(data['teamsData'])
        stats_div.append({
            'division': div_nombre, 'total_goles': total_g, 'total_tarjetas': total_t,
            'num_equipos': n_eq,
            'promedio_goles_equipo': round(total_g / n_eq, 1) if n_eq else 0,
            'promedio_tarjetas_equipo': round(total_t / n_eq, 1) if n_eq else 0,
        })

    all_g.sort(key=lambda x: -x['goles'])
    all_c.sort(key=lambda x: -x['_p'])
    fp_torneo.sort(key=lambda x: x['total_tarjetas'])
    stats_div.sort(key=lambda x: -x['total_goles'])
    top_c = [{k: v for k,v in c.items() if k != '_p'} for c in all_c[:15]]

    global_torneo = {
        'top_15_goleadores': all_g[:15],
        'top_15_amonestados': top_c,
        'fair_play': fp_torneo[:20],
        'ranking_divisiones': stats_div,
    }

    for d in resultados.values():
        d.pop('_goles_raw', None); d.pop('_tarjetas_raw', None)

    return {
        'generado': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'divisiones': resultados,
        'globalTorneo': global_torneo,
    }


def generar_html(datos):
    json_str = json.dumps(datos, ensure_ascii=False)
    generado = datos['generado']

    html = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Estadísticas LSMI - Liga San Miguel de Ibarra</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Roboto', sans-serif; background-color: #f4f4f5; }
        .card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 1.5rem; }
        .header-bg { background: linear-gradient(90deg, #991b1b 0%, #7f1d1d 100%); }
        .tab-btn { transition: all .2s; cursor: pointer; white-space: nowrap; }
        .tab-btn.active { background: #991b1b; color: white; }
        .tab-btn:not(.active) { background: white; color: #374151; }
        .tab-btn:not(.active):hover { background: #fef2f2; color: #991b1b; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        ::-webkit-scrollbar { height: 4px; } 
        ::-webkit-scrollbar-thumb { background: #991b1b; border-radius: 4px; }
    </style>
</head>
<body class="text-gray-800">

<!-- HEADER -->
<div class="header-bg p-5 shadow-md border-b-4 border-red-500">
    <div class="max-w-7xl mx-auto text-center">
        <h1 class="text-3xl font-black uppercase tracking-wider text-white">⚽ Liga San Miguel de Ibarra</h1>
        <p class="text-red-200 text-sm font-bold mt-1">ESTADÍSTICAS OFICIALES DEL TORNEO · """ + generado + """</p>
    </div>
</div>

<!-- TABS DIVISIONES -->
<div class="bg-white shadow-sm border-b border-gray-200 sticky top-0 z-30">
    <div class="max-w-7xl mx-auto px-4 overflow-x-auto">
        <div class="flex gap-1 py-2" id="tabBar">
            <button class="tab-btn active px-4 py-2 rounded font-black text-sm uppercase tracking-wide" onclick="switchTab('global')">🏆 Global LSMI</button>
            <button class="tab-btn px-4 py-2 rounded font-black text-sm uppercase tracking-wide" onclick="switchTab('1era División')">1era División</button>
            <button class="tab-btn px-4 py-2 rounded font-black text-sm uppercase tracking-wide" onclick="switchTab('2da División')">2da División</button>
            <button class="tab-btn px-4 py-2 rounded font-black text-sm uppercase tracking-wide" onclick="switchTab('3era División C1')">3era Div. C1</button>
            <button class="tab-btn px-4 py-2 rounded font-black text-sm uppercase tracking-wide" onclick="switchTab('3era División C2')">3era Div. C2</button>
        </div>
    </div>
</div>

<!-- TAB: GLOBAL TORNEO -->
<div id="tab-global" class="tab-content active max-w-7xl mx-auto px-4 py-6">
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-4 mb-6 text-center">
        <h2 class="text-3xl font-black text-red-800 uppercase tracking-tight">🏆 Estadística Global LSMI</h2>
        <p class="text-gray-500 font-bold mt-1">Resumen de las 4 divisiones del torneo</p>
    </div>

    <!-- Ranking divisiones -->
    <div class="card border-t-4 border-red-800 mb-6">
        <h3 class="text-lg font-black mb-4 text-gray-800 uppercase tracking-wide border-b-2 border-gray-100 pb-2">📊 Ranking de Divisiones</h3>
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4" id="rankingDivisiones"></div>
    </div>

    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
        <div class="card border-t-4 border-gray-800">
            <h3 class="text-lg font-black mb-4 text-gray-800 uppercase tracking-wide border-b-2 border-gray-100 pb-2">⚽ Top 15 Goleadores</h3>
            <ul id="globalGoleadores" class="space-y-2 text-sm font-medium"></ul>
        </div>
        <div class="card border-t-4 border-red-600">
            <h3 class="text-lg font-black mb-4 text-gray-800 uppercase tracking-wide border-b-2 border-gray-100 pb-2">⚖️ Top 15 Indisciplina</h3>
            <ul id="globalTarjetas" class="space-y-2 text-sm font-medium"></ul>
        </div>
        <div class="card border-t-4 border-green-600">
            <h3 class="text-lg font-black mb-4 text-gray-800 uppercase tracking-wide border-b-2 border-gray-100 pb-2">🤝 Fair Play General</h3>
            <p class="text-xs text-gray-400 mb-3 italic">Equipos con menor acumulación de tarjetas.</p>
            <ul id="globalFairPlay" class="space-y-2 text-sm font-medium"></ul>
        </div>
    </div>
</div>

<!-- TABS DIVISIONES (se generan dinámicamente) -->
<div id="tab-1era División" class="tab-content max-w-7xl mx-auto px-4 py-6"></div>
<div id="tab-2da División" class="tab-content max-w-7xl mx-auto px-4 py-6"></div>
<div id="tab-3era División C1" class="tab-content max-w-7xl mx-auto px-4 py-6"></div>
<div id="tab-3era División C2" class="tab-content max-w-7xl mx-auto px-4 py-6"></div>

<div class="text-center text-xs text-gray-400 font-bold mt-4 mb-8 pb-4">
    * Datos calculados de los registros oficiales de la Liga San Miguel de Ibarra · """ + generado + """
</div>

<script>
// ============================================================
//  DATOS INCRUSTADOS — NO editar. Ejecutar generar_datos.py
// ============================================================
const DATA = """ + json_str + """;

const COLORES_DIV = {
    '1era División':    '#991b1b',
    '2da División':     '#1e40af',
    '3era División C1': '#065f46',
    '3era División C2': '#92400e',
};

// ---- TABS ----
function switchTab(id) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById('tab-' + id).classList.add('active');
    event.currentTarget.classList.add('active');
    window.scrollTo({top:0, behavior:'smooth'});
}

// ---- BADGES TARJETAS ----
function badges(c, small=true) {
    const sz = small ? 'text-[10px] px-1.5 py-0.5' : 'text-xs px-2 py-1';
    let b = '';
    if(c.rojas   > 0) b += `<span class="text-white bg-red-600 ${sz} rounded ml-1 font-black">${c.rojas}🔴</span>`;
    if(c.dobles  > 0) b += `<span class="text-white bg-orange-500 ${sz} rounded ml-1 font-black">${c.dobles}🟠</span>`;
    if(c.amarillas>0) b += `<span class="text-gray-800 bg-yellow-400 ${sz} rounded ml-1 font-black">${c.amarillas}🟡</span>`;
    return b;
}

// ---- RENDER GLOBAL TORNEO ----
function renderGlobalTorneo(g) {
    // Ranking divisiones
    const rd = document.getElementById('rankingDivisiones');
    rd.innerHTML = '';
    g.ranking_divisiones.forEach((d, i) => {
        const color = COLORES_DIV[d.division] || '#991b1b';
        rd.innerHTML += `
        <div class="bg-gray-50 rounded-lg p-4 border-t-4 text-center" style="border-color:${color}">
            <p class="font-black text-xs uppercase mb-2" style="color:${color}">${d.division}</p>
            <p class="text-3xl font-black text-gray-800">${d.total_goles}</p>
            <p class="text-xs text-gray-500 uppercase font-bold">Goles totales</p>
            <div class="border-t border-gray-200 mt-3 pt-3 grid grid-cols-2 gap-2 text-center">
                <div><p class="text-lg font-black text-yellow-600">${d.total_tarjetas}</p><p class="text-[10px] text-gray-400 uppercase font-bold">Tarjetas</p></div>
                <div><p class="text-lg font-black text-gray-700">${d.num_equipos}</p><p class="text-[10px] text-gray-400 uppercase font-bold">Equipos</p></div>
            </div>
        </div>`;
    });

    // Top goleadores
    const lg = document.getElementById('globalGoleadores');
    lg.innerHTML = '';
    g.top_15_goleadores.forEach((p, i) => {
        const medal = i===0?'🥇':i===1?'🥈':i===2?'🥉':`<span class="text-gray-400 w-5 inline-block text-center font-black text-xs">${i+1}</span>`;
        const color = COLORES_DIV[p.division] || '#991b1b';
        lg.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm hover:bg-gray-50">
            <div class="flex items-center gap-2 overflow-hidden">
                ${medal}
                <div class="flex flex-col truncate">
                    <span class="text-gray-800 font-bold text-xs truncate">${p.nombre}</span>
                    <span class="text-[10px] font-black truncate" style="color:${color}">${p.equipo} · ${p.division}</span>
                </div>
            </div>
            <span class="font-black text-sm text-gray-800 bg-gray-100 px-2 py-1 rounded ml-2 whitespace-nowrap">${p.goles} ⚽</span>
        </li>`;
    });

    // Top indisciplina
    const lt = document.getElementById('globalTarjetas');
    lt.innerHTML = '';
    g.top_15_amonestados.forEach((c, i) => {
        const color = COLORES_DIV[c.division] || '#991b1b';
        lt.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm hover:bg-gray-50">
            <div class="flex flex-col overflow-hidden w-3/5">
                <span class="text-gray-800 font-bold text-xs truncate">${c.nombre}</span>
                <span class="text-[10px] font-black truncate" style="color:${color}">${c.equipo} · ${c.division}</span>
            </div>
            <div class="flex flex-wrap justify-end">${badges(c)}</div>
        </li>`;
    });

    // Fair play
    const fp = document.getElementById('globalFairPlay');
    fp.innerHTML = '';
    g.fair_play.forEach((eq, i) => {
        const color = COLORES_DIV[eq.division] || '#991b1b';
        const brd = i < 3 ? 'border-l-4 border-green-500' : 'border-l-4 border-gray-300';
        fp.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm ${brd}">
            <div class="flex items-center gap-2 overflow-hidden">
                <span class="text-gray-400 text-xs font-black w-4">${i+1}</span>
                <div class="flex flex-col truncate">
                    <span class="text-gray-800 font-bold text-xs uppercase truncate">${eq.equipo}</span>
                    <span class="text-[10px] font-black truncate" style="color:${color}">${eq.division}</span>
                </div>
            </div>
            <span class="font-black text-xs text-gray-600 bg-gray-100 px-2 py-1 rounded whitespace-nowrap">${eq.total_tarjetas} tarj.</span>
        </li>`;
    });
}

// ---- RENDER DIVISION ----
function renderDivision(divNombre, divData) {
    const container = document.getElementById('tab-' + divNombre);
    const meta = divData.meta;
    const color = COLORES_DIV[divNombre] || '#991b1b';

    container.innerHTML = `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-4 mb-6 text-center" style="border-top:4px solid ${color}">
        <h2 class="text-3xl font-black uppercase tracking-tight" style="color:${color}">${divNombre}</h2>
        <p class="text-gray-500 font-bold mt-1">Fecha ${meta.ultima_fecha} · Actualizado: ${meta.generado}</p>
    </div>

    <!-- Selector equipo -->
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-4 mb-6 flex flex-col sm:flex-row items-center gap-4">
        <span class="font-black text-gray-700 uppercase text-sm">Ver equipo:</span>
        <select id="select-${divNombre}" class="flex-1 text-gray-900 font-bold p-2 rounded border-2 cursor-pointer outline-none focus:border-red-500"
            style="border-color:${color}" onchange="renderEquipo('${divNombre}')">
            <option value="GLOBAL">📊 Estadística de la División</option>
        </select>
    </div>

    <!-- Global de la división -->
    <div id="divglobal-${divNombre}">
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
            <div class="card border-t-4 border-gray-800">
                <h3 class="text-base font-black mb-4 text-gray-800 uppercase border-b-2 border-gray-100 pb-2">⚽ Top 10 Goleadores</h3>
                <ul id="divgol-${divNombre}" class="space-y-2 text-sm font-medium"></ul>
            </div>
            <div class="card border-t-4 border-red-600">
                <h3 class="text-base font-black mb-4 text-gray-800 uppercase border-b-2 border-gray-100 pb-2">⚖️ Top 10 Indisciplina</h3>
                <ul id="divtarj-${divNombre}" class="space-y-2 text-sm font-medium"></ul>
            </div>
            <div class="card border-t-4 border-green-600">
                <h3 class="text-base font-black mb-4 text-gray-800 uppercase border-b-2 border-gray-100 pb-2">🤝 Fair Play</h3>
                <ul id="divfp-${divNombre}" class="space-y-2 text-sm font-medium"></ul>
            </div>
        </div>
    </div>

    <!-- Detalle equipo -->
    <div id="divclub-${divNombre}" class="hidden">
        <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-4 mb-6" style="border-top:4px solid ${color}">
            <h2 id="clubname-${divNombre}" class="text-3xl font-black text-center uppercase tracking-tight" style="color:${color}"></h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
            <div class="card border-t-4 border-gray-800">
                <h3 class="text-base font-black mb-4 text-gray-800 uppercase border-b-2 border-gray-100 pb-2">⚽ Rendimiento Ofensivo</h3>
                <div class="flex justify-around text-center mb-6 bg-gray-50 p-4 rounded-lg">
                    <div><p class="text-5xl font-black text-gray-800" id="clubgoles-${divNombre}">0</p><p class="text-xs font-bold text-gray-500 uppercase mt-1">Goles Totales</p></div>
                    <div class="border-l-2 border-gray-200 pl-8"><p class="text-5xl font-black text-red-700" id="clubpromgol-${divNombre}">0</p><p class="text-xs font-bold text-gray-500 uppercase mt-1">Promedio/Partido</p></div>
                </div>
                <h4 class="font-bold text-red-800 mb-3 bg-red-50 p-2 rounded text-sm uppercase">Top Goleadores</h4>
                <ul id="clublistagol-${divNombre}" class="space-y-2 text-sm font-medium"></ul>
            </div>
            <div class="card border-t-4 border-red-600">
                <h3 class="text-base font-black mb-4 text-gray-800 uppercase border-b-2 border-gray-100 pb-2">⚖️ Disciplina</h3>
                <div class="flex justify-around text-center mb-6 bg-gray-50 p-4 rounded-lg">
                    <div><p class="text-4xl font-black text-yellow-500" id="clubamar-${divNombre}">0</p><p class="text-xs font-bold text-gray-500 uppercase mt-1">Amarillas</p></div>
                    <div class="border-l border-gray-200 pl-4"><p class="text-4xl font-black text-red-600" id="clubrojas-${divNombre}">0</p><p class="text-xs font-bold text-gray-500 uppercase mt-1">Rojas/Dobles</p></div>
                    <div class="border-l border-gray-200 pl-4"><p class="text-4xl font-black text-gray-800" id="clubpromtarj-${divNombre}">0</p><p class="text-xs font-bold text-gray-500 uppercase mt-1">Promedio</p></div>
                </div>
                <h4 class="font-bold text-red-800 mb-3 bg-red-50 p-2 rounded text-sm uppercase">Mayor Acumulación</h4>
                <ul id="clublistatar-${divNombre}" class="space-y-2 text-sm font-medium"></ul>
            </div>
        </div>
    </div>`;

    // Poblar selector
    const sel = document.getElementById(`select-${divNombre}`);
    Object.keys(divData.teamsData).sort().forEach(t => {
        const o = document.createElement('option');
        o.value = t; o.textContent = '⚽ ' + t;
        sel.appendChild(o);
    });

    // Render global división
    const gd = divData.globalData;
    const lg = document.getElementById(`divgol-${divNombre}`);
    lg.innerHTML = '';
    gd.top_10_goleadores.forEach((p, i) => {
        const medal = i===0?'🥇':i===1?'🥈':i===2?'🥉':`<span class="text-gray-400 w-5 inline-block text-center font-black text-xs">${i+1}</span>`;
        lg.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm hover:bg-gray-50">
            <div class="flex items-center gap-2 overflow-hidden">${medal}
                <div class="flex flex-col truncate">
                    <span class="text-gray-800 font-bold text-xs truncate">${p.nombre}</span>
                    <span class="text-[10px] font-black truncate" style="color:${color}">${p.equipo}</span>
                </div>
            </div>
            <span class="font-black text-sm text-gray-800 bg-gray-100 px-2 py-1 rounded ml-2 whitespace-nowrap">${p.goles} ⚽</span>
        </li>`;
    });

    const lt = document.getElementById(`divtarj-${divNombre}`);
    lt.innerHTML = '';
    gd.top_10_amonestados.forEach(c => {
        lt.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm hover:bg-gray-50">
            <div class="flex flex-col overflow-hidden w-3/5">
                <span class="text-gray-800 font-bold text-xs truncate">${c.nombre}</span>
                <span class="text-[10px] font-black truncate" style="color:${color}">${c.equipo}</span>
            </div>
            <div class="flex flex-wrap justify-end">${badges(c)}</div>
        </li>`;
    });

    const fp = document.getElementById(`divfp-${divNombre}`);
    fp.innerHTML = '';
    gd.fair_play.forEach((eq, i) => {
        const brd = i<3?'border-l-4 border-green-500':'border-l-4 border-gray-300';
        fp.innerHTML += `<li class="flex justify-between items-center bg-white p-2 rounded border border-gray-200 shadow-sm ${brd}">
            <div class="flex items-center gap-2">
                <span class="text-gray-400 text-xs font-black w-4">${i+1}</span>
                <span class="text-gray-800 font-bold text-xs uppercase">${eq.equipo}</span>
            </div>
            <span class="font-black text-xs text-gray-600 bg-gray-100 px-2 py-1 rounded">${eq.total_tarjetas} tarj.</span>
        </li>`;
    });
}

function renderEquipo(divNombre) {
    const sel   = document.getElementById(`select-${divNombre}`);
    const val   = sel.value;
    const divGl = document.getElementById(`divglobal-${divNombre}`);
    const divCl = document.getElementById(`divclub-${divNombre}`);
    if (val === 'GLOBAL') { divCl.classList.add('hidden'); divGl.classList.remove('hidden'); return; }
    divGl.classList.add('hidden'); divCl.classList.remove('hidden');

    const d = DATA.divisiones[divNombre].teamsData[val];
    const color = COLORES_DIV[divNombre] || '#991b1b';
    document.getElementById(`clubname-${divNombre}`).textContent = val;
    document.getElementById(`clubgoles-${divNombre}`).textContent     = d.total_goles;
    document.getElementById(`clubpromgol-${divNombre}`).textContent   = d.promedio_goles;
    document.getElementById(`clubamar-${divNombre}`).textContent      = d.total_amarillas;
    document.getElementById(`clubrojas-${divNombre}`).textContent     = d.total_rojas + d.total_doble;
    document.getElementById(`clubpromtarj-${divNombre}`).textContent  = d.promedio_tarjetas;

    const lg = document.getElementById(`clublistagol-${divNombre}`);
    lg.innerHTML = '';
    if(!d.top_scorers.length) {
        lg.innerHTML = '<li class="text-gray-400 italic p-2 text-center">Sin goles registrados</li>';
    } else {
        d.top_scorers.forEach(g => {
            lg.innerHTML += `<li class="flex justify-between items-center bg-white p-3 rounded border border-gray-200 shadow-sm hover:shadow-md">
                <span class="text-gray-700 font-bold text-sm">${g.nombre}</span>
                <span class="font-black text-lg text-gray-800 bg-gray-100 px-3 py-1 rounded">${g.goles} ⚽</span>
            </li>`;
        });
    }

    const lt = document.getElementById(`clublistatar-${divNombre}`);
    lt.innerHTML = '';
    if(!d.top_cards.length) {
        lt.innerHTML = '<li class="text-gray-400 italic p-2 text-center">Sin tarjetas registradas</li>';
    } else {
        d.top_cards.forEach(c => {
            lt.innerHTML += `<li class="flex justify-between items-center bg-white p-3 rounded border border-gray-200 shadow-sm hover:shadow-md">
                <span class="text-gray-700 font-bold text-sm">${c.nombre}</span>
                <span class="flex flex-wrap justify-end">${badges(c, false)}</span>
            </li>`;
        });
    }
}

// ---- INICIALIZAR ----
renderGlobalTorneo(DATA.globalTorneo);
Object.keys(DATA.divisiones).forEach(div => renderDivision(div, DATA.divisiones[div]));
</script>
</body>
</html>"""
    return html


if __name__ == '__main__':
    carpeta = os.path.dirname(os.path.abspath(__file__))
    datos   = procesar_todo(carpeta)
    html    = generar_html(datos)
    salida  = os.path.join(carpeta, 'index.html')
    with open(salida, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n✅ ¡Listo! Generado: {salida}")
    print(f"👉 Abre index.html con doble clic — funciona sin internet ni servidor.")
