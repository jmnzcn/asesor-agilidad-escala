# -*- coding: utf-8 -*-
"""Gráficos del resultado de la matriz, en HTML/CSS (se insertan con st.html).

Diseño: el marco recomendado lleva el único color de acento y los demás van en gris,
así el color cuenta la historia (quién gana) y la identidad la da siempre el nombre.
El desglose usa la paleta categórica validada para daltonismo (7 criterios, orden fijo)
y el perfil es una matriz de calor con una sola tonalidad (azul) de 0 a 5.
"""
from __future__ import annotations

from html import escape

import motor_matriz as motor

CORTO_CRIT = {"C1": "Escala", "C2": "Producto", "C3": "Dependencias", "C4": "Gobernanza",
              "C5": "Portafolio", "C6": "Autonomía", "C7": "Heterogeneidad"}

# Paleta categórica (orden fijo C1…C7) validada con validate_palette.js: CVD ΔE ≥ 8,4 y visión normal ≥ 19 entre vecinos.
CRIT = {"claro": ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6"],
        "oscuro": ["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181", "#008300", "#9085e9"]}
# Texto dentro de un relleno categórico: blanco, salvo en amarillo.
TINTA_EN = {"#eda100": "#0b0b0b"}
# Rampa secuencial azul 0→5 para la matriz de compatibilidad (texto: oscuro en celdas claras).
RAMPA = {"claro": ["#f0efec", "#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#1c5cab"],
         "oscuro": ["#262624", "#0d366b", "#184f95", "#256abf", "#2a78d6", "#5598e7"]}
TEXTO_RAMPA = {"claro": ["#52514e", "#0b0b0b", "#0b0b0b", "#0b0b0b", "#ffffff", "#ffffff"],
               "oscuro": ["#898781", "#f0efec", "#f0efec", "#ffffff", "#ffffff", "#ffffff"]}

_CSS = """
<style>
.viz{--ink:#f0efec;--ink2:#c3c2b7;--muted:#898781;--grid:rgba(255,255,255,.07);--track:rgba(255,255,255,.045);
  --accent:#3987e5;--rest:#5d5c58;--hatch:rgba(255,255,255,.16);--tipbg:#f0efec;--tipink:#0b0b0b;--card:rgba(255,255,255,.025);
  color:var(--ink);font-family:"Source Sans Pro","Source Sans 3",Inter,system-ui,sans-serif;font-size:14px;line-height:1.35;
  margin:4px 0 6px}
.viz.claro{--ink:#0b0b0b;--ink2:#52514e;--muted:#898781;--grid:rgba(0,0,0,.08);--track:rgba(0,0,0,.045);
  --accent:#2a78d6;--rest:#b9b8b0;--hatch:rgba(0,0,0,.18);--tipbg:#1a1a19;--tipink:#f0efec;--card:rgba(0,0,0,.02)}
.viz *{box-sizing:border-box}
.viz .tip{position:relative}
.viz .tip:hover::after{content:attr(data-tip);position:absolute;z-index:20;left:50%;bottom:calc(100% + 8px);transform:translateX(-50%);
  white-space:pre-line;width:max-content;max-width:280px;background:var(--tipbg);color:var(--tipink);font-size:12.5px;
  padding:7px 10px;border-radius:8px;font-weight:400;line-height:1.4;box-shadow:0 6px 24px rgba(0,0,0,.28);pointer-events:none;text-align:left}
.viz .eje{display:grid;grid-template-columns:var(--izq) 1fr var(--der);gap:0 14px;font-size:11.5px;color:var(--muted);
  font-variant-numeric:tabular-nums;margin-top:6px}
.viz .eje .esc{position:relative;height:14px}
.viz .eje .esc span{position:absolute;transform:translateX(-50%)}
.viz .eje .esc span:first-child{transform:none}.viz .eje .esc span:last-child{transform:translateX(-100%)}
.viz .grid5{background-image:linear-gradient(to right,var(--grid) 1px,transparent 1px);background-size:20% 100%;
  background-position:-1px 0}
/* ranking */
.viz .fila{display:grid;grid-template-columns:var(--izq) 1fr var(--der);align-items:center;gap:0 14px;padding:9px 0;
  border-top:1px solid var(--grid)}
.viz .fila:first-child{border-top:0}
.viz .nom{font-weight:600;font-size:14.5px;min-width:0}
.viz .nom small{display:block;font-weight:400;font-size:12px;color:var(--muted);margin-top:1px}
.viz .val{text-align:right;font-weight:700;font-size:20px;font-variant-numeric:tabular-nums;letter-spacing:-.01em}
.viz .val small{display:block;font-weight:400;font-size:11.5px;color:var(--muted)}
.viz .pista{height:30px;border-radius:0 6px 6px 0;background-color:var(--track);display:flex;align-items:center}
.viz .barra{height:14px;border-radius:0 4px 4px 0;background:var(--rest);transition:filter .15s}
.viz .barra.gana{height:18px;background:var(--accent);box-shadow:0 0 0 1px rgba(255,255,255,.04),0 6px 18px -6px var(--accent)}
.viz .barra.excl{background:repeating-linear-gradient(135deg,var(--hatch) 0 2px,transparent 2px 7px);border:0}
.viz .fila:hover .barra{filter:brightness(1.12)}
.viz .chip{display:inline-block;font-size:10.5px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;
  padding:1px 7px;border-radius:999px;margin-left:6px;vertical-align:2px}
.viz .chip.a{background:color-mix(in srgb,var(--accent) 18%,transparent);color:var(--ink)}
.viz .chip.x{background:var(--track);color:var(--ink2)}
/* desglose */
.viz .ley{display:flex;flex-wrap:wrap;gap:6px 14px;margin:0 0 10px;font-size:12.5px;color:var(--ink2)}
.viz .ley span{display:inline-flex;align-items:center;gap:6px;cursor:default}
.viz .ley i{width:10px;height:10px;border-radius:3px;display:inline-block}
.viz .segs{display:flex;gap:2px;height:22px;align-items:stretch}
.viz .seg{border-radius:2px;transition:opacity .15s}
.viz .seg:last-child{border-radius:2px 4px 4px 2px}
.viz .segs:hover .seg{opacity:.55}.viz .segs .seg:hover{opacity:1}
/* matriz */
.viz table.mx{width:100%;border-collapse:separate;border-spacing:3px;font-variant-numeric:tabular-nums}
.viz .mx th{font-size:12px;font-weight:600;color:var(--ink2);text-align:center;padding:4px 2px 6px}
.viz .mx th.gana{color:var(--ink)}
.viz .mx th.gana::after{content:"";display:block;height:3px;border-radius:2px;background:var(--accent);margin:5px auto 0;width:60%}
.viz .mx td.cr{padding:6px 8px 6px 0;text-align:left;font-size:13px}
.viz .mx td.cr b{font-weight:600}.viz .mx td.cr small{display:block;color:var(--muted);font-size:11.5px}
.viz .mx td.pe{color:var(--muted);font-size:12px;text-align:right;padding-right:6px;white-space:nowrap}
.viz .mx td.c{height:44px;min-width:52px;border-radius:6px;text-align:center;font-weight:700;font-size:15px}
.viz .mx td.c.col{outline:2px solid var(--accent);outline-offset:-2px}
.viz .mx td.c.cero{background-image:repeating-linear-gradient(135deg,var(--hatch) 0 2px,transparent 2px 7px)!important}
.viz .mx tr.tot td{padding-top:8px;font-weight:700;font-size:15px;text-align:center;border-top:1px solid var(--grid)}
.viz .mx tr.tot td.cr{text-align:left;font-size:13px}
.viz .mx tr.tot td.pe{font-weight:400;font-size:12px;color:var(--muted);text-align:right}
.viz .fila.excl{opacity:.6}
.viz .escala{display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--muted);margin-top:8px}
.viz .escala i{width:22px;height:10px;border-radius:3px;display:inline-block}
/* comparación */
.viz .pista.cmp{height:30px;position:relative;display:block}
.viz .cmp .une{position:absolute;top:50%;height:2px;transform:translateY(-50%);background:var(--ink2);opacity:.45}
.viz .cmp .pt{position:absolute;top:50%;width:14px;height:14px;border-radius:50%;transform:translate(-50%,-50%)}
.viz .cmp .pt.a{background:transparent;border:2px solid var(--ink2)}
.viz .cmp .pt.b{background:var(--accent);box-shadow:0 0 0 3px color-mix(in srgb,var(--accent) 25%,transparent)}
.viz .delta{font-size:12px;color:var(--ink2);font-weight:600}
/* robustez */
.viz{--crit:#d03b3b}
.viz h4{font-size:14px;font-weight:700;margin:4px 0 2px}
.viz .nota{font-size:12.5px;color:var(--muted);margin:0 0 8px}
.viz .estado{display:flex;gap:10px;align-items:flex-start;padding:10px 12px;border-radius:8px;background:var(--card);margin:2px 0 12px;font-size:13.5px}
.viz .estado b{font-weight:700}
.viz .ico{flex:0 0 auto;width:20px;height:20px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:800;color:#fff}
.viz .ico.ok{background:#0ca30c}.viz .ico.no{background:#fab219;color:#0b0b0b}
.viz .pista.rb{position:relative;display:block;background-image:none}
.viz .rb .cero{position:absolute;top:-4px;bottom:-4px;width:2px;background:var(--ink2);opacity:.6;border-radius:1px}
.viz .rb .act{position:absolute;top:8px;height:14px;background:var(--rest);border-radius:4px}
.viz .rb .act.gana{background:var(--accent)}
.viz .rb .perd{position:absolute;top:8px;height:14px;border-radius:4px;
  background:repeating-linear-gradient(135deg,var(--hatch) 0 2px,transparent 2px 6px);box-shadow:inset 0 0 0 1px var(--hatch)}
.viz .rb .perd.cruza{background:repeating-linear-gradient(135deg,color-mix(in srgb,var(--crit) 70%,transparent) 0 2px,transparent 2px 6px);
  box-shadow:inset 0 0 0 1px var(--crit)}
.viz .rb .peor{position:absolute;top:4px;width:3px;height:22px;border-radius:2px;background:var(--ink);transform:translateX(-50%)}
.viz .val.chico{font-size:15px}
.viz .chip.w{background:color-mix(in srgb,var(--crit) 18%,transparent);color:var(--ink)}
.viz .rb .act.riesgo{background:color-mix(in srgb,var(--crit) 60%,var(--rest))}
.viz .rb .lim{position:absolute;top:2px;bottom:2px;width:0;border-left:2px dashed var(--ink2);opacity:.8}
</style>
"""


def _q(texto: str) -> str:
    """Escapa para un atributo HTML; los saltos de línea pasan al tooltip."""
    return escape(texto, quote=True).replace("\n", "&#10;")


def _pct(x: float) -> str:
    return f"{max(0.0, min(5.0, x)) / 5 * 100:.2f}%"


def _orden(r: dict) -> list[str]:
    """Recomendado arriba; luego elegibles por puntaje; excluidos al final."""
    return sorted(motor.MARCOS, key=lambda m: (m != r.get("recomendacion"), r["excluido"][m], -r["totales"][m]))


def _eje(izq: str, der: str) -> str:
    marcas = "".join(f'<span style="left:{i * 20}%">{i}</span>' for i in range(6))
    return f'<div class="eje" style="--izq:{izq};--der:{der}"><span></span><div class="esc">{marcas}</div><span></span></div>'


def _motivo(r: dict, m: str) -> str:
    if r["auto_excl"][m]:
        return f"R3 · compatibilidad 0 en {', '.join(r['auto_excl'][m])}"
    if m in r["excl_manual"]:
        return f"R4 · {r['excl_manual'][m]}"
    return ""


def puntaje_total(r: dict, tema: str = "oscuro") -> str:
    rec = r.get("recomendacion")
    filas = []
    for m in _orden(r):
        t, excl, gana = r["totales"][m], r["excluido"][m], m == rec
        if gana:
            sub, chip = "Recomendado", '<span class="chip a">★ 1.º</span>'
        elif excl:
            sub, chip = "Excluido · " + _motivo(r, m), '<span class="chip x">excluido</span>'
        else:
            sub, chip = (f"a {motor._fmt(r['totales'][rec] - t)} de {rec}" if rec else "elegible"), ""
        tip = f"{m}\n{motor._fmt(t)} de 5,00 · {sub}"
        clase = "gana" if gana else ("excl" if excl else "")
        filas.append(
            f'<div class="fila tip" data-tip="{_q(tip)}">'
            f'<div class="nom">{escape(m)}{chip}<small>{escape(sub)}</small></div>'
            f'<div class="pista grid5"><div class="barra {clase}" style="width:{_pct(t)}"></div></div>'
            f'<div class="val">{motor._fmt(t)}<small>de 5,00</small></div></div>')
    return (_CSS + f'<div class="viz {tema}" style="--izq:minmax(120px,30%);--der:64px">' + "".join(filas)
            + _eje("minmax(120px,30%)", "64px") + "</div>")


def desglose(r: dict, tema: str = "oscuro") -> str:
    colores = CRIT[tema]
    ids = [f["id"] for f in r["filas"]]
    ley = "".join(f'<span data-c="{c}"><i style="background:{colores[i]}"></i>{c} {CORTO_CRIT.get(c, "")}</span>'
                  for i, c in enumerate(ids))
    # pasar el cursor por un criterio de la leyenda resalta ese criterio en las cuatro barras
    resaltar = "".join(f'.viz:has(.ley [data-c="{c}"]:hover) .seg:not([data-c="{c}"]){{opacity:.18}}' for c in ids)
    filas = []
    for m in _orden(r):
        segs = []
        for i, f in enumerate(r["filas"]):
            a = f["aportes"][m]
            if a <= 0:
                continue
            tip = (f"{m} · {f['id']} {f['criterio']}\nNivel: {f['nivel']} ({f['estado']})\n"
                   f"{motor._num(f['puntos'][m])} × {f['peso'] * 100:.0f} % = {motor._fmt(a)}")
            segs.append(f'<div class="seg tip" data-c="{f["id"]}" data-tip="{_q(tip)}" '
                        f'style="flex:0 0 calc({_pct(a)} - 2px);background:{colores[i]}"></div>')
        sub = "Recomendado" if m == r.get("recomendacion") else ("Excluido · " + _motivo(r, m) if r["excluido"][m] else "")
        filas.append(f'<div class="fila{" excl" if r["excluido"][m] else ""}"><div class="nom">{escape(m)}<small>{escape(sub)}</small></div>'
                     f'<div class="pista grid5" style="background-color:transparent"><div class="segs" style="width:100%">'
                     f'{"".join(segs)}</div></div><div class="val">{motor._fmt(r["totales"][m])}</div></div>')
    return (_CSS + f"<style>{resaltar}</style>" + f'<div class="viz {tema}" style="--izq:minmax(120px,30%);--der:64px">'
            f'<div class="ley">{ley}</div>' + "".join(filas) + _eje("minmax(120px,30%)", "64px") + "</div>")


def perfil(r: dict, tema: str = "oscuro") -> str:
    """Matriz de calor: compatibilidad 0–5 de cada marco en cada criterio."""
    rampa, tinta = RAMPA[tema], TEXTO_RAMPA[tema]
    rec = r.get("recomendacion")
    cab = "".join(f'<th class="{"gana" if m == rec else ""}">{escape(motor.CORTO[m])}{" ★" if m == rec else ""}</th>'
                  for m in motor.MARCOS)
    filas = []
    for f in r["filas"]:
        celdas = []
        for m in motor.MARCOS:
            p = f["puntos"][m]
            i = max(0, min(5, round(p)))
            cero = p == 0 and f["excluyente"] and f["estado"] != "Faltante"
            tip = (f"{m} · {f['id']} {f['criterio']}\nNivel: {f['nivel']} ({f['estado']})\n"
                   f"Compatibilidad {motor._num(p)} de 5 · aporta {motor._fmt(f['aportes'][m])}"
                   + ("\nCriterio excluyente: excluye el marco (R3)" if cero else ""))
            celdas.append(f'<td class="c tip{" col" if m == rec else ""}{" cero" if cero else ""}" data-tip="{_q(tip)}" '
                          f'style="background-color:{rampa[i]};color:{tinta[i]}">{motor._num(p)}</td>')
        nivel = f"{f['valor']} equipos → {f['nivel']}" if f["id"] == "C1" and f["valor"] else f["nivel"]
        filas.append(f'<tr><td class="cr"><b>{f["id"]} {escape(CORTO_CRIT.get(f["id"], f["criterio"]))}</b>'
                     f'<small>{escape(nivel)} · {escape(f["estado"])}</small></td>'
                     f'<td class="pe">{f["peso"] * 100:.0f} %</td>{"".join(celdas)}</tr>')
    tot = "".join(f'<td>{motor._fmt(r["totales"][m])}</td>' for m in motor.MARCOS)
    filas.append(f'<tr class="tot"><td class="cr">Total ponderado</td><td class="pe">100 %</td>{tot}</tr>')
    escala = "".join(f'<i style="background:{c}" title="{i}"></i>' for i, c in enumerate(rampa))
    return (_CSS + f'<div class="viz {tema}"><table class="mx"><thead><tr><th></th><th></th>{cab}</tr></thead>'
            f'<tbody>{"".join(filas)}</tbody></table>'
            f'<div class="escala">0{escala}5 · rayado: 0 en un criterio excluyente</div></div>')


def datos_decisivos(r: dict, criterios: dict) -> list[tuple]:
    """Para cada criterio faltante o inferido: cuánto podría reducir, por sí solo, la ventaja del recomendado
    sobre algún competidor elegible si su nivel resultara otro. Devuelve (fila, reducción, competidor, nivel), de mayor a menor."""
    rec, out = r.get("recomendacion"), []
    if not rec:
        return out
    for f in r["filas"]:
        if f["estado"] not in ("Faltante", "Inferido"):
            continue
        mejor = None
        for x in motor.MARCOS:
            if x == rec or r["excluido"][x]:
                continue
            actual = f["puntos"][rec] - f["puntos"][x]
            for nivel, p in criterios[f["id"]]["niveles"].items():
                red = f["peso"] * (actual - (p[rec] - p[x]))
                if red > 0.005 and (mejor is None or red > mejor[0]):
                    mejor = (red, x, nivel)
        if mejor:
            out.append((f, *mejor))
    return sorted(out, key=lambda t: -t[1])


def robustez(r: dict, criterios: dict, tema: str = "oscuro") -> str:
    """Ventaja del recomendado sobre cada competidor: actual y en el peor caso de los datos faltantes (R7),
    y qué dato faltante o inferido podría, por sí solo, cambiar más el resultado."""
    rec = r.get("recomendacion")
    if not rec:
        return _CSS + f'<div class="viz {tema}"><p class="nota">Sin recomendación: los cuatro marcos quedaron excluidos (R9).</p></div>'
    rivales = [x for x in _orden(r) if x != rec and not r["excluido"][x]]
    peor = {x: r["peor_brecha"].get(x, {"brecha": r["totales"][rec] - r["totales"][x], "peor": r["totales"][rec] - r["totales"][x]})
            for x in rivales}
    sensible = r["robustez"].startswith("SENSIBLE")
    if not r.get("faltantes"):
        estado = ('<div class="estado"><span class="ico ok">✓</span><div><b>Robusta.</b> No hay datos faltantes: '
                  f'la ventaja de {escape(rec)} no depende de supuestos sin evidencia.</div></div>')
    elif sensible:
        cruzan = [x for x in rivales if peor[x]["peor"] <= 0]
        lista = cruzan[0] if len(cruzan) == 1 else ", ".join(cruzan[:-1]) + " y " + cruzan[-1]
        estado = ('<div class="estado"><span class="ico no">!</span><div><b>Provisional.</b> Según cómo se resuelvan los datos '
                  f'faltantes ({", ".join(r["faltantes"])}), podría{"n" if len(cruzan) > 1 else ""} superar a {escape(rec)}: '
                  f'{escape(lista)}. Conviene aclarar esos datos antes de decidir.</div></div>')
    else:
        estado = ('<div class="estado"><span class="ico ok">✓</span><div><b>Robusta.</b> Aun con el peor valor posible de '
                  f'{", ".join(r["faltantes"])}, {escape(rec)} sigue adelante de todos.</div></div>')
    # escala común: de min(0, peor caso) a la mayor brecha
    vals = [0.0] + [v for x in rivales for v in (peor[x]["brecha"], peor[x]["peor"])]
    lo, hi = min(vals), max(vals)
    lo, hi = lo - (hi - lo) * 0.06 if lo < 0 else 0.0, hi * 1.08 or 1.0
    pos = lambda v: f"{(v - lo) / (hi - lo) * 100:.2f}%"  # noqa: E731
    filas = []
    for x in rivales:
        b, w = peor[x]["brecha"], peor[x]["peor"]
        tip = (f"Ventaja de {rec} sobre {x}\nActual: {motor._fmt(b)}"
               + (f"\nPeor caso de los faltantes: {motor._fmt(w)}" if abs(w - b) > 0.005 else "\nNo cambia con los faltantes")
               + ("\nPodría perder la recomendación" if w <= 0 else ""))
        # sólido: la ventaja que se conserva aun en el peor caso; rayado: lo que se podría perder
        firme = max(0.0, min(w, b))
        perd = (f'<div class="perd{" cruza" if w <= 0 else ""}" style="left:{pos(min(w, b) if w < 0 else firme)};'
                f'width:calc({pos(b if w >= 0 else 0)} - {pos(min(w, b) if w < 0 else firme)})"></div>' if abs(w - b) > 0.005 else "")
        if w < 0:
            firme = b
        sub = "podría superar al recomendado" if w <= 0 else ("se mantiene detrás" if r.get("faltantes") else "")
        filas.append(
            f'<div class="fila tip" data-tip="{_q(tip)}"><div class="nom">{escape(rec)} vs {escape(x)}'
            + (' <span class="chip w">en riesgo</span>' if w <= 0 else "") + f'<small>{escape(sub)}</small></div>'
            f'<div class="pista rb"><div class="act gana" style="left:{pos(0)};width:calc({pos(firme)} - {pos(0)})"></div>{perd}'
            f'<div class="cero" style="left:{pos(0)}"></div>'
            + (f'<div class="peor" style="left:{pos(w)}"></div>' if abs(w - b) > 0.005 else "")
            + f'</div><div class="val chico">{motor._fmt(b)}<small>{"peor " + motor._fmt(w) if abs(w - b) > 0.005 else "sin cambio"}</small></div></div>')
    eje_v = "".join(f'<span style="left:{pos(t)}">{motor._fmt(t)}</span>' for t in sorted({round(lo, 2), 0.0, round(hi, 2)}))
    eje_html = (f'<div class="eje"><span></span><div class="esc">{eje_v}</div><span></span></div>' if rivales else "")
    parte1 = ("<h4>Ventaja del recomendado sobre cada marco</h4>"
              '<p class="nota">Barra: ventaja actual. Rayado y marca: hasta dónde podría bajar en el peor caso de los datos faltantes (R7). '
              "La línea vertical es el empate: si el rayado la cruza, ese marco podría ganar.</p>"
              + ("".join(filas) + eje_html if rivales else '<p class="nota">No hay otros marcos elegibles.</p>'))
    dec = datos_decisivos(r, criterios)
    if dec:
        tope = max(max(d[1] for d in dec), min((peor[x]["brecha"] for x in rivales), default=0)) * 1.1
        filas2 = []
        for f, red, x, nivel in dec[:5]:
            lim = peor.get(x, {}).get("brecha", 0)
            decide = red >= lim - 1e-9
            tip = (f"{f['id']} {f['criterio']} ({f['estado']}, hoy: {f['nivel']})\nSi fuera «{nivel}», la ventaja de {rec} "
                   f"sobre {x} bajaría {motor._fmt(red)} (ventaja actual {motor._fmt(lim)})"
                   + ("\nPor sí solo puede cambiar la recomendación" if decide else ""))
            filas2.append(
                f'<div class="fila tip" data-tip="{_q(tip)}"><div class="nom">{f["id"]} {escape(CORTO_CRIT.get(f["id"], f["criterio"]))}'
                + (' <span class="chip w">decisivo</span>' if decide else "")
                + f'<small>{escape(f["estado"])} · peor caso: «{escape(nivel)}» favorece a {escape(x)}</small></div>'
                f'<div class="pista rb"><div class="act{" riesgo" if decide else ""}" style="left:0;width:{red / tope * 100:.2f}%"></div>'
                f'<div class="lim" style="left:{min(lim / tope, 1) * 100:.2f}%"></div></div>'
                f'<div class="val chico">−{motor._fmt(red)}<small>vs {escape(motor.CORTO[x])}</small></div></div>')
        parte2 = ('<h4 style="margin-top:16px">Qué dato decide</h4>'
                  '<p class="nota">Cuánto podría bajar la ventaja si un dato faltante o inferido resultara en su nivel menos favorable. '
                  "La línea discontinua es la ventaja actual sobre ese marco: si la barra la alcanza, ese dato por sí solo puede cambiar la recomendación. "
                  "Son los datos que conviene preguntar primero.</p>" + "".join(filas2))
    else:
        parte2 = ('<h4 style="margin-top:16px">Qué dato decide</h4><p class="nota">Ningún dato faltante o inferido puede reducir la ventaja: '
                  "todos los criterios que afectan el resultado vienen explícitos en el caso.</p>")
    return _CSS + f'<div class="viz {tema}" style="--izq:minmax(120px,34%);--der:64px">{estado}{parte1}{parte2}</div>'


def comparar(a: dict, b: dict, etiqueta_a: str, etiqueta_b: str, tema: str = "oscuro") -> str:
    """Dos ejecuciones de la bitácora, marco por marco: anillo = A, punto de acento = B."""
    filas = []
    for m in motor.MARCOS:
        ta, tb = a["totales"][m], b["totales"][m]
        lo, hi = sorted((ta, tb))
        d = tb - ta
        flecha = "▲" if d > 0.005 else ("▼" if d < -0.005 else "=")
        tip = (f"{m}\n{etiqueta_a}: {motor._fmt(ta)}{' (excluido)' if a['excluido'][m] else ''}\n"
               f"{etiqueta_b}: {motor._fmt(tb)}{' (excluido)' if b['excluido'][m] else ''}")
        gana_b = m == b["recomendacion"]
        sub = " · ".join(x for x in ("★ recomendado en B" if gana_b else "", "excluido en B" if b["excluido"][m] else "") if x)
        filas.append(
            f'<div class="fila tip" data-tip="{_q(tip)}"><div class="nom">{escape(m)}<small>{escape(sub)}</small></div>'
            f'<div class="pista cmp grid5"><div class="une" style="left:{_pct(lo)};width:calc({_pct(hi)} - {_pct(lo)})"></div>'
            f'<div class="pt a" style="left:{_pct(ta)}"></div><div class="pt b" style="left:{_pct(tb)}"></div></div>'
            f'<div class="val">{motor._fmt(tb)}<small class="delta">{flecha} {"" if abs(d) < 0.005 else motor._fmt(abs(d))}</small></div></div>')
    ley = (f'<div class="ley"><span><i style="border:2px solid var(--ink2);border-radius:50%"></i>{escape(etiqueta_a)}</span>'
           f'<span><i style="background:var(--accent);border-radius:50%"></i>{escape(etiqueta_b)}</span></div>')
    return (_CSS + f'<div class="viz {tema}" style="--izq:minmax(120px,30%);--der:64px">{ley}' + "".join(filas)
            + _eje("minmax(120px,30%)", "64px") + "</div>")
