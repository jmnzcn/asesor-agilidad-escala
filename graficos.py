# -*- coding: utf-8 -*-
"""Gráficos interactivos (Plotly) del resultado de la matriz."""
from __future__ import annotations

import plotly.graph_objects as go

import motor_matriz as motor

COLOR_MARCO = {"Nexus": "#4F7CFF", "LeSS": "#1FB57F", "SAFe": "#FF5A6E", "Disciplined Agile": "#9B6BFF"}
# Un color por criterio para el desglose (C1…C7)
COLOR_CRIT = ["#2EC4B6", "#5B6CFF", "#FFB020", "#8D99AE", "#FF6B9A", "#3FA7FF", "#9BD14B"]
CORTO_CRIT = {"C1": "Escala", "C2": "Producto", "C3": "Dependencias", "C4": "Gobernanza",
              "C5": "Portafolio", "C6": "Autonomía", "C7": "Heterogeneidad"}
FUENTE = dict(family="Inter, Segoe UI, system-ui, sans-serif", size=13)


def _base(fig: go.Figure, alto: int) -> go.Figure:
    fig.update_layout(
        height=alto, margin=dict(l=10, r=20, t=30, b=10), font=FUENTE,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        hoverlabel=dict(font=dict(family=FUENTE["family"], size=13), bordercolor="rgba(0,0,0,0)"),
        separators=",.",  # coma decimal
    )
    return fig


def _orden(r: dict) -> list[str]:
    """Recomendado arriba; luego elegibles por puntaje; excluidos al final."""
    return sorted(motor.MARCOS, key=lambda m: (m != r.get("recomendacion"), r["excluido"][m], -r["totales"][m]))


def puntaje_total(r: dict) -> go.Figure:
    orden = list(reversed(_orden(r)))  # Plotly dibuja de abajo hacia arriba
    rec = r.get("recomendacion")
    textos, hovers = [], []
    for m in orden:
        t = r["totales"][m]
        estado = "excluido" if r["excluido"][m] else ("★ recomendado" if m == rec else "elegible")
        textos.append(f"<b>{motor._fmt(t)}</b>" + ("  excluido" if r["excluido"][m] else ""))
        brecha = ""
        if rec and m != rec and not r["excluido"][m]:
            brecha = f"<br>A {motor._fmt(r['totales'][rec] - t)} de {rec}"
        motivo = ""
        if r["auto_excl"][m]:
            motivo = f"<br>R3: compatibilidad 0 en {', '.join(r['auto_excl'][m])}"
        elif m in r["excl_manual"]:
            motivo = f"<br>R4: {r['excl_manual'][m]}"
        hovers.append(f"<b>{m}</b><br>{motor._fmt(t)} de 5,00 · {estado}{brecha}{motivo}")
    fig = go.Figure(go.Bar(
        y=[("<b>★ " + m + "</b>") if m == rec else m for m in orden],
        x=[r["totales"][m] for m in orden], orientation="h",
        marker=dict(color=[COLOR_MARCO[m] for m in orden],
                    opacity=[0.35 if r["excluido"][m] else 1 for m in orden],
                    pattern=dict(shape=["/" if r["excluido"][m] else "" for m in orden], fgcolor="rgba(255,255,255,0.55)", size=7),
                    line=dict(width=[3 if m == rec else 0 for m in orden], color="rgba(255,255,255,0.9)"),
                    cornerradius=10),
        text=textos, textposition="outside", cliponaxis=False,
        hovertext=hovers, hoverinfo="text", width=0.62,
    ))
    fig.add_vline(x=5, line=dict(color="rgba(140,150,170,0.6)", width=1, dash="dot"))
    fig.add_annotation(x=5, y=1.04, yref="paper", text="máx. 5,00", showarrow=False, font=dict(size=11, color="#8C96AA"))
    fig.update_xaxes(range=[0, 5.9], tickvals=[0, 1, 2, 3, 4, 5], gridcolor="rgba(140,150,170,0.18)", zeroline=False,
                     title=dict(text="Puntaje ponderado", font=dict(size=12)))
    fig.update_yaxes(tickfont=dict(size=14))
    return _base(fig, 300)


def desglose(r: dict) -> go.Figure:
    orden = list(reversed(_orden(r)))
    etiquetas = [motor.CORTO[m] if m != "Disciplined Agile" else "Disciplined Agile" for m in orden]
    fig = go.Figure()
    for i, f in enumerate(r["filas"]):
        custom = [[f["criterio"], f["nivel"], motor._num(f["puntos"][m]), f"{f['peso'] * 100:.0f}", f["estado"], motor._fmt(f["aportes"][m])]
                  for m in orden]
        fig.add_trace(go.Bar(
            name=f"{f['id']} {CORTO_CRIT.get(f['id'], '')}", y=etiquetas, x=[f["aportes"][m] for m in orden], orientation="h",
            marker=dict(color=COLOR_CRIT[i % len(COLOR_CRIT)], line=dict(width=1, color="rgba(255,255,255,0.35)")),
            customdata=custom,
            hovertemplate=("<b>%{y} · " + f["id"] + "</b> %{customdata[0]}<br>Nivel: %{customdata[1]} (%{customdata[4]})"
                           "<br>%{customdata[2]} × %{customdata[3]} % = <b>%{customdata[5]}</b><extra></extra>"),
        ))
    for m, et in zip(orden, etiquetas):
        fig.add_annotation(x=r["totales"][m], y=et, text=f"<b>{motor._fmt(r['totales'][m])}</b>", xanchor="left",
                           xshift=8, showarrow=False, font=dict(size=13))
    fig.update_layout(barmode="stack", bargap=0.38,
                      legend=dict(orientation="h", y=-0.14, yanchor="top", x=0, title=None, font=dict(size=12), traceorder="normal"))
    fig.update_xaxes(range=[0, 5.6], tickvals=[0, 1, 2, 3, 4, 5], gridcolor="rgba(140,150,170,0.18)", zeroline=False)
    _base(fig, 400)
    fig.update_layout(margin=dict(l=10, r=20, t=10, b=10))
    return fig


def perfil(r: dict) -> go.Figure:
    ejes = [f"{f['id']} {CORTO_CRIT.get(f['id'], f['id'])}" for f in r["filas"]]
    rec = r.get("recomendacion")
    fig = go.Figure()
    for m in sorted(motor.MARCOS, key=lambda x: x == rec):  # recomendado al final, queda encima
        valores = [f["puntos"][m] for f in r["filas"]]
        fig.add_trace(go.Scatterpolar(
            r=valores + valores[:1], theta=ejes + ejes[:1], name=m + (" ★" if m == rec else ""),
            fill="toself", fillcolor=_rgba(COLOR_MARCO[m], 0.28 if m == rec else 0.08),
            line=dict(color=COLOR_MARCO[m], width=3.5 if m == rec else 1.6, dash="dot" if r["excluido"][m] else "solid"),
            marker=dict(size=6 if m == rec else 4),
            hovertemplate="<b>" + m + "</b><br>%{theta}: %{r} / 5<extra></extra>",
        ))
    fig.update_layout(
        polar=dict(bgcolor="rgba(0,0,0,0)",
                   radialaxis=dict(range=[0, 5], tickvals=[1, 2, 3, 4, 5], gridcolor="rgba(140,150,170,0.25)", angle=90,
                                   tickfont=dict(size=10)),
                   angularaxis=dict(gridcolor="rgba(140,150,170,0.25)", tickfont=dict(size=11))),
        legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center"),
    )
    _base(fig, 440)
    fig.update_layout(margin=dict(l=70, r=70, t=30, b=10))
    return fig


def comparar(a: dict, b: dict, etiqueta_a: str, etiqueta_b: str) -> go.Figure:
    """Totales de dos ejecuciones de la bitácora, marco por marco (A atenuada, B sólida)."""
    marcos = list(motor.MARCOS)
    fig = go.Figure()
    for ej, et, opac in ((a, etiqueta_a, 0.4), (b, etiqueta_b, 1)):
        fig.add_trace(go.Bar(
            name=et, x=[motor.CORTO[m] if m != "Disciplined Agile" else "DA" for m in marcos],
            y=[ej["totales"][m] for m in marcos],
            marker=dict(color=[COLOR_MARCO[m] for m in marcos], opacity=opac, cornerradius=8,
                        pattern=dict(shape=["/" if ej["excluido"][m] else "" for m in marcos],
                                     fgcolor="rgba(255,255,255,0.55)", size=7)),
            text=[f"<b>{motor._fmt(ej['totales'][m])}</b>" + ("★" if m == ej["recomendacion"] else "") for m in marcos],
            textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{x}</b> · " + et + "<br>%{y:.2f} de 5,00<extra></extra>",
        ))
    fig.update_layout(barmode="group", bargap=0.3, bargroupgap=0.08,
                      legend=dict(orientation="h", y=-0.15, x=0, title=None))
    fig.update_yaxes(range=[0, 5.6], tickvals=[0, 1, 2, 3, 4, 5], gridcolor="rgba(140,150,170,0.18)", zeroline=False)
    return _base(fig, 320)


def _rgba(hexcolor: str, alfa: float) -> str:
    h = hexcolor.lstrip("#")
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{alfa})"
