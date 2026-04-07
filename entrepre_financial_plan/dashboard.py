"""
InfoCursos 2025 — Portuguese Higher Education & Business Plan Dashboard
Run with:  python dashboard.py   (then open http://127.0.0.1:8050)
"""

import pathlib
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, callback

# ── paths ──────────────────────────────────────────────────────────────────────
XL = pathlib.Path(__file__).parent / "final_Infocursos_2025_3.xlsx"

# ── colour palette ─────────────────────────────────────────────────────────────
C = {
    "bg":       "#0f1117",
    "card":     "#1a1d27",
    "border":   "#2a2d3e",
    "txt":      "#e8eaf0",
    "muted":    "#8b92a9",
    "blue":     "#4f8ef7",
    "teal":     "#31c6b0",
    "amber":    "#f5a623",
    "red":      "#e05c5c",
    "purple":   "#9b7ff7",
    "green":    "#52c97f",
}
SCENARIOS   = ["conservative", "neutral", "optimistic"]
SC_COLORS   = [C["amber"], C["blue"], C["green"]]
YEARS       = [2027, 2028, 2029, 2030, 2031]

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=C["txt"], family="Inter, system-ui, sans-serif", size=12),
    xaxis=dict(gridcolor=C["border"], linecolor=C["border"]),
    yaxis=dict(gridcolor=C["border"], linecolor=C["border"]),
    margin=dict(l=40, r=20, t=36, b=40),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor=C["border"]),
)

# ═══════════════════════════════════════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════════════════════════════════════

def load_df(sheet, skiprows=4):
    return pd.read_excel(XL, sheet_name=sheet, skiprows=skiprows, engine="openpyxl")

# ── Education sheets ───────────────────────────────────────────────────────────
df_sex      = load_df("Inscritos por Sexo")
df_nat      = load_df("Inscritos por Nacionalidade")
df_age      = load_df("Inscritos por Idade")
df_num      = load_df("Numero Inscritos")
df_desemp   = load_df("Desemprego Registado")
df_taxa     = load_df("Taxa de Conclusao")
df_situ     = load_df("Situacao Apos 1 Ano")

# ── Aggregated helper sheets ───────────────────────────────────────────────────
df_gender_agg = load_df("gender", skiprows=0)
df_age_agg    = load_df("age", skiprows=0)

# ── Financial sheets ───────────────────────────────────────────────────────────
df_costs      = load_df("costs", skiprows=1)
df_is         = load_df("income statement", skiprows=0)
df_cash       = load_df("cash budget", skiprows=0)

# ── clean up column names ──────────────────────────────────────────────────────
def snake(s):
    return str(s).strip().lower().replace(" ", "_").replace("/", "_")

df_sex.columns    = [snake(c) for c in df_sex.columns]
df_nat.columns    = [snake(c) for c in df_nat.columns]
df_age.columns    = [snake(c) for c in df_age.columns]
df_num.columns    = [snake(c) for c in df_num.columns]
df_desemp.columns = [snake(c) for c in df_desemp.columns]
df_taxa.columns   = [snake(c) for c in df_taxa.columns]
df_situ.columns   = [snake(c) for c in df_situ.columns]

# ═══════════════════════════════════════════════════════════════════════════════
#  PRE-COMPUTED FIGURES
# ═══════════════════════════════════════════════════════════════════════════════

# ── KPI numbers ────────────────────────────────────────────────────────────────
total_enrolled  = int(df_sex["numeroinscritos"].sum(skipna=True))
total_male      = int(df_sex["numerohomens"].sum(skipna=True))
total_female    = int(df_sex["numeromulheres"].sum(skipna=True))
pct_female      = 100 * total_female / (total_male + total_female)

# nationality — sum from df_nat
nat_col_pt  = [c for c in df_nat.columns if "portuguese" in c or "portugues" in c.lower()]
nat_col_for = [c for c in df_nat.columns if "estrang" in c.lower() or "foreign" in c.lower()]
if nat_col_pt:
    n_pt  = int(df_nat[nat_col_pt[0]].sum(skipna=True))
    n_for = int(df_nat[nat_col_for[0]].sum(skipna=True)) if nat_col_for else 0
else:
    n_pt, n_for = 338283, 58439  # fallback to observed values

pct_foreign = 100 * n_for / (n_pt + n_for)
n_institutions = df_sex["estabelecimento"].nunique()

# avg completion on time
completao_col = "percentagemconclusaotempoesperado"
if completao_col in df_taxa.columns:
    avg_completion = df_taxa[completao_col].mean(skipna=True)
else:
    avg_completion = float("nan")

# ── Figure 1 – Gender donut ────────────────────────────────────────────────────
fig_gender = go.Figure(go.Pie(
    labels=["Female", "Male"],
    values=[total_female, total_male],
    hole=0.62,
    marker_colors=[C["purple"], C["blue"]],
    textfont_size=13,
))
fig_gender.update_layout(
    **PLOTLY_LAYOUT, title_text="Gender split", title_x=0.5,
    showlegend=True,
    annotations=[dict(text=f"{pct_female:.1f}%<br>female", x=0.5, y=0.5,
                      font_size=15, showarrow=False, font_color=C["txt"])],
)

# ── Figure 2 – Nationality donut ──────────────────────────────────────────────
fig_nat = go.Figure(go.Pie(
    labels=["Portuguese", "Foreign"],
    values=[n_pt, n_for],
    hole=0.62,
    marker_colors=[C["teal"], C["amber"]],
    textfont_size=13,
))
fig_nat.update_layout(
    **PLOTLY_LAYOUT, title_text="Nationality", title_x=0.5,
    showlegend=True,
    annotations=[dict(text=f"{pct_foreign:.1f}%<br>foreign", x=0.5, y=0.5,
                      font_size=15, showarrow=False, font_color=C["txt"])],
)

# ── Figure 3 – Teaching type donut ────────────────────────────────────────────
tipo_counts = df_sex.groupby("tipoensino")["numeroinscritos"].sum(skipna=True).reset_index()
tipo_counts.columns = ["tipo", "total"]
fig_tipo = go.Figure(go.Pie(
    labels=tipo_counts["tipo"],
    values=tipo_counts["total"],
    hole=0.62,
    marker_colors=[C["blue"], C["green"], C["amber"], C["red"]],
    textfont_size=13,
))
fig_tipo.update_layout(
    **PLOTLY_LAYOUT, title_text="Type of teaching", title_x=0.5, showlegend=True,
)

# ── Figure 4 – Age distribution bar ───────────────────────────────────────────
age_cols = [c for c in df_age.columns if "numero" in c and "anos" in c.lower()]
age_labels_map = {
    "numero1_18anos": "≤18", "numero19anos": "19", "numero20anos": "20",
    "numero21anos": "21", "numero22anos": "22", "numero23anos": "23",
    "numero24anos": "24", "numero25_27anos": "25-27",
    "numero28_30anos": "28-30", "numero31_39anos": "31-39",
    "numero40_99anos": "40+",
}
age_sums = {age_labels_map.get(c, c): df_age[c].sum(skipna=True) for c in age_cols if c in age_labels_map}
fig_age = go.Figure(go.Bar(
    x=list(age_sums.keys()), y=list(age_sums.values()),
    marker_color=C["blue"], opacity=0.85,
))
fig_age.update_layout(
    **PLOTLY_LAYOUT, title_text="Students by age (national total)", title_x=0,
    xaxis_title="Age bracket", yaxis_title="Enrolled students",
)

# ── Figure 5 – Top 20 institutions ────────────────────────────────────────────
top20 = (
    df_sex.groupby("estabelecimento")["numeroinscritos"]
    .sum(skipna=True)
    .nlargest(20)
    .sort_values()
    .reset_index()
)
# shorten names
top20["short"] = top20["estabelecimento"].str.replace(
    r"(Universidade|Instituto Politécnico|Escola Superior)", "", regex=True
).str.strip().str[:45]
fig_top20 = go.Figure(go.Bar(
    x=top20["numeroinscritos"], y=top20["short"],
    orientation="h",
    marker_color=C["teal"], opacity=0.85,
))
fig_top20.update_layout(
    **{**PLOTLY_LAYOUT, "margin": dict(l=230, r=20, t=36, b=40)},
    title_text="Top 20 institutions by enrollment (2023/24)", title_x=0,
    xaxis_title="Students", yaxis_title="", height=520,
)

# ── Figure 6 – Enrollment trend 2019→2024 ────────────────────────────────────
year_cols = [c for c in df_num.columns if c.startswith("201") or c.startswith("202")]
if year_cols:
    totals = df_num[year_cols].sum(skipna=True)
    # rename to display year
    labels = [c.replace("_", "/") for c in year_cols]
    fig_trend = go.Figure(go.Scatter(
        x=labels, y=totals.values,
        mode="lines+markers",
        line=dict(color=C["blue"], width=3),
        marker=dict(size=8, color=C["blue"]),
        fill="tozeroy", fillcolor=f"rgba(79,142,247,0.15)",
    ))
    fig_trend.update_layout(
        **PLOTLY_LAYOUT, title_text="Total enrollment trend", title_x=0,
        xaxis_title="Academic year", yaxis_title="Students enrolled",
    )
else:
    fig_trend = go.Figure()

# ── Figure 7 – Completion rate histogram ─────────────────────────────────────
if completao_col in df_taxa.columns:
    fig_completion = go.Figure(go.Histogram(
        x=df_taxa[completao_col].dropna(),
        nbinsx=30, marker_color=C["teal"], opacity=0.8,
        xbins=dict(start=0, end=100, size=5),
    ))
    fig_completion.update_layout(
        **PLOTLY_LAYOUT, title_text="On-time completion rate distribution", title_x=0,
        xaxis_title="Completion rate (%)", yaxis_title="Number of courses",
    )
else:
    fig_completion = go.Figure()

# ── Figure 8 – Top/bottom unemployment rates ──────────────────────────────────
desemp_col = "taxadesempregocurso" if "taxadesempregocurso" in df_desemp.columns else [c for c in df_desemp.columns if "taxa" in c.lower()][0]
df_d = df_desemp[["nomecurso", desemp_col]].dropna()
df_d = df_d.rename(columns={"nomecurso": "Course", desemp_col: "Unemployment %"})
df_d_worst = df_d.nlargest(15, "Unemployment %").sort_values("Unemployment %")
df_d_worst["short"] = df_d_worst["Course"].str[:50]
fig_desemp = go.Figure(go.Bar(
    x=df_d_worst["Unemployment %"], y=df_d_worst["short"],
    orientation="h",
    marker_color=C["red"], opacity=0.8,
))
fig_desemp.update_layout(
    **{**PLOTLY_LAYOUT, "margin": dict(l=280, r=20, t=36, b=40)},
    title_text="Top 15 courses by graduate unemployment rate",
    title_x=0, xaxis_title="Unemployment rate (%)", yaxis_title="", height=480,
)

# ── Figure 9 – Situation after 1 year (stacked by degree type) ───────────────
situ_cols = {
    "Graduated": "percentagemdiplomadoscurso",
    "Stayed same course": "percentageminscritosmesmoestabelecimentocurso",
    "Changed course": "percentagemoutrocursomesmoestabelecimento",
    "Left for another institution": "percentagemnoutroestabelecimento",
    "Not found in HE": "percentagemnaoencradosensinosuperior",
}
real_situ_cols = {k: v for k, v in situ_cols.items() if v in df_situ.columns}
if real_situ_cols:
    situ_grau = df_situ.groupby("grau")[list(real_situ_cols.values())].mean(skipna=True).reset_index()
    situ_grau["grau_short"] = situ_grau["grau"].str[:40]
    fig_situ = go.Figure()
    palette = [C["green"], C["blue"], C["teal"], C["amber"], C["red"]]
    for i, (label, col) in enumerate(real_situ_cols.items()):
        fig_situ.add_trace(go.Bar(
            name=label,
            x=situ_grau[col],
            y=situ_grau["grau_short"],
            orientation="h",
            marker_color=palette[i % len(palette)],
        ))
    fig_situ.update_layout(
        **{**PLOTLY_LAYOUT, "margin": dict(l=250, r=20, t=80, b=40),
           "legend": dict(bgcolor="rgba(0,0,0,0)", bordercolor=C["border"],
                          orientation="h", y=1.12, x=0)},
        barmode="stack",
        title_text="Student situation 1 year after enrollment (avg % by degree level)",
        title_x=0,
        xaxis_title="Average %", yaxis_title="", height=420,
    )
else:
    fig_situ = go.Figure()

# ── Figure 10 – Revenue by scenario (line) ────────────────────────────────────
revenue_data = {
    "conservative": [731.08, 7464.17, 18360.78, 29414.64, 43066.71],
    "neutral":      [1827.71, 18660.43, 45901.96, 73536.61, 107666.77],
    "optimistic":   [3655.42, 37320.87, 91803.92, 147073.22, 215333.53],
}
fig_revenue = go.Figure()
for sc, color in zip(SCENARIOS, SC_COLORS):
    fig_revenue.add_trace(go.Scatter(
        x=YEARS, y=revenue_data[sc],
        name=sc.capitalize(),
        mode="lines+markers",
        line=dict(color=color, width=2.5),
        marker=dict(size=7),
    ))
fig_revenue.update_layout(
    **PLOTLY_LAYOUT, title_text="Annual subscription revenue (€)", title_x=0,
    xaxis_title="Year", yaxis_title="Revenue (€)",
)

# ── Figure 11 – Net profit by scenario (line) ─────────────────────────────────
profit_data = {
    "conservative": [-5603.58, -2812.30, 6201.82, 11655.01, 19791.04],
    "neutral":      [-4591.94, 7516.25, 31608.56, 52357.52, 79384.59],
    "optimistic":   [-2905.88, 24730.50, 73953.12, 120195.04, 178707.18],
}
fig_profit = go.Figure()
for sc, color in zip(SCENARIOS, SC_COLORS):
    fig_profit.add_trace(go.Scatter(
        x=YEARS, y=profit_data[sc],
        name=sc.capitalize(),
        mode="lines+markers",
        line=dict(color=color, width=2.5),
        marker=dict(size=7),
    ))
fig_profit.add_hline(y=0, line_dash="dot", line_color=C["muted"])
fig_profit.update_layout(
    **PLOTLY_LAYOUT, title_text="Net profit by scenario (€)", title_x=0,
    xaxis_title="Year", yaxis_title="Net profit (€)",
)

# ── Figure 12 – Cost breakdown bar ────────────────────────────────────────────
cost_col_names   = [2027, 2028, 2029, 2030, 2031]
cost_category_col = [c for c in df_costs.columns if "cost" in str(c).lower() or "categ" in str(c).lower()]
if cost_category_col:
    df_c = df_costs[[cost_category_col[0]] + cost_col_names].dropna(subset=[cost_category_col[0]])
    df_c.columns = ["Category"] + [str(y) for y in cost_col_names]
    cost_col_names = [str(y) for y in cost_col_names]
    df_c = df_c[df_c["Category"].notna()]
    # filter only rows with numeric values
    def any_num(row):
        for v in row[1:]:
            try:
                if v is not None and float(v) > 0:
                    return True
            except (ValueError, TypeError):
                pass
        return False
    df_c = df_c[df_c.apply(any_num, axis=1)]
    fig_costs_chart = go.Figure()
    year_colors = [C["blue"], C["teal"], C["green"], C["amber"], C["purple"]]
    for yr, color in zip(cost_col_names, year_colors):
        vals = pd.to_numeric(df_c[yr], errors="coerce")
        fig_costs_chart.add_trace(go.Bar(
            name=str(yr), x=df_c["Category"].str[:40], y=vals.abs(),
            marker_color=color,
        ))
    fig_costs_chart.update_layout(
        **PLOTLY_LAYOUT, barmode="group",
        title_text="Operating cost breakdown by year (€)", title_x=0,
        xaxis_title="", yaxis_title="Cost (€)",
        xaxis_tickangle=-30,
    )
else:
    fig_costs_chart = go.Figure()

# ── Figure 13 – Cash position ─────────────────────────────────────────────────
cash_scenarios = {
    "Conservative": [6396.42, 584.12, 3785.95, 12440.96, 29231.99],
    "Neutral":      [6396.42, 14631.92, 43600.48, 93318.00, 171062.59],
    "Optimistic":   [6396.42, 30861.92, 101990.78, 219838.82, 395896.00],
}
fig_cash = go.Figure()
for (sc, vals), color in zip(cash_scenarios.items(), SC_COLORS):
    fig_cash.add_trace(go.Scatter(
        x=YEARS, y=vals, name=sc,
        mode="lines+markers",
        line=dict(color=color, width=2.5),
        marker=dict(size=7),
        fill="tozeroy" if sc == "Conservative" else None,
        fillcolor="rgba(245,166,35,0.08)",
    ))
fig_cash.update_layout(
    **PLOTLY_LAYOUT, title_text="Year-end cash position (€)", title_x=0,
    xaxis_title="Year", yaxis_title="Cash (€)",
)

# ═══════════════════════════════════════════════════════════════════════════════
#  LAYOUT HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def card(children, style_extra=None):
    style = {
        "background": C["card"],
        "border": f"1px solid {C['border']}",
        "borderRadius": "10px",
        "padding": "18px 20px",
    }
    if style_extra:
        style.update(style_extra)
    return html.Div(children, style=style)

def kpi(label, value, color=None):
    return card([
        html.P(label, style={"color": C["muted"], "fontSize": "12px", "margin": "0 0 4px 0", "textTransform": "uppercase", "letterSpacing": "0.08em"}),
        html.P(value,  style={"color": color or C["txt"], "fontSize": "26px", "fontWeight": "700", "margin": 0}),
    ], {"flex": "1", "minWidth": "160px"})

def section_title(text):
    return html.H3(text, style={"color": C["txt"], "fontWeight": "600", "fontSize": "15px",
                                "margin": "28px 0 10px 0", "borderLeft": f"3px solid {C['blue']}",
                                "paddingLeft": "10px"})

def row(*children, gap="16px"):
    return html.Div(list(children), style={"display": "flex", "gap": gap, "flexWrap": "wrap"})

def graph_card(fig, style_extra=None):
    s = {"flex": "1", "minWidth": "340px"}
    if style_extra:
        s.update(style_extra)
    return card(dcc.Graph(figure=fig, config={"displayModeBar": False}), s)

# ═══════════════════════════════════════════════════════════════════════════════
#  APP
# ═══════════════════════════════════════════════════════════════════════════════

app = Dash(__name__, title="InfoCursos 2025 Dashboard")

TAB_STYLE = {
    "background": C["card"],
    "color": C["muted"],
    "border": f"1px solid {C['border']}",
    "borderRadius": "6px 6px 0 0",
    "padding": "8px 18px",
    "fontSize": "13px",
    "fontWeight": "500",
}
TAB_SEL = {**TAB_STYLE, "background": C["blue"], "color": "#fff", "borderColor": C["blue"]}

PAGE_STYLE = {
    "fontFamily": "Inter, system-ui, sans-serif",
    "background": C["bg"],
    "minHeight": "100vh",
    "padding": "24px 32px",
    "color": C["txt"],
}

app.layout = html.Div([

    # ── header ────────────────────────────────────────────────────────────────
    html.Div([
        html.H1("InfoCursos 2025", style={"margin": 0, "fontSize": "22px", "fontWeight": "700", "color": C["txt"]}),
        html.P("Portuguese Higher Education Analytics · Business Plan 2027–2031",
               style={"margin": "2px 0 0 0", "color": C["muted"], "fontSize": "13px"}),
    ], style={"marginBottom": "24px"}),

    # ── tabs ──────────────────────────────────────────────────────────────────
    dcc.Tabs(id="tabs", value="edu", children=[
        dcc.Tab(label="📊  Education Overview",   value="edu",     style=TAB_STYLE, selected_style=TAB_SEL),
        dcc.Tab(label="🎓  Student Outcomes",     value="outcome", style=TAB_STYLE, selected_style=TAB_SEL),
        dcc.Tab(label="💼  Business Plan",        value="biz",     style=TAB_STYLE, selected_style=TAB_SEL),
    ], style={"marginBottom": "0"}),

    html.Div(id="tab-content"),

], style=PAGE_STYLE)


@callback(Output("tab-content", "children"), Input("tabs", "value"))
def render_tab(tab):

    # ── TAB 1 — Education Overview ────────────────────────────────────────────
    if tab == "edu":
        return html.Div([
            section_title("Key metrics  ·  Academic year 2023/24"),
            row(
                kpi("Total enrolled",        f"{total_enrolled:,}", C["blue"]),
                kpi("% female",              f"{pct_female:.1f}%",  C["purple"]),
                kpi("% foreign students",    f"{pct_foreign:.1f}%", C["amber"]),
                kpi("Institutions",          str(n_institutions),   C["teal"]),
                kpi("Avg on-time completion",
                    f"{avg_completion:.1f}%" if avg_completion == avg_completion else "–", C["green"]),
            ),

            section_title("Composition"),
            row(graph_card(fig_gender), graph_card(fig_nat), graph_card(fig_tipo)),

            section_title("Age distribution — all institutions"),
            card(dcc.Graph(figure=fig_age, config={"displayModeBar": False})),

            section_title("Enrollment trend — 2019/20 → 2023/24"),
            card(dcc.Graph(figure=fig_trend, config={"displayModeBar": False})),

            section_title("Top 20 institutions by total enrollment"),
            card(dcc.Graph(figure=fig_top20, config={"displayModeBar": False})),
        ])

    # ── TAB 2 — Student Outcomes ──────────────────────────────────────────────
    if tab == "outcome":
        return html.Div([
            section_title("Situation 1 year after first enrollment"),
            card(dcc.Graph(figure=fig_situ, config={"displayModeBar": False})),

            section_title("On-time completion rate distribution (all courses)"),
            card(dcc.Graph(figure=fig_completion, config={"displayModeBar": False})),

            section_title("Top 15 courses by registered unemployment rate"),
            card(dcc.Graph(figure=fig_desemp, config={"displayModeBar": False})),
        ])

    # ── TAB 3 — Business Plan ─────────────────────────────────────────────────
    if tab == "biz":
        fin_kpis_2031 = {
            "conservative": {"rev": 43_067, "profit": 19_791, "cash": 29_232},
            "neutral":      {"rev": 107_667, "profit": 79_385, "cash": 171_063},
            "optimistic":   {"rev": 215_334, "profit": 178_707, "cash": 395_896},
        }
        sc_colors_map = {"conservative": C["amber"], "neutral": C["blue"], "optimistic": C["green"]}
        return html.Div([
            section_title("2031 projections — by scenario"),
            row(*[
                card([
                    html.P(sc.capitalize(),
                           style={"color": sc_colors_map[sc], "fontWeight": "700", "fontSize": "14px",
                                  "margin": "0 0 10px 0", "textTransform": "uppercase", "letterSpacing": "0.06em"}),
                    html.Div([
                        html.Span("Revenue: ", style={"color": C["muted"], "fontSize": "12px"}),
                        html.Span(f"€{fin_kpis_2031[sc]['rev']:,}", style={"fontWeight": "600"}),
                    ], style={"marginBottom": "4px"}),
                    html.Div([
                        html.Span("Net profit: ", style={"color": C["muted"], "fontSize": "12px"}),
                        html.Span(f"€{fin_kpis_2031[sc]['profit']:,}", style={"fontWeight": "600", "color": C["green"]}),
                    ], style={"marginBottom": "4px"}),
                    html.Div([
                        html.Span("Cash position: ", style={"color": C["muted"], "fontSize": "12px"}),
                        html.Span(f"€{fin_kpis_2031[sc]['cash']:,}", style={"fontWeight": "600"}),
                    ]),
                ], {"flex": "1", "minWidth": "200px", "borderTop": f"3px solid {sc_colors_map[sc]}"})
                for sc in SCENARIOS
            ]),

            section_title("Revenue 2027–2031"),
            card(dcc.Graph(figure=fig_revenue, config={"displayModeBar": False})),

            section_title("Net profit 2027–2031"),
            card(dcc.Graph(figure=fig_profit, config={"displayModeBar": False})),

            row(
                graph_card(fig_cash, {"minWidth": "400px"}),
                graph_card(fig_costs_chart, {"minWidth": "400px"}),
            ),
        ])


# ───────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=False, port=8050)
