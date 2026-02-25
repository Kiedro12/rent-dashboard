import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from shiny import App, ui, reactive, render


candidates = []
env_path = os.environ.get("RENT_DATA_PATH")
if env_path:
    candidates.append(Path(env_path))

candidates += [
    Path("rent-poznan.xlsx"),
    Path("/content/rent-poznan.xlsx"),
    Path("/mnt/data/rent-poznan.xlsx"),
]

data_path = next((p for p in candidates if p.exists()), None)
if data_path is None:
    raise FileNotFoundError(
        "Nie znaleziono pliku rent-poznan.xlsx. "
        "Dodaj go do repo obok app.py albo ustaw zmienną środowiskową RENT_DATA_PATH."
    )

df = pd.read_excel(data_path)

DISTRICT_COL = "quarter"
DISTRICT_LABEL = "Dzielnica"

df["date_activ"] = pd.to_datetime(df.get("date_activ"), errors="coerce")

for c in ["price", "flat_rent", "flat_area", "flat_rooms"]:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")

if DISTRICT_COL in df.columns:
    df[DISTRICT_COL] = df[DISTRICT_COL].astype("string").str.strip()
    df.loc[df[DISTRICT_COL].isin(["", "nan", "NaN", "None"]), DISTRICT_COL] = pd.NA

if "price" not in df.columns:
    df["price"] = np.nan
if "flat_rent" not in df.columns:
    df["flat_rent"] = np.nan

df["total_cost"] = df["price"].fillna(0) + df["flat_rent"].fillna(0)
area = df.get("flat_area").replace(0, np.nan)
df["cost_per_sqm"] = df["total_cost"] / area
df["activ_month"] = df["date_activ"].dt.to_period("M").dt.to_timestamp()

ind = df.get("individual")
if ind is None:
    ind = pd.Series(False, index=df.index)
else:
    ind = ind.fillna(False)
df["seller_type"] = np.where(ind, "Prywatne", "Pośrednik")

rooms_order = ["1", "2", "3", "4"]

metric_meta = {
    "cost_per_sqm": {"label": "Stawka za m²", "unit": "PLN/m²/mies."},
    "total_cost": {"label": "Czynsz całkowity", "unit": "PLN/mies."},
}

AGG_LABEL = {"mean": "Średnia", "median": "Mediana"}


def agg_pl(k: str) -> str:
    return AGG_LABEL.get(k, k)


def tickfmt(metric: str) -> str:
    return ",.0f"


COLOR_LINE = "#2F6FED"
COLOR_BAR = "rgba(0,0,0,0.14)"
GRID_COL = "rgba(0,0,0,0.07)"

FIG_CONFIG = dict(
    displaylogo=False,
    responsive=True,
    displayModeBar=False,
    modeBarButtonsToRemove=[
        "select2d",
        "lasso2d",
        "autoScale2d",
        "toggleSpikelines",
        "hoverCompareCartesian",
        "hoverClosestCartesian",
    ],
)


def apply_business_layout(fig, height=560, margin=None):
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=margin or dict(l=26, r=26, t=74, b=30),
        font=dict(family="Arial", size=14, color="#1f2a44"),
        title=dict(x=0.02, xanchor="left", font=dict(size=22)),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1.0,
            bgcolor="rgba(255,255,255,0.0)",
        ),
        hovermode="closest",
        separators=", ",
    )
    fig.update_xaxes(showgrid=True, gridcolor=GRID_COL, zeroline=False, ticks="outside")
    fig.update_yaxes(showgrid=True, gridcolor=GRID_COL, zeroline=False, ticks="outside")
    return fig


def fig_html(fig):
    return ui.HTML(fig.to_html(full_html=False, include_plotlyjs="cdn", config=FIG_CONFIG))


if "flat_area" in df.columns and df["flat_area"].notna().any():
    q01, q99 = df["flat_area"].dropna().quantile([0.01, 0.99]).tolist()
    area_min = round(float(np.floor(q01 * 2) / 2), 1)
    area_max = round(float(np.ceil(q99 * 2) / 2), 1)
else:
    area_min, area_max = 10.0, 120.0

if "flat_rooms" in df.columns and df["flat_rooms"].notna().any():
    rooms_min = int(df["flat_rooms"].dropna().min())
    rooms_max = int(df["flat_rooms"].dropna().quantile(0.99))
    rooms_max = max(rooms_max, rooms_min)
else:
    rooms_min, rooms_max = 1, 6

districts_all = sorted(df[DISTRICT_COL].dropna().unique().tolist()) if DISTRICT_COL in df.columns else []
default_districts = df[DISTRICT_COL].value_counts().head(10).index.tolist() if DISTRICT_COL in df.columns else []


app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.tags.style(
            r"""
          .bslib-sidebar-layout .sidebar { padding: 10px !important; }
          .sidebar h4 { font-size: 1.05rem; margin: 4px 0 6px 0; }
          .sidebar .shiny-input-container { margin-bottom: 10px !important; }
          .sidebar .form-label { font-weight: 600; font-size: 0.92rem; margin-bottom: 2px; }
          .sidebar .form-control, .sidebar .selectize-input { font-size: 0.95rem; }
          .selectize-control.multi .selectize-input { max-height: 78px; overflow-y: auto; }
          .selectize-input > div { margin: 2px 4px 2px 0; }
          .sidebar hr { margin: 8px 0; }
          .irs--shiny .irs-line { height: 6px; }
          .irs--shiny .irs-bar { height: 6px; }
          .irs--shiny .irs-handle { top: 18px; }
        """
        ),
        ui.h4("Dzielnice"),
        ui.input_select(
            "district_mode",
            "Zakres dzielnic",
            choices={"all": "Wszystkie", "top": "Top N (najwięcej ofert)", "custom": "Własny wybór"},
            selected="all",
        ),
        ui.panel_conditional(
            "input.district_mode == 'top'",
            ui.input_slider("district_top_n_filter", "Top N dzielnic", min=5, max=30, value=12, step=1),
        ),
        ui.panel_conditional(
            "input.district_mode == 'custom'",
            ui.input_action_button("district_clear", "Wyczyść wybór", class_="btn btn-outline-secondary btn-sm"),
            ui.input_selectize(
                "district_sel",
                "Wybierz dzielnice",
                choices=districts_all,
                selected=default_districts,
                multiple=True,
            ),
        ),
        ui.hr(),
        ui.h4("Metodologia"),
        ui.input_select(
            "metric",
            "Miara",
            choices={k: f'{v["label"]} ({v["unit"]})' for k, v in metric_meta.items()},
            selected="cost_per_sqm",
        ),
        ui.input_select(
            "agg",
            "Agregacja",
            choices={"median": "Mediana", "mean": "Średnia"},
            selected="median",
        ),
        ui.hr(),
        ui.panel_conditional(
            "input.view == 'trend'",
            ui.input_slider("trend_min_n", "Min. liczność miesiąca (n)", min=10, max=250, value=50, step=10),
        ),
        ui.panel_conditional(
            "input.view == 'ranking'",
            ui.input_slider("top_n", "Top N", min=5, max=25, value=15, step=1),
            ui.input_slider("rank_min_n", "Min. liczność dzielnicy (n)", min=10, max=500, value=50, step=10),
        ),
        ui.panel_conditional(
            "input.view == 'heatmap'",
            ui.input_slider("heat_top_n", "Top N dzielnic na mapie", min=5, max=20, value=12, step=1),
            ui.input_slider("heat_min_n", "Min. liczność w komórce (n)", min=10, max=200, value=40, step=10),
        ),
        ui.panel_conditional(
            "input.view == 'scatter'",
            ui.input_slider(
                "area_rng",
                "Metraż",
                min=area_min,
                max=area_max,
                value=(area_min, area_max),
                step=0.5,
                post=" m²",
            ),
            ui.input_slider(
                "rooms_rng",
                "Liczba pokoi",
                min=rooms_min,
                max=rooms_max,
                value=(rooms_min, rooms_max),
                step=1,
            ),
            ui.input_checkbox("scatter_facet_seller", "Rozdziel wg sprzedającego (2 panele)", value=True),
            ui.input_checkbox("scatter_median_line", "Pokaż trend mediany (krok 2 m²)", value=True),
        ),
        width=290,
    ),
    ui.navset_tab(
        ui.nav_panel("Trend w czasie", ui.output_ui("trend_plot"), value="trend"),
        ui.nav_panel("Ranking dzielnic", ui.output_ui("ranking_plot"), value="ranking"),
        ui.nav_panel("Relacja metraż × stawka", ui.output_ui("scatter_plot"), value="scatter"),
        ui.nav_panel("Struktura stawek", ui.output_ui("heatmap_plot"), value="heatmap"),
        id="view",
        selected="trend",
    ),
    title="Raport rynku najmu mieszkań - Poznań",
)


def server(input, output, session):
    def get_int(name: str, default: int) -> int:
        try:
            return int(getattr(input, name)())
        except Exception:
            return default

    def get_bool(name: str, default: bool) -> bool:
        try:
            return bool(getattr(input, name)())
        except Exception:
            return default

    def esc_html(s: str) -> str:
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    @reactive.effect
    def _clear_districts():
        _ = input.district_clear()
        if input.district_mode() != "custom":
            return
        session.send_input_message("district_sel", {"value": []})

    @reactive.calc
    def district_list():
        if DISTRICT_COL not in df.columns:
            return None
        mode = input.district_mode()
        if mode == "all":
            return districts_all
        if mode == "top":
            n = get_int("district_top_n_filter", 12)
            return df[DISTRICT_COL].value_counts(dropna=True).head(n).index.tolist()
        sel = input.district_sel()
        return list(sel) if sel else []

    @reactive.calc
    def filtered():
        if DISTRICT_COL not in df.columns:
            return df.copy()
        dl = district_list()
        if dl is None:
            return df.copy()
        if isinstance(dl, list) and len(dl) == 0:
            return df.iloc[0:0].copy()
        return df[df[DISTRICT_COL].isin(dl)].copy()

    @reactive.calc
    def monthly_with_n():
        metric = input.metric()
        d = filtered().dropna(subset=["activ_month", metric]).copy()
        if d.empty:
            return pd.DataFrame(columns=["activ_month", "value", "n"])
        g = d.groupby("activ_month", observed=True)
        val = g[metric].mean() if input.agg() == "mean" else g[metric].median()
        n = g.size()
        out = pd.DataFrame({"activ_month": val.index, "value": val.values, "n": n.values}).sort_values("activ_month")
        out["activ_month"] = pd.to_datetime(out["activ_month"], errors="coerce")
        return out.dropna(subset=["activ_month"])

    @reactive.calc
    def district_rank():
        metric = input.metric()
        d = filtered().dropna(subset=[DISTRICT_COL, metric]).copy()
        if d.empty:
            return pd.DataFrame(columns=[DISTRICT_COL, "value", "n"])
        g = d.groupby(DISTRICT_COL, observed=True)
        val = g[metric].mean() if input.agg() == "mean" else g[metric].median()
        n = g.size()
        out = pd.DataFrame({DISTRICT_COL: val.index, "value": val.values, "n": n.values})
        out = out[out["n"] >= get_int("rank_min_n", 50)].copy()
        return out.sort_values("value", ascending=False).head(get_int("top_n", 15))

    @reactive.calc
    def scatter_data():
        d = filtered().dropna(subset=["flat_area", "flat_rooms", "cost_per_sqm", "seller_type"]).copy()
        if d.empty:
            return d
        a0, a1 = input.area_rng()
        r0, r1 = input.rooms_rng()
        d = d[(d["flat_area"].between(a0, a1)) & (d["flat_rooms"].between(r0, r1))].copy()
        if len(d) > 400:
            lo, hi = d["cost_per_sqm"].quantile([0.01, 0.99])
            d = d[d["cost_per_sqm"].between(lo, hi)]
        return d

    @reactive.calc
    def heatmap_data():
        metric = input.metric()
        d = filtered().dropna(subset=[DISTRICT_COL, "flat_rooms", metric]).copy()
        if d.empty:
            return None
        d = d[d["flat_rooms"].between(1, 4)].copy()
        if d.empty:
            return None
        d["rooms_grp"] = d["flat_rooms"].astype(int).astype(str)
        d["rooms_grp"] = pd.Categorical(d["rooms_grp"], categories=rooms_order, ordered=True)
        g = d.groupby([DISTRICT_COL, "rooms_grp"], observed=True)
        val = g[metric].mean() if input.agg() == "mean" else g[metric].median()
        n = g.size()
        out = pd.DataFrame(
            {
                DISTRICT_COL: val.index.get_level_values(0),
                "rooms_grp": val.index.get_level_values(1),
                "value": val.values,
                "n": n.values,
            }
        )
        top_n = get_int("heat_top_n", 12)
        keep = (
            out.groupby(DISTRICT_COL, observed=True)["n"]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index.tolist()
        )
        out = out[out[DISTRICT_COL].isin(keep)].copy()
        if out.empty:
            return None
        order = (
            out.groupby(DISTRICT_COL, observed=True)["value"]
            .mean()
            .sort_values(ascending=False)
            .index.tolist()
        )
        pivot_val = out.pivot(index=DISTRICT_COL, columns="rooms_grp", values="value").reindex(index=order, columns=rooms_order)
        pivot_n = out.pivot(index=DISTRICT_COL, columns="rooms_grp", values="n").reindex(index=order, columns=rooms_order)
        return metric, pivot_val, pivot_n

    @output
    @render.ui
    def trend_plot():
        try:
            metric = input.metric()
            meta = metric_meta[metric]
            agg_name = agg_pl(input.agg())
            m = monthly_with_n()
            if m.empty:
                return ui.HTML("<div style='padding:10px'>Brak danych po filtrach</div>")
            min_n = get_int("trend_min_n", 50)
            m2 = m[m["n"] >= min_n].copy()
            if m2.empty:
                return ui.HTML(f"<div style='padding:10px'>Brak miesięcy spełniających próg n ≥ {min_n}</div>")
            x = m2["activ_month"].dt.to_pydatetime().tolist()
            fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.06, row_heights=[0.65, 0.35])
            fig.add_trace(
                go.Scatter(
                    x=x,
                    y=m2["value"],
                    mode="lines+markers",
                    name=f'{meta["label"]} ({agg_name})',
                    line=dict(color=COLOR_LINE, width=2.6),
                    marker=dict(size=6),
                    hovertemplate="Miesiąc: %{x|%m.%Y}<br>Wartość: %{y:,.0f}<extra></extra>",
                ),
                row=1,
                col=1,
            )
            fig.add_trace(
                go.Bar(
                    x=x,
                    y=m2["n"],
                    name="Liczba ofert (n)",
                    marker_color=COLOR_BAR,
                    hovertemplate="Miesiąc: %{x|%m.%Y}<br>n=%{y}<extra></extra>",
                ),
                row=2,
                col=1,
            )
            fig.update_yaxes(title_text=f'{meta["label"]} ({meta["unit"]})', tickformat=tickfmt(metric), row=1, col=1)
            fig.update_yaxes(title_text="Liczba ofert (n)", rangemode="tozero", row=2, col=1)
            fig.update_xaxes(title_text="Miesiąc aktywacji ogłoszenia", tickformat="%m.%Y", dtick="M3", row=2, col=1)
            fig.update_layout(title=f'{meta["label"]} - {agg_name}, miesięcznie<br><sup>Pokazano miesiące z n ≥ {min_n}</sup>')
            apply_business_layout(fig, height=640, margin=dict(l=26, r=26, t=80, b=30))
            return fig_html(fig)
        except Exception as e:
            return ui.HTML(
                f"<div style='padding:10px;color:#b00020'><b>Błąd wykresu trendu:</b><br>"
                f"<pre style='white-space:pre-wrap;margin:8px 0 0 0'>{esc_html(e)}</pre></div>"
            )

    @output
    @render.ui
    def ranking_plot():
        try:
            metric = input.metric()
            meta = metric_meta[metric]
            agg_name = agg_pl(input.agg())
            r = district_rank()
            if r.empty:
                return ui.HTML("<div style='padding:10px'>Brak danych po filtrach / progu n</div>")
            fig = px.bar(
                r,
                x="value",
                y=DISTRICT_COL,
                orientation="h",
                title=f'{meta["label"]} - {agg_name} (Top {get_int("top_n", 15)})',
                hover_data={"n": True, "value": ":.0f"},
            )
            fig.update_yaxes(autorange="reversed", title=DISTRICT_LABEL)
            fig.update_xaxes(title=f'{meta["label"]} ({meta["unit"]})', tickformat=tickfmt(metric))
            fig.update_traces(
                marker_color="rgba(47,111,237,0.78)",
                text=r["value"].round(0).astype(int),
                texttemplate="%{text}",
                textposition="outside",
                cliponaxis=False,
            )
            apply_business_layout(fig, height=620, margin=dict(l=26, r=90, t=74, b=30))
            return fig_html(fig)
        except Exception as e:
            return ui.HTML(
                f"<div style='padding:10px;color:#b00020'><b>Błąd rankingu:</b><br>"
                f"<pre style='white-space:pre-wrap;margin:8px 0 0 0'>{esc_html(e)}</pre></div>"
            )

    @output
    @render.ui
    def scatter_plot():
        try:
            d = scatter_data()
            if d.empty:
                return ui.HTML("<div style='padding:10px'>Brak danych po filtrach</div>")

            def median_line_data(dd):
                bins = np.arange(dd["flat_area"].min(), dd["flat_area"].max() + 2, 2)
                b = pd.cut(dd["flat_area"], bins=bins, include_lowest=True)
                med = dd.groupby(b, observed=True)["cost_per_sqm"].median().dropna()
                x_mid = [iv.mid for iv in med.index]
                return x_mid, med.values

            facet = get_bool("scatter_facet_seller", True)

            if not facet:
                fig = go.Figure()
                fig.add_trace(
                    go.Histogram2d(
                        x=d["flat_area"],
                        y=d["cost_per_sqm"],
                        nbinsx=28,
                        nbinsy=24,
                        coloraxis="coloraxis",
                        hovertemplate="Powierzchnia: %{x}<br>Stawka: %{y}<br>Liczba ofert: %{z}<extra></extra>",
                    )
                )
                fig.update_layout(
                    title="Metraż × stawka za m² - gęstość ofert",
                    xaxis_title="Powierzchnia (m²)",
                    yaxis_title="Stawka za m² (PLN/m²/mies.)",
                    coloraxis=dict(colorscale="Blues", colorbar=dict(title="Liczba ofert")),
                )
                fig.update_xaxes(dtick=10)
                if get_bool("scatter_median_line", True):
                    x_mid, y_med = median_line_data(d)
                    fig.add_trace(
                        go.Scatter(
                            x=x_mid,
                            y=y_med,
                            mode="lines",
                            name="Trend mediany (2 m²)",
                            line=dict(color=COLOR_LINE, width=2.4),
                            hovertemplate="Metraż (bin): %{x:.1f} m²<br>Mediana: %{y:.0f}<extra></extra>",
                        )
                    )
                apply_business_layout(fig, height=640)
                return fig_html(fig)

            fig = make_subplots(
                rows=1,
                cols=2,
                shared_xaxes=True,
                shared_yaxes=True,
                subplot_titles=["Prywatne", "Pośrednik"],
                horizontal_spacing=0.06,
            )

            for i, seller in enumerate(["Prywatne", "Pośrednik"], start=1):
                dd = d[d["seller_type"] == seller].copy()
                if dd.empty:
                    continue

                fig.add_trace(
                    go.Histogram2d(
                        x=dd["flat_area"],
                        y=dd["cost_per_sqm"],
                        nbinsx=28,
                        nbinsy=24,
                        coloraxis="coloraxis",
                        hovertemplate="Powierzchnia: %{x}<br>Stawka: %{y}<br>Liczba ofert: %{z}<extra></extra>",
                        showlegend=False,
                    ),
                    row=1,
                    col=i,
                )

                if get_bool("scatter_median_line", True):
                    x_mid, y_med = median_line_data(dd)
                    fig.add_trace(
                        go.Scatter(
                            x=x_mid,
                            y=y_med,
                            mode="lines",
                            name="Trend mediany (2 m²)" if i == 1 else None,
                            showlegend=(i == 1),
                            line=dict(color=COLOR_LINE, width=2.4),
                            hovertemplate="Metraż (bin): %{x:.1f} m²<br>Mediana: %{y:.0f}<extra></extra>",
                        ),
                        row=1,
                        col=i,
                    )

            fig.update_layout(
                title="Metraż × stawka za m² - gęstość ofert (Prywatne vs Pośrednik)",
                xaxis_title="Powierzchnia (m²)",
                yaxis_title="Stawka za m² (PLN/m²/mies.)",
                coloraxis=dict(colorscale="Blues", colorbar=dict(title="Liczba ofert")),
            )
            fig.update_xaxes(dtick=10)
            apply_business_layout(fig, height=640, margin=dict(l=26, r=26, t=90, b=30))
            return fig_html(fig)
        except Exception as e:
            return ui.HTML(
                f"<div style='padding:10px;color:#b00020'><b>Błąd relacji:</b><br>"
                f"<pre style='white-space:pre-wrap;margin:8px 0 0 0'>{esc_html(e)}</pre></div>"
            )

    @output
    @render.ui
    def heatmap_plot():
        try:
            metric = input.metric()
            meta = metric_meta[metric]
            agg_name = agg_pl(input.agg())
            hm = heatmap_data()
            if hm is None:
                return ui.HTML("<div style='padding:10px'>Brak danych po filtrach</div>")

            _, pivot_val, pivot_n = hm
            min_n = get_int("heat_min_n", 40)

            z = pivot_val.values
            n_mat = pivot_n.fillna(0).astype(int).values

            has_data = n_mat > 0
            low_n = has_data & (n_mat < min_n)
            ok_n = has_data & (n_mat >= min_n)

            z_ok = np.where(ok_n, z, np.nan)
            z_low = np.where(low_n, 1, np.nan)

            cd_ok = np.where(ok_n, n_mat, 0)
            cd_low = np.where(low_n, n_mat, 0)

            fig = go.Figure()
            fig.add_trace(
                go.Heatmap(
                    z=z_ok,
                    x=pivot_val.columns.astype(str),
                    y=pivot_val.index.astype(str),
                    customdata=cd_ok,
                    hovertemplate=(
                        f"{DISTRICT_LABEL}: %{{y}}<br>"
                        "Pokoje: %{x}<br>"
                        "Wartość: %{z:,.0f}<br>"
                        "n=%{customdata}<extra></extra>"
                    ),
                    colorscale="Blues",
                    colorbar={"title": f'{meta["label"]} ({meta["unit"]})'},
                    xgap=1,
                    ygap=1,
                )
            )
            fig.add_trace(
                go.Heatmap(
                    z=z_low,
                    x=pivot_val.columns.astype(str),
                    y=pivot_val.index.astype(str),
                    customdata=cd_low,
                    hovertemplate=(
                        f"{DISTRICT_LABEL}: %{{y}}<br>"
                        "Pokoje: %{x}<br>"
                        f"Za mało danych (n < {min_n})<br>"
                        "n=%{customdata}<extra></extra>"
                    ),
                    colorscale=[[0, "rgba(210,210,210,0.75)"], [1, "rgba(210,210,210,0.75)"]],
                    showscale=False,
                    xgap=1,
                    ygap=1,
                )
            )

            fig.update_layout(
                title=(
                    f'{meta["label"]} - {agg_name} (1–4 pokoje)<br>'
                    f"<sup>Niebieskie: n ≥ {min_n} • Szare: n &lt; {min_n} • Brak koloru: brak danych</sup>"
                ),
                xaxis_title="Liczba pokoi",
                yaxis_title=f"{DISTRICT_LABEL} (Top N)",
            )
            apply_business_layout(fig, height=700)
            return fig_html(fig)
        except Exception as e:
            return ui.HTML(
                f"<div style='padding:10px;color:#b00020'><b>Błąd heatmapy:</b><br>"
                f"<pre style='white-space:pre-wrap;margin:8px 0 0 0'>{esc_html(e)}</pre></div>"
            )


app = App(app_ui, server)
