# forecast_demo.py
import pandas as pd
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
from prophet import Prophet
from datetime import datetime

# ----------------------------
# Helpers: lectura y tipos
# ----------------------------
def _read_params(sht, start_cell="A1", max_rows=30):
    top = sht.range(start_cell).expand("table").value
    if not top or not isinstance(top, list):
        return {}
    if not isinstance(top[0], list):
        top = [top]
    params = {}
    for row in top[:max_rows]:
        if not row or all(v is None for v in row):
            continue
        if len(row) < 2:
            continue
        k = str(row[0]).strip().lower()
        v = row[1]
        if k:
            params[k] = v
    return params

def _coerce_bool(v, default=False):
    if isinstance(v, bool): return v
    if v is None: return default
    s = str(v).strip().lower()
    if s in ("true","1","yes","si","sí"): return True
    if s in ("false","0","no"): return False
    return default

def _mk_df_from_selection(sel_vals):
    if sel_vals is None:
        raise ValueError("No se detectó selección.")
    vals = sel_vals
    if not isinstance(vals, list): vals = [vals]
    if vals and not isinstance(vals[0], list): vals = [vals]

    headers = [str(c).strip() if c is not None else "" for c in vals[0]]
    if all(isinstance(c, str) and c for c in headers):
        data = vals[1:]
    else:
        headers = ["Producto","Fecha","Ventas"][:len(vals[0])]
        data = vals

    df = pd.DataFrame(data, columns=headers)
    df.columns = [c.strip().lower() for c in df.columns]
    rename_map = {}
    for c in df.columns:
        if "product" in c or "producto" in c: rename_map[c] = "producto"
        elif "fecha" in c or "date" in c:    rename_map[c] = "fecha"
        elif "venta" in c or "sales" in c:   rename_map[c] = "ventas"
    df = df.rename(columns=rename_map)

    if "fecha" in df.columns:  df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    if "ventas" in df.columns: df["ventas"] = pd.to_numeric(df["ventas"], errors="coerce")
    df = df.dropna(subset=["fecha","ventas"]).sort_values("fecha").reset_index(drop=True)
    if "producto" not in df.columns: df["producto"] = "Producto A"
    return df[["producto","fecha","ventas"]]

def _freq_label(freq: str) -> str:
    if not freq: return "steps"
    f = str(freq).upper()
    if f.startswith("W"): return "W"
    if f in ("M","MS","ME"): return "M"
    return f

def _is_weekly(freq: str) -> bool:
    return str(freq).upper().startswith("W")

# ----------------------------
# Helpers: Prophet y gráficos
# ----------------------------
def _prophet_for(df_prod, yearly=True, weekly=False, periods=12, freq="W", sens=0.8):
    m = Prophet(
        weekly_seasonality=weekly,
        yearly_seasonality=yearly,
        changepoint_prior_scale=float(sens) if sens else 0.8
    )
    dfx = df_prod.rename(columns={"fecha":"ds","ventas":"y"}).copy()
    m.fit(dfx)
    futuro = m.make_future_dataframe(periods=int(periods), freq=freq)
    fc = m.predict(futuro)
    out_future = fc[["ds","yhat","yhat_lower","yhat_upper"]].tail(int(periods)).reset_index(drop=True)
    return m, fc, out_future

def _add_picture(sheet, fig, name, left_cell, top_cell, x_offset=0, y_offset=0):
    for p in list(sheet.pictures):
        if p.name == name: p.delete()
    sheet.pictures.add(
        fig, name=name,
        left=sheet.range(left_cell).left + x_offset,
        top=sheet.range(top_cell).top + y_offset
    )
    plt.close(fig)

# ----------------------------
# Piso de crecimiento (growth floor)
# ----------------------------
def _shift_ly_dates(dates, freq):
    """Fechas del mismo periodo del año pasado."""
    if str(freq).upper().startswith("W"):
        return pd.to_datetime(dates) - pd.DateOffset(weeks=52)
    else:
        return pd.to_datetime(dates) - pd.DateOffset(years=1)

def _prepare_hist_series(hist_df):
    """
    Series indexada por fecha, sin duplicados: suma por fecha y ordena.
    """
    if hist_df.empty:
        return pd.Series(dtype=float)

    df = hist_df.copy()
    val_col = "ventas_total" if "ventas_total" in df.columns else "ventas"
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.dropna(subset=["fecha", val_col])

    s = (df.groupby("fecha", as_index=True)[val_col]
           .sum()
           .sort_index())
    return s

def _calc_growth_floor(df_hist_total, freq, mode="blend",
                       growth_pct=0.0, momentum_window=8, min_growth_default=0.05):
    """
    Escalar g_floor (>= min_default si sale negativo).
    df_hist_total: ['fecha','ventas_total'].
    """
    s = df_hist_total.copy()
    if s.empty:
        return float(min_growth_default or 0.0)

    mode = (mode or "momentum").lower()

    if mode == "manual":
        g = float(growth_pct or 0.0)

    elif mode == "momentum":
        n = int(momentum_window or 8)
        mean_hist = s["ventas_total"].mean()
        mean_last = s["ventas_total"].tail(n).mean() if len(s)>=n else s["ventas_total"].mean()
        g = (mean_last / mean_hist - 1.0) if mean_hist else 0.0

    elif mode == "yoy":
        n = int(momentum_window or 8)
        lastN = s.tail(n).copy()
        lastN_ly_idx = _shift_ly_dates(lastN["fecha"], freq)
        s_idx = s.set_index("fecha")
        match_ly = s_idx.reindex(lastN_ly_idx)
        if match_ly["ventas_total"].notna().any():
            g = (lastN["ventas_total"].mean() / match_ly["ventas_total"].mean() - 1.0)
        else:
            g = 0.0

    elif mode == "cagr":
        shift = pd.DateOffset(weeks=52) if _is_weekly(freq) else pd.DateOffset(years=1)
        end_mean = s["ventas_total"].tail(8).mean()
        start_mean = s[s["fecha"] <= s["fecha"].max() - shift]["ventas_total"].tail(8).mean()
        g = (end_mean / start_mean - 1.0) if (start_mean and not np.isnan(start_mean)) else 0.0

    elif mode == "blend":
        g_m = _calc_growth_floor(s, freq, "momentum", momentum_window=momentum_window,
                                 min_growth_default=min_growth_default)
        g_y = _calc_growth_floor(s, freq, "yoy", momentum_window=momentum_window,
                                 min_growth_default=min_growth_default)
        g = 0.5*g_m + 0.5*g_y

    else:  # none
        g = 0.0

    if g < 0:
        g = float(min_growth_default or 0.0)
    return float(g)

def _apply_growth_floor_to_future(fut_df, hist_df, freq, g_floor):
    """
    Aplica piso: yhat = max(yhat, LY*(1+g_floor)) por fecha.
    Si no existe LY para una fecha, no se aplica (piso = -inf).
    """
    if fut_df.empty or hist_df.empty or g_floor is None:
        return fut_df

    s_hist = _prepare_hist_series(hist_df)
    if s_hist.empty:
        return fut_df

    ly_idx = _shift_ly_dates(fut_df["ds"], freq)
    ly_idx = pd.DatetimeIndex(pd.to_datetime(ly_idx))

    ly_vals = s_hist.reindex(ly_idx).to_numpy()
    floor_vals = np.where(np.isnan(ly_vals), -np.inf, ly_vals * (1.0 + float(g_floor)))

    fut = fut_df.copy()
    fut["yhat"] = np.maximum(fut["yhat"].values, floor_vals)
    if "yhat_lower" in fut.columns:
        fut["yhat_lower"] = np.minimum(fut["yhat"].values,
                                       np.maximum(fut["yhat_lower"].values, floor_vals))
    if "yhat_upper" in fut.columns:
        fut["yhat_upper"] = np.maximum(fut["yhat"].values,
                                       np.maximum(fut["yhat_upper"].values, floor_vals))
    return fut

# ----------------------------
# Forecast por producto y totales (por suma) + piso
# ----------------------------
def _forecast_all_products(
    df, yearly, weekly, periods, freq, sens,
    growth_mode="blend", growth_pct=0.0, momentum_window=8,
    min_growth_default=0.05, apply_growth="per_product"
):
    """
    Devuelve:
      per_prod: lista de dicts {'prod','hist','model','fc_full','out_future'}
      total_hist: DataFrame ['fecha','ventas_total']
      total_full: DataFrame ['ds','yhat','yhat_lower','yhat_upper'] (hist+futuro por suma)
      total_future: DataFrame ['ds','yhat','yhat_lower','yhat_upper'] (solo futuro)
      g_floor: escalar aplicado
    """
    productos = df["producto"].unique().tolist()
    per_prod = []
    full_concat = []
    fut_concat = []

    # histórico total simple
    total_hist = df.groupby("fecha", as_index=False)["ventas"].sum().sort_values("fecha")
    total_hist = total_hist.rename(columns={"ventas": "ventas_total"})

    # piso a partir del total
    g_floor = _calc_growth_floor(
        total_hist, freq, mode=growth_mode, growth_pct=growth_pct,
        momentum_window=momentum_window, min_growth_default=min_growth_default
    )

    for prod in productos:
        dfx = df[df["producto"] == prod].copy().reset_index(drop=True)
        m, fc_full, out_future = _prophet_for(
            dfx, yearly=yearly, weekly=weekly, periods=periods, freq=freq, sens=sens
        )

        # piso per_product (recomendado)
        if (apply_growth or "per_product").lower() == "per_product" and growth_mode != "none":
            out_future = _apply_growth_floor_to_future(out_future, dfx, freq, g_floor)

        per_prod.append({"prod": prod, "hist": dfx, "model": m, "fc_full": fc_full, "out_future": out_future})

        tmp_full = fc_full[["ds","yhat","yhat_lower","yhat_upper"]].copy(); tmp_full["prod"] = prod
        full_concat.append(tmp_full)

        tmp_fut = out_future.copy(); tmp_fut["prod"] = prod
        fut_concat.append(tmp_fut)

    # suma de TODOS los productos (full y futuro)
    if full_concat:
        conc_full = pd.concat(full_concat, ignore_index=True)
        total_full = conc_full.groupby("ds", as_index=False).agg(
            yhat=("yhat","sum"),
            yhat_lower=("yhat_lower","sum"),
            yhat_upper=("yhat_upper","sum"),
        )
    else:
        total_full = pd.DataFrame(columns=["ds","yhat","yhat_lower","yhat_upper"])

    if fut_concat:
        conc_fut = pd.concat(fut_concat, ignore_index=True)
        total_future = conc_fut.groupby("ds", as_index=False).agg(
            yhat=("yhat","sum"),
            yhat_lower=("yhat_lower","sum"),
            yhat_upper=("yhat_upper","sum"),
        )
    else:
        total_future = pd.DataFrame(columns=["ds","yhat","yhat_lower","yhat_upper"])

    # piso a nivel total (si se eligió total)
    if (apply_growth or "per_product").lower() == "total" and growth_mode != "none":
        total_future = _apply_growth_floor_to_future(total_future, total_hist, freq, g_floor)

    return per_prod, total_hist, total_full, total_future, g_floor

# --------------------- Comparativos de negocio (Resumen) ---------------------
def _comparativos_resumen(total_hist: pd.DataFrame, total_future: pd.DataFrame, periods: int, freq: str):
    mean_fcst = float(total_future["yhat"].mean()) if len(total_future) else np.nan
    mean_hist = float(total_hist["ventas_total"].mean()) if len(total_hist) else np.nan
    growth_hist = (mean_fcst/mean_hist - 1.0) if (mean_hist and not np.isnan(mean_hist)) else np.nan

    if _is_weekly(freq):
        window = 8; label_n = f"Promedio últimas {window} semanas"
    else:
        window = 3 if _freq_label(freq) == "M" else 8
        label_n = f"Promedio últimas {window} observaciones"
    base_lastN = float(total_hist["ventas_total"].tail(window).mean()) if len(total_hist) else np.nan
    growth_lastN = (mean_fcst/base_lastN - 1.0) if (base_lastN and not np.isnan(base_lastN)) else np.nan

    if len(total_future):
        shift = pd.DateOffset(weeks=52) if _is_weekly(freq) else pd.DateOffset(years=1)
        ly_dates = total_future["ds"] - shift
        hist_matched = total_hist.set_index("fecha").reindex(ly_dates.values)
        base_lastyear = float(hist_matched["ventas_total"].mean()) if len(hist_matched) else np.nan
    else:
        base_lastyear = np.nan
    growth_lastyear = (mean_fcst/base_lastyear - 1.0) if (base_lastyear and not np.isnan(base_lastyear)) else np.nan

    return [
        ["Promedio histórico",          mean_hist,     mean_fcst, growth_hist],
        [label_n,                       base_lastN,    mean_fcst, growth_lastN],
        ["Promedio mismo periodo LY",   base_lastyear, mean_fcst, growth_lastyear],
    ]

# --------------------- Hoja Resumen (gráfico híbrido) ----------------------
def _build_summary_sheet(wb, total_hist, total_full, total_future, periods, freq,
                         growth_note: str):
    ulabel = _freq_label(freq)
    try:
        wb.sheets["Resumen"].delete()
    except Exception:
        pass
    sht = wb.sheets.add("Resumen", after=wb.sheets[-1])

    # Título
    sht["A1"].value = f"Resumen total — {datetime.now():%Y-%m-%d %H:%M}"
    sht["A1"].font.bold = True; sht["A1"].font.size = 14

    # --------- Gráfica 1: híbrida (histórico + futuro ajustado) ----------
    img_row = 3
    fig1, ax1 = plt.subplots(figsize=(7, 3))

    if len(total_hist):
        ax1.scatter(total_hist["fecha"], total_hist["ventas_total"],
                    s=18, c="black", label="Histórico (puntos)")
        ax1.plot(total_hist["fecha"], total_hist["ventas_total"],
                 linewidth=2, label="Histórico (línea)")
        cut = total_hist["fecha"].max()
        ax1.axvline(cut, linestyle="--", linewidth=1, alpha=0.7)

    if len(total_future):
        ax1.fill_between(total_future["ds"], total_future["yhat_lower"], total_future["yhat_upper"],
                         alpha=0.25, label="Rango (fut.)", zorder=1)
        ax1.plot(total_future["ds"], total_future["yhat"],
                 linewidth=2, label="FCST base (ajustado)", zorder=2)

    ax1.set_title("Forecast total (histórico + pronóstico)", fontsize=12, fontweight="bold")
    ax1.set_xlabel("Fecha"); ax1.set_ylabel("Ventas")
    ax1.legend(loc="best")

    fig1.tight_layout()
    _add_picture(sht, fig1, "RES_FC_FULL",
                 left_cell=f"A{img_row}", top_cell=f"A{img_row}")

    # --------- Gráfica 2: solo futuro (barras, sin set_ticklabels) ----------
    fig2, ax2 = plt.subplots(figsize=(7, 3))
    if len(total_future):
        fut2 = total_future.copy()
        fut2["ds"] = pd.to_datetime(fut2["ds"])
        x = fut2["ds"].dt.strftime("%Y-%m-%d")
        y = fut2["yhat"]
        ax2.bar(x, y)
        ax2.tick_params(axis="x", rotation=45)
        ax2.set_title(f"Proyección total — futuro ({periods} {ulabel})", fontsize=12, fontweight="bold")
        ax2.set_xlabel("Fecha"); ax2.set_ylabel("Ventas proyectadas")
    fig2.tight_layout()
    _add_picture(sht, fig2, "RES_FC_FUT", left_cell=f"H{img_row}", top_cell=f"H{img_row}")

    # --------- Comparativos de negocio ----------
    sht["A12"].value = "Comparativos (promedios)"
    sht["A12"].font.bold = True
    comp_rows = _comparativos_resumen(total_hist, total_future, periods, freq)
    sht["A13"].options(index=False, header=False).value = [["Concepto","Promedio base","Promedio forecast","Crec. (fcst/base - 1)"]]
    sht["A14"].options(index=False, header=False).value = [
        [
            r[0],
            (None if pd.isna(r[1]) else round(r[1], 4)),
            (None if pd.isna(r[2]) else round(r[2], 4)),
            (None if pd.isna(r[3]) else round(r[3], 4)),
        ] for r in comp_rows
    ]

    # Nota de regla de crecimiento aplicada
    sht["A18"].value = growth_note

    # Tablas compactas
    sht["A20"].value = "Histórico total (suma de productos)"
    sht["A21"].options(index=False).value = total_hist
    sht["D20"].value = f"Proyección total (suma por producto) — {periods} {ulabel}"
    sht["D21"].options(index=False).value = total_future[["ds","yhat"]]

    sht.autofit()

# -------------------- Hoja por producto -------------------
def _build_products_sheet(wb, per_prod, periods, freq, yearly, weekly, sens):
    ulabel = _freq_label(freq)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    sht = wb.sheets.add(f"FCST_{ts}", after=wb.sheets[-1])

    # Header
    sht["A1"].value = "Parámetros usados"
    sht["A2"].options(index=False, header=False).value = [
        ["weekly seasonality", _coerce_bool(weekly)],
        ["yearly seasonality", _coerce_bool(yearly)],
        ["periods", periods],
        ["freq", freq],
        ["sensibility", sens],
    ]

    # Histórico global (selección)
    all_hist = pd.concat([i["hist"].assign(_p=i["prod"]) for i in per_prod], ignore_index=True)
    sht["A7"].value = "Histórico (selección)"
    sht["A8"].options(index=False).value = all_hist[["producto","fecha","ventas"]]

    # Bloques hacia la derecha
    base_col = 6; block_width = 8; chart_gap_y = 360
    for i, item in enumerate(per_prod):
        prod = item["prod"]; dfx = item["hist"]; m = item["model"]; fc = item["fc_full"]; out = item["out_future"]
        col0 = base_col + i * block_width; c_fc = col0 + 3

        sht.cells(7, col0).value = f"Histórico {prod}"
        sht.cells(7, col0).font.bold = True
        sht.cells(8, col0).options(index=False).value = dfx

        sht.cells(7, c_fc).value = f"Forecast {prod} ({periods} {_freq_label(freq)})"
        sht.cells(7, c_fc).font.bold = True
        sht.cells(8, c_fc).options(index=False).value = out

        # -------- Forecast completo AJUSTADO: dibujado manualmente ----------
        fc_adj = fc.copy()
        fc_adj["ds"] = pd.to_datetime(fc_adj["ds"]).dt.normalize()
        out_plot = out.copy()
        out_plot["ds"] = pd.to_datetime(out_plot["ds"]).dt.normalize()
        out_plot = out_plot.sort_values("ds")

        mask = fc_adj["ds"].isin(out_plot["ds"])
        fc_adj.loc[mask, ["yhat","yhat_lower","yhat_upper"]] = \
            out_plot[["yhat","yhat_lower","yhat_upper"]].to_numpy()

        fig1, ax1 = plt.subplots(figsize=(6.5, 3))
        ax1.scatter(dfx["fecha"], dfx["ventas"], s=18, c="black",
                    label="Histórico (puntos)", zorder=3)
        ax1.fill_between(out_plot["ds"], out_plot["yhat_lower"], out_plot["yhat_upper"],
                         alpha=0.25, label="Rango (fut.)", zorder=1)
        ax1.plot(fc_adj["ds"], fc_adj["yhat"], lw=2, label="FCST base (ajustado)", zorder=2)

        cut = dfx["fecha"].max()
        ax1.axvline(cut, ls="--", lw=1, alpha=0.7)

        ax1.set_title(f"Forecast completo – {prod}", fontsize=12, fontweight="bold")
        ax1.set_xlabel("Fecha"); ax1.set_ylabel("Ventas")
        ax1.legend(loc="best")

        ymax = float(max(out_plot["yhat_upper"].max(), dfx["ventas"].max()))
        ax1.set_ylim(bottom=0, top=ymax * 1.05)

        fig1.tight_layout()
        _add_picture(sht, fig1, f"FC_{i}_full",
                     left_cell=sht.cells(8, c_fc + 3).get_address(),
                     top_cell=sht.cells(8, c_fc + 3).get_address())

        # Barras de futuro (x string + rotación)
        fut = out_plot.copy()
        x = fut["ds"].dt.strftime("%Y-%m-%d"); y = fut["yhat"]
        fig2, ax2 = plt.subplots(figsize=(6.5, 3))
        ax2.bar(x, y)
        ax2.tick_params(axis="x", rotation=45)
        ax2.set_title(f"Predicción futura ({periods} {_freq_label(freq)}) – {prod}", fontsize=12, fontweight="bold")
        ax2.set_xlabel("Fecha"); ax2.set_ylabel("Ventas proyectadas")
        fig2.tight_layout()
        _add_picture(sht, fig2, f"FC_{i}_bars",
                     left_cell=sht.cells(8, c_fc + 3).get_address(),
                     top_cell=sht.cells(8, c_fc + 3).get_address(),
                     y_offset=chart_gap_y)
    sht.autofit()

# ---------------------------- Entry ------------------------------
@xw.sub
def run_from_selection():
    wb = xw.Book.caller()
    sht_in = wb.app.selection.sheet
    sel = wb.app.selection.value

    df = _mk_df_from_selection(sel)
    if df.empty:
        xw.apps.active.api.MsgBox("La selección no contiene datos válidos (Fecha/Ventas).")
        return

    raw = _read_params(sht_in, start_cell="A1", max_rows=30)
    yearly = _coerce_bool(raw.get("yearly seasonality", True))
    weekly = _coerce_bool(raw.get("weekly seasonality", False))
    periods = int(raw.get("periods", 12) or 12)
    freq = str(raw.get("freq", "W") or "W").strip()
    sens = float(raw.get("sensibility", 0.8) or 0.8)

    # --- parámetros de piso de crecimiento ---
    growth_mode = str(raw.get("growth_mode", "blend") or "blend").strip().lower()
    growth_pct = float(raw.get("growth_pct", 0.0) or 0.0)
    momentum_window = int(raw.get("momentum_window", 8) or 8)
    min_growth_default = float(raw.get("min_growth_default", 0.05) or 0.05)
    apply_growth = str(raw.get("apply_growth", "per_product") or "per_product").strip().lower()
    if apply_growth not in ("per_product","total"):
        apply_growth = "per_product"

    mode = str(raw.get("output_mode", "ambos") or "ambos").strip().lower()
    if mode not in ("resumen","productos","ambos"): mode = "ambos"

    # Forecast por SKU + suma total (full + futuro) + piso
    per_prod, total_hist, total_full, total_future, g_floor = _forecast_all_products(
        df, yearly=yearly, weekly=weekly, periods=periods, freq=freq, sens=sens,
        growth_mode=growth_mode, growth_pct=growth_pct, momentum_window=momentum_window,
        min_growth_default=min_growth_default, apply_growth=apply_growth
    )

    # Nota para el resumen
    growth_note = (
        f"Regla aplicada: el forecast futuro no puede ser menor a LY*(1 + {g_floor:.2%}). "
        f"Modo={growth_mode}, ventana={momentum_window}, mínimo por defecto={min_growth_default:.0%}, "
        f"aplicación={'por producto' if apply_growth=='per_product' else 'total'}."
        if growth_mode != "none" else
        "Sin regla de crecimiento mínima (growth_mode=none)."
    )

    if mode in ("ambos","resumen"):
        _build_summary_sheet(wb, total_hist=total_hist, total_full=total_full,
                             total_future=total_future, periods=periods, freq=freq,
                             growth_note=growth_note)
    if mode in ("ambos","productos"):
        _build_products_sheet(wb, per_prod=per_prod, periods=periods, freq=freq,
                              yearly=yearly, weekly=weekly, sens=sens)
