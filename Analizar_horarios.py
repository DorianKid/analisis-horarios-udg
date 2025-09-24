# Calcular_mejor_horario_por_semestres_y_dia.py
from pathlib import Path
import pandas as pd
import os
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.patches import Patch
import win32com.client as win32

from constantes import SEMESTRES_POR_CARRERA  # Debe mapear carrera -> {semestre: [claves_materia,...]}

DIAS_MAP = {'L':'Lunes','M':'Martes','I':'Miércoles','J':'Jueves','V':'Viernes','S':'Sábado'}
DIAS_ORDEN = ['L','M','I','J','V','S']
SEMESTRES = {1:"Primero", 2:"Segundo", 3:"Tercero", 4:"Cuarto",5:"Quinto", 6: "Sexto", 7:"Séptimo", 8:"Octavo", 9:"Noveno", 10:"Décimo"}

# Parámetros de negocio
EXCLUIR_7_AM = True                          # Excluir 07:00 como candidato
CANDIDATOS_NONES = [700, 900, 1100, 1300, 1500]  # Puedes extender (1700, etc.) si aplica

# -----------------Helpers -------------------
def nombre_excel_por_dia(filename_base: str, dia_abrev: str) -> str:
    parts = filename_base.split('_', 1)
    prefix = parts[0] if parts else "HORARIOS"
    rest   = parts[1] if len(parts) > 1 else ""
    dia_upper = DIAS_MAP[dia_abrev].upper()
    return f"{prefix}_{dia_upper}_{rest}" if rest else f"{prefix}_{dia_upper}"

def hhmm_int_to_str(x: int) -> str:
    s = f"{int(x):04d}"
    return f"{s[:2]}:{s[2:]}"

def to_minutes(hhmm: int) -> int:
    h = hhmm // 100
    m = hhmm % 100
    return h*60 + m

def hhmm_to_am_pm(val) -> str:
    s = str(val).strip()
    # normaliza: "15:00" -> "1500", "7:05" -> "705"
    if ":" in s:
        h, m = s.split(":")
        s = int(h)*100 + int(m)
    else:
        s = int(s)

    h = s // 100
    m = s % 100
    ampm = "am" if h < 12 else "pm"
    h12 = ((h - 1) % 12) + 1  # 0->12, 13->1, etc.

    return f"{h12} {ampm}" if m == 0 else f"{h12}:{m:02d} {ampm}"

def AutoFit_columns_width(ruta_archivo):
    # Crear una instancia de la aplicación de Excel
    xlApp = win32.Dispatch('Excel.Application')

    # Abrir el libro de Excel
    wb = xlApp.Workbooks.Open(ruta_archivo)

    # Iterar sobre todas las hojas del libro
    for ws in wb.Worksheets:
        # Ajustar automáticamente el ancho de todas las columnas en la hoja actual
        ws.Columns.AutoFit()

    # Guardar los cambios y cerrar el libro
    wb.Save()
    wb.Close()

    # Cerrar la aplicación de Excel
    xlApp.Quit()

# ----------------- Métricas por tiempo -----------------
def activos_en_t(df_sem_dia: pd.DataFrame, t_min: int) -> int:
    """Suma 'Alumnos' de clases activas en t: inicio_min <= t < fin_min"""
    if df_sem_dia.empty:
        return 0
    ini = df_sem_dia['_ini_min']
    fin = df_sem_dia['_fin_min']
    return int(df_sem_dia[(ini <= t_min) & (t_min < fin)]['Alumnos'].sum())

def base_en_ventana(df_sem_dia: pd.DataFrame, t_min: int, delta_min: int = 240) -> int:
    """
    Máximo simultáneo del semestre dentro de la ventana [t_min - delta_min, t_min).
    """
    if df_sem_dia.empty:
        return 0

    win_lo = t_min - delta_min
    win_hi = t_min

    # Puntos de cambio (inicios/finales) restringidos a la ventana
    puntos = sorted(set(df_sem_dia['Hora inicio'].tolist() + df_sem_dia['Hora fin'].tolist()))
    tmins = []
    for h in puntos:
        try:
            h = int(h)
        except Exception:
            continue
        m = (h // 100) * 60 + (h % 100)
        if win_lo <= m < win_hi:
            tmins.append(m)

    # Si la ventana no incluye puntos de cambio, muestrea algunos momentos
    if not tmins:
        tmins = [win_lo, (win_lo + win_hi) // 2, win_hi - 1]

    vals = [activos_en_t(df_sem_dia, tau) for tau in tmins]
    return max(vals) if vals else 0

def plot_top5_bar(df, hora_col_name, libres_col_name, outdir: Path,
                  n_top=5,
                  color_top1="#2563eb",   # azul vivo
                  color_top2="#10b981",   # verde
                  color_rest="#9ca3af"):  # gris
    """
    Crea una barra comparativa estética del Top-N (default=5).
    Resalta top1/top2, difumina el resto, añade leyenda con parámetros del cálculo.
    """
    # --- Datos ---
    top = (df.sort_values(libres_col_name, ascending=False)
             .head(n_top)
             .copy())
    top[hora_col_name] = top[hora_col_name].map(hhmm_to_am_pm)
    top["Etiqueta"] = top["Día"].astype(str) + " (" + top[hora_col_name].astype(str) + ")"
    y = top[libres_col_name].astype(float).values
    xlabs = top["Etiqueta"].tolist()

    # --- Colores por ranking ---
    colors = [color_top1, color_top2] + [color_rest]*(len(top)-2)
    alphas = [1.0, 0.95] + [0.50]*(len(top)-2)

    # --- Plot ---
    fig, ax = plt.subplots(figsize=(11.5, 6.5))
    bars = ax.bar(range(len(top)), y,
                  color=colors, alpha=1.0, edgecolor="none")
    # aplica alphas a partir de 3.º
    for i, rect in enumerate(bars):
        rect.set_alpha(alphas[i])

    # Ejes y formato
    ax.set_title("Top días para conferencia", pad=14, fontsize=16, weight="bold")
    ax.set_ylabel("Alumnos libres (estimados)", fontsize=12, weight="semibold")
    ax.set_xlabel("Día (Mejor hora)", fontsize=12, weight="semibold")
    ax.set_xticks(range(len(top)))
    ax.set_xticklabels(xlabs, rotation=0, fontsize=11)

    # Grid y limpieza de spines
    #ax.yaxis.set_major_formatter(mtick.FuncFormatter(_fmt_miles))
    ax.grid(axis="y", linestyle="--", linewidth=0.7, alpha=0.35)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    # Anotaciones (dentro si la barra es “alta”, si no encima)
    ymax = max(y) if len(y) else 0
    for i, rect in enumerate(bars):
        val = y[i]
        inside = val >= 0.18 * ymax
        txt_color = "white" if inside and i < 2 else "#111827"  # blanco sólo si resalta y es alto
        va = "center" if inside else "bottom"
        ytext = rect.get_y() + rect.get_height()/2 if inside else rect.get_height()
        ax.annotate(f"{int(val):,}",
                    xy=(rect.get_x() + rect.get_width()/2, ytext),
                    xytext=(0, 8 if not inside else 0),
                    textcoords="offset points",
                    ha="center", va=va,
                    fontsize=11, color=txt_color,
                    weight="bold" if i < 2 else "semibold")

    # Leyenda (NO caption): colores + parámetros del cálculo
    legend_handles = [
        Patch(facecolor=color_top1, label="Mejor día"),
        Patch(facecolor=color_top2, label="2.º mejor"),
        Patch(facecolor=color_rest, alpha=0.50, label="Resto Top-5"),
        Patch(facecolor="none", edgecolor="none")
    ]
    leg = ax.legend(handles=legend_handles,
                    loc="upper right", frameon=True, fontsize=10,
                    title="Interpretación", title_fontsize=11)
    leg.get_frame().set_alpha(0.9)

    fig.tight_layout()
    out_png = outdir / "Top5_mejores_horarios.png"
    fig.savefig(out_png, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return out_png

# ----------------- Main -----------------
def main():
    ruta_excel = input("Ruta completa al Excel ETL (debe incluir columna 'Clave'): ").strip().strip('"').strip("'")
    carrera = input("Clave de carrera (ej. LIFI, LQFB...): ").strip().upper()

    excel_path = Path(os.path.expanduser(os.path.expandvars(ruta_excel))).resolve()
    if not excel_path.is_file():
        print("ERROR: no se encontró el Excel.")
        return

    if carrera not in SEMESTRES_POR_CARRERA:
        print("ERROR: esa carrera no está definida en SEMESTRES_POR_CARRERA de constantes.py")
        return

    try:
        xls = pd.ExcelFile(excel_path)
    except Exception as e:
        print("ERROR al abrir Excel:", e)
        return

    # Cargar y normalizar por día
    per_day = {}
    for d in DIAS_ORDEN:
        if d not in xls.sheet_names:
            continue
        df = pd.read_excel(excel_path, sheet_name=d)
        if df.empty:
            continue

        # Validaciones básicas
        needed = {'Clave','Hora inicio','Hora fin','Alumnos'}
        missing = [c for c in needed if c not in df.columns]
        if missing:
            print(f"AVISO: Hoja {d} falta columnas {missing}; se omite.")
            continue

        for col in ['Hora inicio','Hora fin','Alumnos']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['Hora inicio','Hora fin'])
        df['Hora inicio'] = df['Hora inicio'].astype(int)
        df['Hora fin']    = df['Hora fin'].astype(int)
        df['Alumnos']     = df['Alumnos'].fillna(0).astype(int)
        df['_ini_min']    = df['Hora inicio'].apply(lambda x: to_minutes(int(x)))
        df['_fin_min']    = df['Hora fin'].apply(lambda x: to_minutes(int(x)))
        df['Clave']       = df['Clave'].astype(str)

        per_day[d] = df

    if not per_day:
        print("No hay hojas con datos utilizables.")
        return

    # Derivar carpeta de salida a partir del nombre del Excel
    filename_base = excel_path.stem  # p.ej. "HORARIOS_LIFI_CUCEI_202520"
    analysis_folder = f"ANALISIS {filename_base.replace('_', ' ')}"
    outdir = (excel_path.parent / analysis_folder).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"Carpeta de salida: {outdir}")

    # 1) Generar 6 archivos Excel (uno por día) con hojas por semestre
    for d, df_dia in per_day.items():
        day_filename = nombre_excel_por_dia(filename_base, d)
        out_day = outdir / f"{day_filename}.xlsx"
        with pd.ExcelWriter(out_day, engine="xlsxwriter") as w:
            for sem, claves in SEMESTRES_POR_CARRERA[carrera].items():
                claves_str = set(str(c) for c in claves)
                df_sem = df_dia[df_dia['Clave'].isin(claves_str)].copy()
                if not df_sem.empty:
                    df_sem = df_sem[['Hora inicio','Hora fin','Día','Salón','Materia','Profesor','Alumnos']]
                    df_sem = df_sem.sort_values(['Hora inicio'])
                    df_sem["Hora inicio"] = pd.to_datetime(
                        df_sem["Hora inicio"].astype(str).str.zfill(4), format="%H%M"
                    ).dt.strftime("%H:%M") # 700 -> 7:00
                    df_sem["Hora fin"] = pd.to_datetime(
                        df_sem["Hora fin"].astype(str).str.zfill(4), format="%H%M"
                    ).dt.strftime("%H:%M") # 700 -> 7:00
                df_sem.to_excel(w, sheet_name=f"{SEMESTRES[sem]} ", index=False)
        print("Archivo por día:", out_day)

        # (Opcional) AutoFit con COM de Excel si estás en Windows con Office
        try:
            AutoFit_columns_width(out_day)
        except Exception:
            pass

    # 2) Calcular mejor horario por día (sumando semestres sin clase)
    resumen_rows = []
    detalle_rows = []

    for d, df_dia in per_day.items():
        # Precalcular por semestre: DF y base_del_dia
        sem_data = {}
        for sem, claves in SEMESTRES_POR_CARRERA[carrera].items():
            claves_str = set(str(c) for c in claves)
            df_sem = df_dia[df_dia['Clave'].isin(claves_str)].copy()
            sem_data[sem] = df_sem  # guarda solo el DF; la base se calculará por candidato

        # Candidatos (nones) y exclusión 07:00 si aplica
        candidatos = CANDIDATOS_NONES[:]
        if EXCLUIR_7_AM and 700 in candidatos:
            candidatos = [h for h in candidatos if h != 700]

        best_t, best_score = None, -1

        for t in candidatos:
            t_min = to_minutes(t)
            aportes = []
            for sem, df_sem in sem_data.items():
                if df_sem.empty:
                    continue
                # 1) ¿Tiene clase en t?
                activos = activos_en_t(df_sem, t_min)
                if activos > 0:
                    libres = 0
                else:
                    # 2) Base en la ventana de 4 horas previa a t
                    base_win = base_en_ventana(df_sem, t_min, delta_min=240)
                    libres = base_win if base_win > 0 else 0
                aportes.append(libres)

            score = sum(aportes) if aportes else 0

            detalle_rows.append({
                "Día": DIAS_MAP[d],
                "Hora (candidato)": hhmm_int_to_str(t),
                "Suma libres": score
            })

            # Empate: elige la más tardía
            if (score > best_score) or (score == best_score and (best_t is None or t < best_t)):    
                best_t, best_score = t, score


        resumen_rows.append({
            "Día": DIAS_MAP[d],
            "Mejor hora": hhmm_int_to_str(best_t) if best_t is not None else None,
            "Libres estimados": best_score if best_score >= 0 else None
        })

    # 3) Exportar resumen global
    resumen_df = pd.DataFrame(resumen_rows, columns=[
        "Día", "Mejor hora", "Libres estimados"
    ])

    detalle_df = pd.DataFrame(detalle_rows, columns=[
        "Día", "Hora (candidato)", "Suma libres"
    ])

    out_summary = outdir / "Mejor_horario.xlsx"

    with pd.ExcelWriter(out_summary, engine="xlsxwriter") as writer:
        detalle_df.to_excel(writer, sheet_name="Detalle candidatos", index=False)
        resumen_df.to_excel(writer, sheet_name="Horarios recomendados", index=False)

        wb = writer.book
        bold = wb.add_format({"bold": True})
        for name, df in [("Detalle candidatos", detalle_df), ("Horarios recomendados", resumen_df)]:
            ws = writer.sheets[name]
            # Encabezados en negrita y autoancho
            for col, col_name in enumerate(df.columns):
                ws.write(0, col, col_name, bold)
            for i, col in enumerate(df.columns):
                width = min(max(len(str(x)) for x in [col] + df[col].astype(str).tolist()) + 2, 50)
                ws.set_column(i, i, width)

    print("\n[OK] Resumen global generado:", out_summary)
    print("Listo.\n")

    # ------- Graficas ------- #
    libres_col_name = "Libres estimados" 
    hora_col_name = "Mejor hora" 

    # Generar visualizaciones
    p1 = plot_top5_bar(resumen_df, hora_col_name, libres_col_name, outdir)

    print("\n[OK] Gráficas generadas en:", outdir)

if __name__ == "__main__":
    main()
