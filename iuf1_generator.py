import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import calendar
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Generador IUF1 – Resolución 9995/2021",
    page_icon="⚡",
    layout="wide"
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 1.7rem; font-weight: 700; color: #1a3a5c;
        border-bottom: 3px solid #f0a500; padding-bottom: 8px; margin-bottom: 4px;
    }
    .sub-title {
        font-size: 0.9rem; color: #6b7280; margin-bottom: 24px;
    }
    .section-header {
        font-size: 1rem; font-weight: 600; color: #1a3a5c;
        background: #eef3fa; padding: 6px 12px; border-radius: 6px;
        border-left: 4px solid #1a3a5c; margin: 16px 0 10px 0;
    }
    .stButton>button {
        background-color: #1a3a5c; color: white; border: none;
        font-weight: 600; border-radius: 6px; padding: 0.5rem 1.5rem;
    }
    .stButton>button:hover { background-color: #f0a500; color: #1a3a5c; }
    .info-box {
        background: #fffbea; border: 1px solid #f0a500;
        border-radius: 6px; padding: 10px 14px; margin-bottom: 12px;
        font-size: 0.88rem; color: #7a5500;
    }
    .success-box {
        background: #ecfdf5; border: 1px solid #10b981;
        border-radius: 6px; padding: 10px 14px; font-size: 0.88rem; color: #065f46;
    }
    .error-box {
        background: #fef2f2; border: 1px solid #ef4444;
        border-radius: 6px; padding: 10px 14px; font-size: 0.88rem; color: #7f1d1d;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚡ Generador Formato IUF1</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Resolución 9995 de 2021 · Sistemas Individuales de Generación Solar Fotovoltaica (SISFV) – ZNI</div>', unsafe_allow_html=True)

# ── Helpers ──────────────────────────────────────────────────────────────────
MESES_ES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}

def limpiar_guiones(valor):
    """Elimina todos los guiones de un string."""
    return str(valor).replace("-", "").strip()

def primer_dia(anio, mes):
    return date(anio, mes, 1)

def ultimo_dia(anio, mes):
    return date(anio, mes, calendar.monthrange(anio, mes)[1])

def dias_mes(anio, mes):
    return calendar.monthrange(anio, mes)[1]

def fmt_fecha(d):
    """DD-MM-YYYY"""
    return d.strftime("%d-%m-%Y")

def redondear_5dec(val):
    try:
        return round(float(val), 5)
    except:
        return val

def redondear_4dec(val):
    try:
        return round(float(val), 4)
    except:
        return val

# ── Sidebar: parámetros del mes ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📅 Parámetros del formato")
    anio_sel = st.selectbox("Año", list(range(2021, date.today().year + 2)),
                            index=list(range(2021, date.today().year + 2)).index(date.today().year))
    mes_sel = st.selectbox("Mes", list(MESES_ES.keys()),
                           format_func=lambda x: MESES_ES[x],
                           index=date.today().month - 1)

    n_dias = dias_mes(anio_sel, mes_sel)
    fecha_ini = primer_dia(anio_sel, mes_sel)
    fecha_exp = ultimo_dia(anio_sel, mes_sel)

    st.markdown("---")
    st.caption(f"**Primer día:** {fmt_fecha(fecha_ini)}")
    st.caption(f"**Último día:** {fmt_fecha(fecha_exp)}")
    st.caption(f"**Días del mes:** {n_dias}")
    st.markdown("---")
    st.caption("DISPOWER S.A.S. E.S.P. · ZNI Colombia")

# ── Carga de archivos ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">1 · Carga de archivos</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    st.caption("📋 **Base de Usuarios** (BD_Usuarios)")
    file_usuarios = st.file_uploader("Usuarios.xlsx", type=["xlsx"], key="usuarios")
with col2:
    st.caption("🧾 **Consolidado de Facturación** del mes")
    file_fact = st.file_uploader("Fact_MM_YYYY.xlsx", type=["xlsx"], key="fact")
with col3:
    st.caption("💳 **Cartera por NIU** (opcional)")
    file_cartera = st.file_uploader("Cartera_NIU.xlsx", type=["xlsx"], key="cartera")

# ── Procesamiento ─────────────────────────────────────────────────────────────
if file_usuarios and file_fact:
    try:
        # --- Leer Usuarios ---
        df_usr = pd.read_excel(file_usuarios, sheet_name="BD_Usuarios")
        cols_req_usr = ["NIU_SUI","COD_LOCALIDAD","Whd","LONGITUD","LATITUD","VEREDA"]
        faltantes = [c for c in cols_req_usr if c not in df_usr.columns]
        if faltantes:
            st.markdown(f'<div class="error-box">❌ Faltan columnas en Usuarios: {faltantes}</div>', unsafe_allow_html=True)
            st.stop()

        # Limpiar NIU_SUI en usuarios (convertir a str sin guiones)
        df_usr["NIU_KEY"] = df_usr["NIU_SUI"].astype(str).str.replace("-","",regex=False).str.strip()

        # --- Leer Facturación ---
        df_fact = pd.read_excel(file_fact, sheet_name=0)
        cols_req_fac = ["nui","nfacturasiigo","cantidad"]
        faltantes2 = [c for c in cols_req_fac if c not in df_fact.columns]
        if faltantes2:
            st.markdown(f'<div class="error-box">❌ Faltan columnas en Facturación: {faltantes2}</div>', unsafe_allow_html=True)
            st.stop()

        # Limpiar NIU en facturación
        df_fact["NIU_KEY"] = df_fact["nui"].astype(str).str.replace("-","",regex=False).str.strip()
        df_fact["ID_FACTURA_CLEAN"] = df_fact["nfacturasiigo"].astype(str).str.replace("-","",regex=False).str.strip()

        # --- Leer Cartera (opcional) ---
        df_cartera = None
        if file_cartera:
            df_cartera = pd.read_excel(file_cartera, sheet_name="Cartera_Total_NIU")
            cols_req_car = ["NIU","TARIFA TOTAL"]
            faltantes3 = [c for c in cols_req_car if c not in df_cartera.columns]
            if faltantes3:
                st.markdown(f'<div class="error-box">❌ Faltan columnas en Cartera: {faltantes3}</div>', unsafe_allow_html=True)
                st.stop()
            df_cartera["NIU_KEY"] = df_cartera["NIU"].astype(str).str.replace("-","",regex=False).str.strip()
            df_cartera = df_cartera[["NIU_KEY","TARIFA TOTAL"]].rename(columns={"TARIFA TOTAL":"CARTERA_VALOR"})

        # --- Cruce ---
        # Base: todos los usuarios → se cruza con facturación (left desde usuarios)
        # Usuarios sin factura ese mes quedan con ID_FACTURA_CLEAN y cantidad en NaN
        df_fact_red = df_fact[["NIU_KEY","ID_FACTURA_CLEAN","cantidad"]].copy()
        df = df_usr[["NIU_KEY","NIU_SUI","COD_LOCALIDAD","Whd","LONGITUD","LATITUD","VEREDA"]].merge(
            df_fact_red, on="NIU_KEY", how="left"
        )

        sin_fact = df["ID_FACTURA_CLEAN"].isna().sum()
        en_fact  = df["ID_FACTURA_CLEAN"].notna().sum()
        if sin_fact > 0:
            st.markdown(
                f'<div class="info-box">⚠️ <strong>{sin_fact}</strong> usuario(s) de la base '
                f'<strong>no tienen factura</strong> en este mes (se incluyen en el IUF1 con '
                f'campos de facturación vacíos). <strong>{en_fact}</strong> usuario(s) sí tienen factura.</div>',
                unsafe_allow_html=True
            )

        # ── Construir columnas IUF1 ──────────────────────────────────────────

        # Campos directos
        df["NIU"]            = df["NIU_SUI"].astype("Int64")
        df["COD LOCALIDAD"]  = df["COD_LOCALIDAD"].astype(str)
        df["dane"]           = df["COD LOCALIDAD"].str[:5]
        df["Whd"]            = df["Whd"]
        df["ALTITUD"]        = 0
        df["LONGITUD"]       = df["LONGITUD"].apply(redondear_5dec)
        df["LATITUD"]        = df["LATITUD"].apply(redondear_5dec)
        # ID FACTURA: si tiene factura → nfacturasiigo sin guiones
        #             si NO tiene factura → FV + NIU + MMYYYY
        mes_str = str(mes_sel).zfill(2)
        anio_str = str(anio_sel)
        df["ID FACTURA"] = df.apply(
            lambda r: r["ID_FACTURA_CLEAN"]
            if pd.notna(r["ID_FACTURA_CLEAN"]) and r["ID_FACTURA_CLEAN"] != ""
            else f"FV{str(int(r['NIU_SUI'])).strip()}{mes_str}{anio_str}",
            axis=1
        )
        df["TIPO CORRI SALIDA"] = 1
        # DIAS PRES MES: desde facturación; 0 si no tiene factura
        df["DIAS PRES MES"]  = df["cantidad"].fillna(0).astype(int)

        # ENERGIA GEN MES = Whd * DIAS PRES MES / 1000; 0 si no tiene factura
        df["ENERGIA GEN MES"] = (df["Whd"] * df["DIAS PRES MES"] / 1000).round(3).fillna(0)

        # disp promedio = DIAS PRES MES / dias del mes (fórmula; aquí calculado para previsualización)
        df["disp promedio"]   = (df["DIAS PRES MES"] / n_dias).round(4)

        # Campos en 0 con 4 decimales
        df["IUC"]  = 0.0000
        df["MP"]   = 0.0000
        df["GIO"]  = 0.0000
        df["PE"]   = 0.0000

        # Campos vacíos
        df["COR"]  = ""
        df["GAOM"] = ""

        # Dirección = VEREDA
        df["DIRECCION"] = df["VEREDA"]

        # Fechas
        df["FECH EXP FACT"]  = fmt_fecha(fecha_exp)
        df["FECH INI PERIO"] = fmt_fecha(fecha_ini)

        df["DIAS FACT"]  = df["DIAS PRES MES"]  # 0 para NIU sin factura
        df["ESTRATO"]    = 1
        df["TIPO LECT"]  = 3

        # Campos vacíos (usuario los llena)
        df["FACT CONSUMO"] = df["ID_FACTURA_CLEAN"].apply(lambda x: "" if pd.notna(x) and x != "" else 0)
        df["VAL REFACT"]   = 0.0000
        # VAL MORA: desde cartera si se cargó, cruzando por NIU; 0 si no hay cartera o no hay match
        if df_cartera is not None:
            df = df.merge(df_cartera, on="NIU_KEY", how="left")
            df["VAL MORA"] = df["CARTERA_VALOR"].fillna(0.0).round(4)
            df.drop(columns=["CARTERA_VALOR"], inplace=True)
            con_mora = (df["VAL MORA"] > 0).sum()
            st.markdown(
                f'<div class="info-box">💳 Cartera cargada: <strong>{con_mora}</strong> NIU con saldo en VAL MORA.</div>',
                unsafe_allow_html=True
            )
        else:
            df["VAL MORA"] = 0.0000
        df["INT MORA"]     = 0.0000
        df["VAL SUBS"]     = df["ID_FACTURA_CLEAN"].apply(lambda x: "" if pd.notna(x) and x != "" else 0)

        # PORCE SUBS = VAL SUBS / FACT CONSUMO (vacío, se calcula cuando el usuario llene)
        df["PORCE SUBS"]   = df["ID_FACTURA_CLEAN"].apply(lambda x: "" if pd.notna(x) and x != "" else 0)

        df["TARIFA"]        = df["ID_FACTURA_CLEAN"].apply(lambda x: "" if pd.notna(x) and x != "" else 0)
        df["VAL TOTAL FACT"] = ""  # = FACT CONSUMO

        # ── Orden de columnas IUF1 ───────────────────────────────────────────
        COLS_IUF1 = [
            "NIU","COD LOCALIDAD","dane","Whd","ALTITUD","LONGITUD","LATITUD",
            "ID FACTURA","ENERGIA GEN MES","TIPO CORRI SALIDA","DIAS PRES MES",
            "disp promedio","IUC","MP","COR","GIO","PE","GAOM","DIRECCION",
            "FECH EXP FACT","FECH INI PERIO","DIAS FACT","ESTRATO","TIPO LECT",
            "FACT CONSUMO","VAL REFACT","VAL MORA","INT MORA","VAL SUBS",
            "PORCE SUBS","TARIFA","VAL TOTAL FACT"
        ]
        df_out = df[COLS_IUF1].copy()

        # ── Vista previa ──────────────────────────────────────────────────────
        st.markdown('<div class="section-header">2 · Vista previa del formato IUF1</div>', unsafe_allow_html=True)
        st.caption(f"**{len(df_out):,}** registros · Mes: **{MESES_ES[mes_sel]} {anio_sel}**")

        # Mostrar solo primeras 50 filas para performance
        st.dataframe(df_out.head(50), use_container_width=True, height=280)
        if len(df_out) > 50:
            st.caption(f"Mostrando 50 de {len(df_out):,} filas. El archivo exportado contiene todos los registros.")

        # ── Estadísticas rápidas ──────────────────────────────────────────────
        with st.expander("📊 Resumen del proceso"):
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Registros facturación", len(df_fact))
            c2.metric("Registros en IUF1", len(df_out))
            c3.metric("Sin factura este mes", int(sin_fact))
            c4.metric("Días del mes", n_dias)

        # ── Exportar XLSX ─────────────────────────────────────────────────────
        st.markdown('<div class="section-header">3 · Descargar formato IUF1</div>', unsafe_allow_html=True)

        def generar_xlsx(df_out, anio, mes, n_dias, fecha_ini_dt, fecha_exp_dt):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Hoja1"

            # Colores
            HEADER_FILL  = PatternFill("solid", fgColor="1A3A5C")
            HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=9)
            DATA_FONT    = Font(name="Arial", size=9)
            ALT_FILL     = PatternFill("solid", fgColor="EEF3FA")
            BORDER_SIDE  = Side(style="thin", color="CCCCCC")
            CELL_BORDER  = Border(left=BORDER_SIDE, right=BORDER_SIDE,
                                  bottom=BORDER_SIDE, top=BORDER_SIDE)

            # Cabecera
            headers = list(df_out.columns)
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.fill   = HEADER_FILL
                cell.font   = HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = CELL_BORDER

            ws.row_dimensions[1].height = 30

            # Índices de columnas (1-based) para formatos especiales
            col_idx = {h: i+1 for i, h in enumerate(headers)}

            # Formatos de número para Excel
            FMT_5DEC   = '#,##0.00000'
            FMT_4DEC   = '#,##0.0000'
            FMT_3DEC   = '#,##0.000'
            FMT_INT    = '0'
            FMT_PCTG   = '0.0000'   # porcentaje guardado como fracción

            COLS_4DEC  = {"IUC","MP","GIO","PE","VAL REFACT","VAL MORA","INT MORA"}
            COLS_5DEC  = {"LONGITUD","LATITUD"}

            for ri, row in enumerate(df_out.itertuples(index=False), 2):
                fill = ALT_FILL if ri % 2 == 0 else None
                row_vals = list(row)
                for ci, (h, val) in enumerate(zip(headers, row_vals), 1):
                    cell = ws.cell(row=ri, column=ci)
                    cell.font   = DATA_FONT
                    cell.border = CELL_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if fill:
                        cell.fill = fill

                    # Fórmulas para celdas calculadas
                    col_disp   = col_idx["disp promedio"]
                    col_dp     = col_idx["DIAS PRES MES"]
                    col_porce  = col_idx["PORCE SUBS"]
                    col_fc     = col_idx["FACT CONSUMO"]
                    col_vs     = col_idx["VAL SUBS"]
                    col_vtf    = col_idx["VAL TOTAL FACT"]

                    if h == "disp promedio":
                        dp_cell = get_column_letter(col_dp) + str(ri)
                        cell.value = f"={dp_cell}/{n_dias}"
                        cell.number_format = FMT_4DEC
                    elif h == "PORCE SUBS":
                        fc_cell = get_column_letter(col_fc) + str(ri)
                        vs_cell = get_column_letter(col_vs) + str(ri)
                        cell.value = f'=IFERROR({vs_cell}/{fc_cell},"")'
                        cell.number_format = FMT_4DEC
                    elif h == "VAL TOTAL FACT":
                        fc_cell = get_column_letter(col_fc) + str(ri)
                        cell.value = f'=IFERROR({fc_cell},"")'
                        cell.number_format = FMT_4DEC
                    elif h in COLS_5DEC:
                        cell.value = float(val) if val != "" else ""
                        cell.number_format = FMT_5DEC
                    elif h in COLS_4DEC:
                        cell.value = float(val) if val != "" else 0.0
                        cell.number_format = FMT_4DEC
                    elif h == "ENERGIA GEN MES":
                        cell.value = float(val) if val != "" else ""
                        cell.number_format = FMT_3DEC
                    elif h == "NIU":
                        cell.value = int(val) if pd.notna(val) else ""
                        cell.number_format = FMT_INT
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    elif h in {"COD LOCALIDAD","dane"}:
                        cell.value = str(val) if val != "" else ""
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    elif h in {"FECH EXP FACT","FECH INI PERIO"}:
                        cell.value = str(val)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.value = val if val != "" else (None if isinstance(val, str) else val)

            # Anchos de columna aproximados
            COL_WIDTHS = {
                "NIU":14,"COD LOCALIDAD":17,"dane":8,"Whd":6,"ALTITUD":7,
                "LONGITUD":12,"LATITUD":12,"ID FACTURA":14,"ENERGIA GEN MES":14,
                "TIPO CORRI SALIDA":8,"DIAS PRES MES":10,"disp promedio":11,
                "IUC":9,"MP":9,"COR":9,"GIO":9,"PE":9,"GAOM":9,
                "DIRECCION":18,"FECH EXP FACT":13,"FECH INI PERIO":13,
                "DIAS FACT":9,"ESTRATO":8,"TIPO LECT":9,"FACT CONSUMO":13,
                "VAL REFACT":11,"VAL MORA":10,"INT MORA":10,"VAL SUBS":12,
                "PORCE SUBS":10,"TARIFA":12,"VAL TOTAL FACT":13,
            }
            for ci, h in enumerate(headers, 1):
                ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(h, 12)

            # Freeze header
            ws.freeze_panes = "A2"

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf.getvalue()

        def generar_csv(df_out):
            """CSV separado por comas, fechas y NIU como string."""
            df_csv = df_out.copy()
            buf = BytesIO()
            df_csv.to_csv(buf, index=False, sep=",", encoding="utf-8-sig")
            buf.seek(0)
            return buf.getvalue()

        nombre_base = f"IUF1_{str(mes_sel).zfill(2)}_{anio_sel}"

        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            xlsx_bytes = generar_xlsx(df_out, anio_sel, mes_sel, n_dias, fecha_ini, fecha_exp)
            st.download_button(
                label=f"⬇️ Descargar {nombre_base}.xlsx",
                data=xlsx_bytes,
                file_name=f"{nombre_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption("Incluye fórmulas Excel para disp promedio, PORCE SUBS y VAL TOTAL FACT")

        with col_dl2:
            csv_bytes = generar_csv(df_out)
            st.download_button(
                label=f"⬇️ Descargar {nombre_base}.csv",
                data=csv_bytes,
                file_name=f"{nombre_base}.csv",
                mime="text/csv",
                use_container_width=True
            )
            st.caption("CSV separado por comas · codificación UTF-8")

        st.markdown(f'<div class="success-box">✅ Formato IUF1 generado correctamente para <strong>{MESES_ES[mes_sel]} {anio_sel}</strong> con <strong>{len(df_out):,} registros</strong>.</div>', unsafe_allow_html=True)

    except Exception as e:
        st.markdown(f'<div class="error-box">❌ Error al procesar los archivos: {e}</div>', unsafe_allow_html=True)
        st.exception(e)

else:
    st.markdown('<div class="info-box">👆 Carga el archivo de <strong>Usuarios</strong> y el <strong>Consolidado de Facturación</strong> para comenzar. Selecciona el mes en el panel izquierdo.</div>', unsafe_allow_html=True)

    # Instrucciones
    with st.expander("📖 Instrucciones de uso"):
        st.markdown("""
**Paso 1 · Selecciona el mes** en el panel izquierdo (año y mes del formato a generar).

**Paso 2 · Carga los archivos:**
- **Usuarios.xlsx** → debe tener la hoja `BD_Usuarios` con las columnas: `NIU_SUI`, `COD_LOCALIDAD`, `Whd`, `LONGITUD`, `LATITUD`, `VEREDA`.
- **Facturación (Fact_MM_YYYY.xlsx)** → hoja 1, con columnas: `nui sui`, `nfacturasiigo`, `cantidad`.

**Paso 3 · Descarga** el formato en `.xlsx` (con fórmulas) o `.csv`.

---
**Campos calculados automáticamente:**
| Campo | Regla |
|---|---|
| `dane` | Primeros 5 dígitos de `COD LOCALIDAD` |
| `ENERGIA GEN MES` | `Whd × DIAS PRES MES ÷ 1000` |
| `disp promedio` | `DIAS PRES MES ÷ días del mes` (fórmula Excel) |
| `PORCE SUBS` | `VAL SUBS ÷ FACT CONSUMO` (fórmula Excel) |
| `VAL TOTAL FACT` | `= FACT CONSUMO` (fórmula Excel) |
| `FECH EXP FACT` | Último día del mes seleccionado |
| `FECH INI PERIO` | Primer día del mes seleccionado |
| `ID FACTURA` | `nfacturasiigo` sin guiones |

**Campos que debes completar manualmente:**
`FACT CONSUMO`, `VAL SUBS`, `TARIFA`, `COR`, `GAOM`
        """)
