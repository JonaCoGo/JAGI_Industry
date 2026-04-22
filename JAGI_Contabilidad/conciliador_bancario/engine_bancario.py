"""
Motor de Conciliación Bancaria — JAGI CAPS v1.0
Cruza el auxiliar contable WorldOffice contra el extracto bancario.

Proceso:
  - Auxiliar: registros de la cuenta (débitos y créditos) por día
  - Extracto: movimientos del banco (créditos = abonos, débitos = cargos)
  - Cruce: por fecha y valor, con tolerancia configurable

Bancos soportados en esta versión:
  - Davivienda (Corriente / Ahorros) — formato Excel estándar de descarga
  [Agregar lectores: leer_extracto_bancolombia(), leer_extracto_bogota()]
"""

import pandas as pd
import re
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paleta (consistente con engine de datafonos) ─────────────────────────────
COLOR = {
    "header_bg":    "1F3864", "header_font": "FFFFFF",
    "cuadra":       "C6EFCE", "cuadra_font": "276221",
    "dif_menor":    "FFEB9C", "dif_font":    "9C5700",
    "diferencia":   "FFC7CE", "sin_font":    "9C0006",
    "title_bg":     "2E75B6", "title_font":  "FFFFFF",
    "total_bg":     "1F3864",
    "solo_aux":     "FCE4D6", "solo_aux_f":  "843C0C",
    "solo_banco":   "E2EFDA", "solo_banco_f":"375623",
}
THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_COP  = '#,##0'
FMT_DATE = 'DD/MM/YYYY'
FMT_PCT  = '0.00%'

TOLERANCIA_DEFAULT = 1.0   # diferencia < $1 → CUADRA
TOLERANCIA_MENOR   = 0.05  # diferencia ≤ 5% → DIF_MENOR


# ════════════════════════════════════════════════════════════════════════════
# 1. LECTORES DE EXTRACTO BANCARIO (uno por banco)
# ════════════════════════════════════════════════════════════════════════════

def leer_extracto_davivienda(path: str) -> pd.DataFrame:
    """
    Lee extracto bancario Davivienda en formato TXT posicional (impresión).
    Compatible con:
      - Cuenta Corriente (2346): columnas Día Mes Oficina Descripción Doc Débito Crédito Saldo
      - Cuenta Ahorros   (4677): columnas Día Mes Oficina Descripción Doc Débito Crédito
        (sin columna Saldo — la omite en ahorros)
    Detecta automáticamente el tipo de cuenta desde el encabezado del archivo.
    Encoding: latin-1 (caracteres especiales Davivienda: é→?, ó→?)
    """
    import re

    # Patrón A: doc = solo dígitos (mayoría de líneas)
    PAT_A = re.compile(
        r'^\s{1,20}(\d{1,2})\s{1,5}(\d{1,2})\s{1,6}(\d{3,6})\s{2,}'
        r'(.+?)\s{2,}'
        r'(\d+)\s+'
        r'\$\s*([\d,\.]+)\s+'
        r'\$\s*([\d,\.]+)'
        r'(?:\s+\$\s*([\d,\.]+[\+\-]?))?'
    )
    # Patrón B: doc tiene espacios (transferencias a otras entidades)
    # En estas líneas SIEMPRE: crédito=$0.00 y el débito es el 1er valor monetario
    PAT_B = re.compile(
        r'^\s{1,20}(\d{1,2})\s{1,5}(\d{1,2})\s{1,6}(\d{3,6})\s{2,}'
        r'(.+?)\s{3,}'                           # descripcion con nombre de banco
        r'\$\s*([\d,\.]+)\s+'                    # débito
        r'\$\s*([\d,\.]+)'                       # crédito ($0.00)
        r'(?:\s+\$\s*([\d,\.]+[\+\-]?))?'       # saldo
    )

    with open(path, 'r', encoding='latin-1') as f:
        lines = f.readlines()

    # Extraer metadata del encabezado (primeras 40 líneas)
    texto_cab = ' '.join(l.strip() for l in lines[:40] if l.strip())
    tipo_cuenta = 'Corriente' if 'CORRIENTE' in texto_cab.upper() else 'Ahorros'

    # Extraer año del encabezado para construir fechas completas
    m_anio = re.search(r'INFORME DEL MES:\s+\w+\s*/(\d{4})', texto_cab, re.IGNORECASE)
    anio   = int(m_anio.group(1)) if m_anio else datetime.now().year

    # Parsear líneas de datos
    registros = []
    for linea in lines:
        lc = linea.rstrip()
        if not re.match(r'^\s{1,20}\d{1,2}\s+\d{1,2}\s+\d{3,}', lc):
            continue

        m = PAT_A.match(lc)
        if m:
            dia_str, mes_str, oficina, desc, doc, deb_str, cred_str, saldo_str = m.groups()
        else:
            m = PAT_B.match(lc)
            if not m:
                continue
            dia_str, mes_str, oficina, desc, deb_str, cred_str, saldo_str = m.groups()
            doc = ''   # doc embebido en descripción

        dia  = int(dia_str)
        mes  = int(mes_str)
        fecha = pd.Timestamp(year=anio, month=mes, day=dia)
        desc_clean = re.sub(r'\s{2,}', ' ', desc).strip()

        def _parse_money(s):
            if not s: return 0.0
            return float(s.replace(',', '').rstrip('+-'))

        deb   = _parse_money(deb_str)
        cred  = _parse_money(cred_str)
        saldo = _parse_money(saldo_str) if saldo_str else None

        registros.append({
            'Fecha':       fecha,
            'Dia':         dia,
            'Mes':         mes,
            'Oficina':     oficina,
            'Descripcion': desc_clean,
            'Referencia':  doc,
            'Debitos':     deb,
            'Creditos':    cred,
            'Saldo':       saldo,
            '_Banco':      'Davivienda',
            '_TipoCuenta': tipo_cuenta,
            '_Archivo':    os.path.basename(path),
        })

    if not registros:
        raise ValueError(
            f"No se encontraron movimientos en '{os.path.basename(path)}'.\n"
            "Verifique que el archivo es un extracto Davivienda en formato TXT."
        )

    df = pd.DataFrame(registros)
    # Sanidad: filtrar filas con valores inconsistentes
    df = df[df['Fecha'].dt.year >= 2020].copy()
    return df.reset_index(drop=True)


def leer_extracto_bancolombia(path: str) -> pd.DataFrame:
    """
    Lector para extracto Bancolombia — PENDIENTE implementación.
    Completar cuando se reciba el archivo real.
    """
    raise NotImplementedError(
        "Lector Bancolombia pendiente. Por favor comparta un extracto de ejemplo "
        "de la cuenta Bancolombia para implementar el lector específico."
    )


def leer_extracto_bogota(path: str) -> pd.DataFrame:
    """
    Lector para extracto Banco de Bogotá — PENDIENTE implementación.
    Completar cuando se reciba el archivo real.
    """
    raise NotImplementedError(
        "Lector Banco de Bogotá pendiente. Por favor comparta un extracto de ejemplo "
        "para implementar el lector específico."
    )


# Registro de lectores por banco (extensible sin modificar engine)
LECTORES_EXTRACTO = {
    "Davivienda":     leer_extracto_davivienda,
    "Bancolombia":    leer_extracto_bancolombia,
    "Banco de Bogotá": leer_extracto_bogota,
}


def leer_extracto(path: str, banco: str) -> pd.DataFrame:
    """
    Punto de entrada unificado. Selecciona el lector según el banco.
    Davivienda: soporta TXT (formato impresión) y Excel.
    Bancolombia / Banco de Bogotá: pendiente — agregar lector cuando lleguen archivos.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Archivo no encontrado: {path}")

    lector = LECTORES_EXTRACTO.get(banco)
    if lector is None:
        raise ValueError(
            f"Banco '{banco}' no tiene lector implementado.\n"
            f"Bancos disponibles: {[b for b,fn in LECTORES_EXTRACTO.items() if not fn.__doc__ or 'PENDIENTE' not in fn.__doc__]}"
        )
    return lector(path)


def _mapear_columnas_extracto(cols_reales: list, mapa_esperado: dict) -> dict:
    """
    Mapea columnas reales del extracto a nombres estándar.
    Retorna {nombre_std: nombre_real_o_None}
    """
    cols_lower = {c.lower().strip(): c for c in cols_reales}
    resultado  = {}
    for col_std, alias_list in mapa_esperado.items():
        encontrado = None
        for alias in alias_list:
            if alias in cols_lower:
                encontrado = cols_lower[alias]
                break
        resultado[col_std] = encontrado
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# 2. LECTURA DEL AUXILIAR (reutiliza lógica del engine de datafonos)
# ════════════════════════════════════════════════════════════════════════════

def leer_auxiliar_bancario(path: str, year: int, mes: int) -> pd.DataFrame:
    """
    Lee el auxiliar WorldOffice para conciliación bancaria.
    A diferencia del engine de datafonos, aquí NO filtramos solo DATAFONO:
    incluimos TODOS los movimientos de la cuenta en el período (débitos y créditos).
    Retorna DataFrame con: Fecha_Aux, Nota, Doc_Num, Debitos, Creditos, Saldo, _Dia
    """
    raw  = pd.read_excel(path, sheet_name=0, header=None)
    data = raw.iloc[4:].copy()
    data.columns = raw.iloc[3].tolist()
    data = data.reset_index(drop=True)

    # Normalizar columnas numéricas
    for col in ['Debitos', 'Creditos', 'Saldo']:
        data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)

    # Normalizar fecha contable (aquí SÍ usamos la columna Fecha — es conciliación bancaria)
    data['Fecha'] = pd.to_datetime(data['Fecha'], errors='coerce')

    # Filtrar por período
    mask_periodo = (
        (data['Fecha'].dt.year  == year) &
        (data['Fecha'].dt.month == mes)
    )
    df = data[mask_periodo].copy()

    # Eliminar filas completamente vacías
    df = df[~(df['Debitos'].eq(0) & df['Creditos'].eq(0))].copy()

    df['_Dia']      = df['Fecha'].dt.day
    df['_Mes']      = df['Fecha'].dt.month
    df['_Year']     = df['Fecha'].dt.year
    df['Nota']      = df['Nota'].astype(str).str.strip()
    df['Doc Num']   = df['Doc Num'].astype(str).str.strip() if 'Doc Num' in df.columns else ''

    return df.reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════════════
# 3. CRUCE AUXILIAR ↔ EXTRACTO BANCARIO
# ════════════════════════════════════════════════════════════════════════════

def cruzar_auxiliar_extracto(df_aux: pd.DataFrame,
                              df_ext: pd.DataFrame,
                              year: int, mes: int,
                              tolerancia_abs: float = TOLERANCIA_DEFAULT,
                              tolerancia_pct: float = TOLERANCIA_MENOR) -> dict:
    """
    Cruza el auxiliar contra el extracto bancario.

    Estrategia de cruce (en orden de prioridad):
      1. Mismo día + mismo valor (exacto o dentro de tolerancia)
      2. Mismo valor en ventana ±3 días (para abonos con diferencia de fecha)
      3. Sin match

    Retorna dict con:
      'cruces':        lista de registros cruzados (auxiliar ↔ banco)
      'solo_aux':      registros del auxiliar sin contrapartida en extracto
      'solo_banco':    registros del extracto sin contrapartida en auxiliar
      'resumen':       totales y estadísticas
    """
    # Filtrar extracto por período
    ext_periodo = df_ext[
        (df_ext['Fecha'].dt.year  == year) &
        (df_ext['Fecha'].dt.month == mes)
    ].copy()

    # Separar débitos y créditos del auxiliar
    aux_deb = df_aux[df_aux['Debitos']  > 0].copy()
    aux_cred= df_aux[df_aux['Creditos'] > 0].copy()

    # Separar créditos y débitos del extracto bancario
    # En extracto: Créditos = dinero que entra a la cuenta (coincide con Débitos del auxiliar de ingresos)
    # En extracto: Débitos  = dinero que sale de la cuenta (coincide con Créditos del auxiliar de gastos)
    ext_cred = ext_periodo[ext_periodo['Creditos'] > 0].copy()
    ext_deb  = ext_periodo[ext_periodo['Debitos']  > 0].copy()

    cruces      = []
    usados_ext  = set()

    def _cruzar_grupo(df_a: pd.DataFrame, df_e: pd.DataFrame,
                       col_a: str, col_e: str, tipo: str):
        """Cruza un grupo (débitos o créditos) del auxiliar contra el extracto."""
        usados_local = set()
        for idx_a, ra in df_a.iterrows():
            v_aux = float(ra[col_a])
            dia   = int(ra['_Dia'])
            mejor = None; mejor_score = -1

            for idx_e, re_ in df_e.iterrows():
                if idx_e in usados_ext or idx_e in usados_local:
                    continue
                v_ext = float(re_[col_e])
                dif   = abs(v_aux - v_ext)
                dif_e = int(abs((re_['Fecha'] - ra['Fecha']).days)) if pd.notna(re_['Fecha']) and pd.notna(ra['Fecha']) else 999

                if dif < tolerancia_abs and dif_e == 0:
                    # Exacto mismo día
                    mejor = (idx_e, re_, dif, dif_e, 'CUADRA'); break
                elif dif <= v_aux * tolerancia_pct and dif_e == 0:
                    score = 90 - dif_e
                    if score > mejor_score:
                        mejor = (idx_e, re_, dif, dif_e, 'DIF_MENOR'); mejor_score = score
                elif dif < tolerancia_abs and dif_e <= 3:
                    score = 80 - dif_e * 10
                    if score > mejor_score:
                        mejor = (idx_e, re_, dif, dif_e, 'CUADRA_FECHA_DIF'); mejor_score = score
                elif dif <= v_aux * tolerancia_pct and dif_e <= 3:
                    score = 70 - dif_e * 10
                    if score > mejor_score:
                        mejor = (idx_e, re_, dif, dif_e, 'DIF_MENOR_FECHA_DIF'); mejor_score = score

            if mejor:
                idx_e, re_, dif, dif_e, estado = mejor
                usados_ext.add(idx_e); usados_local.add(idx_e)
                cruces.append({
                    'tipo':         tipo,
                    'row_aux':      ra,
                    'row_ext':      re_,
                    'valor_aux':    v_aux,
                    'valor_ext':    float(re_[col_e]),
                    'diferencia':   v_aux - float(re_[col_e]),
                    'dias_offset':  dif_e,
                    'estado':       estado,
                    'fecha_aux':    ra['Fecha'],
                    'fecha_ext':    re_['Fecha'],
                })
            else:
                cruces.append({
                    'tipo':         tipo,
                    'row_aux':      ra,
                    'row_ext':      None,
                    'valor_aux':    v_aux,
                    'valor_ext':    0.0,
                    'diferencia':   v_aux,
                    'dias_offset':  None,
                    'estado':       'SIN_MATCH',
                    'fecha_aux':    ra['Fecha'],
                    'fecha_ext':    None,
                })

    _cruzar_grupo(aux_deb,  ext_cred, 'Debitos',  'Creditos', 'DEBITO_AUX')
    _cruzar_grupo(aux_cred, ext_deb,  'Creditos', 'Debitos',  'CREDITO_AUX')

    # Registros del extracto sin match en el auxiliar
    idxs_cruzados = {c['row_ext'].name for c in cruces if c['row_ext'] is not None}
    solo_banco    = ext_periodo[~ext_periodo.index.isin(idxs_cruzados)].copy()

    # Construir resumen
    cuadra    = sum(1 for c in cruces if c['estado'] in ('CUADRA','CUADRA_FECHA_DIF'))
    dif_menor = sum(1 for c in cruces if 'DIF_MENOR' in c['estado'])
    sin_match = sum(1 for c in cruces if c['estado'] == 'SIN_MATCH')
    total_aux = sum(c['valor_aux'] for c in cruces)
    total_ext = sum(c['valor_ext'] for c in cruces if c['row_ext'] is not None)

    resumen = {
        'total_registros': len(cruces),
        'cuadra':          cuadra,
        'dif_menor':       dif_menor,
        'sin_match':       sin_match,
        'solo_banco':      len(solo_banco),
        'total_aux':       total_aux,
        'total_ext':       total_ext,
        'diferencia_neta': total_aux - total_ext,
    }

    return {
        'cruces':     cruces,
        'solo_banco': solo_banco,
        'resumen':    resumen,
    }


# ════════════════════════════════════════════════════════════════════════════
# 4. GENERACIÓN DEL EXCEL DE RESULTADO
# ════════════════════════════════════════════════════════════════════════════

def _sc(cell, bg=None, fc="000000", bold=False, fmt=None, ha="left", sz=10):
    if bg: cell.fill = PatternFill("solid", start_color=bg, fgColor=bg)
    cell.font      = Font(color=fc, bold=bold, name="Arial", size=sz)
    cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=False)
    cell.border    = BORDER
    if fmt: cell.number_format = fmt


def _titulo(ws, texto, subtexto="", ncols=14):
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    t = ws["A1"]
    t.value = texto
    t.fill  = PatternFill("solid", start_color=COLOR["title_bg"])
    t.font  = Font(color="FFFFFF", bold=True, name="Arial", size=12)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    if subtexto:
        ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
        g = ws["A2"]
        g.value = subtexto
        g.font  = Font(color="595959", italic=True, name="Arial", size=8)
        g.alignment = Alignment(horizontal="center")


def generar_excel_bancario(resultado: dict,
                            cuenta: str,
                            banco: str,
                            output_path: str,
                            periodo: str = "Enero 2026",
                            info_empresa: dict = None) -> str:
    """
    Genera el Excel de conciliación bancaria.
    Hojas: RESUMEN | DETALLE_CRUCE | SOLO_EN_BANCO | LOG_AUDITORIA
    """
    wb   = Workbook()
    cruces     = resultado['cruces']
    solo_banco = resultado['solo_banco']
    resumen    = resultado['resumen']

    emp_label = f"{info_empresa['razon_social']} — {cuenta}" if info_empresa else cuenta

    # ── Hoja RESUMEN ─────────────────────────────────────────────────────────
    ws_r = wb.active; ws_r.title = "RESUMEN"
    _titulo(ws_r,
            f"CONCILIACIÓN BANCARIA — {cuenta} — {periodo.upper()}",
            f"JAGI CAPS | {emp_label} | Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            ncols=3)
    ws_r.column_dimensions['A'].width = 48
    ws_r.column_dimensions['B'].width = 16
    ws_r.column_dimensions['C'].width = 22

    def put(r, c, v, bg=None, bold=False, fmt=None, ha="left"):
        cell = ws_r.cell(row=r, column=c, value=v)
        _sc(cell, bg=bg, bold=bold, fmt=fmt, ha=ha)

    put(4,1,"INDICADOR",bg=COLOR["header_bg"],bold=True,ha="center")
    put(4,2,"CANTIDAD", bg=COLOR["header_bg"],bold=True,ha="center")
    put(4,3,"VALOR ($)",bg=COLOR["header_bg"],bold=True,ha="center")
    ws_r.row_dimensions[4].height = 20

    stats = [
        ("Total movimientos en auxiliar",                   resumen['total_registros'], resumen['total_aux']),
        ("✔  Cuadran exacto (mismo día y valor)",           resumen['cuadra'],          None),
        ("⚠  Diferencia menor / fecha diferente",           resumen['dif_menor'],       None),
        ("❌  Sin match en extracto bancario",              resumen['sin_match'],       None),
        ("ℹ  Solo en extracto bancario (sin auxiliar)",     resumen['solo_banco'],      None),
        ("Diferencia neta ($)",                             "",                         resumen['diferencia_neta']),
    ]
    bgs = [None, COLOR["cuadra"], COLOR["dif_menor"], COLOR["diferencia"], COLOR["solo_banco"], None]
    for i, ((lbl, cant, val), bg) in enumerate(zip(stats, bgs), 1):
        put(4+i, 1, lbl,  bg=bg)
        put(4+i, 2, cant, bg=bg, ha="center")
        put(4+i, 3, val if val is not None else "", bg=bg,
            fmt=FMT_COP if val is not None else None, ha="right")
        ws_r.row_dimensions[4+i].height = 18

    # ── Hoja DETALLE_CRUCE ───────────────────────────────────────────────────
    ws_d = wb.create_sheet("DETALLE_CRUCE")
    _titulo(ws_d,
            f"DETALLE DE CRUCE — {cuenta} — {periodo.upper()}",
            "Auxiliar WorldOffice ↔ Extracto Bancario", ncols=12)

    HDRS = ["Tipo","Fecha Aux","Nota Auxiliar","Doc Num",
            "Valor Aux ($)","Fecha Banco","Descripción Banco","Ref. Banco",
            "Valor Banco ($)","Diferencia ($)","Días Offset","Estado"]
    WIDS = [12, 12, 36, 16, 16, 12, 32, 16, 16, 14, 12, 18]
    for ci,(h,w) in enumerate(zip(HDRS,WIDS),1):
        cell = ws_d.cell(row=3, column=ci, value=h)
        _sc(cell, bg=COLOR["header_bg"], fc=COLOR["header_font"], bold=True, ha="center")
        ws_d.column_dimensions[get_column_letter(ci)].width = w
    ws_d.row_dimensions[3].height = 20

    row_n = 4
    for c in cruces:
        estado = c['estado']
        if 'CUADRA' in estado:
            bg, fc = COLOR["cuadra"],    COLOR["cuadra_font"]
        elif 'DIF_MENOR' in estado:
            bg, fc = COLOR["dif_menor"], COLOR["dif_font"]
        else:
            bg, fc = COLOR["diferencia"],COLOR["sin_font"]

        re_ = c['row_ext']
        vals = [
            c['tipo'],
            c['fecha_aux'].date() if pd.notna(c['fecha_aux']) else None,
            str(c['row_aux'].get('Nota',''))[:60],
            str(c['row_aux'].get('Doc Num','')),
            c['valor_aux'],
            c['fecha_ext'].date() if re_ is not None and pd.notna(c['fecha_ext']) else None,
            str(re_['Descripcion'])[:50] if re_ is not None and 'Descripcion' in re_ else None,
            str(re_['Referencia'])       if re_ is not None and 'Referencia'  in re_ else None,
            c['valor_ext'] if c['valor_ext'] > 0 else None,
            c['diferencia'] if c['valor_ext'] > 0 else None,
            c['dias_offset'],
            estado,
        ]
        fmts = [None,FMT_DATE,None,None,FMT_COP,FMT_DATE,None,None,FMT_COP,FMT_COP,None,None]
        for ci,(v,fmt) in enumerate(zip(vals,fmts),1):
            cell = ws_d.cell(row=row_n, column=ci, value=v)
            _sc(cell, bg=bg, fc=fc, fmt=fmt,
                ha="right" if ci in [5,9,10] else "center" if ci in [1,2,6,11] else "left")
        ws_d.row_dimensions[row_n].height = 16
        row_n += 1
    ws_d.freeze_panes = "A4"

    # ── Hoja SOLO_EN_BANCO ───────────────────────────────────────────────────
    ws_s = wb.create_sheet("SOLO_EN_BANCO")
    _titulo(ws_s,
            f"MOVIMIENTOS SOLO EN EXTRACTO BANCARIO — {cuenta}",
            "Movimientos que el banco registró pero no aparecen en el auxiliar contable", ncols=6)
    HDRS_S = ["Fecha","Descripción","Referencia","Débitos ($)","Créditos ($)","Saldo ($)"]
    WIDS_S = [12, 44, 20, 18, 18, 18]
    for ci,(h,w) in enumerate(zip(HDRS_S,WIDS_S),1):
        cell = ws_s.cell(row=3, column=ci, value=h)
        _sc(cell, bg=COLOR["header_bg"], fc=COLOR["header_font"], bold=True, ha="center")
        ws_s.column_dimensions[get_column_letter(ci)].width = w
    ws_s.row_dimensions[3].height = 18

    row_n = 4
    for _, re_ in solo_banco.iterrows():
        vals = [
            re_['Fecha'].date() if pd.notna(re_['Fecha']) else None,
            str(re_.get('Descripcion','')),
            str(re_.get('Referencia','')),
            re_.get('Debitos',  0) or None,
            re_.get('Creditos', 0) or None,
            re_.get('Saldo',    0),
        ]
        fmts = [FMT_DATE,None,None,FMT_COP,FMT_COP,FMT_COP]
        for ci,(v,fmt) in enumerate(zip(vals,fmts),1):
            cell = ws_s.cell(row=row_n, column=ci, value=v)
            _sc(cell, bg=COLOR["solo_banco"], fc=COLOR["solo_banco_f"],
                fmt=fmt, ha="right" if ci in [4,5,6] else "center" if ci==1 else "left")
        ws_s.row_dimensions[row_n].height = 15
        row_n += 1
    ws_s.freeze_panes = "A4"

    # ── Hoja LOG_AUDITORIA ───────────────────────────────────────────────────
    ws_a = wb.create_sheet("LOG_AUDITORIA")
    ws_a.column_dimensions['A'].width = 40
    ws_a.column_dimensions['B'].width = 66
    _titulo(ws_a, "LOG DE AUDITORÍA — CONCILIACIÓN BANCARIA", ncols=2)

    registros_log = [
        ("Fecha / Hora proceso",       datetime.now().strftime('%d/%m/%Y %H:%M:%S')),
        ("Empresa",                    info_empresa.get('razon_social','') if info_empresa else ''),
        ("Marca comercial",            "JAGI CAPS"),
        ("Cuenta bancaria",            cuenta),
        ("Banco",                      banco),
        ("Período",                    periodo),
        ("Versión motor",              "1.0"),
        ("Normativa",                  "NIIF para PYMES (Decreto 2420 de 2015) / DIAN Colombia"),
        ("Lógica de cruce",            "Auxiliar WorldOffice ↔ Extracto bancario por fecha y valor"),
        ("Estrategia de match",        "1) Mismo día + valor exacto  2) ±3 días + tolerancia 5%"),
        ("Tolerancia diferencia",      f"< ${TOLERANCIA_DEFAULT:.0f} → CUADRA | ≤ {TOLERANCIA_MENOR*100:.0f}% → DIF_MENOR"),
        ("Archivos originales",        "NO modificados — motor opera sobre copias en memoria"),
        ("Auditoría externa",          "KPMG Ltda."),
    ]
    for i,(k,v) in enumerate(registros_log, 2):
        ck = ws_a.cell(row=i, column=1, value=k)
        cv = ws_a.cell(row=i, column=2, value=v)
        _sc(ck, bg="F2F2F2", bold=True)
        _sc(cv)
        ws_a.row_dimensions[i].height = 16

    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    wb.save(output_path)
    return output_path