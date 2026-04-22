"""
Interfaz Gráfica — Sistema de Conciliación de Datafonos  v1.5
Tkinter · Compatible Windows / Mac / Linux

NUEVO en v1.5:
  • Coincidencia difusa de nombres (AMERICAS → PLAZA DE LAS AMERICAS)
  • Aviso de archivos datafono huérfanos (sin sede en el auxiliar)
  • Tres modos de salida: individual | archivos separados | todo en un Excel
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, subprocess, sys
from datetime import datetime
from pathlib import Path

from conciliador_engine_1_5 import (
    leer_auxiliar, sedes_disponibles,
    cargar_multiples_datafonos, construir_mapa_nombres,
    agrupar_datafono_por_dia_vale, cruzar_auxiliar_datafono,
    generar_excel_resultado, generar_excel_unificado,
    generar_resumen_consolidado,
)

C = {"bg":"#F5F7FA","sidebar":"#1F3864","accent":"#2E75B6",
     "label":"#333333","entry":"#FFFFFF","hover":"#1A5A9F",
     "verde":"#276221","naranja":"#843C0C"}
FT = ("Arial",15,"bold"); FH = ("Arial",11,"bold")
FB = ("Arial",10);        FS = ("Arial",9)

MODO_IND    = "individual"
MODO_SEP    = "separados"
MODO_UNI    = "unificado"


class ConciliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conciliador de Datafonos — v1.5")
        self.geometry("1000x780"); self.minsize(860,640)
        self.configure(bg=C["bg"])
        self._auxiliar   = tk.StringVar()
        self._df_paths   = []
        self._out_dir    = tk.StringVar(value=str(Path.home()/"Desktop"))
        self._modo       = tk.StringVar(value=MODO_IND)
        self._sede       = tk.StringVar(value="")
        self._periodo    = tk.StringVar(value="Enero 2026")
        self._year       = tk.StringVar(value="2026")
        self._mes        = tk.StringVar(value="1")
        self._status     = tk.StringVar(value="Listo")
        self._prog       = tk.IntVar(value=0)
        self._last_out   = None; self._last_dir = None
        self._build_ui()

    # ── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Sidebar
        sb = tk.Frame(self, bg=C["sidebar"], width=210)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)
        tk.Label(sb, text="⚙ Conciliador\nDatafonos",
                 bg=C["sidebar"], fg="white",
                 font=("Arial",13,"bold"), justify="center").pack(pady=(24,6), padx=12)
        tk.Label(sb, text="──────────────────",
                 bg=C["sidebar"], fg="#445577", font=FS).pack()
        for lbl, cmd in [
            ("📂  Cargar archivos",   lambda: self._canvas.yview_moveto(0)),
            ("⚙   Configuración",    lambda: self._canvas.yview_moveto(0.32)),
            ("▶   Ejecutar",         self._ejecutar),
            ("📋  Ver resultados",   self._abrir),
        ]:
            b = tk.Button(sb, text=lbl, command=cmd,
                          bg=C["sidebar"], fg="white",
                          activebackground=C["accent"],
                          relief="flat", anchor="w", font=FB,
                          padx=16, pady=9, cursor="hand2")
            b.pack(fill="x")
            b.bind("<Enter>", lambda e,btn=b: btn.config(bg=C["accent"]))
            b.bind("<Leave>", lambda e,btn=b: btn.config(bg=C["sidebar"]))
        tk.Label(sb, text="\nv1.5 | Colombia\nNIIF PYMES / DIAN\nMatch difuso\nFormatos multi-año",
                 bg=C["sidebar"], fg="#8899BB",
                 font=FS, justify="center").pack(side="bottom", pady=16)

        # Main scrollable
        main = tk.Frame(self, bg=C["bg"]); main.pack(side="left", fill="both", expand=True)
        canvas = tk.Canvas(main, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
        self._canvas = canvas
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        self._frame = tk.Frame(canvas, bg=C["bg"])
        self._win   = canvas.create_window((0,0), window=self._frame, anchor="nw")
        self._frame.bind("<Configure>",
                          lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                     lambda e: canvas.itemconfig(self._win, width=e.width))
        self._build_content()

    def _sec(self, parent, title, pad=None):
        f = tk.LabelFrame(parent, text=f"  {title}  ",
                           bg=C["bg"], fg=C["accent"],
                           font=FH, padx=4, pady=8, relief="groove", bd=1)
        f.pack(fill="x", padx=24, pady=10); return f

    def _btn(self, parent, text, cmd, side="left", pri=False, sec=False, w=None):
        bg = C["accent"] if (pri or not sec) else "#E0E0E0"
        fg = "white"     if (pri or not sec) else "#333333"
        b  = tk.Button(parent, text=text, command=cmd,
                        bg=bg, fg=fg, font=FB, relief="flat",
                        cursor="hand2", padx=10, pady=5, width=w)
        b.pack(side=side, padx=4)
        b.bind("<Enter>", lambda e: b.config(bg=C["hover"] if (pri or not sec) else "#CCCCCC"))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _log(self, msg):
        self._log_txt.config(state="normal")
        self._log_txt.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self._log_txt.see("end"); self._log_txt.config(state="disabled")
        self.update_idletasks()

    def _set_prog(self, v, msg):
        self._prog.set(v); self._status.set(msg); self.update_idletasks()

    # ── CONTENIDO ────────────────────────────────────────────────────────────

    def _build_content(self):
        c = self._frame

        # Banner
        bn = tk.Frame(c, bg=C["accent"], height=74); bn.pack(fill="x")
        tk.Label(bn, text="Sistema de Conciliación de Datafonos",
                 bg=C["accent"], fg="white", font=FT).pack(side="left", padx=24, pady=20)
        tk.Label(bn, text="GIRALDO GIRALDO JAIME WILSON\nCuenta 2346 · Banco Davivienda",
                 bg=C["accent"], fg="#CCE0F5", font=FS, justify="right").pack(side="right", padx=20)

        # ── Sec 1: Archivos
        s1 = self._sec(c, "📂  1. Carga de archivos")
        # Auxiliar
        af = tk.Frame(s1, bg=C["bg"]); af.pack(fill="x", padx=24, pady=4)
        tk.Label(af, text="Auxiliar WorldOffice (.xlsx):", bg=C["bg"],
                 fg=C["label"], font=FB, width=28, anchor="w").pack(side="left")
        tk.Entry(af, textvariable=self._auxiliar, font=FB,
                 bg=C["entry"], relief="solid", bd=1, width=36).pack(side="left")
        self._btn(af, "📁 Buscar", self._browse_aux, "left")
        tk.Label(s1, text="ℹ Exportado directamente desde WorldOffice — sin modificar",
                 bg=C["bg"], fg="#888888", font=FS).pack(anchor="w", padx=24)

        # Datafonos
        df_lf = tk.LabelFrame(s1, text=" Archivos de Datafono (1 ó varios .xlsx) ",
                               bg=C["bg"], fg=C["label"], font=FB, padx=12, pady=8)
        df_lf.pack(fill="x", padx=24, pady=8)

        # Aviso matching difuso
        aviso = tk.Frame(df_lf, bg="#EFF4FB"); aviso.pack(fill="x", pady=(0,6))
        tk.Label(aviso,
                 text="ℹ  El sistema detecta coincidencias parciales de nombres automáticamente.\n"
                      "   Ejemplo: datafono_AMERICAS.xlsx  →  sede 'PLAZA DE LAS AMERICAS' en el auxiliar.",
                 bg="#EFF4FB", fg="#1F3864", font=FS, justify="left",
                 padx=8, pady=6).pack(anchor="w")

        btns = tk.Frame(df_lf, bg=C["bg"]); btns.pack(fill="x")
        self._btn(btns, "➕ Agregar datafono(s)", self._browse_df, "left")
        self._btn(btns, "🗑 Limpiar",             self._clear_df,  "left", sec=True)
        tk.Label(btns, text="Hasta 20 archivos — el nombre debe contener el nombre de la sede",
                 bg=C["bg"], fg="#777777", font=FS).pack(side="left", padx=10)

        lf = tk.Frame(df_lf, bg=C["bg"]); lf.pack(fill="x", pady=(6,0))
        self._lb = tk.Listbox(lf, height=5, font=FS,
                               bg=C["entry"], relief="flat",
                               highlightthickness=1, highlightcolor=C["accent"])
        sb_lb = ttk.Scrollbar(lf, orient="vertical", command=self._lb.yview)
        self._lb.configure(yscrollcommand=sb_lb.set)
        self._lb.pack(side="left", fill="x", expand=True); sb_lb.pack(side="right", fill="y")
        self._lbl_cnt = tk.Label(df_lf, text="0 archivos cargados",
                                  bg=C["bg"], fg="#777777", font=FS)
        self._lbl_cnt.pack(anchor="e")

        # Carpeta salida
        of = tk.Frame(s1, bg=C["bg"]); of.pack(fill="x", padx=24, pady=4)
        tk.Label(of, text="Carpeta de salida:", bg=C["bg"],
                 fg=C["label"], font=FB, width=22, anchor="w").pack(side="left")
        tk.Entry(of, textvariable=self._out_dir, font=FB,
                 bg=C["entry"], relief="solid", bd=1, width=42).pack(side="left")
        self._btn(of, "📁", lambda: self._out_dir.set(
            filedialog.askdirectory(title="Carpeta de salida") or self._out_dir.get()),
            "left", w=4)

        # ── Sec 2: Configuración
        s2 = self._sec(c, "⚙  2. Configuración de conciliación")

        # Radio modo
        tk.Label(s2, text="Modo de procesamiento:", bg=C["bg"],
                 fg=C["label"], font=("Arial",10,"bold")).pack(anchor="w", padx=24, pady=(8,2))

        rb_frame = tk.Frame(s2, bg=C["bg"]); rb_frame.pack(fill="x", padx=28, pady=(0,8))
        for val, lbl, bold in [
            (MODO_IND, "Una sede específica",                     False),
            (MODO_SEP, "🚀 Todas — un archivo por sede",          False),
            (MODO_UNI, "📦 Todas — un solo Excel consolidado",    True),
        ]:
            rb = tk.Radiobutton(rb_frame, text=lbl, variable=self._modo, value=val,
                                 command=self._on_modo,
                                 bg=C["bg"], fg=C["label"],
                                 activebackground=C["bg"],
                                 font=("Arial",10,"bold") if bold else FB,
                                 cursor="hand2")
            rb.pack(side="left", padx=(0,20))

        # Panel INDIVIDUAL
        self._p_ind = tk.Frame(s2, bg="#F0F4FA", relief="flat", bd=1)
        self._p_ind.pack(fill="x", padx=24, pady=(0,4))
        g = tk.Frame(self._p_ind, bg="#F0F4FA"); g.pack(fill="x", padx=12, pady=8)
        campos = [
            ("Sede a procesar:",   self._sede,    24, "Debe coincidir con la Nota del auxiliar (ej: ENVIGADO)"),
            ("Período (texto):",   self._periodo, 24, "Aparece en el encabezado del reporte"),
            ("Año del período:",   self._year,     8, "Ej: 2026"),
            ("Mes del período:",   self._mes,      8, "Ej: 1=Enero  2=Febrero  …"),
        ]
        for i,(lbl,var,w,hint) in enumerate(campos):
            tk.Label(g, text=lbl, bg="#F0F4FA", fg=C["label"], font=FB).grid(
                row=i, column=0, sticky="w", padx=4, pady=5)
            tk.Entry(g, textvariable=var, font=FB, width=w,
                     bg=C["entry"], relief="solid", bd=1).grid(
                row=i, column=1, sticky="w", padx=4, pady=5)
            tk.Label(g, text=hint, bg="#F0F4FA", fg="#777777", font=FS).grid(
                row=i, column=2, sticky="w", padx=8)
        tk.Button(g, text="🔍 Detectar sedes del auxiliar",
                   command=self._detectar_sedes,
                   bg=C["accent"], fg="white", font=FB, relief="flat",
                   cursor="hand2", padx=10, pady=4).grid(
            row=len(campos), column=0, columnspan=3, sticky="w", padx=4, pady=8)

        # Panel TODOS (sep y uni comparten los mismos campos de período)
        self._p_todos = tk.Frame(s2, bg="#EFF6E0", relief="flat", bd=1)
        self._p_todos.pack(fill="x", padx=24, pady=(0,4))
        tk.Label(self._p_todos, text="🚀  Modo «Generar Todos»",
                 bg="#EFF6E0", fg="#375623",
                 font=("Arial",11,"bold")).pack(anchor="w", padx=14, pady=(10,2))
        tk.Label(self._p_todos,
                 text="El sistema detecta automáticamente todas las sedes del auxiliar,\n"
                      "hace matching difuso con los archivos cargados, e informa los huérfanos.\n"
                      "📦 Modo consolidado: todo en un Excel (una hoja por sede + resumen + mapa de nombres).",
                 bg="#EFF6E0", fg="#375623", font=FS, justify="left").pack(anchor="w", padx=14)
        gt = tk.Frame(self._p_todos, bg="#EFF6E0"); gt.pack(fill="x", padx=14, pady=(4,10))
        tk.Label(gt, text="Período:", bg="#EFF6E0", fg=C["label"], font=FB).grid(
            row=0, column=0, sticky="w", padx=4, pady=4)
        tk.Entry(gt, textvariable=self._periodo, font=FB, width=18,
                 bg=C["entry"], relief="solid", bd=1).grid(row=0,column=1,sticky="w",padx=4,pady=4)
        tk.Label(gt, text="Año:", bg="#EFF6E0", fg=C["label"], font=FB).grid(
            row=0,column=2,sticky="w",padx=(16,4),pady=4)
        tk.Entry(gt, textvariable=self._year, font=FB, width=6,
                 bg=C["entry"], relief="solid", bd=1).grid(row=0,column=3,sticky="w",padx=4,pady=4)
        tk.Label(gt, text="Mes:", bg="#EFF6E0", fg=C["label"], font=FB).grid(
            row=0,column=4,sticky="w",padx=(12,4),pady=4)
        tk.Entry(gt, textvariable=self._mes, font=FB, width=4,
                 bg=C["entry"], relief="solid", bd=1).grid(row=0,column=5,sticky="w",padx=4,pady=4)

        self._on_modo()

        # Info reglas
        info = tk.Frame(s2, bg="#EAF0FB"); info.pack(fill="x", padx=24, pady=(4,8))
        tk.Label(info,
                 text="✔  Reglas activas:  cruce por Día(Nota)=Día(FechaVale)  |  valor=Neto  |  "
                      "matching difuso automático  |  anti-duplicado de hojas  |  log de auditoría",
                 bg="#EAF0FB", fg="#1F3864", font=FS, justify="left",
                 padx=12, pady=8).pack(anchor="w")

        # ── Sec 3: Ejecutar
        s3 = self._sec(c, "▶  3. Ejecutar conciliación")
        er = tk.Frame(s3, bg=C["bg"]); er.pack(fill="x", padx=24, pady=8)
        self._btn_ej = self._btn(er, "▶  EJECUTAR CONCILIACIÓN", self._ejecutar, "left", pri=True, w=28)
        self._btn_ab = self._btn(er, "📂  Abrir resultados",      self._abrir,    "left", sec=True, w=20)
        self._btn_ab.config(state="disabled")

        pf = tk.Frame(s3, bg=C["bg"]); pf.pack(fill="x", padx=24, pady=(2,0))
        self._pbar = ttk.Progressbar(pf, variable=self._prog, maximum=100, length=450, mode="determinate")
        self._pbar.pack(side="left")
        tk.Label(pf, textvariable=self._status,
                 bg=C["bg"], fg=C["label"], font=FS).pack(side="left", padx=10)

        # ── Sec 4: Log
        s4 = self._sec(c, "📋  4. Log de proceso")
        lf2 = tk.Frame(s4, bg=C["bg"]); lf2.pack(fill="x", padx=24, pady=4)
        self._log_txt = tk.Text(lf2, height=12, font=("Courier",9),
                                 bg="#1A1A2E", fg="#A8D8EA",
                                 insertbackground="white", relief="flat",
                                 wrap="word", state="disabled")
        sb2 = ttk.Scrollbar(lf2, orient="vertical", command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=sb2.set)
        self._log_txt.pack(side="left", fill="x", expand=True); sb2.pack(side="right", fill="y")
        tk.Label(c, text="", bg=C["bg"]).pack(pady=16)

    def _on_modo(self):
        modo = self._modo.get()
        if modo == MODO_IND:
            self._p_ind.pack(fill="x", padx=24, pady=(0,4))
            self._p_todos.pack_forget()
        else:
            self._p_ind.pack_forget()
            self._p_todos.pack(fill="x", padx=24, pady=(0,4))

    # ── Acciones ─────────────────────────────────────────────────────────────

    def _browse_aux(self):
        p = filedialog.askopenfilename(
            title="Seleccionar Auxiliar WorldOffice",
            filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
        if p: self._auxiliar.set(p); self._log(f"Auxiliar: {os.path.basename(p)}")

    def _browse_df(self):
        paths = filedialog.askopenfilenames(
            title="Seleccionar datafono(s)",
            filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
        for p in paths:
            if p not in self._df_paths:
                self._df_paths.append(p)
                self._lb.insert("end", f"  {os.path.basename(p)}")
                self._log(f"Datafono: {os.path.basename(p)}")
        self._lbl_cnt.config(text=f"{len(self._df_paths)} archivo(s)")

    def _clear_df(self):
        self._df_paths.clear(); self._lb.delete(0,"end")
        self._lbl_cnt.config(text="0 archivos cargados"); self._log("Lista limpiada")

    def _detectar_sedes(self):
        if not self._auxiliar.get():
            messagebox.showinfo("Aviso","Primero cargue el Auxiliar."); return
        try:
            sedes = sedes_disponibles(self._auxiliar.get())
            self._log(f"Sedes: {', '.join(sedes)}")
            messagebox.showinfo("Sedes disponibles",
                                "Sedes con DATAFONO en el auxiliar:\n\n" +
                                "\n".join(f"  • {s}" for s in sedes))
        except Exception as e: messagebox.showerror("Error",str(e))

    def _ejecutar(self):
        if not self._auxiliar.get():
            messagebox.showerror("Error","Seleccione el Auxiliar."); return
        if not self._df_paths:
            messagebox.showerror("Error","Agregue al menos un archivo de datafono."); return
        try: year,mes = int(self._year.get()), int(self._mes.get())
        except ValueError:
            messagebox.showerror("Error","Año y Mes deben ser enteros."); return
        self._btn_ej.config(state="disabled"); self._btn_ab.config(state="disabled")
        self._prog.set(0)
        modo = self._modo.get()
        target = self._run_uno if modo==MODO_IND else self._run_todos
        threading.Thread(target=target, args=(year,mes), daemon=True).start()

    # ── Run individual ───────────────────────────────────────────────────────

    def _run_uno(self, year, mes):
        try:
            sede_req = self._sede.get().strip().upper()
            if not sede_req:
                self.after(0,lambda: messagebox.showerror("Error","Ingrese el nombre de la sede.")); return
            self._log("="*52)
            self._log(f"MODO INDIVIDUAL — {sede_req} | {self._periodo.get()}")
            self._log("="*52)

            self._set_prog(10,"Leyendo auxiliar…")
            df_aux_full = leer_auxiliar(self._auxiliar.get())
            sedes       = sedes_disponibles(self._auxiliar.get())

            # Matching difuso para la sede solicitada
            self._set_prog(20,"Calculando matching difuso…")
            mapa = construir_mapa_nombres([sede_req], sedes)
            info = mapa[sede_req]
            sede_aux = info['sede']

            if not sede_aux:
                self.after(0,lambda: messagebox.showwarning(
                    "Sin coincidencia",
                    f"'{sede_req}' no coincide con ninguna sede del auxiliar.\n\n"
                    f"Sedes disponibles:\n" + "\n".join(f"  • {s}" for s in sedes)))
                self._set_prog(0,"Sin coincidencia"); self._btn_ej.config(state="normal"); return

            if not info['confirmado']:
                self._log(f"  ⚠ Match difuso: '{sede_req}' → '{sede_aux}' (score={info['score']:.0f}, tipo={info['tipo']})")
            else:
                self._log(f"  ✔ Match exacto: '{sede_aux}'")

            df_aux = df_aux_full[df_aux_full['Sede'].str.upper() == sede_aux.upper()].copy()
            self._log(f"Auxiliar: {len(df_aux)} registros DATAFONO {sede_aux}")

            self._set_prog(30,"Cargando datafono…")
            df_dat, info_arch = cargar_multiples_datafonos(self._df_paths)

            # Buscar el archivo que corresponde a la sede
            nombre_arch = next(
                (n for n,inf in construir_mapa_nombres(list(info_arch.keys()),[sede_aux]).items()
                 if inf['sede']), None)

            if not nombre_arch:
                self._log(f"  ⚠ Ningún archivo cargado corresponde a '{sede_aux}'")
                nombre_arch = list(info_arch.keys())[0]

            self._set_prog(55,"Cruzando…")
            por_dia, por_bolruta = agrupar_datafono_por_dia_vale(df_dat)
            resultados = cruzar_auxiliar_datafono(df_aux, por_dia, por_bolruta, nombre_arch, year, mes)

            cuadra  = sum(1 for r in resultados if r['estado']=='CUADRA')
            dif_m   = sum(1 for r in resultados if r['estado']=='DIF_MENOR')
            revisar = sum(1 for r in resultados if r['estado'] in ('DIFERENCIA','SIN_MATCH'))
            self._log(f"Resultado: {cuadra} cuadran | {dif_m} dif.menor | {revisar} revisar")

            self._set_prog(80,"Generando Excel…")
            ts   = datetime.now().strftime("%Y%m%d_%H%M")
            name = f"CONCILIACION_DF_{sede_aux}_{ts}.xlsx"
            out  = os.path.join(self._out_dir.get(), name)
            info_match = {**info, 'nombre_archivo': sede_req}
            generar_excel_resultado(resultados, sede_aux, out, self._periodo.get(), info_match)
            self._last_out = out; self._last_dir = self._out_dir.get()
            self._set_prog(100,f"✔ {name}")
            self._log(f"✔ {out}")
            self.after(0,lambda: self._btn_ab.config(state="normal"))
            self.after(0,lambda: messagebox.showinfo("Completado",
                f"✔ {sede_aux} — {self._periodo.get()}\n\n"
                f"  Cuadran:   {cuadra}\n  Dif menor: {dif_m}\n  Revisar:   {revisar}\n\n"
                f"Archivo:\n{out}"))
        except Exception as exc:
            import traceback
            self._log(f"❌ {exc}\n{traceback.format_exc()}")
            self._set_prog(0,f"Error: {str(exc)[:60]}")
            self.after(0,lambda: messagebox.showerror("Error",str(exc)))
        finally:
            self.after(0,lambda: self._btn_ej.config(state="normal"))

    # ── Run TODOS ─────────────────────────────────────────────────────────────

    def _run_todos(self, year, mes):
        try:
            modo = self._modo.get()
            self._log("="*60)
            self._log(f"MODO {'SEPARADOS' if modo==MODO_SEP else 'UNIFICADO'} — {self._periodo.get()}")
            self._log("="*60)

            self._set_prog(5,"Detectando sedes…")
            df_aux_full = leer_auxiliar(self._auxiliar.get())
            sedes       = sedes_disponibles(self._auxiliar.get())
            self._log(f"Sedes en auxiliar ({len(sedes)}): {', '.join(sedes)}")

            self._set_prog(10,"Cargando y mapeando datafonos…")
            df_dat, info_arch = cargar_multiples_datafonos(self._df_paths)
            nombres_arch = list(info_arch.keys())
            mapa = construir_mapa_nombres(nombres_arch, sedes)

            # Reportar matching
            self._log("─── MAPA DE NOMBRES ───")
            for nombre, info in sorted(mapa.items()):
                icono = "✔" if info['confirmado'] else ("⚠" if info['sede'] else "❌")
                self._log(f"  {icono} '{nombre}' → '{info['sede'] or 'SIN ASIGNAR'}' "
                           f"(score={info['score']:.0f}, tipo={info['tipo']})")

            huerfanos = [n for n,i in mapa.items() if not i['sede']]
            if huerfanos:
                self._log(f"  ⚠ HUÉRFANOS ({len(huerfanos)}): {', '.join(huerfanos)}")

            por_dia, por_bolruta = agrupar_datafono_por_dia_vale(df_dat)

            # Sedes sin datafono asignado
            sedes_asignadas = {i['sede'] for i in mapa.values() if i['sede']}
            sedes_sin_df    = [s for s in sedes if s not in sedes_asignadas]
            if sedes_sin_df:
                self._log(f"  ℹ Sedes sin datafono cargado: {', '.join(sedes_sin_df)}")

            ts = datetime.now().strftime("%Y%m%d_%H%M")
            if modo == MODO_SEP:
                out_dir = os.path.join(self._out_dir.get(), f"CONCILIACION_TODOS_{ts}")
                os.makedirs(out_dir, exist_ok=True)
                self._last_dir = out_dir

            resultados_por_sede = []
            resumen_sedes       = []
            n = len(sedes_asignadas)

            for i, (nombre_arch, info) in enumerate(mapa.items()):
                sede_aux = info['sede']
                if not sede_aux: continue
                prog = 15 + int(i/max(n,1)*65)
                self._set_prog(prog, f"Procesando {sede_aux} ({i+1}/{n})…")
                self._log(f"[{i+1}/{n}] {sede_aux}")

                df_aux = df_aux_full[df_aux_full['Sede'].str.upper() == sede_aux.upper()].copy()
                if df_aux.empty:
                    self._log(f"    ⚠ Sin registros — omitida"); continue

                resultados = cruzar_auxiliar_datafono(
                    df_aux, por_dia, por_bolruta, nombre_arch, year, mes)

                cuadra  = sum(1 for r in resultados if r['estado']=='CUADRA')
                dif_m   = sum(1 for r in resultados if r['estado']=='DIF_MENOR')
                revisar = sum(1 for r in resultados if r['estado'] in ('DIFERENCIA','SIN_MATCH','SIN_DIA'))
                self._log(f"    ✔ {cuadra}/{len(resultados)} cuadran | revisar: {revisar}")

                info_match = {**info, 'nombre_archivo': nombre_arch}
                resultados_por_sede.append({
                    'sede': sede_aux, 'nombre_archivo': nombre_arch,
                    'resultados': resultados, 'info_match': info_match
                })
                resumen_sedes.append({
                    'sede': sede_aux, 'nombre_df': nombre_arch,
                    'total_reg': len(resultados),
                    'cuadra': cuadra, 'dif_menor': dif_m, 'revisar': revisar,
                    'sin_df': sum(1 for r in resultados if r['estado']=='SIN_MATCH'),
                    'total_aux': sum(r['valor_aux']   for r in resultados),
                    'total_df':  sum(r['suma_grupos'] for r in resultados),
                    'diferencia':sum(r['valor_aux']   for r in resultados) -
                                  sum(r['suma_grupos'] for r in resultados),
                    'score':     info['score'], 'tipo_match': info['tipo'],
                })

                if modo == MODO_SEP:
                    out_xlsx = os.path.join(out_dir, f"DF_{sede_aux}.xlsx")
                    generar_excel_resultado(resultados, sede_aux, out_xlsx,
                                             self._periodo.get(), info_match)

            self._set_prog(88,"Generando archivo(s) de salida…")

            if modo == MODO_SEP:
                res_path = os.path.join(out_dir, f"RESUMEN_EJECUTIVO_{ts}.xlsx")
                generar_resumen_consolidado(resumen_sedes, res_path,
                                             self._periodo.get(), huerfanos)
                self._last_out = res_path
                self._log(f"✔ {len(resultados_por_sede)} Excel individuales + resumen en {out_dir}")
            else:
                out_file = os.path.join(self._out_dir.get(),
                                         f"CONCILIACION_UNIFICADA_{ts}.xlsx")
                generar_excel_unificado(resultados_por_sede, out_file,
                                         self._periodo.get(), mapa, huerfanos)
                self._last_out = out_file; self._last_dir = self._out_dir.get()
                self._log(f"✔ Excel unificado: {out_file}")

            sedes_ok  = sum(1 for r in resumen_sedes if r['revisar']==0)
            sedes_rev = sum(1 for r in resumen_sedes if r['revisar']>0)
            total_dif = sum(r['diferencia'] for r in resumen_sedes)

            self._set_prog(100,f"✔ {len(resumen_sedes)} sedes procesadas")
            self.after(0,lambda: self._btn_ab.config(state="normal"))

            msg_extra = ""
            if huerfanos:
                msg_extra = f"\n\n⚠  {len(huerfanos)} archivo(s) sin sede:\n" + "\n".join(f"  • {h}" for h in huerfanos)
            if sedes_sin_df:
                msg_extra += f"\n\nℹ  {len(sedes_sin_df)} sede(s) sin datafono cargado:\n" + "\n".join(f"  • {s}" for s in sedes_sin_df)

            self.after(0,lambda: messagebox.showinfo(
                "Completado",
                f"✔ {len(resumen_sedes)} sedes procesadas\n\n"
                f"  Sin diferencias: {sedes_ok}\n"
                f"  Con diferencias: {sedes_rev}\n"
                f"  Diferencia neta: ${total_dif:,.0f}"
                f"{msg_extra}"))

        except Exception as exc:
            import traceback
            self._log(f"❌ {exc}\n{traceback.format_exc()}")
            self._set_prog(0,f"Error: {str(exc)[:60]}")
            self.after(0,lambda: messagebox.showerror("Error",str(exc)))
        finally:
            self.after(0,lambda: self._btn_ej.config(state="normal"))

    def _abrir(self):
        target = self._last_dir or self._last_out or self._out_dir.get()
        if not target or not os.path.exists(target):
            target = self._out_dir.get()
        if sys.platform=="win32": os.startfile(target)
        elif sys.platform=="darwin": subprocess.call(["open",target])
        else: subprocess.call(["xdg-open",target])


if __name__ == "__main__":
    app = ConciliadorApp(); app.mainloop()
