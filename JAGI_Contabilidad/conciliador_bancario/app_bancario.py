"""
Interfaz Gráfica — Conciliación Bancaria JAGI CAPS v1.2
Tkinter · Compatible Windows / Mac / Linux
Cruza Auxiliar WorldOffice ↔ Extracto Bancario
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, subprocess, sys
from datetime import datetime
from pathlib import Path

import sys; sys.path.insert(0, str(Path(__file__).parent.parent))

from config.empresas import EMPRESAS, ORDEN_EMPRESAS, opciones_ui, cuentas_empresa
from conciliador_bancario.engine_bancario import (
    leer_auxiliar_bancario, leer_extracto, cruzar_auxiliar_extracto,
    generar_excel_bancario
)

C = {"bg":"#F5F7FA","sidebar":"#1F3864","accent":"#2E75B6",
     "label":"#333333","entry":"#FFFFFF","hover":"#1A5A9F",
     "verde":"#276221","naranja":"#843C0C"}
FT = ("Arial",15,"bold"); FH = ("Arial",11,"bold")
FB = ("Arial",10);        FS = ("Arial",9)


class ConciliacionBancariaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JAGI CAPS — Conciliación Bancaria v1.2")
        self.geometry("960x740"); self.minsize(820,600)
        self.configure(bg=C["bg"])

        self._empresa_key  = tk.StringVar(value=ORDEN_EMPRESAS[0])
        self._cuenta_sel   = tk.StringVar(value="")
        self._banco_sel    = tk.StringVar(value="")
        self._auxiliar     = tk.StringVar()
        self._extracto     = tk.StringVar()
        self._out_dir      = tk.StringVar(value=str(Path.home()/"Desktop"))
        self._periodo      = tk.StringVar(value="Enero 2025")
        self._year         = tk.StringVar(value="2025")
        self._mes          = tk.StringVar(value="1")
        self._status       = tk.StringVar(value="Listo")
        self._prog         = tk.IntVar(value=0)
        self._last_out     = None
        self._opciones_emp = opciones_ui()
        self._keys_emp     = {lbl: k for lbl, k in self._opciones_emp}

        self._build_ui()

    def _build_ui(self):
        # Sidebar
        sb = tk.Frame(self, bg=C["sidebar"], width=210)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)
        tk.Label(sb, text="🏦 Conciliación\nBancaria",
                 bg=C["sidebar"], fg="white",
                 font=("Arial",13,"bold"), justify="center").pack(pady=(24,6), padx=12)
        tk.Label(sb, text="──────────────────",
                 bg=C["sidebar"], fg="#445577", font=FS).pack()
        for lbl, cmd in [
            ("📂  Cargar archivos",  lambda: self._canvas.yview_moveto(0)),
            ("▶   Ejecutar",        self._ejecutar),
            ("📋  Ver resultados",  self._abrir),
        ]:
            b = tk.Button(sb, text=lbl, command=cmd,
                          bg=C["sidebar"], fg="white",
                          activebackground=C["accent"],
                          relief="flat", anchor="w", font=FB,
                          padx=16, pady=9, cursor="hand2")
            b.pack(fill="x")
            b.bind("<Enter>", lambda e, btn=b: btn.config(bg=C["accent"]))
            b.bind("<Leave>", lambda e, btn=b: btn.config(bg=C["sidebar"]))
        tk.Label(sb, text="\nv1.2 | JAGI CAPS\nNIIF PYMES / DIAN\nAuxiliar ↔ Extracto",
                 bg=C["sidebar"], fg="#8899BB",
                 font=FS, justify="center").pack(side="bottom", pady=16)

        # Main scrollable
        main = tk.Frame(self, bg=C["bg"]); main.pack(side="left", fill="both", expand=True)
        canvas = tk.Canvas(main, bg=C["bg"], highlightthickness=0)
        vsb    = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
        self._canvas = canvas
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        self._frame = tk.Frame(canvas, bg=C["bg"])
        win = canvas.create_window((0,0), window=self._frame, anchor="nw")
        self._frame.bind("<Configure>",
                          lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        self._build_content()

    def _sec(self, parent, title):
        f = tk.LabelFrame(parent, text=f"  {title}  ",
                           bg=C["bg"], fg=C["accent"],
                           font=FH, padx=4, pady=8, relief="groove", bd=1)
        f.pack(fill="x", padx=24, pady=10); return f

    def _log(self, msg):
        self._log_txt.config(state="normal")
        self._log_txt.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self._log_txt.see("end"); self._log_txt.config(state="disabled")
        self.update_idletasks()

    def _set_prog(self, v, msg):
        self._prog.set(v); self._status.set(msg); self.update_idletasks()

    def _build_content(self):
        c = self._frame

        # Banner
        bn = tk.Frame(c, bg=C["accent"], height=74); bn.pack(fill="x")
        tk.Label(bn, text="Conciliación Bancaria — Auxiliar ↔ Extracto",
                 bg=C["accent"], fg="white", font=FT).pack(side="left", padx=24, pady=20)
        self._lbl_banner = tk.Label(bn, text="",
                 bg=C["accent"], fg="#CCE0F5", font=FS, justify="right")
        self._lbl_banner.pack(side="right", padx=20)

        # Sec 0 — Empresa y cuenta
        s0 = self._sec(c, "🏢  0. Empresa y cuenta bancaria")
        ef = tk.Frame(s0, bg=C["bg"]); ef.pack(fill="x", padx=24, pady=6)
        tk.Label(ef, text="Razón social:", bg=C["bg"], fg=C["label"],
                 font=FB, width=18, anchor="w").pack(side="left")
        self._combo_emp = ttk.Combobox(ef,
            values=[lbl for lbl,_ in self._opciones_emp],
            state="readonly", width=46, font=FB)
        self._combo_emp.set(self._opciones_emp[0][0])
        self._combo_emp.pack(side="left", padx=4)
        self._combo_emp.bind("<<ComboboxSelected>>", lambda e: self._on_empresa())

        cf = tk.Frame(s0, bg=C["bg"]); cf.pack(fill="x", padx=24, pady=4)
        tk.Label(cf, text="Cuenta bancaria:", bg=C["bg"], fg=C["label"],
                 font=FB, width=18, anchor="w").pack(side="left")
        self._combo_cuenta = ttk.Combobox(cf, textvariable=self._cuenta_sel,
                                            state="readonly", width=30, font=FB)
        self._combo_cuenta.pack(side="left", padx=4)
        self._combo_cuenta.bind("<<ComboboxSelected>>", lambda e: self._on_cuenta())
        self._lbl_tipo = tk.Label(cf, text="", bg=C["bg"], fg="#555555", font=FS)
        self._lbl_tipo.pack(side="left", padx=8)
        self._on_empresa()

        # Sec 1 — Archivos
        s1 = self._sec(c, "📂  1. Carga de archivos")

        af = tk.Frame(s1, bg=C["bg"]); af.pack(fill="x", padx=24, pady=4)
        tk.Label(af, text="Auxiliar WorldOffice:", bg=C["bg"], fg=C["label"],
                 font=FB, width=22, anchor="w").pack(side="left")
        tk.Entry(af, textvariable=self._auxiliar, font=FB,
                 bg=C["entry"], relief="solid", bd=1, width=34).pack(side="left")
        tk.Button(af, text="📁", command=self._browse_aux,
                  bg=C["accent"], fg="white", font=FB, relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=4)

        xf = tk.Frame(s1, bg=C["bg"]); xf.pack(fill="x", padx=24, pady=4)
        tk.Label(xf, text="Extracto bancario (.xlsx):", bg=C["bg"], fg=C["label"],
                 font=FB, width=22, anchor="w").pack(side="left")
        tk.Entry(xf, textvariable=self._extracto, font=FB,
                 bg=C["entry"], relief="solid", bd=1, width=34).pack(side="left")
        tk.Button(xf, text="📁", command=self._browse_ext,
                  bg=C["accent"], fg="white", font=FB, relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=4)

        of = tk.Frame(s1, bg=C["bg"]); of.pack(fill="x", padx=24, pady=4)
        tk.Label(of, text="Carpeta de salida:", bg=C["bg"], fg=C["label"],
                 font=FB, width=22, anchor="w").pack(side="left")
        tk.Entry(of, textvariable=self._out_dir, font=FB,
                 bg=C["entry"], relief="solid", bd=1, width=34).pack(side="left")
        tk.Button(of, text="📁",
                  command=lambda: self._out_dir.set(
                      filedialog.askdirectory(title="Carpeta de salida") or self._out_dir.get()),
                  bg=C["accent"], fg="white", font=FB, relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=4)

        # Sec 2 — Configuración
        s2 = self._sec(c, "⚙  2. Período")
        gf = tk.Frame(s2, bg=C["bg"]); gf.pack(fill="x", padx=24, pady=6)
        for lbl, var, w in [
            ("Período (texto):", self._periodo, 18),
            ("Año:",             self._year,     6),
            ("Mes:",             self._mes,      4),
        ]:
            tk.Label(gf, text=lbl, bg=C["bg"], fg=C["label"], font=FB).pack(side="left", padx=(0,4))
            tk.Entry(gf, textvariable=var, font=FB, width=w,
                     bg=C["entry"], relief="solid", bd=1).pack(side="left", padx=(0,16))

        # Sec 3 — Ejecutar
        s3 = self._sec(c, "▶  3. Ejecutar conciliación bancaria")
        er = tk.Frame(s3, bg=C["bg"]); er.pack(fill="x", padx=24, pady=8)
        self._btn_ej = tk.Button(er, text="▶  EJECUTAR CONCILIACIÓN BANCARIA",
                                  command=self._ejecutar,
                                  bg=C["accent"], fg="white", font=FB,
                                  relief="flat", cursor="hand2", padx=10, pady=5, width=34)
        self._btn_ej.pack(side="left", padx=4)
        self._btn_ab = tk.Button(er, text="📂  Abrir resultado",
                                  command=self._abrir,
                                  bg="#E0E0E0", fg="#333333", font=FB,
                                  relief="flat", cursor="hand2", padx=10, pady=5,
                                  state="disabled")
        self._btn_ab.pack(side="left", padx=4)

        pf = tk.Frame(s3, bg=C["bg"]); pf.pack(fill="x", padx=24, pady=(2,0))
        ttk.Progressbar(pf, variable=self._prog, maximum=100, length=420,
                         mode="determinate").pack(side="left")
        tk.Label(pf, textvariable=self._status,
                 bg=C["bg"], fg=C["label"], font=FS).pack(side="left", padx=10)

        # Sec 4 — Log
        s4 = self._sec(c, "📋  4. Log de proceso")
        lf = tk.Frame(s4, bg=C["bg"]); lf.pack(fill="x", padx=24, pady=4)
        self._log_txt = tk.Text(lf, height=10, font=("Courier",9),
                                 bg="#1A1A2E", fg="#A8D8EA",
                                 insertbackground="white", relief="flat",
                                 wrap="word", state="disabled")
        sb2 = ttk.Scrollbar(lf, orient="vertical", command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=sb2.set)
        self._log_txt.pack(side="left", fill="x", expand=True)
        sb2.pack(side="right", fill="y")
        tk.Label(c, text="", bg=C["bg"]).pack(pady=16)

    # ── Callbacks empresa/cuenta ─────────────────────────────────────────────

    def _on_empresa(self):
        lbl = self._combo_emp.get()
        key = self._keys_emp.get(lbl, ORDEN_EMPRESAS[0])
        self._empresa_key.set(key)
        emp = EMPRESAS[key]
        cuentas = list(emp["cuentas"].keys())
        self._combo_cuenta["values"] = cuentas
        self._combo_cuenta.set(emp["cuenta_default"])
        self._on_cuenta()

    def _on_cuenta(self):
        key  = self._empresa_key.get()
        csel = self._cuenta_sel.get()
        emp  = EMPRESAS[key]
        info = emp["cuentas"].get(csel, {})
        banco = info.get("banco","")
        tipo  = info.get("tipo","")
        self._banco_sel.set(banco)
        self._lbl_tipo.config(text=f"{banco} · {tipo}" if banco else "")
        self._lbl_banner.config(
            text=f"{emp['razon_social']}\n{csel}")

    # ── Acciones ─────────────────────────────────────────────────────────────

    def _browse_aux(self):
        p = filedialog.askopenfilename(
            title="Auxiliar WorldOffice",
            filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
        if p: self._auxiliar.set(p); self._log(f"Auxiliar: {os.path.basename(p)}")

    def _browse_ext(self):
        p = filedialog.askopenfilename(
            title="Extracto bancario",
            filetypes=[("Excel","*.xlsx *.xls"),("CSV","*.csv"),("Todos","*.*")])
        if p: self._extracto.set(p); self._log(f"Extracto: {os.path.basename(p)}")

    def _ejecutar(self):
        if not self._auxiliar.get():
            messagebox.showerror("Error","Seleccione el Auxiliar WorldOffice."); return
        if not self._extracto.get():
            messagebox.showerror("Error","Seleccione el extracto bancario."); return
        if not self._cuenta_sel.get():
            messagebox.showerror("Error","Seleccione la cuenta bancaria."); return
        try:
            year = int(self._year.get()); mes = int(self._mes.get())
        except ValueError:
            messagebox.showerror("Error","Año y Mes deben ser enteros."); return
        self._btn_ej.config(state="disabled"); self._btn_ab.config(state="disabled")
        self._prog.set(0)
        threading.Thread(target=self._run, args=(year, mes), daemon=True).start()

    def _run(self, year: int, mes: int):
        try:
            banco   = self._banco_sel.get()
            cuenta  = self._cuenta_sel.get()
            emp_key = self._empresa_key.get()
            emp     = EMPRESAS[emp_key]

            self._log("="*52)
            self._log(f"CONCILIACIÓN BANCARIA — {cuenta} | {self._periodo.get()}")
            self._log("="*52)

            self._set_prog(15, "Leyendo auxiliar…")
            df_aux = leer_auxiliar_bancario(self._auxiliar.get(), year, mes)
            self._log(f"Auxiliar: {len(df_aux)} movimientos en el período")

            self._set_prog(35, "Leyendo extracto bancario…")
            df_ext = leer_extracto(self._extracto.get(), banco)
            self._log(f"Extracto {banco}: {len(df_ext)} movimientos")

            self._set_prog(60, "Cruzando auxiliar ↔ extracto…")
            resultado = cruzar_auxiliar_extracto(df_aux, df_ext, year, mes)
            res       = resultado['resumen']
            self._log(f"Datáfonos retirados: {res['datafonos_retirados']} movimientos "
                      f"(${res['datafonos_valor']:,.0f})")
            self._log(f"Cruce general: {res['cuadra']} cuadran | "
                      f"{res['dif_menor']} dif.≤1% | "
                      f"{res['sin_match']} sin match | "
                      f"{res['solo_banco']} solo en banco")
            self._log(f"Nómina — {res['nom_fechas']} fechas: "
                      f"{res['nom_cuadra']} cuadran | "
                      f"{res['nom_dif']} con dif. | "
                      f"{res['nom_sin_match']} sin match | "
                      f"dif. neta: ${res['nom_diferencia']:,.0f}")

            self._set_prog(80, "Generando Excel…")
            ts      = datetime.now().strftime("%Y%m%d_%H%M")
            cuenta_f = cuenta.replace(" ","_")
            name    = f"CONC_BANCARIA_{emp['nombre_corto'].replace(' ','_')}_{cuenta_f}_{ts}.xlsx"
            out     = os.path.join(self._out_dir.get(), name)
            info_emp = {"razon_social": emp["razon_social"], "cuenta": cuenta}
            generar_excel_bancario(resultado, cuenta, banco, out,
                                    self._periodo.get(), info_empresa=info_emp)
            self._last_out = out
            self._set_prog(100, f"✔ {name}")
            self._log(f"✔ {out}")
            self.after(0, lambda: self._btn_ab.config(state="normal"))
            self.after(0, lambda: messagebox.showinfo(
                "Completado",
                f"✔ Conciliación bancaria finalizada\n\n"
                f"── CRUCE GENERAL ──\n"
                f"  Cuadran:        {res['cuadra']}\n"
                f"  Dif. ≤1%:       {res['dif_menor']}\n"
                f"  Sin match:      {res['sin_match']}\n"
                f"  Solo en banco:  {res['solo_banco']}\n"
                f"  Dif. neta:      ${res['diferencia_neta']:,.0f}\n\n"
                f"── NÓMINA ──\n"
                f"  Fechas:         {res['nom_fechas']}\n"
                f"  Cuadran:        {res['nom_cuadra']}\n"
                f"  Con diferencia: {res['nom_dif']}\n"
                f"  Sin match:      {res['nom_sin_match']}\n"
                f"  Dif. neta:      ${res['nom_diferencia']:,.0f}\n\n"
                f"── DATÁFONOS ──\n"
                f"  Retirados:      {res['datafonos_retirados']} movimientos\n"
                f"  Valor:          ${res['datafonos_valor']:,.0f}\n\n"
                f"Archivo:\n{out}"))

        except NotImplementedError as nie:
            self._log(f"⚠ Lector pendiente: {nie}")
            self.after(0, lambda: messagebox.showwarning("Lector pendiente", str(nie)))
            self._set_prog(0, "Pendiente")
        except Exception as exc:
            import traceback
            self._log(f"❌ {exc}\n{traceback.format_exc()}")
            self._set_prog(0, f"Error: {str(exc)[:60]}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, lambda: self._btn_ej.config(state="normal"))

    def _abrir(self):
        target = self._last_out or self._out_dir.get()
        if not target or not os.path.exists(target):
            target = self._out_dir.get()
        if sys.platform == "win32":   os.startfile(target)
        elif sys.platform == "darwin": subprocess.call(["open", target])
        else:                          subprocess.call(["xdg-open", target])


if __name__ == "__main__":
    app = ConciliacionBancariaApp(); app.mainloop()