"""
launcher.py — JAGI CAPS — Menú principal de sistemas de conciliación
Abre el módulo seleccionado como proceso independiente.
"""
import tkinter as tk
from tkinter import ttk
import subprocess, sys
from pathlib import Path

C = {"bg":"#1F3864","accent":"#2E75B6","hover":"#1A5A9F","white":"#FFFFFF"}

class Launcher(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("JAGI CAPS — Sistema de Conciliación")
        self.geometry("900x450"); self.resizable(False, False)
        self.configure(bg=C["bg"])
        self._build()

    def _build(self):
        tk.Label(self, text="JAGI CAPS", bg=C["bg"], fg=C["white"],
                 font=("Arial",22,"bold")).pack(pady=(28,4))
        tk.Label(self, text="Sistema de Conciliación Contable",
                 bg=C["bg"], fg="#8899BB", font=("Arial",11)).pack()
        tk.Label(self, text="─"*42, bg=C["bg"], fg="#2E4070",
                 font=("Arial",10)).pack(pady=12)

        modulos = [
            ("🏪  Conciliación de Datafonos",  "app_conciliador.py"),
            ("🏦  Conciliación Bancaria",       "conciliador_bancario/app_bancario.py"),
        ]
        for lbl, script in modulos:
            b = tk.Button(self, text=lbl,
                          command=lambda s=script: self._abrir(s),
                          bg=C["accent"], fg=C["white"],
                          font=("Arial",12), relief="flat",
                          cursor="hand2", padx=20, pady=10, width=30)
            b.pack(pady=6)
            b.bind("<Enter>", lambda e, btn=b: btn.config(bg=C["hover"]))
            b.bind("<Leave>", lambda e, btn=b: btn.config(bg=C["accent"]))

        tk.Label(self, text="Distribuidora · Jaime Wilson · Jagi Industry",
                 bg=C["bg"], fg="#445577", font=("Arial",9)).pack(side="bottom", pady=10)

    def _abrir(self, script: str):
        path = Path(__file__).parent / script
        subprocess.Popen([sys.executable, str(path)])


if __name__ == "__main__":
    app = Launcher(); app.mainloop()