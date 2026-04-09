# agent.py – Subagente de Contabilidad
"""
Subagente especializado para procesos contables de JAGI Industry.
Actualmente expone la funcionalidad de conciliación bancaria mediante
el script `conciliacion_bancaria.py` bajo `Scripts/`.

Uso básico:
    from agent import run_conciliacion
    run_conciliacion(
        aux_path="path/al/auxiliar.xlsx",
        ext_path="path/al/extracto.xlsx",
        out_path="path/al/resultados.xlsx",
        mes=None  # o número de mes 1-12
    )

El agente puede ampliarse con nuevas funciones (causación, IVA, etc.)
consultando la documentación en `Agente_Contabilidad/CLAUDE.md`.
"""

import os
import sys
from pathlib import Path

# Añadir el directorio Scripts al PYTHONPATH
ROOT_DIR = Path(__file__).resolve().parents[2]  # JAGI Industry
SCRIPTS_DIR = ROOT_DIR / "Scripts"
sys.path.append(str(SCRIPTS_DIR))

try:
    import conciliacion_bancaria
    conciliar = conciliacion_bancaria.concilia
    exportar_excel = conciliacion_bancaria.exportar_excel
    cargar_auxiliar = conciliacion_bancaria.cargar_auxiliar
    cargar_extracto = conciliacion_bancaria.cargar_extracto
except Exception as e:
    raise ImportError(f"No se pudo cargar conciliacion_bancaria: {e}")

def run_conciliacion(aux_path: str, ext_path: str, out_path: str, mes: int | None = None):
    """Ejecuta la conciliación bancaria.

    Args:
        aux_path: Ruta al archivo del Libro Auxiliar (.xlsx).
        ext_path: Ruta al archivo del Extracto Bancario (.xlsx).
        out_path: Ruta donde se guardará el resultado Excel.
        mes: Opcional, número de mes a filtrar (1‑12). ``None`` para todos.
    """
    aux_path = str(Path(aux_path).resolve())
    ext_path = str(Path(ext_path).resolve())
    out_path = str(Path(out_path).resolve())

    if not Path(aux_path).is_file():
        raise FileNotFoundError(f"Archivo auxiliar no encontrado: {aux_path}")
    if not Path(ext_path).is_file():
        raise FileNotFoundError(f"Archivo de extracto no encontrado: {ext_path}")

    # Cargar datos
    df_aux = cargar_auxiliar(aux_path)
    df_ext = cargar_extracto(ext_path)

    # Ejecutar lógica de conciliación
    df_conc, df_paux, df_pext, df_err = conciliar(df_aux, df_ext, mes)

    # Exportar a Excel
    exportar_excel(df_conc, df_paux, df_pext, df_err, out_path)
    return {
        "conciliadas": len(df_conc),
        "pendientes_auxiliar": len(df_paux),
        "pendientes_banco": len(df_pext),
        "errores": len(df_err),
        "output": out_path,
    }

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Ejecutar subagente de conciliación bancaria")
    parser.add_argument("--aux", required=True, help="Ruta al Libro Auxiliar (xlsx)")
    parser.add_argument("--ext", required=True, help="Ruta al Extracto Bancario (xlsx)")
    parser.add_argument("--out", required=True, help="Ruta del archivo de salida (xlsx)")
    parser.add_argument("--mes", type=int, choices=range(1,13), help="Mes a filtrar (1‑12), por defecto todos")
    args = parser.parse_args()
    try:
        result = run_conciliacion(args.aux, args.ext, args.out, args.mes)
        print("Resultado:")
        for k, v in result.items():
            print(f"  {k}: {v}")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)