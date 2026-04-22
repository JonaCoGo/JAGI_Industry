# renombrar_datafonos.py

"""
  Script para renombrar, en la carpeta donde se ejecuta, todos los archivos
  de Excel que contengan la palabra “datafono” en su nombre.

  Los archivos serán renombrados con el formato:

      datafono_<NOMBRE_DE_LA_TIENDA>.xlsx

  donde <NOMBRE_DE_LA_TIENDA> se obtiene de la lista oficial de sedes

  Si el nombre de la tienda no se reconoce, el archivo se deja sin cambios
  y se registra una advertencia.

  Requisitos
  ----------
  - Python 3.9+ (el entorno del proyecto usa `python`).
  - Sólo librerías estándar (pathlib, logging, sys, re, unicodedata).

  Uso
  ----
  Coloca este archivo en la carpeta que desea procesar y ejecútalo:

      python renombrar_datafonos.py

  El script es idempotente: si el archivo ya tiene el nombre esperado,
  no se realiza ninguna operación.
  """

import logging
import re
import sys
import unicodedata
from pathlib import Path
from typing import Optional

# ----------------------------------------------------------------------
# Configuración de logging (cumple con el estándar de “log obligatorio”)
# ----------------------------------------------------------------------
LOG_FORMAT = "%(asctime)s | %(levelname)s | %(message)s"
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler("renombrar_datafonos.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

logger = logging.getLogger(__name__)

# ----------------------------------------------------------------------
# Función auxiliar para quitar acentos ( insensible a tilde )
# ----------------------------------------------------------------------
def quitar_acentos(texto: str) -> str:
    """
    Devuelve una copia de `texto` con todos los acentos eliminados.
    """
    texto_normalizado = unicodedata.normalize('NFD', texto)
    return ''.join(
        c for c in texto_normalizado
        if not unicodedata.combining(c)
    )

# ----------------------------------------------------------------------
# Lista oficial de sedes (tiendas) – extraída de conciliador_engine.py
# ----------------------------------------------------------------------
SEDES_OFICIALES = [
      "Barranquilla", "Buenavista", "Centro Mayor", "Chipichape",
      "Eden", "Envigado", "Fabricato", "Plaza Imperial",
      "Jardin Plaza", "Mercurio", "Molinos", "Nuestro Bogota",
      "Pasto", "Puerta del Norte", "Santa Marta", "Santa Fe",
      "Sincelejo", "Parque Alegra", "Americas", "Cacique",
      "Tesoro", "Titan Plaza", "Cali", "Pereira", "Titan","Serrezuela"
  ]

# patrón para identificar la palabra “datafono” (ignorando mayúsculas)
PAT_DATAFONO = re.compile(r"datafono", re.IGNORECASE)

def buscar_tienda(nombre_archivo: str) -> Optional[str]:
    """
    Busca, dentro del nombre del archivo, alguna coincidencia con los
    nombres de las tiendas declaradas en `SEDES_OFICIALES`.

    Se hace una comparación insensible a mayúsculas, espacios, guiones,
    guiones bajos y acentos (tilde).

    Retorna el nombre exacto de la tienda tal como aparece en la lista,
    o `None` si no se encuentra coincidencia.
    """
    # Quitamos acentos y pasamos a minúsculas para la comparación
    nombre_limpio = quitar_acentos(nombre_archivo).lower()
    for sede in SEDES_OFICIALES:
        # normalizamos la sede para la comparación (sin acentos, minúsculas, sin espacios/guiones)
        sede_normal = quitar_acentos(sede).lower().replace(" ", "").replace("-", "").replace("_", "")
        # también limpiamos el nombre del archivo
        archivo_normal = nombre_limpio.replace(" ", "").replace("-", "").replace("_", "")
        if sede_normal in archivo_normal:
            return sede  # devolvemos la versión original (con mayúsculas y espacios)
    return None


def renombrar_archivo(ruta: Path) -> None:
    """
    Renombra un archivo que contiene la palabra “datafono”.
    El nuevo nombre será: datafono_<NOMBRE_DE_LA_TIENDA>.xlsx
    """
    if not ruta.is_file():
        logger.warning(f"No es un archivo válido: {ruta}")
        return

    # Solo procesamos archivos de Excel
    if ruta.suffix.lower() not in (".xlsx", ".xls"):
        logger.debug(f"Se omite archivo no Excel: {ruta.name}")
        return

    # Verificamos que el nombre contenga “datafono”
    if not PAT_DATAFONO.search(ruta.name):
        logger.debug(f"Archivo sin 'datafono' en el nombre: {ruta.name}")
        return

    # Intentamos obtener el nombre de la tienda
    nombre_tienda = buscar_tienda(ruta.stem)
    if not nombre_tienda:
        logger.warning(
            f"No se pudo identificar la tienda en el archivo: {ruta.name}. "
            "Se deja sin renombrar."
        )
        return

    nuevo_nombre = f"datafono_{nombre_tienda}.xlsx"
    nuevo_path = ruta.with_name(nuevo_nombre)

    # Evitamos sobrescribir un archivo existente
    if nuevo_path.exists():
        logger.error(
            f"Destino ya existe: {nuevo_path}. "
            f"Archivo original {ruta.name} NO será renombrado."
        )
        return

    try:
        ruta.rename(nuevo_path)
        logger.info(f"Renombrado: {ruta.name} → {nuevo_nombre}")
    except OSError as exc:
        logger.exception(
            f"Error al renombrar {ruta.name} a {nuevo_nombre}: {exc}"
        )


def main() -> None:
    """
    Recorre la carpeta donde se ejecuta el script y renombra los archivos
    que cumplan los criterios.
    """
    # La carpeta objetivo es la misma donde está este script
    carpeta_objetivo = Path(__file__).parent.resolve()

    if not carpeta_objetivo.is_dir():
        logger.critical(f"Carpeta objetivo no encontrada: {carpeta_objetivo}")
        sys.exit(1)

    logger.info(f"Iniciando proceso en: {carpeta_objetivo}")

    for elemento in carpeta_objetivo.iterdir():
        if elemento.is_file():
            renombrar_archivo(elemento)

    logger.info("Proceso completado.")


if __name__ == "__main__":
    main()
