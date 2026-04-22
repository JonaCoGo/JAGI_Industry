"""
config/empresas.py — JAGI CAPS
Fuente de verdad única para razones sociales, cuentas bancarias y tiendas.
NO contiene lógica de negocio.
"""

EMPRESAS = {
    "JAIME_WILSON": {
        "razon_social":   "Jaime Wilson Giraldo Giraldo",
        "nombre_corto":   "JAIME WILSON",
        "marca":          "JAGI CAPS",
        "nit":            "",          # completar
        "cuentas": {
            "Bancolombia 8821": {"banco": "Bancolombia", "tipo": "Ahorros",   "ultimos": "8821"},
            "Davivienda 2346":  {"banco": "Davivienda",  "tipo": "Corriente", "ultimos": "2346"},
        },
        "cuenta_default": "Davivienda 2346",
        "tiendas": [
            "Chipichape", "Jardin Plaza", "Nuestro Bogota", "Mercurio",
            "Barranquilla", "Fabricato", "Puerta del Norte", "Envigado",
            "Sincelejo", "Molinos", "Santa Fe", "Eden",
            "Buenavista", "Pereira", "Palmas Mall", "Serrezuela",
        ],
    },
    "DISTRIBUIDORA": {
        "razon_social":   "Distribuidora de Accesorios Giraldo SAS",
        "nombre_corto":   "DISTRIBUIDORA",
        "marca":          "JAGI CAPS",
        "nit":            "",          # completar
        "cuentas": {
            "Bancolombia 3263": {"banco": "Bancolombia", "tipo": "Ahorros", "ultimos": "3263"},
            "Davivienda 4677":  {"banco": "Davivienda",  "tipo": "Ahorros", "ultimos": "4677"},
        },
        "cuenta_default": "Davivienda 4677",
        "tiendas": [
            "Unicentro Cali", "Plaza Imperial", "Centro Mayor",
            "Titan Plaza", "Cacique", "Tesoro", "Plaza de las Americas",
        ],
    },
    "JAGI_INDUSTRY": {
        "razon_social":   "Jagi Industry SAS",
        "nombre_corto":   "JAGI INDUSTRY",
        "marca":          "JAGI CAPS",
        "nit":            "",          # completar
        "cuentas": {
            "Bancolombia 7162":  {"banco": "Bancolombia",   "tipo": "Ahorros", "ultimos": "7162"},
            "Davivienda 6816":   {"banco": "Davivienda",    "tipo": "Ahorros", "ultimos": "6816"},
            "Banco Bogotá 1014": {"banco": "Banco de Bogotá","tipo": "Ahorros", "ultimos": "1014"},
        },
        "cuenta_default": "Bancolombia 7162",
        "tiendas": [
            "Tienda en Linea",
        ],
    },
}

# Orden de presentación en la UI
ORDEN_EMPRESAS = ["JAIME_WILSON", "DISTRIBUIDORA", "JAGI_INDUSTRY"]


def get_empresa(key: str) -> dict:
    return EMPRESAS[key]


def opciones_ui() -> list[tuple]:
    """Retorna [(label_display, key), ...] para el selector de la UI."""
    return [
        (f"{EMPRESAS[k]['razon_social']}", k)
        for k in ORDEN_EMPRESAS
    ]


def cuentas_empresa(key: str) -> list[str]:
    return list(EMPRESAS[key]["cuentas"].keys())


def label_banner(key: str) -> str:
    e = EMPRESAS[key]
    return f"{e['marca']} — {e['razon_social']}"


def tiendas_empresa(key: str) -> list[str]:
    return EMPRESAS[key]["tiendas"]