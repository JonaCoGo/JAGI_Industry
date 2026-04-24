import re
import pandas as pd
from pathlib import Path


def limpiar_texto(texto):
    return re.sub(r"\s+", " ", texto).strip()


def limpiar_numero(valor):
    if not valor:
        return 0.0
    valor = re.sub(r"[^\d,\.]", "", valor)
    return float(valor.replace(",", "")) if valor else 0.0


def procesar_extracto(ruta_txt, ruta_salida):
    try:
        with open(ruta_txt, "r", encoding="latin-1") as f:
            lineas = f.readlines()

        rows = []

        for linea in lineas:
            linea = linea.strip()

            # 🔥 Solo líneas que empiezan con fecha (evita resúmenes)
            if not linea or not re.match(r"\d{2}\s+\d{2}", linea):
                continue

            try:
                # =========================
                # 🔥 1. EXTRAER VALORES $
                # =========================
                valores = re.findall(r"\$\s*[\d,\.]+", linea)

                if len(valores) < 2:
                    continue

                # 🔥 Detectar si hay saldo o no
                if len(valores) >= 3:
                    debito = limpiar_numero(valores[-3])
                    credito = limpiar_numero(valores[-2])
                    saldo = limpiar_numero(valores[-1])
                    cantidad = 3
                else:
                    debito = limpiar_numero(valores[-2])
                    credito = limpiar_numero(valores[-1])
                    saldo = None
                    cantidad = 2

                # =========================
                # 🔥 2. ELIMINAR VALORES $
                # =========================
                texto_base = linea
                for v in valores[-cantidad:]:
                    texto_base = texto_base.replace(v, "")

                partes = texto_base.split()

                # Validación mínima
                if len(partes) < 3:
                    continue

                # =========================
                # 🔥 3. FECHA
                # =========================
                dia = int(partes[0])
                mes = int(partes[1])

                # =========================
                # 🔥 4. OFICINA
                # =========================
                if len(partes) > 2 and re.match(r"^\d{3,4}$", partes[2]):
                    oficina = partes[2]
                    resto = partes[3:]
                else:
                    oficina = "0000"
                    resto = partes[2:]

                # =========================
                # 🔥 5. DOCUMENTO + DESCRIPCIÓN
                # =========================
                documento = None
                descripcion = ""

                for i in range(len(resto) - 1, -1, -1):
                    posible_doc = re.sub(r"\D", "", resto[i])  # quitar todo lo que no sea número

                    if posible_doc:  # si queda algo numérico
                        documento = posible_doc
                        descripcion = " ".join(resto[:i])
                        break

                if not documento:
                    continue

                # =========================
                # 🔥 6. GUARDAR
                # =========================
                rows.append({
                    "dia": dia,
                    "mes": mes,
                    "oficina": oficina,
                    "descripcion": limpiar_texto(descripcion),
                    "documento": documento,
                    "debito": debito,
                    "credito": credito,
                    "saldo": saldo
                })

            except Exception:
                print(f"❌ Error parseando: {linea}")

        if not rows:
            print(f"⚠️ No se encontraron datos en: {ruta_txt.name}")
            return

        df = pd.DataFrame(rows)

        df = df.sort_values(by=["mes", "dia"]).reset_index(drop=True)

        df.to_excel(ruta_salida, index=False)

        print(f"✅ Convertido: {ruta_txt.name} → {ruta_salida.name}")
        print(f"📊 Registros capturados: {len(df)}")

    except Exception as e:
        print(f"❌ Error procesando {ruta_txt.name}: {e}")


def procesar_carpeta(carpeta):
    carpeta = Path(carpeta)

    if not carpeta.exists():
        print("❌ La carpeta no existe")
        return

    archivos = [
        f for f in carpeta.glob("*.txt")
        if "extracto" in f.name.lower()
    ]

    if not archivos:
        print("⚠️ No se encontraron archivos tipo extracto")
        return

    print(f"📂 Archivos encontrados: {len(archivos)}\n")

    for archivo in archivos:
        salida = archivo.with_suffix(".xlsx")

        if salida.exists():
            print(f"⏭️ Saltado (ya existe): {salida.name}")
            continue

        print(f"🔄 Procesando: {archivo.name}")
        procesar_extracto(archivo, salida)

    print("\n🏁 Proceso terminado")


# =========================
# 🚀 EJECUCIÓN
# =========================

if __name__ == "__main__":
    while True:
        ruta = input("\n📁 Ingresa la ruta (o 'salir'): ").strip()

        if ruta.lower() == "salir":
            break

        procesar_carpeta(ruta)