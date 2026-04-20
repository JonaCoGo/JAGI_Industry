# REGLAS DE CONCILIACIÓN DE DATAFONOS
## GIRALDO GIRALDO JAIME WILSON — Cuenta Corriente 2346 Davivienda
**Motor:** `conciliador_engine.py` | **Versión:** 1.5 | **Normativa:** NIIF PYMES / DIAN Colombia

---

## 1. CONTEXTO DEL SISTEMA

- Empresa con **25 sedes** en Colombia, cada una con terminales datafono (Redeban / Credibanco).
- El sistema cruza el **auxiliar contable WorldOffice** contra los **reportes de datafono** de cada sede.
- Cuenta bancaria conciliada: **Davivienda CC 2346**.
- Proceso: **mensual**, por sede o consolidado.
- Auditoría externa: **KPMG Ltda.**

---

## 2. FUENTES DE DATOS

### 2.1 Auxiliar Contable (WorldOffice)
- Archivo Excel exportado directamente desde WorldOffice **sin modificar**.
- Encabezados reales inician en **fila 4** (índice 3). Los datos desde fila 5.
- Se filtran **solo** las filas donde la columna `Nota` contenga la palabra `DATAFONO`.
- Columnas relevantes: `Nota`, `Doc Num`, `Debitos`, `Creditos`, `Saldo`, `Tercero`.

### 2.2 Reporte Datafono (Redeban / Credibanco)
- Uno o varios archivos Excel por sede.
- Columnas requeridas: `Fecha Vale`, `Fecha de Abono`, `Bol. Ruta`, `Valor Neto`.
- Columnas opcionales: `Valor Comisión`, `Ret. Fuente`, `Ret. IVA`, `Ret. ICA`, `Valor Consumo`.
- Un archivo puede tener múltiples hojas. El sistema detecta el encabezado buscando la fila que contenga `Fecha Vale` y `Valor Neto` (máximo primeras 8 filas).

---

## 3. REGLA PRINCIPAL DE CRUCE

```
auxiliar.Dia_Operacion == datafono.Fecha_Vale.day
```

**Para la misma sede, mismo mes y mismo año.**

- El valor que se compara es `auxiliar.Debitos` vs `SUM(datafono.Valor_Neto)` del día.
- El datafono se agrupa: todos los registros del mismo archivo (sede) + mismo día de `Fecha Vale` → suma de `Valor Neto`.

> **NUNCA** se usa la columna `Fecha` del auxiliar para extraer el día de operación.  
> **NUNCA** se usa `Fecha de Abono` como criterio de cruce principal.

---

## 4. FECHA VALE vs. FECHA DE ABONO

| | Fecha Vale | Fecha de Abono |
|---|---|---|
| ¿Qué es? | Fecha en que el cliente pagó con tarjeta | Fecha en que el banco acreditó el dinero |
| ¿Cuándo ocurre? | El mismo día de la venta | D+1 a D+3 (días hábiles después) |
| ¿Dónde la usa el sistema? | **Cruce principal** auxiliar vs. datafono | Solo informativo (detalle Bol. Ruta) |

**Por qué Fecha Vale y no Fecha de Abono:**
1. El auxiliar WorldOffice registra la venta en la fecha de la transacción (principio de devengo NIIF).
2. Con Fecha de Abono los últimos días del mes quedan sin cruce (el abono cae en el mes siguiente).
3. Los abonos agrupados (varios días en un solo abono) generan ambigüedad irresoluble.
4. Los festivos y fines de semana producen diferencias artificiales si se usa Fecha de Abono.

**Fecha de Abono** se usa únicamente para:
- Mostrar el campo `Abono D+N` en el detalle del informe.
- Conciliación bancaria (proceso separado, fuera del alcance de este sistema).

---

## 5. EXTRACCIÓN DEL DÍA DE OPERACIÓN DESDE LA NOTA

El día se extrae del campo `Nota` del auxiliar. Se prueban los siguientes patrones **en orden de prioridad**:

| Prioridad | Patrón | Ejemplo de Nota | Día extraído |
|---|---|---|---|
| 1 | Fecha completa `dd/mm/aaaa` o `dd-mm-aaaa` | `"01/03/2025 DATAFONO"` | 1 |
| 2 | Fecha parcial `dd/mm` o `dd-mm` | `"15/03 DATAFONO"` | 15 |
| 3 | Número al inicio seguido de `-` o espacio | `"2- DATAFONO ENVIGADO"` | 2 |
| 4 | Cualquier número válido 1-31 en el texto | `"DATAFONO DIA 7"` | 7 |

Si ningún patrón extrae un día → estado `SIN_DIA`.

---

## 6. DETECCIÓN DE FORMATO DEL AUXILIAR (por fila)

El sistema detecta el formato **fila por fila** (un mismo auxiliar puede mezclar ambos):

### Formato NUEVO (Enero 2026 en adelante)
- La `Nota` empieza con dígito(s) seguido de `-` o espacio + `DATAFONO` + nombre de sede.
- Ejemplo: `"2- DATAFONO ENVIGADO"`, `"10 DATAFONO MOLINOS"`
- **Sede:** se extrae de la Nota, lo que viene después de `DATAFONO`.
- **Día:** número al inicio de la Nota.

### Formato ANTIGUO (años anteriores)
- La `Nota` contiene fecha completa o parcial + `DATAFONO`.
- Ejemplo: `"01/03/2025 DATAFONO"`, `"15-03 DATAFONO"`
- **Sede:** se extrae del campo `Tercero`, limpiando prefijos.
- **Día:** fecha extraída de la Nota.

### Limpieza del campo Tercero (formato antiguo)
Se eliminan los siguientes prefijos del campo `Tercero` (sin quitar el resto del nombre):
- `CLIENTES VENTAS `
- `CLIENTES VENTA `
- `CLIENTES `

**No se eliminan:** `PLAZA`, `PARQUE`, `VIVA`, `LOCAL` — son parte del nombre de la sede.

Ejemplo: `"CLIENTES VENTAS VIVA ENVIGADO"` → sede = `"VIVA ENVIGADO"`

---

## 7. SEDES OFICIALES (lista maestra — 25 sedes)

```python
SEDES_OFICIALES = [
    "Barranquilla", "Buenavista", "Centro Mayor", "Chipichape",
    "Eden", "Envigado", "Fabricato", "Plaza Imperial",
    "Jardin Plaza", "Mercurio", "Molinos", "Nuestro",
    "Pasto", "Puerta del Norte", "Santa Marta", "Santa Fe",
    "Sincelejo", "Parque Alegra", "Americas", "Cacique",
    "Tesoro", "Titan Plaza", "Cali", "Pereira", "Serrezuela",
]
```

Esta lista es la **fuente de verdad única** para normalización de nombres.

---

## 8. MATCHING DIFUSO DE NOMBRES (archivo → sede)

El nombre del archivo de datafono se compara con los nombres de sedes del auxiliar.

**Algoritmo:**
1. Normalizar ambos nombres: mayúsculas, sin tildes, sin caracteres especiales.
2. Calcular puntuación de similitud (0-100).
3. Umbral mínimo: **60/100**. Por debajo → archivo huérfano (sin asignar).
4. Asignación greedy: cada sede solo puede usarse una vez.

**Tipos de coincidencia:**

| Tipo | Puntuación | ¿Requiere revisión manual? |
|---|---|---|
| Exacto | 100 | No |
| Contenido (`A in B`) | 95 | Recomendable |
| Contenido inverso (`B in A`) | 90 | Recomendable |
| Palabras significativas en común | 60-85 | Sí — confirmar |
| Sin match | 0 | Sí — renombrar archivo o verificar sede |

**Stopwords excluidas del matching de palabras:**
`DE, DEL, LA, LAS, LOS, EL, EN, Y, A, MALL, CC, ENERO...DICIEMBRE, THE, OF, VIVA, LOCAL, TIENDA, PARQUE, PLAZA, UNICENTRO, CLIENTES, VENTAS, VENTA`

**Nombre del archivo:** se extrae quitando el prefijo `datafono_` y la extensión.  
Ejemplo: `datafono_ENVIGADO.xlsx` → nombre = `ENVIGADO`

---

## 9. ANTI-DUPLICADO DE HOJAS EN EL DATAFONO

Si un archivo de datafono tiene múltiples hojas:
- Se calcula el total de `Valor Neto` por día para cada hoja.
- Si dos hojas tienen totales idénticos en **más del 80%** de los días en común → se descarta la hoja secundaria.
- Solo se usa la hoja principal para evitar doble conteo.

---

## 10. ESTADOS DE RESULTADO DEL CRUCE

| Estado | Condición | Color en informe |
|---|---|---|
| `CUADRA` | `abs(diferencia) < 1` | Verde |
| `DIF_MENOR` | `abs(diferencia) <= valor_aux * 0.05` | Amarillo |
| `DIFERENCIA` | `abs(diferencia) > valor_aux * 0.05` | Rojo |
| `SIN_MATCH` | No se encontró ningún registro en el datafono para ese día | Rojo |
| `SIN_DIA` | No se pudo extraer el día de la Nota | Rojo |

**Tolerancia diferencia menor:** ≤ 5% del valor del auxiliar.  
Causa típica de `DIF_MENOR`: comisiones, Ret. Fuente, Ret. IVA, Ret. ICA, redondeos.

---

## 11. AGRUPACIÓN DEL DATAFONO

### Por día (para el cruce principal)
```
GROUP BY: nombre_archivo, año(Fecha Vale), mes(Fecha Vale), día(Fecha Vale)
AGG: SUM(Valor Neto), SUM(Valor Comisión), SUM(Ret. Fuente), SUM(Ret. IVA), SUM(Ret. ICA)
```

### Por Bol. Ruta (para el detalle en el informe)
```
GROUP BY: nombre_archivo, año, mes, día, Fecha de Abono, Bol. Ruta
AGG: SUM(Valor Neto)
```

---

## 12. ESTRUCTURA DEL INFORME EXCEL

### Modos de generación
| Modo | Salida |
|---|---|
| `individual` | Un Excel para una sede específica |
| `separados` | Un Excel por sede + resumen ejecutivo |
| `unificado` | Un solo Excel con todas las sedes (recomendado para gerencia) |

### Hojas del informe unificado
| Hoja | Contenido |
|---|---|
| `RESUMEN_EJECUTIVO` | Tabla con todas las sedes, totales y % de cuadre |
| `MAPA_NOMBRES` | Asignación archivo → sede, tipo de match, score |
| `[NOMBRE_SEDE]` | Detalle completo del cruce por sede (máx 28 chars) |
| `PDTE_[SEDE]` | Pendientes: registros sin conciliar (sección A) + días en datafono sin registro en auxiliar (sección B) |
| `LOG_AUDITORIA` | Registro inmutable del proceso |

### Columnas del detalle por sede
`Tipo Fila` | `Día` | `Nota Auxiliar` | `Doc Num` | `Valor Auxiliar ($)` | `Sum Valor Neto Datafono ($)` | `Diferencia ($)` | `% Dif` | `Estado` | `Fecha Abono` | `Bol_Ruta` | `Valor Neto Bol_Ruta ($)` | `Comisión + Ret ($)` | `# Bol_Rutas` | `Observación`

- Filas `AUXILIAR`: un registro del auxiliar contable.
- Filas `GRUPO_BOL_RUTA`: detalle del datafono para ese día (subordinadas al AUXILIAR).

---

## 13. LOG DE AUDITORÍA (inmutable)

Cada informe incluye una hoja `LOG_AUDITORIA` con:
- Fecha y hora del proceso
- Sede y período
- Versión del motor
- Normativa aplicada
- Lógica de cruce usada
- Resultado del matching de nombres
- Umbral de matching difuso
- Nota sobre archivos originales no modificados
- Referencia a auditoría externa (KPMG Ltda.)

---

## 14. REGLAS DE NEGOCIO ADICIONALES

- Los archivos originales (auxiliar y datafono) **nunca se modifican**. El motor trabaja sobre copias en memoria.
- El valor a conciliar es siempre **`Valor Neto`** del datafono (ya descontadas comisiones y retenciones).
- Si una sede del auxiliar no tiene archivo de datafono cargado → se omite del proceso pero se reporta en el resumen.
- Si un archivo de datafono no tiene sede correspondiente en el auxiliar → queda como **huérfano** y se reporta.
- El campo `Fecha` del auxiliar WorldOffice **no se usa en ningún cálculo**. Solo se usa la fecha embebida en `Nota`.

---

## 15. PARÁMETROS DE CONFIGURACIÓN

| Parámetro | Valor actual | Ubicación en código |
|---|---|---|
| Umbral matching difuso | 60 / 100 | `UMBRAL_MATCH` en `conciliador_engine.py` |
| Tolerancia diferencia menor | 5% del valor auxiliar | `cruzar_auxiliar_datafono()` |
| Umbral anti-duplicado hojas | 80% días con totales iguales | `leer_datafono()` |
| Fila encabezado auxiliar | Fila 4 (índice 3), datos desde fila 5 | `leer_auxiliar()` |
| Máx filas a escanear para encabezado datafono | 8 | `_detectar_encabezado()` |

---

*Documento generado desde `conciliador_engine.py` v1.5 + `app_conciliador.py` v1.4*  
*GIRALDO GIRALDO JAIME WILSON — Uso interno — Confidencial*
