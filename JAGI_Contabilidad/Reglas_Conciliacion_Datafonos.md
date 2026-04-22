# REGLAS DE CONCILIACIÓN DE DATAFONOS
## GIRALDO GIRALDO JAIME WILSON — Cuenta Corriente 2346 Davivienda
**Motor:** `conciliador_engine.py` | **Versión:** 1.6 | **Normativa:** NIIF PYMES / DIAN Colombia

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

El sistema soporta dos criterios de cruce. La auxiliar contable selecciona el criterio
según cómo registró las operaciones en WorldOffice para el período a conciliar.

### Criterio A — Fecha Vale (NIIF estándar)
```
auxiliar.Dia_Operacion == datafono.Fecha_Vale.day
```
- Usar cuando la Nota del auxiliar contiene la **fecha en que ocurrió la venta**.
- Consistente con NIIF PYMES (principio de devengo: el ingreso se reconoce cuando ocurre la transacción).
- El auxiliar registra la venta el mismo día que el cliente pagó.

### Criterio B — Fecha de Abono
```
auxiliar.Dia_Operacion == datafono.Fecha_Abono.day
```
- Usar cuando la Nota del auxiliar contiene la **fecha en que el banco acreditó el dinero**.
- La auxiliar tomó la fecha de abono del extracto del datafono para registrar en WorldOffice.
- Válido operativamente cuando la empresa decide registrar por fecha de acreditación bancaria.

**Para la misma sede, mismo mes y mismo año.**

- El valor que se compara siempre es `auxiliar.Debitos` vs `SUM(datafono.Valor_Neto)`.
- El criterio seleccionado queda registrado en el `LOG_AUDITORIA` del informe.
- La selección del criterio la hace la auxiliar contable desde la interfaz antes de ejecutar.

> **REGLA DE CONSISTENCIA:** El criterio debe ser uniforme en todo el período conciliado.
> No mezclar Fecha Vale y Fecha de Abono en el mismo mes. Si el auxiliar tiene registros
> con ambos criterios, reportar antes de continuar.

> El campo `Fecha` del auxiliar WorldOffice **no se usa en ningún cálculo**.
> El día de operación siempre se extrae del campo `Nota`.

---

## 4. FECHA VALE vs. FECHA DE ABONO

| | Fecha Vale | Fecha de Abono |
|---|---|---|
| ¿Qué es? | Fecha en que el cliente pagó con tarjeta | Fecha en que el banco acreditó el dinero |
| ¿Cuándo ocurre? | El mismo día de la venta | D+1 a D+3 (días hábiles después) |
| ¿Cuándo usarlo? | Cuando el auxiliar registra por fecha de venta | Cuando el auxiliar registra por fecha de acreditación bancaria |

### Consideraciones operativas de cada criterio

**Criterio Fecha Vale:**
- Consistente con NIIF PYMES (el ingreso se reconoce cuando ocurre la transacción).
- Los cierres de mes cuadran dentro del mismo período (la venta del 31 cuadra en el período del 31).
- No genera diferencias artificiales en festivos ni fines de semana.

**Criterio Fecha de Abono:**
- Útil cuando la empresa registra contablemente por la fecha de acreditación bancaria.
- Los últimos 1-3 días del mes pueden quedar sin cruce (el abono cae en el mes siguiente).
- Si un abono agrupa varios días de venta, puede generar `SIN_MATCH` si el auxiliar registra día a día.
- El `LOG_AUDITORIA` registra el criterio utilizado en cada proceso para trazabilidad.

### Regla de consistencia (obligatoria)
El criterio seleccionado debe ser **uniforme en todo el período**: si el auxiliar registró
enero bajo Fecha Vale, todo enero debe conciliarse con Fecha Vale. Si gerencia cambia
el criterio de registro a partir de un mes, ese mes se concilia con el nuevo criterio
y se documenta el cambio en el log de auditoría.

### ¿Cuándo usa el sistema Fecha de Abono sin ser criterio de cruce?
La `Fecha de Abono` aparece siempre en el informe como dato informativo:
- Campo `Abono D+N` en el detalle del informe (cuántos días tardó el abono).
- Base para la conciliación bancaria (proceso separado, fuera del alcance de este sistema).

---

## 5. EXTRACCIÓN DEL DÍA DE OPERACIÓN DESDE LA NOTA

El sistema extrae el día del campo `Nota` con los mismos patrones independientemente
del criterio de cruce seleccionado. Lo que cambia entre criterios es **qué representa ese día**
(fecha de venta o fecha de abono) y contra qué columna del datafono se compara.

Se prueban los siguientes patrones **en orden de prioridad**:

| Prioridad | Patrón | Ejemplo de Nota | Día extraído |
|---|---|---|---|
| 1 | Fecha completa `dd/mm/aaaa` o `dd-mm-aaaa` | `"01/03/2026 DATAFONO"` | 1 |
| 2 | Fecha parcial `dd/mm` o `dd-mm` | `"15/03 DATAFONO"` | 15 |
| 3 | Número al inicio seguido de `-` o espacio | `"2- DATAFONO ENVIGADO"` | 2 |
| 4 | Cualquier número válido 1-31 en el texto | `"DATAFONO DIA 7"` | 7 |

Si ningún patrón extrae un día → estado `SIN_DIA`.

---

## 6. DETECCIÓN DE FORMATO DEL AUXILIAR (por fila)

El sistema detecta el formato **fila por fila**. Un mismo auxiliar puede mezclar formatos.

### Formato ESTÁNDAR (vigente desde Marzo 2026 en adelante)
- La `Nota` contiene la fecha completa + `DATAFONO`. Sin nombre de sede en la Nota.
- La fecha en Nota puede ser Fecha Vale o Fecha de Abono — depende del criterio de registro adoptado.
- Ejemplo: `"01/03/2026 DATAFONO"`, `"15-03-2026 DATAFONO"`
- **Sede:** se extrae del campo `Tercero`, limpiando prefijos.
- **Día:** fecha extraída de la Nota.

### Formato NUEVO (transitorio — Enero/Febrero 2026)
- La `Nota` empieza con dígito(s) seguido de `-` o espacio + `DATAFONO` + nombre de sede.
- Ejemplo: `"2- DATAFONO ENVIGADO"`, `"10 DATAFONO MOLINOS"`
- **Sede:** se extrae de la Nota, lo que viene después de `DATAFONO`.
- **Día:** número al inicio de la Nota.

### Formato ANTIGUO (años anteriores a 2026)
- La `Nota` contiene fecha completa o parcial + `DATAFONO`. Sin nombre de sede en la Nota.
- Ejemplo: `"01/03/2025 DATAFONO"`, `"15-03 DATAFONO"`
- **Sede:** se extrae del campo `Tercero`, limpiando prefijos.
- **Día:** fecha extraída de la Nota.

> **Nota técnica:** Los formatos ESTÁNDAR y ANTIGUO son estructuralmente idénticos para el motor.
> El motor los procesa con la misma lógica — la distinción es solo documental para trazabilidad.

### Limpieza del campo Tercero (formatos ESTÁNDAR y ANTIGUO)
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

*Documento generado desde `conciliador_engine.py` v1.6 + `app_conciliador.py` v1.5*  
*GIRALDO GIRALDO JAIME WILSON — Uso interno — Confidencial*  
*Actualizado: Abril 2026 — v1.6: criterio de cruce configurable (Fecha Vale / Fecha de Abono)*
