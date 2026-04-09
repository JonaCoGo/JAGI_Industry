# Instrucciones de Conciliación Datafonos - JAGI Industry

## 🎯 Objetivo
Transformar el auxiliar contable para identificar cómo cada registro fue abonado en banco, descomponiéndolo en uno o varios parciales según:
- Fecha de Abono
- Bol_Ruta
- Valor Neto

## 📋 Archivos Requeridos

### 1. Auxiliar contable
Debe contener:
- Campo CONCEPTO → contiene la fecha real de operación
- Campo valor → valor total contabilizado (agrupado)
- Demás campos no se deben alterar

### 2. Detalle de datafono
Campos obligatorios:
- Fecha Vale → fecha de venta (referencial)
- Fecha de Abono → fecha en extracto bancario (CLAVE)
- Bol. Ruta → agrupador de abonos (CLAVE)
- Valor Neto → valor real abonado (CLAVE)

## 🔧 Lógica del Proceso

### PASO 1 — Extraer fecha real del auxiliar
- Ubicar la fecha dentro del campo CONCEPTO
- Formato esperado: dd/mm/aaaa
- Crear columna: `Fecha_Operacion_Extraida`
- Esta fecha debe coincidir con: `Fecha de Abono` (datafono)

### PASO 2 — Preparar datafono
Agrupar el archivo de datafono así:
`Fecha de Abono + Bol_Ruta → SUM(Valor Neto)`
Resultado: Cada fila representa un abono real en banco

### PASO 3 — Cruce de información
Cruzar: `Auxiliar.Fecha_Operacion_Extraida = Datafono.Fecha de Abono`

### PASO 4 — Construcción del archivo final
REGLA CRÍTICA: **NO eliminar ni modificar filas originales del auxiliar**

#### Estructura final:
Por cada registro del auxiliar:

1. Fila original
   - `Tipo_Fila = AUXILIAR`
   - Mantiene todos los datos originales
   - Agrega:
     - Suma de grupos encontrados
     - Diferencia contra datafono

2. Filas adicionales (debajo)
   Por cada grupo encontrado:
   - `Tipo_Fila = GRUPO_BOL_RUTA`
   - Columnas nuevas:
     - `Fecha_Abono_Encontrada`
     - `Bol_Ruta_Encontrado`
     - `Valor_Grupo_Bol_Ruta`

## ✅ Validaciones Obligatorias
Para cada registro del auxiliar:
`SUM(Valor_Grupo_Bol_Ruta) vs Valor_Auxiliar`

## 📊 Interpretación de Resultados

### Caso 1 — Cuadra exacto
- Registro correcto

### Caso 2 — Diferencia menor
Posibles causas:
- Comisiones
- Retenciones
- Redondeos

### Caso 3 — Sin coincidencias
Problema en:
- Fecha en CONCEPTO
- Falta de datafono
- Error contable

### Caso 4 — Más de un grupo (normal)
Un registro del auxiliar puede corresponder a:
- 1, 2 o 3 Bol_Ruta
- Esto es comportamiento normal del banco

## 🚨 Errores que no se deben cometer
- Cruzar por Fecha Vale
- Usar valor bruto en vez de Valor Neto
- Agrupar por fecha sin Bol_Ruta
- Eliminar filas del auxiliar
- No extraer correctamente la fecha del CONCEPTO

## 📌 Criterio Contable Correcto
- Reconocimiento contable → Fecha de operación (vale)
- Conciliación bancaria → Fecha de abono
- Valor conciliable → Neto

## 🎯 Resultado Final
Este modelo permite:
- Explicar diferencias
- Soportar conciliaciones ante auditoría
- Preparar conciliación bancaria real
- Identificar errores contables