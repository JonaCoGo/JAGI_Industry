# JAGI Industry SAS — Orquestador Principal

## Identidad del proyecto
JAGI Industry SAS opera este repositorio como sustituto funcional
de un departamento de sistemas. El objetivo es construir un ecosistema
de soluciones empresariales autónomas que reemplacen procesos manuales.

---

## Rol de este archivo
Este CLAUDE.md es el orquestador principal. Define el stack, la arquitectura
y los principios que todos los agentes heredan.

Cuando una tarea llega, este orquestador:
1. Clasifica a qué dominio pertenece
2. Verifica si existe agente para ese dominio
3. Delega al agente correcto o declara que no existe aún

---

## Agentes del ecosistema

| Dominio | Carpeta | Archivo agente | Estado |
|---|---|---|---|
| Contabilidad | `JAGI_Contabilidad/` | `.claude/agents/agente-contabilidad.md` | 🔲 Por crear |
| Logística | *(próximamente)* | — | 🔲 Pendiente |
| Recursos Humanos | *(próximamente)* | — | 🔲 Pendiente |
| Producción | *(próximamente)* | — | 🔲 Pendiente |

### Estado actual de agentes
Los agentes aún no están creados como archivos `.claude/agents/*.md`.
Mientras no existan, Claude opera directamente con el CLAUDE.md
de cada subcarpeta de dominio como contexto especializado.

### Cuando se creen los agentes
Cada agente vivirá en `.claude/agents/` con este formato:

~~~markdown
---
name: agente-contabilidad
description: Invocar para tareas de conciliación, causación, IVA, DIAN, NIIF
tools: Read, Write, Bash, Glob, Grep
model: sonnet
---
[system prompt del agente]
~~~

Actualizar la tabla anterior con estado ✅ Activo al crear cada uno.

---

## Protocolo de delegación

### Paso 1 — Clasificar la tarea

| Si la tarea involucra... | Dominio |
|---|---|
| Conciliaciones, causación, IVA, auxiliares, DIAN, NIIF, extractos, datafonos | Contabilidad |
| Inventario, despachos, proveedores, transporte | Logística |
| Nómina, contratos, vacaciones, novedades de personal | Recursos Humanos |
| Órdenes de producción, procesos, control de planta | Producción |

### Paso 2 — Verificar agente

**Si el agente existe** (archivo en `.claude/agents/`):
→ Invocar con `/agents` o dejar que Claude lo active automáticamente por descripción.

**Si el agente NO existe pero hay CLAUDE.md en la subcarpeta**:
→ Navegar a la carpeta del dominio y trabajar con su CLAUDE.md como contexto.
→ Declarar: *"Trabajando en dominio [X] sin agente formal. Usando CLAUDE.md de la carpeta."*

**Si no existe ni agente ni CLAUDE.md**:
→ Declarar: *"El dominio [X] no tiene configuración definida. ¿Construimos el agente o el CLAUDE.md primero?"*

### Paso 3 — Tareas que cruzan dominios
1. Identificar el dominio primario (donde vive el dato principal)
2. Operar desde ese dominio
3. Notificar dependencia con otro dominio
4. NO acoplar lógica entre dominios
5. Si se comparten datos: definir explícitamente el punto de intercambio

---

## Stack tecnológico (heredado por todos los agentes)

- Lenguaje: Python. Alternativas solo con justificación técnica clara.
- Librerías base: pandas, openpyxl
- Entrada/Salida: Excel como medio principal
- Bases de datos: SQLite para proyectos simples. PostgreSQL si el volumen lo justifica.
- UI: Tkinter por defecto. Justificar si se propone otra.
- Normativa: DIAN + NIIF PYMES. Toda solución compatible con regulación colombiana vigente.

---

## Principios de arquitectura (obligatorios para todos los dominios)

- Cada sistema es autónomo — sin dependencias entre dominios
- NO acoplar sistemas entre sí
- NO asumir integraciones implícitas
- Separación estricta: lógica de negocio / procesamiento / interfaz
- Todo módulo debe poder ejecutarse de forma independiente

---

## Modelo de desarrollo

Incremental por módulos. Cada módulo debe:
1. Resolver un problema completo
2. Ser usable en producción desde su entrega
3. Ejecutarse de forma independiente
4. Permitir integración posterior sin reescritura

---

## Estándar de entrega (obligatorio en todos los dominios)

- Código: completo, ejecutable, listo para producción. Sin fragmentos.
- Manejo de errores: siempre presente cuando aplique.
- Validación de datos de entrada: siempre antes de procesar.
- Log obligatorio en todo proceso automatizado: fecha/hora, parámetros, versión.
- Archivos fuente: nunca se modifican. Solo lectura + copia en memoria.

---

## Ciclo obligatorio de validación técnica

Toda salida debe pasar por este ciclo antes de considerarse final:

### Fase 1 — Generación
Construir solución completa según instrucciones.

### Fase 2 — Autorevisión crítica
Evaluar como QA técnico:
- errores lógicos
- edge cases no cubiertos
- supuestos implícitos
- posibles fallos en datos reales

### Fase 3 — Corrección
Si se detectan problemas: corregir directamente.
NO entregar versión defectuosa.

### Fase 4 — Validación de robustez
Confirmar:
- manejo de errores implementado
- validación de inputs presente
- comportamiento ante datos inesperados definido

Regla: ninguna solución se entrega sin pasar por este ciclo completo.

---

## Clasificación de tareas

Antes de ejecutar, clasificar internamente:

- Tipo A — Generación: crear desde cero
- Tipo B — Modificación: cambiar código existente → mostrar ANTES/DESPUÉS
- Tipo C — Diagnóstico: detectar errores → priorizar identificación sobre solución
- Tipo D — Validación: auditar solución → evaluar lógica, no estilo

---

## Comportamiento ante lo desconocido

- Tarea sin dominio claro → declararlo y pedir clarificación
- Sin archivo de reglas para el proceso → construir el archivo de reglas primero, no inventar
- Contradicción entre dos fuentes → exponer la contradicción, no resolverla silenciosamente

---

## Control de calidad inter-dominio

Antes de entregar cualquier solución verificar:
1. No viola autonomía de módulos ni acoplamiento entre dominios
2. Datos críticos tienen mecanismo de validación definido
3. Automatizaciones incluyen logging
4. Incertidumbre en reglas → detener y escalar al usuario

Regla: la velocidad nunca tiene prioridad sobre la consistencia del sistema.

---

*Orquestador principal: JAGI Industry SAS*
*Agentes activos: ninguno aún — en construcción*
*Stack: Python / pandas / openpyxl / Tkinter / Excel*
*Normativa: NIIF PYMES / DIAN Colombia*