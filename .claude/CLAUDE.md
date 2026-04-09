# JAGI Industry - Departamento de Análisis de Datos y Automatizaciones

## 🎯 Nuestro Rol

Somos el **Departamento de Análisis de Datos y Automatizaciones** de JAGI Industry. No somos especialistas de las áreas operativas (contabilidad, producción, etc.), sino **facilitadores técnicos** que:

- Automatizamos procesos de cualquier área de la empresa.
- Analizamos datos y generamos soluciones.
- Creamos herramientas que simplifican el trabajo de otros departamentos.
- Brindamos soporte técnico nivel 1 y 2.

## 👤 Perfil del Equipo

| Rol | Funciones |
|-----|-----------|
| **Analista de Datos** (tú) | Programación, análisis, soporte técnico, mantenimiento |
| **Asistente IA** (yo) | Ejecución de código, automatización, procesamiento de datos |

## 📁 Estructura del Proyecto

```
JAGI Industry/                    ← Carpeta raíz del proyecto
├── JAGI_Analytics/               ← Análisis de datos y métricas
├── JAGI_Contabilidad/           ← Procesos contables y financieros
├── JAGI_Facturacion/            ← (próximamente)
├── JAGI_Produccion/             ← (próximamente)
├── JAGI_Tesoreria/             ← (próximamente)
├── JAGI_RRHH/                  ← (próximamente)
└── JAGI_Soporte/               ← (próximamente)Tickets y soporte técnico
```

### 📌 Regla de Organización

Cada carpeta `JAGI_X` representa un área de la empresa y debe contener:
- `CLAUDE.md` → Reglas específicas de esa área
- Archivos de datos → CSV, Excel, etc.
- Scripts → Códigos de automatización
- Documentación → Procedimientos y reglas

## 🔧 Flujo de Trabajo General

1. **Identificar el área** → ¿Qué área de la empresa necesita análisis/automatización?
2. **Leer el CLAUDE.md de esa carpeta** → Cada área tiene sus propias reglas
3. **Aplicar las reglas específicas** → Procesar según el protocolo de esa área
4. **Generar resultados** → Entregar solución al área solicitante

## 🤖 Modelo de Agentes Especializados

Cada carpeta JAGI_X puede tener su propio "agente" especializado:

| Agente | Área | Estado |
|--------|------|--------|
| `Agente_Analytics` | JAGI_Analytics | Activo |
| `Agente_Contabilidad` | JAGI_Contabilidad | Activo |
| `Agente_Facturacion` | JAGI_Facturacion | Por crear |
| `Agente_Produccion` | JAGI_Produccion | Por crear |
| `Agente_Tesoreria` | JAGI_Tesoreria | Por crear |
| `Agente_RRHH` | JAGI_RRHH | Por crear |
| `Agente_Soporte` | JAGI_Soporte | Por crear |

## 📋 Cómo Crear una Nueva Área

Cuando una dependencia solicite análisis o automatización:

1. Crear carpeta `JAGI_[NombreArea]` en la raíz
2. Crear `CLAUDE.md` dentro con:
   - Objetivo del área
   - Reglas específicas de procesamiento
   - Tipos de archivos esperados
   - Formato de salida requerido
3. Actualizar este CLAUDE.md raíz para incluir la nueva área

## 🔗 Referencias

- **JAGI_Contabilidad**: [JAGI_Contabilidad/CLAUDE.md](JAGI_Contabilidad/CLAUDE.md)
- **JAGI_Analytics**: [JAGI_Analytics/CLAUDE.md](JAGI_Analytics/CLAUDE.md)
- (Agregar más según se creen)