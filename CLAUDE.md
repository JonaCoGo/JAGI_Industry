# JAGI Industry SAS — Sistema de Desarrollo Interno

## Identidad del proyecto
JAGI Industry SAS opera este repositorio como sustituto funcional
de un departamento de sistemas. El objetivo es construir un ecosistema
de soluciones empresariales autónomas que reemplacen procesos manuales.

## Stack tecnológico principal
- Lenguaje: Python (pandas, openpyxl, Tkinter como base; proponer alternativas si aplica)
- Entrada/Salida: Excel como medio principal
- Bases de datos: SQLite para proyectos simples, PostgreSQL si el volumen lo justifica
- UI: Tkinter por defecto; justificar si se propone otra

## Normativa aplicable
- DIAN (facturación electrónica, formatos tributarios colombianos)
- NIIF PYMES (contabilidad)
- Cualquier regulación colombiana relevante debe ser considerada activamente

## Arquitectura del ecosistema
El proyecto contiene sistemas independientes por dominio:

| Dominio | Descripción |
|---|---|
| Contabilidad | Registros, reportes NIIF, DIAN |
| Logística | Inventario, despachos, proveedores |
| Recursos Humanos | Nómina, contratos, novedades |
| Producción | Órdenes, procesos, control |

Principios de arquitectura (obligatorios):
- Cada sistema es autónomo — sin dependencias entre dominios
- NO acoplar sistemas entre sí
- NO asumir integraciones implícitas
- Separación estricta: lógica de negocio / procesamiento / interfaz

## Modelo de desarrollo
Incremental por módulos. Cada módulo debe:
1. Resolver un problema completo
2. Ser usable en producción desde su entrega
3. Ejecutarse de forma independiente
4. Permitir integración posterior sin reescritura

## Prioridad operativa
1. Soluciones funcionales inmediatas
2. Reducción de trabajo manual
3. Impacto en operación

Evitar: sobrearquitectura, complejidad innecesaria.

## Criterio de decisión
Si una solución simple y robusta resuelve el problema → no proponer una más compleja.
Balancear: rapidez de implementación (alta prioridad) + calidad estructural suficiente.