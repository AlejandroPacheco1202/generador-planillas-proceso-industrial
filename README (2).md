# Generador de Planillas de Proceso — Panificación Industrial

**Excel + VBA · En producción desde 2008 · Mantenido por usuarios no programadores**

Automatización que genera las planillas de control de proceso de una planta de panificación industrial (amasado, fermentación, cocción, abatimiento y envasado) a partir de la selección del plan de producción del día. Reemplazó horas de tipeo manual por un flujo de un par de clics, y lleva más de 15 años en uso continuo con mantenimiento mínimo.

> ⚠️ **Versión anonimizada.** El archivo publicado es una versión demo: productos, códigos de artículo, códigos de registro de calidad, marcas de proveedores y logos fueron reemplazados por equivalentes ficticios. La estructura, la lógica y las macros son las originales.

---

## El problema

En una planta de panificación, cada día de producción exige registrar en papel el control de cada etapa del proceso, por exigencia del sistema de gestión de calidad. Cada producto lleva su propia planilla de amasado con sus insumos específicos, y cada amasado debe aparecer también en las planillas de fermentación, cocción, abatimiento y envasado del día.

Armar ese juego de planillas a mano implicaba tipear producto por producto, insumo por insumo, todos los días. Una analista de calidad dedicaba **alrededor de 4 horas por semana** solo a preparar planillas — unas **200 horas al año** — con el riesgo adicional de errores de transcripción en documentos controlados (insumos faltantes, productos mal asignados, paginación incorrecta).

## La solución

Un libro de Excel con macros VBA que funciona como generador: el usuario carga en un menú la fecha, los productos del día, las cantidades de amasados y el turno, y el sistema arma automáticamente todas las planillas del proceso listas para imprimir.

El corazón del diseño no está en el código sino en los datos: una **tabla maestra** (hoja `ART`) define para cada producto su código, qué plantilla de planilla le corresponde, sus insumos (hasta 15), cuántos amasados entran por hoja y si pasa por cocción o se despacha crudo. Las macros no saben nada de productos: recorren la tabla y actúan en consecuencia.

```
Menú (fecha, productos, cantidades, turno)
        │
        ▼
Tabla maestra ART ──── producto → plantilla, insumos, amasados/hoja, ¿cocción?
        │
        ▼
Generación automática:
  ├─ Planillas de amasado (una por producto, con sus insumos)
  ├─ Fermentación   ┐
  ├─ Cocción        │  planillas del día, paginadas
  ├─ Abatimiento    │  (30 renglones por hoja, "Página x de n")
  └─ Envasado       ┘
```

Ese diseño data-driven tiene una consecuencia práctica que explica la longevidad del proyecto: **dar de alta un producto nuevo no requiere tocar código**. Se agrega una fila a la tabla maestra y listo. Por eso el archivo pudo ser mantenido durante años por analistas de calidad sin conocimientos de programación. De hecho, al preparar esta versión pública se verificó que el código VBA no contiene ni un solo nombre de producto: toda la información vive en los datos.

## Qué demuestra técnicamente

El proyecto usa VBA con ciclos, condicionales y punteros de posición para: contar los amasados de la fecha seleccionada, separar productos cocidos de crudos, calcular cuántas páginas necesita cada planilla, clonar las plantillas la cantidad de veces necesaria, numerar páginas, distribuir los amasados en los renglones correctos y reinicializar el libro en cada apertura (`Workbook_Open` limpia los datos del día anterior y oculta las hojas internas). La interfaz es un menú con `UserForm` y validación por listas desplegables.

Es código de 2008, escrito grabando macros y editándolas a mano, y se nota: navegación con `ActiveCell`, bloques repetidos, variables sin declarar. No lo escondo — es parte de la historia. La arquitectura simple fue precisamente lo que lo hizo robusto y mantenible. Con lo que sé hoy lo refactorizaría con procedimientos parametrizados, `Option Explicit` y referencias directas a rangos (o directamente lo migraría a una app con base de datos), pero el criterio de diseño de fondo — separar configuración de lógica — es el mismo que aplicaría hoy.

## Impacto

| | |
|---|---|
| Ahorro de tiempo | ~4 hs/semana → **+3.000 horas acumuladas** desde 2008 |
| Errores de transcripción | Eliminados en la etapa de armado de planillas |
| Vida útil | **2008 → hoy**, con revisiones menores (última revisión formal: 2021) |
| Mantenimiento | Alta de productos sin tocar código, a cargo de usuarios no técnicos |
| Alcance | ~90 productos activos, 5 etapas de proceso, 3 turnos |

## Contenido del repositorio

```
├── Planilla_Demo.xlsm       # Libro completo anonimizado, con macros funcionales
├── src/                     # Módulos VBA exportados, para leer sin abrir Excel
│   ├── Modulo2_PlanillaDeAmasados.bas
│   ├── Modulo3_PlanillaFCA.bas
│   ├── Modulo4_PlanillasDeProceso.bas
│   └── ThisWorkbook.bas
└── README.md
```

Para probarlo: abrir `Planilla_Demo.xlsm` en Excel, habilitar macros, cargar fecha, productos y cantidades en la hoja de menú, y ejecutar la generación desde el botón del formulario.

---

*Desarrollado en 2008 para el sector producción de una planta de panificación industrial. Publicado en versión anonimizada como parte de mi portfolio, como ejemplo de automatización de procesos con impacto medible y diseño orientado a datos.*
