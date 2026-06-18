# 🚀 ExcelEvaluationEngine

Sistema de correccion automatica para examenes de Datos Bivariados (Probabilidad y Estadistica) sobre archivos Excel reales de alumnos.

Procesa planillas caoticas con heuristicas matematicas y semanticas, evaluando no solo el resultado numerico, sino tambien la legitimidad del procedimiento.

Combina recálculo dependiente, deteccion anti-hardcode, tolerancia adaptativa y trazabilidad completa para auditoria docente.

---

## 🎯 El Desafio Tecnico (The Problem)

Este proyecto fue diseñado para operar en un entorno de **alta entropia humana**:

- El layout de Excel no es estable (columnas, filas y orden cambian entre alumnos).
- Los nombres de hoja varian (por ejemplo: Tema A, Hoja1, Hoja2, etc.).
- Existen redondeos analiticos inconsistentes y formatos mixtos.
- Aparecen casos de *hardcoding*: el alumno ingresa el numero final correcto sin formula valida.

En este contexto, una correccion por coordenadas fijas es fragil. El sistema resuelve esto con validacion heuristica y arquitectonica orientada a resiliencia.

---

## 💡 Soluciones de Ingenieria Implementadas (The Core Features)

### 1) Motor de Recalculo en Caliente (Anti-Bola de Nieve)

Cuando una metrica base esta mal (ej. Promedios), el motor evita arrastrar un cero total injusto:

1. Captura el valor entregado por el alumno.
2. Recalcula internamente metricas dependientes con ese valor defectuoso:
   - Covarianza
   - Desvio estandar X/Y
   - Correlacion de Pearson
   - Pendiente y ordenada de la recta
3. Si el flujo procedural es consistente, otorga **puntaje parcial por arrastre (66%)**.

Resultado: se penaliza el error original, pero se reconoce correctamente el dominio del procedimiento subsiguiente.

### 2) Validador Hibrido Anti-Trampa

La deteccion no depende de celdas absolutas; clasifica formulas por estructura:

- **Funciones estadisticas directas**: AVERAGE, PEARSON, CORREL, SLOPE, INTERCEPT, etc.
- **Desarrollo algebraico por pasos**: SUM, SQRT, POWER y operadores relativos (*, /, ^, -).

Si el valor coincide numericamente, pero la celda no contiene formula legitima, la metrica se clasifica como **hardcoded** y recibe **0 puntos** para ese item.

### 3) Busqueda Semantica Adaptativa

- Seleccion de hoja por **score compuesto**:
  - densidad numerica
  - presencia de formulas
  - senales de series de trabajo
- Seleccion de candidatos por ventana heuristica:
  - **tolerancia relativa configurable (5%) + epsilon base (0.005)**

Esto absorbe variaciones de redondeo sin degradar la precision de evaluacion.

### 4) Sanitizador Linguistico de Pearson

La interpretacion textual del coeficiente r se valida con normalizacion difusa:

- Limpieza de diacriticos (tildes) y espacios
- Normalizacion de texto
- Validacion por diccionarios semanticos
- Coherencia entre signo/fuerza textual y valor numerico del alumno

No se valida solo la frase; se valida la consistencia logica con el resultado cuantitativo.

---

## 🏗️ Arquitectura del Sistema (Clean Architecture & SOLID)

El diseno separa reglas de negocio, orquestacion e infraestructura para maximizar mantenibilidad y testabilidad.

| Capa | Responsabilidad | Componentes relevantes |
|---|---|---|
| Core / Domain | Modelos inmutables y reglas matematicas puras | MasterMetrics, StudentEvaluation, validadores |
| Application / Services | Motor de evaluacion desacoplado de I/O | EvaluationEngine |
| Infrastructure / Services | Acceso a archivos, parsing OOXML y reportes | FileScannerService, ExcelReaderService, ReportGeneratorService |
| ConsoleApp | Bootstrap, Generic Host, DI, configuracion | Program + appsettings.json |

### Principios aplicados

- **SRP**: cada servicio tiene una responsabilidad clara.
- **Open/Closed**: reglas y heuristicas extensibles por configuracion.
- **Dependency Inversion**: orquestacion contra interfaces de Core.
- **Boundary Clean**: Application no depende de detalles de EPPlus ni filesystem.

---

## 📊 Modelo de Ponderacion (100 pts)

| Criterio | Puntaje |
|---|---:|
| Columnas intermedias | 20.0 |
| Promedio X | 5.0 |
| Promedio Y | 5.0 |
| Covarianza | 15.0 |
| Desvio estandar X | 7.5 |
| Desvio estandar Y | 7.5 |
| Correlacion Pearson | 15.0 |
| Interpretacion de Pearson | 10.0 |
| Pendiente recta | 7.5 |
| Ordenada al origen | 7.5 |
| **Total** | **100.0** |

---

## 📊 Trazabilidad y Reporting

Cada corrida genera tres artefactos de auditoria:

1. **Consolidado_Notas.xlsx**
   - Vista ejecutiva de resultados
   - Estado aprobado/reprobado
2. **log_detallado_YYYYMMDD_HHMMSS.json**
   - Arbol de decisiones por alumno/metrica
   - Evidencia tecnica de matching y validacion
3. **resumen_YYYYMMDD_HHMMSS.txt**
   - Totales agregados para lectura rapida docente

Esto habilita trazabilidad tecnica, revisiones posteriores y soporte de decisiones academicas.

---

## 🛠️ Configuracion y Ejecucion

### Requisitos

- .NET SDK 8.0+
- Windows/Linux/macOS

### 1) Clonar repositorio

```bash
git clone <URL_DEL_REPOSITORIO>
cd "Correccion Excel con C#"
```

### 2) Restaurar dependencias

```bash
dotnet restore "src/CorreccionExcel.Console/CorreccionExcel.Console.csproj"
```

### 3) (Opcional) Ajustar parametros

Editar:

- src/CorreccionExcel.Console/appsettings.json

Secciones principales:

- Evaluation (pesos, arrastre, tolerancia, ventana de busqueda)
- Scanner (temas, recursion, inclusion de .xls)

### 4) Estructura esperada de entrada

```text
Evaluaciones/
├─ Tema A/
│  ├─ Alumno1.xlsx
│  └─ Alumno2.xlsx
├─ Tema B/
├─ Tema C/
└─ Tema D/
```

### 5) Ejecutar

```bash
dotnet run --project "src/CorreccionExcel.Console/CorreccionExcel.Console.csproj" -- "C:/Evaluaciones"
```

### 6) Salida

Los resultados se escriben en:

```text
C:/Evaluaciones/Resultados
```

---

## 📌 Estado del Proyecto

- Arquitectura en capas implementada
- Generic Host + configuracion tipada implementados
- Heuristicas anti-hardcode y arrastre operativo validadas en lote real
- Reporting multiformato activo (XLSX + JSON + TXT)

---

## 🤝 Nota de Ingenieria

Este proyecto esta orientado a resolver un problema real de evaluacion automatizada en presencia de datos ruidosos, con foco en **justicia pedagogica**, **resiliencia tecnica** y **auditabilidad empresarial**.
