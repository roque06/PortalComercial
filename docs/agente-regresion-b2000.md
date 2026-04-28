# Agente de Regresion B2000

## Objetivo

Crear un agente que reciba una modificacion tecnica del core bancario B2000, identifique las pantallas y objetos afectados, consulte sus dependencias en Oracle `QA3` y entregue una recomendacion de regresion por modulos.

## Caso evaluado

- Expediente: `GDI-8254`
- Aplicacion: `B2000`
- Modulo indicado en documento: `PA`
- Ambiente sugerido: `QA3`
- Usuario Oracle: `admin`
- Contrasena Oracle: `admin`
- Conexion `QA3` encontrada en `dbConnections.json`:
  `(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=172.27.3.10)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=QA3)))`

## Lo que se pudo extraer del documento

Del archivo `Modificacion Tecnica y Pase a Calidad GDI-8254.docx` se identifican estos elementos:

- Pantallas:
  - `pamantra`
  - `pamantrg`
- Objetos PL/SQL mencionados:
  - `genera_ncf_credito`
  - `genera_ncf_mov`
- Contexto funcional:
  - `Incidente 1400612 / 1271579 - Saltos en Secuencia de Comprobantes Fiscales (TF)`

## Propuesta de funcionamiento del agente

El agente debe trabajar en 5 pasos:

1. Leer la modificacion tecnica.
2. Extraer nombres de pantallas, formas, paquetes, procedures, tablas y palabras clave funcionales.
3. Consultar en `QA3` las dependencias tecnicas de cada pantalla y objeto.
4. Traducir esas dependencias a modulos funcionales candidatos a regresion.
5. Entregar una salida corta y accionable:
   - pantalla principal afectada
   - objetos relacionados
   - modulos a probar
   - motivo tecnico
   - prioridad de regresion

## Entradas del agente

- Documento Word o texto pegado de modificacion tecnica.
- Ambiente Oracle a consultar, por defecto `QA3`.
- Credenciales fijas:
  - usuario: `admin`
  - contrasena: `admin`
- Catalogo de conexiones desde `dbConnections.json`.

## Salida esperada

Ejemplo de formato:

```text
Pantallas impactadas:
- pamantra
- pamantrg

Objetos relacionados:
- genera_ncf_credito
- genera_ncf_mov

Modulos recomendados para regresion:
- PA / transacciones con comprobantes fiscales
- flujos que generen NCF por movimiento
- flujos que generen NCF por credito
- consultas/reportes de comprobantes fiscales electronicos

Motivo:
Las pantallas llaman procedimientos que generan secuencias fiscales; cualquier proceso que consuma esas secuencias puede verse afectado por cambios de numeracion, validacion o persistencia.
```

## Consulta tecnica que debe hacer el agente

La parte clave no es solo leer el documento, sino enriquecerlo con Oracle. El agente debe consultar:

- dependencias entre paquetes, procedures, functions y tablas
- referencias cruzadas desde codigo fuente PL/SQL
- formularios/pantallas que reutilicen los mismos objetos
- tablas de comprobantes, movimientos, creditos y secuencias fiscales

## SQL sugerido para Oracle QA3

### 1. Buscar referencias de las pantallas y procedimientos

```sql
select owner, name, type, line, text
from all_source
where upper(text) like '%PAMANTRA%'
   or upper(text) like '%PAMANTRG%'
   or upper(text) like '%GENERA_NCF_CREDITO%'
   or upper(text) like '%GENERA_NCF_MOV%'
order by owner, name, type, line;
```

### 2. Ver dependencias declaradas de objetos PL/SQL

```sql
select owner,
       name,
       type,
       referenced_owner,
       referenced_name,
       referenced_type
from all_dependencies
where upper(name) in ('GENERA_NCF_CREDITO', 'GENERA_NCF_MOV')
   or upper(referenced_name) in ('GENERA_NCF_CREDITO', 'GENERA_NCF_MOV')
order by owner, name, referenced_owner, referenced_name;
```

### 3. Encontrar tablas tocadas por la logica fiscal

```sql
select owner, name, type, line, text
from all_source
where upper(text) like '%COMPROB%'
   or upper(text) like '%NCF%'
   or upper(text) like '%E-CF%'
order by owner, name, type, line;
```

### 4. Confirmar objetos exactos existentes

```sql
select owner, object_name, object_type, status
from all_objects
where upper(object_name) in ('GENERA_NCF_CREDITO', 'GENERA_NCF_MOV', 'PAMANTRA', 'PAMANTRG')
order by owner, object_type, object_name;
```

### 5. Buscar quien invoca la generacion de NCF

```sql
select owner, name, type, line, text
from all_source
where upper(text) like '%GENERA_NCF_CREDITO(%'
   or upper(text) like '%GENERA_NCF_MOV(%'
order by owner, name, type, line;
```

## Regla de decision para recomendar regresion

El agente debe asignar regresion por capas:

- Regresion directa:
  - la pantalla modificada
  - el proceso funcional descrito en la modificacion tecnica
- Regresion por dependencia:
  - otras pantallas que llamen los mismos procedures
  - reportes o cierres que lean las mismas tablas
- Regresion por riesgo:
  - procesos de numeracion o secuencia
  - validaciones fiscales
  - anulaciones, reversos o reprocesos

## Propuesta concreta para este expediente GDI-8254

Sin haber ejecutado aun SQL sobre `QA3`, mi recomendacion inicial de regresion para `GDI-8254` es:

- `PA` transaccional que use `pamantra`
- `PA` transaccional que use `pamantrg`
- cualquier flujo que genere `NCF` por credito
- cualquier flujo que genere `NCF` por movimiento
- consultas o reportes donde se validen secuencias de comprobantes fiscales
- escenarios de reproceso, reverso o repeticion de transacciones que puedan dejar huecos de numeracion

## Arquitectura recomendada del agente

### Opcion recomendada

Un agente con dos componentes:

- Analizador documental:
  - extrae pantallas y objetos del Word o texto pegado
- Analizador Oracle:
  - ejecuta SQL en `QA3`
  - arma el mapa de dependencias
  - clasifica el impacto por modulo

### Resultado final del agente

Debe devolver algo asi:

```json
{
  "expediente": "GDI-8254",
  "ambiente": "QA3",
  "pantallas_detectadas": ["pamantra", "pamantrg"],
  "objetos_detectados": ["genera_ncf_credito", "genera_ncf_mov"],
  "modulos_regresion": [
    "PA - transacciones en pamantra",
    "PA - transacciones en pamantrg",
    "Procesos de generacion de NCF por credito",
    "Procesos de generacion de NCF por movimiento",
    "Reporteria o consultas de comprobantes fiscales"
  ],
  "riesgo": "medio-alto",
  "motivo": "Cambio en logica de secuencia fiscal con posible impacto transversal"
}
```

## Siguiente paso recomendado

Para que esto pase de propuesta a agente util de verdad, hace falta una de estas dos cosas:

- instalar un cliente Oracle en este entorno para ejecutar las consultas a `QA3`
- o exponer un script/servicio interno que consulte Oracle y le devuelva al agente el mapa de dependencias

Con eso, el agente ya puede producir recomendaciones de regresion basadas en evidencia y no solo en lectura del documento.
