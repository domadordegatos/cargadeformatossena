# Generador de Formatos SENA

Carga un Excel con el listado de aprendices y descarga los formatos oficiales de
etapa productiva ya diligenciados: Compromiso del Aprendiz, Autorización de
Tratamiento de Datos y el formato GFPI de Planeación, Seguimiento y Evaluación.

**App en vivo:** https://formatosena.nncoms.com

El trabajo que automatiza es real y repetitivo: un instructor con tres fichas a
cargo puede tener que diligenciar a mano los mismos tres formatos para ochenta
aprendices, copiando nombre, documento, acudiente y programa en cada uno. Esta
herramienta convierte eso en subir un archivo y esperar unos segundos.

## Lo importante primero: los datos no salen del navegador

Los formatos de etapa productiva contienen datos personales de **menores de
edad** — nombre, tipo y número de documento, dirección, teléfonos, datos del
acudiente y, si se usa, la firma escaneada. Por eso la aplicación está construida
como un SPA sin backend:

- No hay servidor que reciba el Excel. El archivo se lee en memoria con
  [`xlsx`](https://www.npmjs.com/package/xlsx) y [`exceljs`](https://www.npmjs.com/package/exceljs),
  y los documentos se arman en el mismo navegador con
  [`docxtemplater`](https://docxtemplater.com/) y [`pizzip`](https://www.npmjs.com/package/pizzip).
- No hay peticiones de red con datos: el único `fetch` del código carga las
  plantillas `.docx` que vienen con la app.
- No se usa `localStorage`, `sessionStorage`, IndexedDB ni cookies de rastreo
  para los datos de los aprendices. Al cerrar la pestaña no queda nada.
- El hosting (Firebase) sirve HTML, CSS, JS y las plantillas. Nunca ve un dato
  personal.

La app incluye su propia política de tratamiento de datos y un aviso que el
instructor debe aceptar antes de cargar el archivo, en línea con la Ley 1581 de
2012 de habeas data.

Esto deja la responsabilidad donde corresponde: **el archivo Excel fuente es del
instructor**. La app no lo custodia, pero tampoco lo expone.

## Qué genera

| Formato | Plantilla | Cuándo aplica |
|---|---|---|
| **Compromiso del Aprendiz** | `compromiso.docx` | Todos los aprendices |
| **Autorización de Tratamiento de Datos** | `tratamiento.docx` | Aprendices menores de edad, firmada por el acudiente |
| **Planeación, Seguimiento y Evaluación de Etapa Productiva** | `formato_seguimiento.xlsx` | Por grupo de aprendices, con los tres momentos del GFPI |

Las tres plantillas en `src/assets/templates/` son los formatos oficiales **en
blanco**, con marcadores (`[nombre]`, `[documento]`, `[acudiente]`…) que la app
reemplaza. No contienen datos de nadie.

Si hay un solo aprendiz, los archivos se descargan sueltos. Si son varios, se
empaquetan en un ZIP.

## El Excel de entrada

La app genera una plantilla de ejemplo desde la misma interfaz, así que no hace
falta armarla a mano. Las columnas son:

```
grupo      nombre           firma         documento  tipo  otro
acudiente  tipopadre        docacudiente  lugardoc
correo     telefono1        telefono2     direccion
ficha      dia              mes
programa   ciudadmunicipio  regional
```

Solo **`nombre`** y **`documento`** son obligatorias. `tipo` acepta `TI`, `CC`,
`CE` o `PPT`; si es otro, se escribe en `otro`. La columna `firma` admite
imágenes incrustadas en la celda, que se insertan en el documento generado.

Los campos que se repiten en toda una ficha —`ficha`, `programa`, `regional`,
`ciudadmunicipio`, fecha— se pueden fijar una sola vez desde la interfaz en vez
de repetirlos en cada fila, y la app avisa si detecta fichas distintas mezcladas
en el mismo archivo.

Antes de generar, valida el archivo y señala filas duplicadas por nombre o
documento, para no emitir dos veces el mismo formato.

## Cómo correrlo

```bash
npm install
npm start
```

Queda en `http://localhost:4200/`. No necesita variables de entorno, base de
datos ni credenciales: no hay backend que configurar.

Para compilar:

```bash
npm run build
```

El resultado queda en `dist/seguidor-formatos/`.

## Estructura

```
src/app/
├── app.component.ts           # Interfaz: carga, edición en tabla, selección, grupos
├── models/
│   └── estudiante.model.ts    # Tipos de aprendiz y de los tres momentos del seguimiento
└── services/
    ├── excel.service.ts       # Lectura y validación del Excel, plantilla de ejemplo, firmas
    ├── document.service.ts    # Llenado de los .docx con docxtemplater
    └── seguimiento.service.ts # Armado del formato GFPI de etapa productiva
```

## Tecnologías

Angular 16 · TypeScript · docxtemplater + PizZip · ExcelJS + SheetJS ·
FileSaver · driver.js (tour guiado la primera vez) · Firebase Hosting

## Contexto

Hecho para el SENA — Centro Agroindustrial y Fortalecimiento Empresarial de
Casanare (CAFEC), para las fichas en articulación con la media técnica. Los
formatos corresponden al procedimiento GFPI de Gestión de Formación Profesional
Integral.
