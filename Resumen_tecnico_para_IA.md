# Resumen técnico — Analytics Carrefour Argentina

Herramienta de análisis de categorías compuesta por **4 archivos HTML** que corren completamente en el navegador (sin backend). Todo el procesamiento ocurre en el cliente mediante JavaScript y Web Workers. Los datos se cargan desde un archivo Excel local que el usuario selecciona en cada sesión.

---

## Estructura de archivos

```
Graficos/
├── Menu.html                        ← Menú principal (índice)
└── Graficos/
    ├── Participacion_Barras.html    ← Dashboard 1
    ├── Piramide_Segmentacion.html   ← Dashboard 2
    └── Cobertura_Mercado.html       ← Dashboard 3
```

Disponible también en GitHub Pages: https://juanmanuelgrosscr4.github.io/Graficos/Menu.html

---

## Librerías externas (CDN)

- **Plotly.js 2.27.0** — gráficos interactivos (Participacion_Barras y Cobertura_Mercado)
- **xlsx@0.18.5** — parseo de Excel dentro de Web Workers
- **xlsx-js-style@1.2.0** — generación de Excel con estilos en el hilo principal (Participacion_Barras y Cobertura_Mercado)
- **SweetAlert2@11** — modales de error/advertencia
- **Montserrat / Segoe UI** — tipografías

---

## Estructura del archivo Excel (compartida por los 3 dashboards)

La hoja se llama obligatoriamente **`Export`**. Los encabezados están en la **fila 3** (índice 2 en 0-based). Los datos comienzan en la **fila 4**. Las celdas A1 y B1 se usan como título/subtítulo del dashboard.

Cada dashboard tiene un array `EXPECTED_HEADERS` con los nombres exactos de columna. Si no coinciden, se rechaza el archivo y se muestra un SweetAlert con el detalle de diferencias. La comparación normaliza tildes.

### Columnas fijas (índices 0–based)

| Índice | Nombre | Notas |
|--------|--------|-------|
| 0 | EAN | Código de barras |
| 1 | SECCIÓN | Sección de categoría |
| 2 | PLANOGRAMA | |
| 3 | UN | Unidad de negocio |
| 4 | DESCRIPCIÓN | |
| 5 | STATUS_FC | 1=activo, 2=baja, 3=pendiente alta |
| 6 | FABRICANTE | |
| 7 | MARCA | |
| 8 | CATEGORÍA | Categoría Scentia |
| 9 | CRF UNIDADES | Total unidades CRF (agregado) |
| 10 | CRF IMPORTE NETO | Total importe CRF (agregado) |
| 11–14 | MDO UNIDADES / MDO IMPORTE NETO / … | Métricas agregadas MDO |
| 15 | PLANO-MARCA | |
| 16 | UN-MARCA | |
| 17–21 | PRECIO / PRECIO PLANO / PRECIO UN / … | |
| 22 | RANGO PLANOGRAMA | Segmento de precio (BAJO/MEDIO/ALTO/ALTA GAMA) |
| 23 | RANGO UN | Segmento por UN |

### Bloques regionales (7 regiones × 12 columnas = 84 columnas, índices 24–107)

```
REGION_BASE_START = 24
REGION_STEP = 12
Regiones: ['TOTAL NACIONAL','NOA','LITORAL','CUYO','CÓRDOBA','BS AS','PATAGONIA']
rb = 24 + i * 12
```

Cada bloque de 12 columnas (en orden fijo):
```
rb+0  CRF $       rb+1  %PG $       rb+2  %UN $
rb+3  CRF uni     rb+4  %PG uni     rb+5  %UN uni
rb+6  MDO $       rb+7  %PG $       rb+8  %UN $
rb+9  MDO uni     rb+10 %PG uni     rb+11 %UN uni
```

### Columnas YTD (índices 108–119)

```
YTD Share $ 2025, YTD Share $ 2026, YTD Dif Sh $,
YTD Share VOL 2025, YTD Share VOL 2026, YTD Dif Sh VOL,
YTD % Prog RM $, YTD % Prog CRF $, YTD GAP CRF vs RM $,
YTD % Prog RM VOL, YTD % Prog CRF VOL, YTD GAP CRF vs RM VOL
```

### Columnas de competidores (índice 120+)

Una columna por competidor con valores 0/1 (presencia). El nombre de la columna es el nombre del competidor. La primera columna cuyo nombre incluya "CARREFOUR" (case-insensitive) se separa y se retorna como `crfCol` (índice de presencia propia de CRF). El resto se carga como `compCols: [{name, col}]`.

Existen dos variantes del Excel: **Retail** y **Maxi**. Las columnas de competidores difieren entre ambas. Las columnas fijas (0–119) son idénticas.

---

## Menu.html

Página de índice estática. Contiene tres tarjetas (`<a class="tool-card">`) con vínculos relativos a:
- `Graficos/Participacion_Barras.html`
- `Graficos/Piramide_Segmentacion.html`
- `Graficos/Cobertura_Mercado.html`

Sin lógica JavaScript relevante.

---

## Participacion_Barras.html

### Propósito
Gráfico de barras horizontales espejo (butterfly chart) de participación de mercado por marca/UN/planograma/categoría. Muestra CRF vs MDO en importe y unidades, con modos opcionales de índice de precio, share YTD y progresión.

### Flujo de datos
1. Usuario selecciona Excel → FileReader → ArrayBuffer → Web Worker
2. Worker parsea con `xlsx@0.18.5`, valida encabezados, construye `allData[]` como array de objetos con campos nombrados (`row[reg+'|CRF_IMP']`, etc.)
3. Worker retorna `{ok, allData, a1, b1}` al hilo principal
4. El hilo principal pobla filtros y llama `renderCharts()`

### Estado global
```js
let allData = [];       // array de objetos parseados del Excel
let prevRegiones = ['TOTAL NACIONAL'];
let viewMode = '';      // '' | 'pct' | 'ip' | 'share' | 'progresion'
```

### Filtros (en cascada)
`Sección → Planograma → UN → Marca` (multi-select)
`Región` (multi-select con lógica de exclusión: TOTAL NACIONAL es mutuamente excluyente con regiones individuales)
`Agrupar por` (MARCA | PLANOGRAMA | UN | CATEGORIA)

### Función principal de cálculo: `calcBrands(plans, uns, marcas, regiones, groupBy, seccion)`
- Filtra `allData` por los parámetros recibidos
- Agrupa por la dimensión elegida (`groupBy`)
- Si es caso simple (1 planograma, 1 región, ≤1 UN) usa los porcentajes pre-calculados del Excel (`CRF_PCT_PG_IMP`, etc.)
- Si no, recalcula los % dividiendo sobre el denominador apropiado
- Añade índice de precio (`crfImp/crfUni`), share YTD y progresión desde los mapas de categoría

### Renderizado: `renderMirrorChart(divId, brands, ...)`
Usa Plotly con `orientation:'h'`, `barmode:'overlay'`. Dos trazas por gráfico (CRF y MDO). Hay dos gráficos: importe y unidades.

Modos de visualización disponibles (`viewMode`):
- **`''`** (vacío): barras de volumen absoluto
- **`'pct'`**: barras de participación porcentual
- **`'ip'`**: índice de precio (precio promedio CRF vs MDO). Solo disponible si `groupBy !== 'CATEGORIA'`
- **`'share'`**: share YTD 2026 y diferencial vs 2025. Solo disponible si `groupBy === 'CATEGORIA'`
- **`'progresion'`**: % progresión RM y CRF, y GAP. Solo disponible si `groupBy === 'CATEGORIA'`

### Export
Botón que genera `.xlsx` con `xlsx-js-style`. Exporta todos los grupos visibles con sus valores de importe, unidades, % participación e índice de precio.

---

## Piramide_Segmentacion.html

### Propósito
Pirámides SVG de segmentación de precios: compara distribución de surtido CRF vs MDO por rangos de precio dentro de un planograma/UN/marca seleccionados. Incluye panel de marcas con recomendaciones.

### Estado global
```js
let META = null;     // datos completos una vez cargado el Excel
const RANGOS = ['BAJO','MEDIO','ALTO','ALTA GAMA'];
const SUBREGIONS = ['NOA','LITORAL','CUYO','CORDOBA','BS AS','PATAGONIA'];
```

### Segmentos y colores
```js
const RCOLS = {'ALTA GAMA':'#6d28d9','ALTO':'#1d4ed8','MEDIO':'#2563eb','BAJO':'#60a5fa'};
const SHAPE = {
  'BAJO':      {topF:0.80, botF:1.00},
  'MEDIO':     {topF:0.58, botF:0.80},
  'ALTO':      {topF:0.34, botF:0.58},
  'ALTA GAMA': {topF:0.12, botF:0.34},
};
```

### Flujo de datos
Web Worker parsea el Excel con `xlsx@0.18.5`. Retorna `allData[]` (array de arrays raw de la hoja). La función `buildNode(rowsArr, region)` construye el nodo de métricas para un conjunto de filas y una región.

### `buildNode(rowsArr, region)` — función central
Retorna objeto con:
- `et` total EANs activos (CRF uni > 0 OR MDO uni > 0)
- `ec` EANs activos en CRF (`rb+3 > 0`)
- `em` EANs activos en MDO (`rb+9 > 0`)
- `ci/cu/di/du` importe/unidades CRF y MDO
- `pci/pmi/pcu/pmu` participaciones porcentuales
- `pc/pm` precio promedio CRF y MDO
- `marcas[]` array con métricas por marca:
  - `pi/pu` participación en importe/unidades dentro del segmento
  - `c` presencia CRF (0|1)
  - `r1` recomendación tipo 1: sin CRF y participación MDO ≥3%
  - `r2` recomendación tipo 2: en CRF con STATUS_FC 2 o 3 y participación ≥3%
  - `r3` recomendación tipo 3: en CRF pero con EANs del top-80% de MDO$ sin surtir
  - `mr/cr` cantidad de sub-regiones con presencia MDO/CRF
  - `mrNames` nombres de sub-regiones con presencia MDO (para export)
  - `brows` filas crudas de la marca en el segmento/región actual (para export EAN por EAN)

### Renderizado
SVG generado inline. Cada pirámide es un `<svg>` con trapecios (`<polygon>`) por segmento. El panel de marcas (`#brand-panel`) se renderiza como tarjetas HTML.

### Filtros
Sección → Planograma → UN → Marca (en cascada). Región (selector único). Vista por Importe o Unidades.

### Recomendaciones R1/R2/R3
Se generan dentro de `buildNode()` por marca y se filtran/muestran en el panel lateral. El export de recomendaciones genera un `.xlsx` con detalle EAN por EAN para R3 y resumen por marca para R1/R2.

### `computeType3(brows, b)`
Identifica EANs ausentes en CRF que representan el top 80% de la venta MDO del segmento para esa marca.

---

## Cobertura_Mercado.html

### Propósito
Gráfico de barras apiladas que compara a CRF y competidores por cobertura de surtido, segmentado por rango de precio. Tres modos de visualización: Referencias, % Cobertura, % Participación.

### Constantes clave
```js
const COL = {
  EAN:0, SECCION:1, PLANOGRAMA:2, UN:3, DESCRIPCION:4,
  STATUS_FC:5, FABRICANTE:6, MARCA:7, CATEGORIA:8,
  RANGO:22, RANGO_UN:23
};
const REGION_BASE_START = 24, REGION_STEP = 12;
const REGIONS = ['TOTAL NACIONAL','NOA','LITORAL','CUYO','CÓRDOBA','BS AS','PATAGONIA'];
const RB = {};  // { 'TOTAL NACIONAL': 24, 'NOA': 36, ... }
const ALL_RBS = REGIONS.map((_,i) => REGION_BASE_START + i * REGION_STEP);
const RANGOS = ['PREMIUM','...'];  // segmentos de precio (leídos dinámicamente del Excel)
const ALL_OPT = '__ALL__';
```

### Estado global
```js
let ALLDATA   = [];       // todas las filas del Excel (sin filtros)
let COMP_COLS = [];       // [{name, col}] competidores detectados dinámicamente
let CRF_COL   = -1;      // índice de la columna de presencia CARREFOUR en el Excel
let viewMode  = 'refs';  // 'refs' | 'cob' | 'part'
let rfMode    = 'none';  // modo de recomendación activo: 'none'|'r1'|'r2'|'r3'|'r4'
let showBarText = false;
```

### Web Worker (parseo)
El worker usa `xlsx@0.18.5`. Al parsear:
- Valida encabezados contra `EH[]` (array de 132 nombres esperados)
- Busca a partir del índice 120 columnas de competidores
- La **primera columna** cuyo nombre incluya `'CARREFOUR'` (case-insensitive) se guarda como `crfCol` (índice de presencia propia de CRF)
- El resto se acumula en `compCols: [{name, col}]`
- Retorna `{ok, allData, compCols, crfCol, a1, b1}`

### Dos conjuntos de filas en render()
```js
// rowsRef: base para Referencias (sin filtro de cadenas ni MDO)
let rowsRef = getSecBase()  // filtrado por Sección
             .filter(planograma)
             .filter(UN)
             .filter(Marca)

// rows: misma base + filtro de presencia en ≥2 cadenas
let rows = getRows()  // incluye filtro de ≥2 cadenas

// rowsMdo: rows + filtro MDO$ > 0 en regiones seleccionadas
const rowsMdo = rows.filter(r => rbs.some(rb => num(r[rb+6]) > 0))
```

La lógica del filtro ≥2 cadenas:
```js
// CRF cuenta si tiene rb+3 (CRF uni) > 0 en alguna región seleccionada
// Cada competidor cuenta si r[c.col] > 0
// EAN se incluye si cnt >= 2
```

### `calcComp(rowsRef, rowsMdo, rbs, isCRF, col)` — función central
Calcula métricas por segmento y totales para un competidor (o CRF).

Retorna objeto con claves `RANGOS[]` + `'_t'` (total), cada una con:
```js
{
  total,        // EANs en rowsMdo en ese segmento (para contexto)
  cov,          // EANs con presencia en el competidor (desde rowsRef) → usado en refs y part
  covMdoCount,  // EANs en universo ≥2 cadenas con presencia → usado en cob
  mdo,          // suma MDO$ (o CRF$ si isCRF) en rowsMdo del segmento
  covMdo        // suma MDO$ (o CRF$) solo de los EANs con presencia
}
```

Detectores de presencia dentro de `calcComp`:
```js
// Para refs (región-independiente): usa col directamente (col = CRF_COL o c.col)
const hasCovRef = r => num(r[col]) > 0;

// Para cob/part (región-dependiente): CRF usa ventas regionales, competitors usan col
const hasCov = r => isCRF
  ? rbs.some(rb => num(r[rb+3]) > 0)   // CRF uni en regiones seleccionadas
  : num(r[col]) > 0;
```

Para CRF, `covMdoCount` se calcula como `rowsMdo.filter(hasCovRef).length` (EANs en universo ≥2 cadenas con presencia CRF), NO usando ventas regionales.

### Modos de visualización

**Referencias** (`viewMode='refs'`):
- Valor por segmento: `segD.cov` (EANs con presencia desde rowsRef)
- Valor total: `d._t.cov`
- Formato anotación: entero con separador de miles (punto)
- No cambia con región

**Cobertura** (`viewMode='cob'`):
- Valor por segmento: `segD.covMdoCount / mktT * 100`
- `mktT = rowsMdo.length` (universo total ≥2 cadenas)
- Mismo denominador para CRF y todos los competidores
- Formato: `x.x%`

**Participación** (`viewMode='part'`):
- Valor por segmento: `segD.cov / fullD._t.cov * 100`
- Denominador = surtido propio del competidor (no el universo)
- Cada barra suma 100%
- Formato segmentos: `x.x%`; formato anotación: entero (refs propias)

### Funciones de valor en render()
```js
// Requieren fullD (objeto completo del competidor, no solo el segmento)
function getBarSegVal(n, segD, fullD) { ... }
function getBarTotVal(n, d) { ... }
```

### Filtros y cascada
```
Región  →  (no afecta Referencias)
Sección →  Planograma → UN → Marca  (afecta Referencias, Cob y Part)
Competidores (selector múltiple, solo afecta qué barras se muestran)
```

**Importante:** el selector de Región NO modifica `rowsRef` (y por tanto no cambia el conteo de Referencias). Solo afecta `rbs` (regiones seleccionadas), que se usa en `rowsMdo` y en `hasCov` para ventas.

### `fmtN(n)` — separador de miles
```js
const fmtN = n => new Intl.NumberFormat('es-AR').format(Math.round(n));
// 13717 → "13.717"
```

### Recomendaciones R1–R4
Todos usan `ALLDATA` como base (sin filtros de rowsMdo).

- **R1**: EANs activos en MDO pero ausentes en CRF (`r[CRF_COL] == 0` o sin datos)
- **R2**: Marcas presentes en CRF con STATUS_FC 2 o 3 y participación ≥3%
- **R3**: EANs faltantes en CRF de marcas que CRF tiene (top-80% MDO$ de la marca no surtidos)
- **R4**: EANs con participación >3% del total MDO presentes en exactamente 1 cadena (CRF o 1 competidor)

### Export
- **Exportar Gráfico visible** (`exportChart()`): genera `.xlsx` con refs/cob/part por segmento para cada competidor visible
- **Exportar Recomendaciones** (`exportRec()`): genera `.xlsx` con detalle del modo rfMode activo

Ambos usan `xlsx-js-style@1.2.0`.

### Renderizado Plotly
```js
Plotly.react('chart', traces, layout, config)
// barmode: 'stack'
// bargap calculado dinámicamente para que cada barra no supere 93px de ancho
// textangle: 0  (evita rotación de etiquetas en segmentos pequeños)
// constraintext: 'none'  (mantiene el tamaño de fuente en segmentos pequeños)
// textfont: { size:10, color:'#fff' }
// CSS: #chart .barlayer text { paint-order: stroke fill; stroke: rgba(0,0,0,0.8); stroke-width: 2.5px }
```

Anotaciones de total (sobre cada barra) usando `layout.annotations[]` con `yshift: showBarText ? 10 : 4`.

---

## Patrones comunes a los 3 dashboards

### Web Worker inline
Cada dashboard crea un worker desde un blob de string para evitar bloquear el hilo principal durante el parseo del Excel (puede tardar varios segundos en archivos grandes):
```js
const blob = new Blob([WORKER_SRC], {type:'application/javascript'});
const w = new Worker(URL.createObjectURL(blob));
w.postMessage(arrayBuffer);
w.onmessage = e => { /* procesar resultado */ };
```

El worker usa `self.importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/...')`.

### Validación de encabezados
Todos los dashboards validan la fila de índice 2 del Excel contra un array `EXPECTED_HEADERS`. La validación normaliza tildes (`normalize('NFD').replace(/[̀-ͯ]/g,'')`) y solo valida las columnas fijas (0–119). Las columnas de competidores (120+) se detectan dinámicamente.

### Botón "Volver al menú"
Los tres dashboards tienen un botón en el header que navega a `../Menu.html` (relativo).

### Separador de miles
`Intl.NumberFormat('es-AR')` — usa punto como separador de miles, coma como decimal.

### SweetAlert2
Usado para mostrar errores de validación del Excel con lista detallada de columnas incorrectas (máximo 5 mostradas en el modal; el resto en la consola del navegador).

---

## Lo que NO hace la herramienta

- No persiste datos entre sesiones (sin localStorage, sin backend)
- No modifica el Excel cargado
- No sube datos a ningún servidor
- No soporta múltiples archivos simultáneos (solo uno por sesión)
