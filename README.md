# Sistema de Extraccion y Consolidacion de Votos E-14
### Elecciones Presidenciales Colombia 2026 — Registraduria Nacional del Estado Civil

Herramienta 100% automatizada que lee los formularios **E-14** (Acta de Escrutinio de Jurados de Votacion) en formato PDF, extrae los votos por candidato usando vision artificial y genera un archivo Excel consolidado con totales verificables.

---

## Que hace

1. Recorre todas las subcarpetas de `E14/` buscando archivos PDF.
2. Convierte cada pagina del PDF a imagen (200 DPI).
3. Envia cada imagen a la API de **Claude claude-haiku-4-5** (Anthropic) con un prompt especializado para leer las marcas tachadas del formulario E-14.
4. Extrae por cada mesa: candidatos, votos, votos en blanco, nulos, no marcados y la suma total del acta.
5. Genera `consolidado_votos_2026.xlsx` con:
   - **Hoja CONSOLIDADO**: tabla resumen con el total de votos por candidato sumado de todos los lugares y mesas, ordenado de mayor a menor, con GRAN TOTAL al final.
   - **Una hoja por lugar** (ej. CARPINTERO): detalle fila a fila de cada mesa con dos filas de verificacion al final de cada PDF:
     - Fila azul `TOTAL ACTA (IA)`: valor leido directamente del formulario por Claude.
     - Fila verde `TOTAL CALCULADO`: formula `=SUM(...)` que Excel calcula para verificar que cuadre.

---

## Estructura del proyecto

```
conteo_votos_2026_presidente_col/
│
├── E14/                          # Carpetas por lugar de votacion
│   └── CARPINTERO/
│       ├── E14_XXX_..._001.pdf
│       └── E14_XXX_..._002.pdf
│
├── Release-26.02.0-0/            # Poppler para Windows (conversion PDF a imagen)
│   └── poppler-26.02.0/
│       └── Library/bin/
│
├── extractor_votos_e14.py        # Script principal
├── consolidado_votos_2026.xlsx   # Resultado generado
├── extraccion_votos.log          # Log de ejecucion
├── .env                          # API Key de Anthropic (NO subir a git)
└── .gitignore
```

---

## Requisitos

### Python
Version recomendada: **Python 3.10 o superior**

### Dependencias
Instalar con:

```bash
pip install anthropic python-dotenv pandas openpyxl Pillow pdf2image
```

| Libreria | Uso |
|---|---|
| `anthropic` | Cliente oficial de la API de Claude (Anthropic) |
| `python-dotenv` | Carga la API Key desde el archivo `.env` |
| `pandas` | Manipulacion y agrupacion de datos |
| `openpyxl` | Escritura del archivo Excel con estilos y formulas |
| `Pillow` | Procesamiento de imagenes PIL |
| `pdf2image` | Conversion de paginas PDF a imagenes |

### Poppler (Windows)
`pdf2image` requiere Poppler para funcionar en Windows.

1. Descargar desde: https://github.com/oschwartz10612/poppler-windows/releases
2. Descomprimir dentro del proyecto (el script lo detecta automaticamente buscando `pdftoppm.exe`).
3. La ruta esperada es:
   ```
   Release-26.02.0-0/poppler-26.02.0/Library/bin/
   ```

---

## Configuracion

### 1. Obtener API Key de Anthropic

1. Crear cuenta en https://console.anthropic.com
2. Ir a **API Keys** → **Create Key**
3. Copiar la clave (formato: `sk-ant-api03-...`)

### 2. Crear el archivo `.env`

En la raiz del proyecto crear el archivo `.env`:

```
ANTHROPIC_API_KEY=sk-ant-api03-TU_CLAVE_AQUI
```

> **IMPORTANTE**: No subir este archivo a GitHub ni compartirlo. Incluirlo en `.gitignore`.

### 3. Organizar los PDFs

Crear una subcarpeta dentro de `E14/` por cada lugar de votacion y copiar ahi los PDFs correspondientes:

```
E14/
├── CARPINTERO/
│   ├── E14_XXX_X_11_052_099_01_000_X_XXX (1).pdf
│   └── E14_XXX_X_11_052_099_01_000_X_XXX (2).pdf
├── OTRO_LUGAR/
│   └── ...
```

---

## Ejecucion

```bash
python extractor_votos_e14.py
```

El script imprime el progreso en consola y en `extraccion_votos.log`:

```
✓ API Key de Anthropic cargada (modelo: claude-haiku-4-5)
✓ Poppler encontrado automaticamente: ...
============================================================
LUGAR: CARPINTERO (2 PDF(s))
============================================================
  ► E14_XXX_X_11_052_099_01_000_X_XXX (1).pdf
    → Rasterizando a 200 DPI...
    → Pagina 1/3
    → Analizando con Claude claude-haiku-4-5: ..._p001
    └─ Pagina 1: 7 registro(s) extraido(s)
    ...
    → Suma total leida del acta: 217
EXPORTANDO 32 REGISTROS A EXCEL...
  ✓ Hoja 'CONSOLIDADO': 17 candidatos unicos
  ✓ Hoja 'CARPINTERO': 32 registro(s) + 4 filas de totales
✓ Excel guardado en: .../consolidado_votos_2026.xlsx
PROCESO FINALIZADO.
```

---

## Como leer el Excel generado

### Hoja CONSOLIDADO
Tabla unica con todos los candidatos y sus votos totales sumados de **todos los lugares y mesas**:

| Candidato | Total Votos |
|---|---|
| IVAN CEPEDA CASTRO | 165 |
| ABELARDO DE LA ESPRIELLA | 155 |
| PALOMA VALENCIA LASERNA | 83 |
| ... | ... |
| VOTOS EN BLANCO | 14 |
| VOTOS NULOS | 12 |
| VOTOS NO MARCADOS | 5 |
| **GRAN TOTAL** | **=SUM(...)** |

### Hoja por lugar (ej. CARPINTERO)
Detalle de cada mesa con sus candidatos y al final de cada PDF:

- Fila **azul** — `TOTAL ACTA (IA)`: valor que leyó Claude directamente de la celda "SUMA TOTAL" del formulario.
- Fila **verde** — `TOTAL CALCULADO`: formula Excel `=SUM(...)` que suma los votos del grupo. Si coincide con la fila azul, el acta esta cuadrada.

---

## Como funciona la lectura de votos (OCR)

El formulario E-14 tiene una columna de votacion con 3 posiciones (centenas, decenas, unidades). Los jurados rellenan las posiciones vacias con simbolos decorativos (`*`, `•`, `-`, `×`). El prompt enviado a Claude le indica:

- Ignorar todos los simbolos que no sean digitos `0-9`
- Concatenar solo los digitos reales de izquierda a derecha
- Ejemplos clave:

| Celda visual | Valor correcto |
|---|---|
| `✱ ✱ ✱` | 0 |
| `✱ ✱ 1` | 1 |
| `✱ ✱ 2` | 2 |
| `✱ 1 0` | 10 |
| `✱ 3 5` | 35 |
| `✱ 7 5` | 75 |
| `✱ 9 0` | 90 |
| `2 1 7` | 217 |

---

## Parametros configurables

En la seccion `CONFIGURACION GLOBAL` del script:

| Variable | Valor por defecto | Descripcion |
|---|---|---|
| `DPI_RASTERIZACION` | `200` | Resolucion de conversion PDF a imagen. Mayor = mas preciso, mas lento. |
| `JPEG_CALIDAD` | `92` | Calidad de compresion de la imagen enviada a la API. |
| `MODELO_CLAUDE` | `claude-haiku-4-5` | Modelo de Anthropic. Haiku es rapido y economico. |
| `MAX_REINTENTOS` | `3` | Intentos por pagina si la API falla. |
| `PAUSA_ENTRE_REINTENTOS` | `5` | Segundos de espera entre reintentos. |

Para mayor precision en casos dificiles se puede cambiar `MODELO_CLAUDE` a `claude-sonnet-4-5` (mas preciso, mayor costo).

---

## Costo estimado

Con `claude-haiku-4-5`:

| Escala | PDFs | Paginas aprox. | Costo estimado |
|---|---|---|---|
| Prueba (1 lugar) | 1 | 3 | ~$0.01 USD |
| Proyecto completo | 77 |231 | ~$0.77 USD |

---

## Tecnologias utilizadas

| Tecnologia | Version | Rol |
|---|---|---|
| Python | 3.10+ | Lenguaje principal |
| Anthropic Claude claude-haiku-4-5 | API v1 | Vision OCR de formularios |
| pdf2image + Poppler | 1.17+ / 26.02 | Conversion PDF a imagen |
| Pillow | 10+ | Procesamiento de imagenes |
| pandas | 2+ | Agrupacion y analisis de datos |
| openpyxl | 3+ | Generacion de Excel con estilos y formulas |
| python-dotenv | 1+ | Gestion segura de credenciales |
