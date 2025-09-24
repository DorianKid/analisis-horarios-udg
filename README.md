# Análisis de Horarios — UDG 

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
![Made with: Python](https://img.shields.io/badge/Made%20with-Python-blue)

Proyecto en Python para **extraer, limpiar y analizar** horarios de UdeG (SIIAU) y **recomendar los mejores horarios por día** para realizar eventos o conferencias con mayor afluencia potencial. Considera **solapamientos**, una **ventana de 4 horas previas** (el intervalo máximo que se asume que un alumno permanece en la universidad desde su última clase) y un **modelo por semestre**: un semestre **aporta** a la afluencia de una hora solo si **no tiene clase en ese momento**. El scrappeo se realiza **una sola vez**; para análisis posteriores se utiliza el **TXT de datos crudos**. Puede ser aplicado a varias carreras a la vez así como a un grupo tan específico como se requiera.

> **Nota sobre optativas**: Las **materias optativas** no se contabilizan en los cálculos por semestre, ya que no viene especificado a qué semestre pertenecen. Dependiendo de su asignación real, podrían aumentar o disminuir el número de alumnos libres.

---

## ✨ Características

- **Extracción SIIAU** (una sola vez) → **TXT** con datos crudos reutilizable.
- **ETL a Excel por día** con hojas por **semestre**.
- **Cálculo de “Mejor horario por día”** usando:
  - Ventana previa de **4 h**.
  - Candidatos en **horas nones** (09:00, 11:00, 13:00, …) con regla para excluir **07:00** si aplica.
  - Aporte por **semestre sin clase** en el instante candidato.
- **Visualizaciones Top 5** (barras y tabla) listas para compartir.
- **Configuración flexible** en `constantes.py` (semestres por carrera, catálogos).

---

## 🧭 Estructura del proyecto (referencia)

```
analisis-horarios-udg/
├─ Extraer_horarios.py                   # Scraping SIIAU → TXT crudo + Excel ordenado
├─ Calcular_mejor_horario_por_semestres_y_dia.py
│                                        # Calcula el mejor horario por día (modelo por semestre)
├─ Analizar_horarios.py                  # Lee “Mejor_horario.xlsx” y genera visualizaciones Top 5
├─ constantes.py                         # Diccionarios de semestres por carrera y catálogos
├─ README.md                             # Este documento
└─ requirements.txt                      # (opcional) Para instalar dependencias

```

*Si tu estructura difiere, ajusta las rutas en este README.*

---

## 🚀 Ejecución rápida

> Requisitos sugeridos: `python 3.10+`, `pandas`, `openpyxl`, `XlsxWriter`, `matplotlib`.
> Para scraping: `selenium` (+ Chromedriver) y, en Windows, opcional `pywin32` para autofit.

### 1) Crear entorno e instalar dependencias
```bash
python -m venv .venv
# Activar:
#   Windows: .venv\Scripts\activate
#   macOS/Linux: source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
```
> Si no quieres scraping aún, puedes omitir `selenium` del requirements.

### 2) Extraer horarios (solo una vez)
```bash
python Extraer_horarios.py
```
- Guarda un **TXT** con datos crudos reutilizables y un **Excel** de horarios por día.
- En corridas posteriores **no necesitas scrappear**: el script puede leer el **TXT**.

### 3) Calcular “Mejor horario por día” Analizar_horarios.py (modelo por semestre)
```bash
python Analizar_horarios.py
```
- Pide el **Excel ETL** (con `Clave`, `Hora inicio`, `Hora fin`, `Alumnos`) y la **carrera**.
- Genera una carpeta `ANALISIS <NOMBRE>` con:
  - **6 Excels por día** (`HORARIOS_LUNES_...xlsx`, etc.) con **hojas por semestre**.
  - Un **Mejor_horario.xlsx** con dos hojas:
    - **Horarios recomendados**: mejor hora por día y alumnos libres estimados.
    - **Detalle candidatos**: score por cada candidato.

### 4) Visualizaciones Top 5
```bash
python Analizar_horarios.py
```
- Pide la ruta a `Mejor_horario.xlsx`.
- Guarda **Top5_alumnos_libres_estetica.png** y **Top5_resumen_tabla.png** en la misma carpeta.

---

## 🧠 Metodología (resumen)

- **Ventana 4 h previas** a la hora candidata *t*: `[t − 240 min, t)`.
- Para cada **semestre** *s*:
  - Si **no tiene clase en t**, aporta su **máximo simultáneo** dentro de la ventana (evita doble conteo intrasemestre).
  - Si **tiene clase en t**, su aporte es **0**.
- La **suma por todos los semestres** es el score de *t*. Se recomienda el **mayor** (empate → más tardío, configurable).
- **Optativas**: no se consideran por no estar atadas a un semestre específico.

---

## ⚙️ Configuración

- `constantes.py` — define `SEMESTRES_POR_CARRERA` y catálogos (centros, ciclos, etc.).
- Candidatos en **horas nones** y exclusión de **07:00** son configurables en los scripts.
- Las rutas de salida se construyen automáticamente a partir del nombre del Excel (`ANALISIS HORARIOS ...`).

---

## 📸 Capturas (sugeridas)

- `Top5_alumnos_libres_estetica.png` — Barra comparativa con mejor/segundo mejor destacados.
- `Top5_resumen_tabla.png` — Tabla con Día, Mejor hora y Libres estimados.

---

## 🤝 Contribución

Las PRs son bienvenidas. Por favor:
- Mantén el estilo del proyecto y los nombres de archivos.
- Evita dependencias innecesarias.
- Documenta cambios en visualización con capturas.

---

## 📄 Licencia

Este proyecto está bajo **MIT**. Consulta [`LICENSE`](./LICENSE).
