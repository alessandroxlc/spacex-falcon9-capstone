# 🚀 Predicción del aterrizaje de la primera etapa del SpaceX Falcon 9

### IBM Applied Data Science Capstone · Proyecto final

![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?logo=pandas&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-F7931E?logo=scikitlearn&logoColor=white)
![Plotly Dash](https://img.shields.io/badge/Plotly%20Dash-3F4F75?logo=plotly&logoColor=white)
![Folium](https://img.shields.io/badge/Folium-77B829)
![License](https://img.shields.io/badge/licencia-MIT-lightgrey)

> **Pregunta:** dado lo que sabemos de una misión (cohete, carga, órbita, sitio de lanzamiento…),
> ¿aterrizará con éxito la primera etapa del Falcon 9?

---

## Índice

1. [Resumen ejecutivo](#resumen-ejecutivo)
2. [Contexto y problema](#contexto-y-problema)
3. [Estructura del repositorio](#estructura-del-repositorio)
4. [Metodología](#metodología)
5. [Resultados](#resultados)
6. [Conclusiones e ideas innovadoras](#conclusiones-e-ideas-innovadoras)
7. [Cómo reproducir el proyecto](#cómo-reproducir-el-proyecto)
8. [Correspondencia con la rúbrica](#correspondencia-con-la-rúbrica)

---

## Resumen ejecutivo

Se construyó una tubería de ciencia de datos completa sobre los lanzamientos del Falcon 9 entre 2010 y 2020:

* **Recolección** desde la API REST de SpaceX y mediante *web scraping* de Wikipedia.
* **Preparación** de los datos y creación de la etiqueta `Class` (1 = aterrizó, 0 = no aterrizó).
* **Análisis exploratorio** con gráficos (Matplotlib/Seaborn) y 10 consultas **SQL**.
* **Analítica interactiva** con mapas de **Folium** y un dashboard de **Plotly Dash**.
* **Modelado predictivo** con cuatro clasificadores (regresión logística, SVM, árbol de decisión y KNN)
  optimizados con `GridSearchCV`.

Los modelos alcanzan alrededor del **83 % de exactitud** sobre los 18 lanzamientos del conjunto de prueba
(83,33 % = 15 de 18 en la configuración del laboratorio de IBM). Las cifras definitivas de cada modelo se
guardan en [`results/model_results.json`](results/) al ejecutar el notebook 07.

## Contexto y problema

SpaceX anuncia lanzamientos del Falcon 9 por unos **62 millones de dólares**, frente a los **más de 165
millones** de otros proveedores. La diferencia se debe, sobre todo, a que **SpaceX recupera y reutiliza la
primera etapa** del cohete. Si podemos predecir si esa etapa aterrizará, podemos estimar el coste de un
lanzamiento: información valiosa para una empresa que quiera competir con SpaceX en una licitación.

Preguntas que responde el proyecto:

1. ¿Qué variables influyen en que la primera etapa aterrice?
2. ¿Cómo ha evolucionado la tasa de éxito con el tiempo?
3. ¿Influye la ubicación del sitio de lanzamiento?
4. ¿Con qué exactitud se puede predecir el resultado?

## Estructura del repositorio

```
spacex-falcon9-capstone/
├── README.md
├── requirements.txt
├── LICENSE
├── notebooks/
│   ├── 01_data_collection_api.ipynb          # Recolección con la API REST de SpaceX
│   ├── 02_data_collection_webscraping.ipynb  # Web scraping de Wikipedia (BeautifulSoup)
│   ├── 03_data_wrangling.ipynb               # Limpieza y etiqueta Class
│   ├── 04_eda_sql.ipynb                      # 10 consultas SQL (SQLite)
│   ├── 05_eda_visualization.ipynb            # Gráficos EDA + One-Hot Encoding
│   ├── 06_interactive_map_folium.ipynb       # Mapas interactivos y distancias
│   └── 07_machine_learning_prediction.ipynb  # 4 modelos, GridSearchCV y matrices de confusión
├── dashboard/
│   └── spacex_dash_app.py                    # Aplicación Plotly Dash
├── presentation/
│   ├── generate_presentation.py              # Genera la presentación con python-pptx
│   ├── SpaceX_Falcon9_Capstone.pptx          # Presentación editable
│   └── Data Science Capstone Project Report.pdf  # PDF para la entrega
├── scripts/
│   ├── run_all.py                            # Ejecuta todo de principio a fin
│   └── capture_screenshots.py                # Capturas de los mapas y del dashboard
├── data/       # CSV generados por los notebooks
├── images/     # Gráficos, capturas de mapas y del dashboard
├── maps/       # Mapas Folium en HTML interactivo
├── results/    # Cifras en JSON (SQL, EDA, modelos, dashboard)
├── docs/
│   └── GUIA_SENCILLA.md                      # Explicación paso a paso y cómo entregar
└── .github/workflows/run-project.yml         # Ejecuta todo en GitHub Actions
```

## Metodología

| Fase | Qué se hizo | Notebook |
|---|---|---|
| Recolección (API) | `GET` a `api.spacexdata.com/v4`, `json_normalize`, resolución de IDs en `/rockets`, `/launchpads`, `/payloads` y `/cores`, filtro de Falcon 9 e imputación de `PayloadMass` con la media | [01](notebooks/01_data_collection_api.ipynb) |
| Recolección (scraping) | `requests` + `BeautifulSoup` sobre la revisión fija de Wikipedia, extracción de cabeceras y filas de las tablas de lanzamientos | [02](notebooks/02_data_collection_webscraping.ipynb) |
| Data wrangling | Análisis de nulos y tipos, recuentos por sitio y órbita, conversión de `Outcome` en `Class` | [03](notebooks/03_data_wrangling.ipynb) |
| EDA con SQL | Sitios únicos, filtros con `LIKE`, `SUM`, `AVG`, `MIN`, `GROUP BY`, subconsulta y ranking por fechas | [04](notebooks/04_eda_sql.ipynb) |
| EDA con visualización | Dispersión (vuelo/carga vs. sitio y órbita), barras (éxito por órbita), línea (tendencia anual) y One-Hot Encoding | [05](notebooks/05_eda_visualization.ipynb) |
| Mapas con Folium | Marcadores de sitios, `MarkerCluster` de éxitos/fallos, `MousePosition` y distancias haversine | [06](notebooks/06_interactive_map_folium.ipynb) |
| Dashboard con Plotly Dash | Desplegable de sitios, gráfico de pastel, control de rango de carga y dispersión carga vs. resultado | [dashboard](dashboard/spacex_dash_app.py) |
| Análisis predictivo | `StandardScaler`, división 80/20 (`random_state=2`), `GridSearchCV` (cv=10) para 4 modelos, exactitud y matriz de confusión | [07](notebooks/07_machine_learning_prediction.ipynb) |

## Resultados

**EDA**

* La tasa de aterrizaje crece con el número de vuelo: la experiencia acumulada mejora el resultado.
* La tasa de éxito anual sube de forma sostenida desde 2013.
* Las órbitas ES-L1, GEO, HEO y SSO muestran 100 % de éxito, pero con muy pocos lanzamientos; GTO, la
  órbita comercial más habitual, ronda el 50 %.

**SQL**

* Hay 4 sitios de lanzamiento: CCAFS LC-40, CCAFS SLC-40, KSC LC-39A y VAFB SLC-4E.
* El primer aterrizaje exitoso en tierra llegó más de cinco años después del primer lanzamiento.
* Casi todas las misiones (poner la carga en órbita) son exitosas: lo difícil es recuperar el booster.

**Analítica interactiva**

* Todos los sitios están junto a la costa, cerca de ferrocarril y carreteras, y lejos de las ciudades.
* El dashboard permite comparar sitios y rangos de carga; KSC LC-39A destaca por su tasa de éxito.

**Modelos**

| Modelo | Exactitud en prueba |
|---|---|
| Regresión logística | ≈ 83 % |
| SVM | ≈ 83 % |
| Árbol de decisión | ≈ 83 % (varía con la semilla) |
| KNN | ≈ 83 % |

Con solo 18 casos de prueba, cada acierto vale 5,6 puntos y varios modelos empatan; se desempata con la
exactitud de validación cruzada. El error más frecuente es el **falso positivo** (predecir que aterriza
cuando no lo hizo). Los valores exactos están en `results/model_results.json` y en la presentación.

## Conclusiones e ideas innovadoras

1. La **experiencia** (número de vuelo), la **órbita** y el **sitio** son las variables más informativas.
2. Un modelo sencillo acierta la gran mayoría de los lanzamientos de prueba: suficiente para una primera
   estimación del coste de una misión.
3. Próximos pasos propuestos: añadir lanzamientos posteriores a 2020, incorporar meteorología y estado del
   mar, estimar probabilidades calibradas (coste esperado), validar de forma temporal, explicar cada
   predicción (SHAP) y publicar el modelo como API conectada al dashboard.

## Cómo reproducir el proyecto

### Opción A · En GitHub, sin instalar nada (recomendada)

1. Pestaña **Actions** → **Ejecutar proyecto y generar presentación** → **Run workflow**.
2. Al terminar (≈ 10–15 min), el repositorio tendrá los notebooks con salidas, los gráficos, los mapas, las
   capturas del dashboard, la presentación y el PDF `presentation/Data Science Capstone Project Report.pdf`.

### Opción B · En tu ordenador

```bash
git clone https://github.com/alessandroxlc/spacex-falcon9-capstone.git
cd spacex-falcon9-capstone
python -m venv .venv
# Windows: .venv\Scripts\activate    ·    macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
python scripts/run_all.py
```

Para abrir el dashboard: `python dashboard/spacex_dash_app.py` y visita <http://127.0.0.1:8050>.

> Los datos proceden de la API pública de SpaceX, de Wikipedia y de los ficheros públicos del curso de IBM
> Skills Network. Si la API de SpaceX no responde, el notebook 01 usa el `dataset_part_1.csv` oficial del
> curso (resultado del mismo proceso) y lo indica en su salida.

## Correspondencia con la rúbrica

| Criterio | Dónde está |
|---|---|
| URL de GitHub | Portada y cada diapositiva de metodología |
| Resumen ejecutivo e introducción | Diapositivas 3 y 4 |
| Recopilación de datos (API y scraping) | Diagramas de flujo + URL de los notebooks 01 y 02 |
| Limpieza de datos | Diapositiva de data wrangling + URL del notebook 03 |
| EDA con visualización y con SQL | Metodología + 6 diapositivas de gráficos + 10 de consultas |
| Análisis visual interactivo | Folium (3 mapas) y Plotly Dash (4 capturas) + URLs |
| Análisis predictivo | Exactitud de los 4 modelos, mejores hiperparámetros y matriz de confusión |
| Conclusión e ideas innovadoras | Diapositivas finales |

---

Autor: [alessandroxlc](https://github.com/alessandroxlc) · Licencia [MIT](LICENSE)
