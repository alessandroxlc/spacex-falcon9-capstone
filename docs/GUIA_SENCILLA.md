# Guía sencilla del proyecto (explicado sin tecnicismos)

## 1. ¿De qué trata el proyecto?

SpaceX fabrica cohetes. Su cohete principal, el **Falcon 9**, tiene dos partes: la de arriba lleva el
satélite al espacio y la de abajo (la **primera etapa** o *booster*) lo empuja durante los primeros minutos.

Lo especial de SpaceX es que esa primera etapa **vuelve y aterriza**, ya sea en tierra o en una barcaza en el
mar, y se reutiliza. Por eso SpaceX cobra unos 62 millones de dólares por lanzamiento y otras empresas, que
tiran su cohete, más de 165 millones.

**La pregunta del proyecto:** con los datos de un lanzamiento (qué cohete, cuánto pesa la carga, a qué órbita
va, desde dónde despega…), ¿podemos **predecir si la primera etapa aterrizará**?

Si sabemos eso, sabemos aproximadamente cuánto costará el lanzamiento. Imagina que trabajas en una empresa
rival (el curso la llama "SpaceY") y quieres saber cuánto ofertar para ganarle un contrato a SpaceX.

## 2. ¿Qué se hace, paso a paso?

| Paso | En palabras sencillas | Archivo |
|---|---|---|
| 1. Recolectar datos (API) | Pedirle a la "base de datos pública" de SpaceX la lista de sus lanzamientos | `notebooks/01_...` |
| 2. Recolectar datos (web) | Copiar automáticamente la tabla de lanzamientos de Wikipedia | `notebooks/02_...` |
| 3. Limpiar datos | Rellenar huecos y crear la columna "¿aterrizó? 1 = sí, 0 = no" | `notebooks/03_...` |
| 4. Preguntas con SQL | Hacer 10 preguntas a los datos en lenguaje de bases de datos | `notebooks/04_...` |
| 5. Gráficos | Dibujar los datos para ver patrones (¿mejora con los años?, ¿influye la órbita?) | `notebooks/05_...` |
| 6. Mapas | Poner los sitios de lanzamiento en un mapa y medir distancias a la costa, etc. | `notebooks/06_...` |
| 7. Dashboard | Una página web interactiva para explorar los datos con filtros | `dashboard/spacex_dash_app.py` |
| 8. Machine learning | Enseñar a 4 modelos a predecir el aterrizaje y ver cuál acierta más | `notebooks/07_...` |
| 9. Presentación | Resumir todo en diapositivas (se genera sola) | `presentation/generate_presentation.py` |

## 3. ¿Qué es "83,33 % de exactitud"?

Se apartan 18 lanzamientos que los modelos **no ven** durante el aprendizaje (el "examen"). Después se les
pide que adivinen si esos 18 aterrizaron. Acertar 15 de 18 es 83,33 %.

La **matriz de confusión** es la tabla de aciertos y errores de ese examen:

* **Verdadero positivo:** dijo "aterriza" y aterrizó. ✅
* **Verdadero negativo:** dijo "no aterriza" y no aterrizó. ✅
* **Falso positivo:** dijo "aterriza" y no aterrizó. ❌ (el error más típico aquí)
* **Falso negativo:** dijo "no aterriza" y sí aterrizó. ❌

## 4. ¿Qué se entrega exactamente? (Opción 1: calificación por IA "Mark")

Se sube **un PDF** llamado **`Data Science Capstone Project Report`** (la presentación convertida a PDF).
Dentro del PDF deben aparecer:

1. La **URL de GitHub** del proyecto (está en la portada y en cada diapositiva de metodología).
2. Las diapositivas que pide la rúbrica (resumen ejecutivo, introducción, metodología, resultados de EDA,
   SQL, Folium, Dash, modelos, conclusión e ideas innovadoras). Cada diapositiva lleva arriba, en naranja,
   el nombre de la sección de la plantilla oficial de IBM para que el evaluador la identifique.

## 5. Pasos para entregar

1. **Ejecuta el proyecto en GitHub** (no hace falta instalar nada):
   pestaña **Actions** → **Ejecutar proyecto y generar presentación** → **Run workflow**. Espera a que
   salga el check verde (≈ 10–15 min).
   *Esto rellena los notebooks con sus resultados y genera todos los gráficos, mapas, capturas, la
   presentación y el PDF con los datos reales.*
2. Revisa que el repositorio sea **público** (Settings → General → Danger Zone → Change visibility).
3. Descarga `presentation/Data Science Capstone Project Report.pdf` (o el artefacto "presentacion-capstone"
   de la ejecución en Actions).
4. Ábrelo y revisa que no quede ningún recuadro "Figura pendiente". Si quieres cambiar tu nombre, edita
   `AUTHOR` al principio de `presentation/generate_presentation.py` y vuelve a ejecutar el paso 1.
5. En Coursera, elige la **Opción 1 (AI-Graded)** y sube el PDF.

## 6. Checklist de la rúbrica (15 puntos)

| # | Criterio | ¿Dónde se cumple? |
|---|---|---|
| 1.1 | URL de GitHub | Portada + diapositivas de metodología |
| 1.2 | Presentación en PDF | `Data Science Capstone Project Report.pdf` |
| 1.3 | Resumen ejecutivo | Diapositiva "Resumen ejecutivo" (métodos + resultados) |
| 1.4 | Introducción | Contexto (62 M$ vs. 165 M$) y preguntas a responder |
| 1.5 | Recopilación — API | Diagrama de flujo de 6 pasos + URL del notebook 01 |
| 1.6 | Recopilación — scraping | Diagrama de flujo de 6 pasos + URL del notebook 02 |
| 1.7 | Limpieza de datos | Pasos + tabla Outcome → Class + URL del notebook 03 |
| 1.8 | EDA con visualización | Tabla de gráficos y su propósito + URL del notebook 05 |
| 1.9 | EDA con SQL | Las 10 consultas resumidas + URL del notebook 04 |
| 1.10 | Análisis visual interactivo | Diapositivas Folium y Dash con ambas URLs |
| 1.11 | Resultados EDA visual | 4 dispersiones, barras por órbita y tendencia anual |
| 1.12 | Resultados SQL | 10 diapositivas: consulta + resultado + conclusión |
| 1.13 | Mapas Folium | Sitios, resultados (verde/rojo) y distancias |
| 1.14 | Plotly Dash | 2 gráficos de pastel + 2 dispersiones (rango completo y filtrado) |
| 1.15 | Análisis predictivo + conclusión | Exactitud de 4 modelos, hiperparámetros, matriz de confusión, mejor modelo, conclusiones e ideas innovadoras |
