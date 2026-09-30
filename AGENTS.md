# Instrucciones del proyecto

## Contexto

- La aplicación es local y su punto de entrada es [app.py](app.py). Mantén la interfaz en español y no introduzcas HTML ni CSS.
- Consulta [README.md](README.md) para instalación, uso y límites analíticos antes de cambiar el dashboard.
- Mantén las dependencias declaradas en [requirements.txt](requirements.txt); no añadas bibliotecas si Streamlit, pandas, NumPy, openpyxl o Plotly resuelven el caso.

## Libro de novedades

- La fuente predeterminada es `Informes.xlsx`, junto a `app.py`. Usa las hojas `NovedadesDeConductores`, `NovedadesDePrecinto`, `UnidadesDeTransporte` y `ParadasNoAutorizadas`; ignora hojas de presentación como `DASHBOARD` y `TablasDinamicas`.
- En el libro actual, los encabezados de `ParadasNoAutorizadas` están en la fila 6 de Excel; las otras hojas de datos tienen encabezados en la primera fila. Vuelve a inspeccionar el esquema si cambia el archivo, no asumas que columnas o filas se mantienen.
- Las coincidencias entre hojas se basan en fecha y placa exactas, no en un identificador de viaje. Trátalas únicamente como contexto, nunca como causalidad.
- `REGISTRO DE TIEMPO` es una marca horaria, no una duración. El libro no aporta cantidades de carga ni un total confiable de viajes: no afirmes pérdidas físicas ni calcules tasas por viaje.
- Las alertas son reglas de priorización para revisión, no prueba de robo o desbalance. Conserva visibles los criterios y sus límites.
- Las hojas de conductores contienen nombres e identificaciones personales. No imprimas ni registres esos datos en diagnósticos; muéstralos solo si una necesidad operativa explícita lo requiere.
- Los duplicados exactos se excluyen de los conteos. Las coordenadas no numéricas o fuera de latitud [-90, 90] y longitud [-180, 180] no deben aparecer en el mapa.

## Comandos

Ejecuta desde la raíz del proyecto en PowerShell usando el entorno local:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
.\.venv\Scripts\python.exe -m compileall -q app.py
```

No hay una suite de pruebas propia. Para una comprobación de ejecución con el libro local:

```powershell
.\.venv\Scripts\python.exe -c "from streamlit.testing.v1 import AppTest; app = AppTest.from_file('app.py').run(); assert not app.exception, [error.value for error in app.exception]"
```