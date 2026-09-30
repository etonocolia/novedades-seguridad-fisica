# Dashboard de paradas no autorizadas

Aplicación local en Streamlit para analizar las novedades de `Informes.xlsx`, con foco en `ParadasNoAutorizadas` y contexto de las hojas de conductores, precintos y unidades.

## Ejecutar

Desde PowerShell, en esta carpeta:

```powershell
\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

El libro `Informes.xlsx` debe estar junto a `app.py`. También se puede cargar otro `.xlsx` desde el panel lateral.

## Lectura de indicadores

- Las filas exactamente duplicadas se excluyen de los conteos y se reportan en calidad.
- La recurrencia se evalúa con un umbral configurable por placa y dentro del periodo filtrado.
- Una coincidencia con otra hoja requiere misma fecha y placa. Es una pista de contexto, no una relación causal.
- La prioridad alta solo indica que coinciden recurrencia y contexto; no significa pérdida confirmada.
- El mapa de Operación muestra hasta los 15 sitios con más paradas según los filtros activos sobre calles OpenStreetMap y usa la coordenada válida más frecuente de cada sitio. Requiere conexión a internet para cargar el mapa base. Los sitios sin coordenadas válidas se omiten; esas coordenadas se cuentan en calidad.

## Limitaciones del libro

No hay cantidades de carga recibida/despachada ni un total confiable de viajes, por lo que no se calcula pérdida física ni tasa por viaje. `REGISTRO DE TIEMPO` es una marca horaria; no permite medir cuánto duró una parada. La hoja de conductores tiene pocos registros y las coincidencias entre hojas son limitadas.