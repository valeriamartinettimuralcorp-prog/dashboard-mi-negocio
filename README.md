# Dashboard Mi Negocio - Mural Corp

Dashboard Streamlit para analizar el reporte diario de Mi Negocio.

## Incluye
- KPIs: universo, registrados, activos, Full Digital y % de venta digital.
- Filtro por vendedor y buscador de cliente.
- Embudo Universo → Registrados → Activos → Full Digital.
- Participación de venta digital.
- Listados de clientes sin registrar, registrados no activos y activos no Full Digital.
- Ranking por vendedor.
- Exportación de oportunidades a Excel.

## Instalación
```bash
pip install -r requirements.txt
```

## Ejecución
```bash
streamlit run app.py
```

Después, cargar desde la barra lateral el Excel diario exportado con la misma estructura de columnas que `data (4).xlsx`.

## Actualización diaria
No es necesario modificar el código. Solo se carga el nuevo Excel y todos los indicadores se recalculan.
