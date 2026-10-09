# 💧 Aplicativo Registro de Verificación PTAR - PROLAN

Aplicación web sencilla para ingresar los datos del **Registro de Verificación PTAR** y generar automáticamente los archivos en **Excel** (mismo diseño oficial) y **PDF**.

## Requisitos

- Python 3.9 o superior
- Conexión a internet solo para la instalación inicial de paquetes

## Instalación (solo una vez)

Abra una terminal en la carpeta donde están los archivos y ejecute:

```bash
pip install -r requirements.txt
```

## Cómo usar

1. Asegúrese de que el archivo `Registro_de_verificacion_PETAR.xlsx` (el formato oficial vacío) esté en la misma carpeta que `app_registro_ptar.py`.

2. Ejecute la aplicación:

```bash
streamlit run app_registro_ptar.py
```

3. Se abrirá automáticamente en su navegador (normalmente en http://localhost:8501).

## Funcionalidades

- **Tabla editable** de hasta 20 filas con todas las columnas del formato oficial.
- Campos para **Elaborado por** y **V°B° del Responsable de Saneamiento**.
- Botón **Cargar datos de ejemplo** para probar rápidamente.
- Botón **Generar Excel** → genera el archivo con el **mismo diseño y formato** del template oficial (colores, encabezados, celdas combinadas, etc.).
- Botón **Generar PDF** → genera un PDF profesional en orientación horizontal con la misma información.
- Botones de **descarga** directa de los archivos generados.
- Botón **Limpiar tabla**.

## Archivos incluidos

| Archivo | Descripción |
|---------|-------------|
| `app_registro_ptar.py` | Aplicación principal (Streamlit) |
| `requirements.txt` | Dependencias necesarias |
| `Registro_de_verificacion_PETAR.xlsx` | Plantilla oficial del formato |
| `README_APLICATIVO_PTAR.md` | Este archivo de instrucciones |

## Notas

- El Excel generado mantiene **exactamente** el diseño del formato PI-R.CE-0031.
- El PDF es una versión limpia y profesional del mismo contenido (orientación horizontal A4).
- Puede modificar el número de filas visibles desde el panel lateral.
- Los datos se mantienen mientras la aplicación esté abierta.

---
Desarrollado para PROLAN · Formato PI-R.CE-0031
