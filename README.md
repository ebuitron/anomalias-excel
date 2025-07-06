# Detector de Saltos en Series Históricas

Aplicación desarrollada en Python con Streamlit para detectar anomalías como:
- Saltos porcentuales
- Rebotes
- Series desactualizadas
- Huecos de datos

## Instrucciones

1. Sube un archivo Excel con series históricas.
2. Ajusta los umbrales.
3. Visualiza incidencias y descarga el Excel anotado.

---

Y un `requirements.txt` con:

```txt
streamlit
pandas
numpy
openpyxl
matplotlib

## Compilar a EXE

OJO: esto NO funciona

Puedes compilar usando PyInstaller:

```bash
pyinstaller --onefile --noconsole launcher.py
