# Automatizador de reportes

## Descripción del Proyecto

Este script está diseñado para generar reportes sobre mineralogía y elementos químicos en el rubro de la minería. El proceso comienza con el equipo "Ecore", que realiza el escaneo de los metros de sondaje. Este escaneo genera múltiples carpetas, cada una correspondiente a un rango de metros de sondaje. Cada carpeta contiene archivos "Sumary.xlsx" en formato .CSV con la información cruda de los minerales y elementos.

El script se encarga de recopilar toda esta información, trasladarla a un DataFrame de Pandas para facilitar su visualización y generar un informe en el formato esperado por el cliente.

## Funcionalidades

- **Recopilación de Datos:** Recolecta datos de múltiples archivos .CSV generados por el equipo "Ecore".
- **Procesamiento de Datos:** Transforma los datos crudos en un formato más accesible y organizado mediante el uso de Pandas.
- **Generación de Reportes:** Produce informes detallados en Excel, incluyendo gráficos y resúmenes.
- **Categorías de Litología y Alteración:** Clasifica los datos de mineralogía en categorías específicas según reglas predefinidas.
- **Zonas Minerales:** Identifica y clasifica las zonas minerales en "Hipogeno" y "Supergeno" basándose en la presencia de ciertos minerales.

## Estructura del Proyecto

### Archivos y Directorios

- `main.py`: Archivo principal del script.
- `data/`: Directorio donde se almacenan los archivos .CSV de entrada.
- `output/`: Directorio donde se guardan los archivos de reporte generados.
- `README.md`: Archivo de documentación del proyecto.

### Requisitos

- Python 3.x
- Pandas
- Openpyxl
- XlsxWriter
- Tkinter (para la interfaz gráfica)

### Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/MarcoArraiz/cem-dataframe
   ```
2. **Navegar al directorio del proyecto:**
   ```bash
   cd cem-dataframe
   ```
3. **Instalar las dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

### Uso

1. **Ejecutar el script:**
   ```bash
   python minerals-entregable1.py
   python minerals-entregable2.py
   ```
2. **Seleccionar el directorio de entrada:**
   Utiliza la interfaz gráfica para seleccionar el directorio que contiene las carpetas con los archivos .CSV.


3. **Procesamiento de datos:**
   El script procesará los datos y generará un archivo de Excel con los informes en el directorio `output`.

### Ejemplo de Estructura de Carpetas

```
project-root/
├── data/
│   ├── 0.00m - 2.00m/
│   │   └── Summary.xlsx
│   ├── 2.00m - 4.00m/
│   │   └── Summary.xlsx
│   └── ...
├── output/
│   └── Report.xlsx
├── main.py
└── README.md
```

### Contacto

Para cualquier consulta o soporte, puedes contactarme a través de:

- Email: marcoarraiz@gmail.com
- GitHub: [MarcoArraiz](https://github.com/MarcoArraiz)

## Licencia

Este proyecto está licenciado bajo la Licencia Creative Commons Zero v1.0 Universal . Ver el archivo `LICENSE` para más detalles.

---

