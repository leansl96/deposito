Sistema de inventario con SQLite y JSON
Sistema creado para el proyecto final de Talento Tech

Se puede destacar principalmente poder usar un lector de codigo de barras para la lectura del mismo imprimiendolo en el sistema de manera automatica
Se crea Automaticamente Tickets cuando se registra una venta
Se crea BackUps manualmente o automaticamente cuando se cierra el sistema
Se utilizo Pandas para poder exportar al Excel de una manera mas legible
Historial de movimiento
Estadisticas
Caja Diaria
Otros

⚙️ Tecnologías utilizadas
Python 3
SQLite (base de datos local)
FPDF (generación de tickets PDF)
Pandas + OpenPyXL (exportación a Excel)
OS / Shutil / Platform (gestión del sistema y backups)

EJECUCIÓN: En caso de no tener dependencias instalar en el CMD:

pip install pandas openpyxl fpdf

Si ya tenes las dependencias o ya la instalaste:

python main.py
