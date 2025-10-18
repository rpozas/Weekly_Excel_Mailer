# Weekly Excel Mailer

Automatiza la generación y el envío por correo de un reporte Excel semanal a partir de un archivo fuente. Permite personalizar la lógica de actualización del Excel y programar la ejecución con el Programador de tareas de Windows.

## Qué hace

- Lee un archivo Excel de entrada (`files.input_xlsx`).
- Ejecuta una función personalizable (`custom_update`) para preparar el reporte.
- Guarda el resultado en `files.output_dir` con un nombre basado en la fecha (`weekly_report_{date}.xlsx`).
- Envía el archivo generado por correo usando SMTP (TLS o SSL), con destinatarios, asunto y cuerpo definidos en `config.yaml`.
- Puede ejecutarse manualmente o de forma automática cada semana.

## Flujo de ejecución

1. Carga configuración desde `config.yaml`.
2. Genera el nombre de salida con la fecha actual (formato `YYYYMMDD`).
3. Llama a `custom_update(input_xlsx, output_xlsx)` para producir el Excel final.
4. Si el envío está habilitado, adjunta el Excel y lo envía por SMTP a `to`/`cc`/`bcc`.
5. Registra mensajes en consola (y en `run.log` si usas el `.bat`).

## Requisitos

- Python 3.9 o superior (recomendado 3.10+).
- Paquetes Python:
  - `pandas`, `openpyxl`, `pyyaml` (ver `requirements.txt`).
- Windows (opcional) para programar con el Programador de tareas.

## Instalación

1. Clona o descarga este repositorio.
2. (Opcional pero recomendado) Crea y activa un entorno virtual:
   - Windows PowerShell:
     ```powershell
     python -m venv .venv
     .\.venv\Scripts\Activate.ps1
     ```
3. Instala dependencias:
   ```bash
   pip install -r requirements.txt
   ```

## Configuración

Edita `config.yaml` (puedes tomar como referencia `config.example.yaml`).

```yaml
files:
  input_xlsx: "C:/ruta/a/tu/source.xlsx"  # Excel fuente
  output_dir:  "C:/ruta/a/salida"          # Carpeta donde escribir
  output_name_pattern: "weekly_report_{date}.xlsx"  # {date} -> YYYYMMDD

email:
  enabled: true
  smtp:
    host: "smtp.gmail.com"
    port: 587
    use_tls: true         # true=STARTTLS, false=SSL directo (465 usualmente)
    username: "tu@correo.com"
    password: "tu-app-password"
  from: "tu@correo.com"
  to: ["destino1@empresa.com", "destino2@empresa.com"]
  cc: []
  bcc: []
  subject: "Weekly XLSX Report"
  body: |
    Hola,
    Te envío el reporte semanal en adjunto.
```

Notas:
- Para Gmail con 2FA, usa un “App password”.
- Si tu servidor usa SSL directo, pon `use_tls: false` y ajusta `port` (465 es común).
- Las rutas en Windows deben usar `/` o `\\`.

## Ejecución manual

- Directo con Python:
  ```bash
  python auto_report.py
  ```

- Con el script por lotes (guarda salida en `run.log`):
  ```cmd
  run_report.bat
  ```

El archivo generado se guarda en la carpeta definida por `files.output_dir`. Si el envío está habilitado, se envía por correo automáticamente.

## Programación semanal (Windows)

Hay un archivo de ejemplo para el Programador de tareas: `TaskScheduler-Weekly-Thursday.xml`.

Pasos sugeridos:

1. Edita el XML y actualiza la ruta del comando para que apunte a tu `run_report.bat` (línea `<Command>`). Ejemplo:
   ```xml
   <Command>D:\TRABAJO\PyFoo101\Weekly_Excel_Mailer\run_report.bat</Command>
   ```
2. Ajusta el `StartBoundary` a la fecha y hora deseadas y verifica el día de la semana.
3. Importa el XML en el Programador de tareas (Acción → Importar tarea...).
4. Configura el usuario que ejecutará la tarea y marca “Ejecutar tanto si el usuario inició sesión como si no”.
5. Prueba “Ejecutar” manualmente y verifica `run.log` y el correo recibido.

## Personalizar la lógica de actualización

Edita la función `custom_update` en `auto_report.py` para aplicar tus reglas de negocio. Comportamiento actual por defecto:

- Lee la primera hoja del Excel de entrada (si existe).
- Escribe esa hoja sin cambios en el archivo de salida.
- Agrega una pestaña `Summary` con:
  - `Last_Updated` (fecha/hora de generación)
  - `Source_File` (nombre del archivo de entrada)
  - `Rows_In_First_Sheet` (conteo de filas)

Puedes reemplazar esa lógica por filtros, agregaciones, uniones a otras fuentes, etc. Punto de entrada: `custom_update(input_xlsx, output_xlsx)`.

## Estructura del proyecto

- `auto_report.py`: script principal; carga config, genera Excel y envía email.
- `config.yaml`: configuración de rutas, nombres y SMTP.
- `config.example.yaml`: ejemplo de configuración.
- `run_report.bat`: ejecuta el script y redirige logs a `run.log`.
- `TaskScheduler-Weekly-Thursday.xml`: plantilla para programar la tarea semanal.
- `requirements.txt`: dependencias Python.

## Solución de problemas

- Error de autenticación SMTP: verifica `host/port`, `use_tls`, usuario/contraseña (usa App Password si aplica).
- Ruta no válida: confirma `files.input_xlsx` y `files.output_dir` existen (la carpeta de salida se crea si falta).
- Adjuntos no llegan: revisa filtros/antivirus del servidor, tamaño del archivo, y destinatarios `to/cc/bcc`.
- Problemas con YAML: respeta la indentación y comillas. Puedes validar el archivo con herramientas YAML.
- Versiones de librerías: si hay conflictos, reinstala según `requirements.txt`.

## License
This project is licensed under the [MIT License](./LICENSE).
