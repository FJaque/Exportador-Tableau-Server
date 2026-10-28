# Exportador Tableau Server / Tableau Cloud

Descarga a una carpeta local todo el contenido de uno o más sitios de Tableau y genera un inventario completo en Excel.

## 🚀 Qué hace
El proceso corre en dos fases:
1. **Inventario**: lista todo el contenido y guarda `inventario_<fecha>.xlsx` *antes* de descargar nada. Aunque la descarga se corte, el inventario ya queda disponible.
2. **Descarga**: baja los archivos y vuelve a guardar el inventario con el estado de descarga de cada elemento.

- Descarga **workbooks** (`.twbx`), **fuentes de datos publicadas** (`.tdsx`) y **flujos de Prep** (`.tflx`), respetando la estructura de proyectos.
- Genera `inventario_<fecha>.xlsx` (con la paleta de colores Backspace) con las hojas:
  | Hoja | Contenido |
  |---|---|
  | Resumen | Servidor, fecha, versión API y totales por sitio |
  | Proyectos | Ruta completa (Padre/Hijo), descripción, modo de permisos |
  | Workbooks | Propietario, fechas, tamaño, etiquetas, URL, **estado de descarga** y archivo local |
  | Vistas | Workbook, tipo de hoja, total de visitas, URL |
  | Fuentes de datos | Tipo, extracto, certificación, URL, **estado de descarga** |
  | Conexiones | Servidor, puerto, usuario BD y si la contraseña está embebida (workbooks y fuentes) |
  | Flujos | Propietario, fechas, URL, **estado de descarga** |
  | Campos | Campos usados por hoja, tipo de campo y tipo de dato (**mismo formato del proyecto Copec**, vía Metadata API) |
  | Usuarios / Grupos / Usuarios por grupo | Roles, autenticación, último login, membresías |
  | Tareas extraccion / Suscripciones | Programaciones de actualización y suscripciones |
  | Pendientes | Lo que falta descargar o revisar a mano |
- Genera `pendientes_<fecha>.xlsx` con un resumen por tipo/acción y el detalle de cada elemento que no se pudo descargar: **URL web** (clic directo) y **acción sugerida** para hacerlo manualmente.
- Deja un log completo en `logs/ejecucion_<fecha>.log`.

## ⚙️ Requisitos
- Python 3.8+
- `pip install -r requirements.txt`
- Un **Personal Access Token (PAT)** de un usuario **Administrador de sitio** (o de servidor si se exportan todos los sitios), para que vea y pueda descargar todo el contenido.
- Para la hoja *Campos* en Tableau Server: Metadata API habilitada (`tsm maintenance metadata-services enable`). En Tableau Cloud ya viene habilitada.

## 📖 Uso
1. Editar `config.ini`: `server_url`, `sites` y opciones de exportación.
   Se recomienda dejar `token_name` / `token_value` vacíos: el programa los pide al ejecutar y no quedan guardados.
2. Ejecutar:
   ```
   python exportador.py
   ```
   o doble clic en `ejecutar.bat`.
   Opciones:
   ```
   python exportador.py --prueba               # modo prueba: solo 5 elementos por tipo
   python exportador.py --solo-inventario      # solo el Excel, sin descargar archivos
   python exportador.py --config cliente.ini   # usar otra configuración
   ```
3. Revisar `pendientes_<fecha>.xlsx` y descargar a mano lo que falte.

### Modo prueba
Con `--prueba` (o `modo_prueba = true` en `config.ini`) se descargan solo `cantidad_prueba` elementos por tipo (5 por defecto), priorizando los más livianos y fuera de espacios personales. El inventario lista todo el contenido, pero las conexiones y los campos solo se leen para esos elementos, así la prueba toma pocos minutos. Sirve para validar acceso, permisos y la estructura de carpetas en una reunión antes de lanzar la exportación completa.

### Si el Excel está abierto
Si `inventario_*.xlsx` o `pendientes_*.xlsx` están abiertos en Excel, el programa no falla: guarda una copia con la hora agregada al nombre.

### Re-ejecutar
Con `sobrescribir = false` (por defecto), los archivos que ya están en la carpeta se saltan, así que si la ejecución se corta o hubo errores de red basta con volver a ejecutar: solo se descarga lo que falta. Las descargas se escriben primero como `*__descargando` y se renombran al terminar, por lo que nunca queda un archivo a medias con el nombre final.

## 📂 Estructura de salida
```
exportacion/
├── <sitio>/
│   ├── Workbooks/<Proyecto>/<Subproyecto>/<workbook>.twbx
│   ├── Workbooks/Espacio personal/<propietario>/<workbook>.twbx
│   ├── Fuentes de datos/<Proyecto>/<fuente>.tdsx
│   └── Flujos/<Proyecto>/<flujo>.tflx
├── inventario_<fecha>.xlsx
├── pendientes_<fecha>.xlsx
└── logs/ejecucion_<fecha>.log
```

## ⚠️ No se exporta (revisar manualmente si aplica)
- Permisos de contenido, vistas personalizadas, alertas, métricas/Pulse, conexiones virtuales y colecciones.
- Contraseñas embebidas de conexiones (Tableau nunca las entrega; la hoja *Conexiones* indica cuáles las tienen para volver a ingresarlas al republicar).
- Workbooks del "Espacio personal" de otros usuarios en Tableau Cloud solo aparecen si el token tiene acceso.
