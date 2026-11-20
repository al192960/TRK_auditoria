# Auditoría de Carga (app Android en Python / Kivy)

## Archivos
| Archivo | Qué hace |
|---|---|
| `main.py` | La app: pantalla del camión, captura por posición, incidencias y envío por correo |
| `reporte.py` | Arma el Excel (mismo acomodo que tu "Borrador de carga") |
| `buildozer.spec` | Configuración para generar el APK |

## Cómo se usa
1. En la pantalla principal escribe **TRK**, **Auditor** y elige **Turno** (auditor y turno se recuerdan).
2. Toca una posición del camión (1–26).
3. Teclea el **Área** → Enter → escanea **No. de parte** → Enter → escanea/teclea **Cantidad** → Enter → pasa solo al siguiente renglón.
   - Si la etiqueta de cantidad viene como `Q100` o `100 PCS`, la app se queda con `100`.
4. Hay 6 renglones. Si tocas **+ Agregar número** cuando ya hay 6, la app pide el **nombre de quien cargó**,
   la categoría (por defecto *Mas de 6 numeros por pallet*) y un comentario. Eso queda como incidencia y se habilitan más renglones (en naranja).
5. **Guardar** regresa al camión. Colores: gris = vacía, verde = capturada (con número de líneas), naranja = tiene incidencia.
6. El botón **Incidencia** (en el camión o dentro de una posición) registra cualquier otra incidencia con las mismas categorías de tu libro.
7. **Generar Excel y enviar** → escribe el/los correos (se recuerdan) → se abre Gmail/Outlook con el Excel adjunto.
   Además queda una copia en `Descargas/AuditoriaCarga`.
8. **Nueva** borra la captura para el siguiente camión. Todo se guarda automáticamente, así que no se pierde nada si se cierra la app.

## El Excel que genera
- **HCARGA**: mismas celdas que tu formato (posición 1 en D12, No. Parte en F12:F17, Cant. en G12:G17, etc.), para que puedas copiar el rango D12:S53 directo a tu libro y funcione la hoja COMPARAR.
- **DETALLE**: todas las líneas, incluyendo las que pasan de 6.
- **RESUMEN**: total por número de parte (para comparar contra SAP).
- **INCIDENCIAS**: columnas Nombre, Tipo, Categoria, Fecha, Mes (en el mismo orden que tu Table3) + Posición, TRK y Auditor.

## Escáner
La app funciona con escáneres que escriben como teclado (Zebra, Honeywell, escáneres Bluetooth HID).
Configura el escáner para que mande **Enter** al final de cada lectura (en Zebra: DataWedge → Keystroke output → "Send ENTER").

## Probar en la computadora
```bash
pip install kivy openpyxl
python main.py
```

## Generar el APK (en Linux o WSL de Windows)
```bash
pip install buildozer cython
sudo apt install -y git zip unzip openjdk-17-jdk autoconf libtool pkg-config zlib1g-dev libncurses-dev libffi-dev libssl-dev
cd auditoria_carga
buildozer -v android debug
```
La primera vez tarda (descarga el SDK/NDK). El APK queda en `bin/`. Cópialo al celular e instálalo
(permite "instalar apps de origen desconocido").
