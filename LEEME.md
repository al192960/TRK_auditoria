# Auditoría de Carga (app Android en Python / Kivy) — v2.0

## Archivos
| Archivo | Qué hace |
|---|---|
| `main.py` | La app: camión, captura por posición, incidencias, menú oculto y envío por correo |
| `reporte.py` | Llena la plantilla original y genera el reporte de incidencias |
| `plantilla.xlsm` | Tu "Borrador de carga" original. **No lo modifiques con otro nombre**; si cambias el formato, reemplaza este archivo |
| `p4a_hook.py` | Permite mandar el Excel por correo sin guardarlo en el dispositivo |
| `buildozer.spec` / `compilar_apk.sh` | Para generar el APK en Ubuntu/WSL (`bash compilar_apk.sh`) |

## Uso diario
1. Escribe **TRK**, **AUDITOR** y elige **Turno** (todo se pasa a MAYÚSCULAS).
2. Desliza el camión y toca una posición. Teclea el **ÁREA** y escanea **No. de parte → Enter → Cantidad → Enter**.
   Los números de parte se guardan tal como los manda el escáner (no se pasan a mayúsculas).
3. Más de 6 números en una posición → **+ Agregar número** pide el nombre de quien cargó y registra la incidencia.
4. **Generar Excel y enviar** → se abre el correo con el archivo. El archivo vive solo en la caché privada de la app y se borra al regresar.
5. La app pregunta si ya se envió; si dices que sí, limpia la troca para la siguiente.

## Excel de carga
- Es **tu archivo original** (`plantilla.xlsm`) con las hojas HCARGA, COMPARAR, INCIDENNOCIAS y GRAFICA intactas: fórmulas, tablas, tablas dinámicas, gráficas y macros.
- La app solo escribe en las celdas de captura de **HCARGA**: Fecha (D5), TRK (H5), nombre (J5), turno (T5), áreas, No. Parte y Cant. de las posiciones 1–26.
- **COMPARAR** se recalcula sola al abrir el archivo en Excel. Solo falta pegar los datos de SAP en `# DE PARTE` / `CANTIDAD` (Table2), como ya lo haces.
- Si una posición tiene **más de 6 números**, los extra se escriben en los bloques **27 y 28** de HCARGA (que el camión no usa), con el área "EXTRA P5", etc. Así COMPARAR también los cuenta. Caben 12 extra por troca.

## Menú oculto (supervisor)
Toca **5 veces seguidas** el título **"AUDITORÍA DE CARGA"**.
- **Reporte de incidencias**: genera un Excel aparte con todas las incidencias pendientes (de todas las trocas) y lo manda por correo. Al confirmar el envío se borran de la lista.
- **Recuperar última troca enviada** *(pide contraseña)*: vuelve a cargar la última troca que se mandó.
- **Modo dos trocas** *(pide contraseña para encender)*: aparecen pestañas TROCA 1 / TROCA 2 para capturar dos al mismo tiempo.
- **Cambiar contraseña**.

Contraseña de fábrica: **2468** (cámbiala en el menú, o en `CONTRASENA_INICIAL` dentro de `main.py`).

## Escáner
Configura el escáner para que mande **Enter** después de cada lectura (Zebra: DataWedge → Keystroke output → Send ENTER).

## Probar en la PC
```
pip install kivy openpyxl
python main.py --size=480x854
```
En la PC no se abre el correo; solo te avisa dónde quedó el archivo temporal.
