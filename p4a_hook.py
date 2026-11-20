# -*- coding: utf-8 -*-
"""
Hook de python-for-android: agrega un FileProvider al AndroidManifest para
poder mandar el Excel por correo desde la caché privada de la app
(así el archivo no se guarda en Descargas ni en ninguna carpeta visible).
"""
import os

PROVIDER = """
        <provider android:name="androidx.core.content.FileProvider"
                  android:authorities="${applicationId}.fileprovider"
                  android:exported="false"
                  android:grantUriPermissions="true">
            <meta-data android:name="android.support.FILE_PROVIDER_PATHS"
                       android:resource="@xml/file_paths" />
        </provider>
"""

FILE_PATHS = """<?xml version="1.0" encoding="utf-8"?>
<paths>
    <cache-path name="reportes" path="reportes/" />
</paths>
"""


def before_apk_assemble(toolchain):
    for manifest in ("src/main/AndroidManifest.xml", "AndroidManifest.xml"):
        if not os.path.exists(manifest):
            continue
        with open(manifest, encoding="utf-8") as fh:
            xml = fh.read()
        if "androidx.core.content.FileProvider" not in xml:
            xml = xml.replace("</application>", PROVIDER + "    </application>", 1)
            with open(manifest, "w", encoding="utf-8") as fh:
                fh.write(xml)
    os.makedirs("src/main/res/xml", exist_ok=True)
    with open("src/main/res/xml/file_paths.xml", "w", encoding="utf-8") as fh:
        fh.write(FILE_PATHS)
    print("[p4a_hook] FileProvider agregado")
