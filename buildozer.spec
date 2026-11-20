[app]
title = Auditoria de Carga
package.name = auditoriacarga
package.domain = com.planta
source.dir = .
source.include_exts = py,png,jpg,kv,json
version = 1.0

# Librerías que se empaquetan dentro del APK
requirements = python3,kivy,openpyxl,et_xmlfile,pyjnius,android

# Se adapta a vertical y horizontal (el camión se gira solo)
orientation = portrait, landscape
fullscreen = 0

# Solo se usa en Android 9 o menor (en 10+ se guarda en Descargas sin permiso)
android.permissions = WRITE_EXTERNAL_STORAGE

android.api = 34
android.minapi = 24
android.archs = arm64-v8a
android.accept_sdk_license = True
android.ndk = 25b

# Versión ESTABLE de python-for-android (la de desarrollo falla con Python 3.14)
p4a.branch = v2024.01.21

[buildozer]
log_level = 2
warn_on_root = 1
