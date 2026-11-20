[app]
title = Auditoria de Carga
package.name = auditoriacarga
package.domain = com.planta
source.dir = .
source.include_exts = py,png,jpg,kv,json,xlsm
source.exclude_patterns = p4a_hook.py,compilar_apk.sh,LEEME.md,build.log
version = 2.0

requirements = python3,kivy,openpyxl,et_xmlfile,pyjnius,android

# Se adapta a vertical y horizontal (el camión se gira solo)
orientation = portrait, landscape
fullscreen = 0

# Ya no se necesitan permisos: el Excel se manda desde la caché privada de la app
# android.permissions =  (no se necesitan)

android.api = 34
android.minapi = 24
android.archs = arm64-v8a
android.accept_sdk_license = True
android.ndk = 25b

# Para mandar archivos por correo sin guardarlos (FileProvider)
android.enable_androidx = True
android.gradle_dependencies = androidx.core:core:1.6.0

# Versión ESTABLE de python-for-android (la de desarrollo falla con Python 3.14)
p4a.branch = v2024.01.21
p4a.hook = p4a_hook.py

[buildozer]
log_level = 2
warn_on_root = 1
