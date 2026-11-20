#!/usr/bin/env bash
# ============================================================
#  Compila el APK de "Auditoría de Carga" en Ubuntu / WSL
#  Uso:  bash compilar_apk.sh  [carpeta_del_proyecto]
#  (si no das carpeta, usa la carpeta actual)
# ============================================================
set -e

# Quitar las rutas de Windows (/mnt/c/...) del PATH: en WSL hacen que la
# compilación tome Python/Java/cmake de Windows o MSYS2 y falle.
export PATH="$(echo "$PATH" | tr ':' '\n' | grep -v '^/mnt/' | paste -sd:)"

ORIGEN="${1:-$(pwd)}"
ORIGEN="$(cd "$ORIGEN" && pwd)"

if [ ! -f "$ORIGEN/main.py" ] || [ ! -f "$ORIGEN/buildozer.spec" ] || [ ! -f "$ORIGEN/plantilla.xlsm" ]; then
  echo "ERROR: en $ORIGEN faltan main.py, buildozer.spec o plantilla.xlsm"
  exit 1
fi

# --- 1. Si el proyecto está en Windows (/mnt/c), copiarlo a Linux ---
if [[ "$ORIGEN" == /mnt/* ]]; then
  PROY="$HOME/apps/auditoria_carga"
  echo ">> Copiando proyecto a $PROY (compilar en /mnt/c es lento y falla)"
  mkdir -p "$PROY"
  cp "$ORIGEN"/*.py "$ORIGEN"/*.xlsm "$ORIGEN"/buildozer.spec "$PROY"/
  # limpiar compilaciones fallidas anteriores
  rm -rf "$PROY/.buildozer/android/platform/build-"* "$PROY/bin"
else
  PROY="$ORIGEN"
fi

# Si cambió la versión de python-for-android en buildozer.spec, borrar la anterior
P4A_RAMA="$(grep -E '^p4a.branch' "$PROY/buildozer.spec" | cut -d= -f2 | tr -d ' ')"
if [ -n "$P4A_RAMA" ] && [ "$(cat "$PROY/.p4a_rama" 2>/dev/null)" != "$P4A_RAMA" ]; then
  echo ">> Cambió python-for-android a $P4A_RAMA: limpiando compilación anterior"
  rm -rf "$PROY/.buildozer" "$PROY/bin"
  echo "$P4A_RAMA" > "$PROY/.p4a_rama"
fi

# --- 2. Dependencias del sistema ---
echo ">> Instalando dependencias del sistema (pide tu contraseña)"
sudo apt update
sudo apt install -y git zip unzip openjdk-17-jdk python3-pip python3-venv \
  autoconf automake libtool pkg-config zlib1g-dev libncurses-dev \
  libffi-dev libssl-dev build-essential ccache cmake adb

# --- 3. Entorno de Python con Buildozer ---
if [ ! -d "$HOME/buildenv" ]; then
  echo ">> Creando entorno ~/buildenv"
  python3 -m venv "$HOME/buildenv"
fi
source "$HOME/buildenv/bin/activate"
pip install --upgrade pip
pip install buildozer "cython==0.29.37" setuptools

# --- 4. Compilar ---
cd "$PROY"
echo ">> Compilando APK (la primera vez tarda 30-60 min y necesita ~10 GB)"
set +e
yes | buildozer -v android debug 2>&1 | tee "$PROY/build.log"
RES=${PIPESTATUS[1]}
set -e
if [ "$RES" -ne 0 ] || ! ls bin/*.apk >/dev/null 2>&1; then
  echo ""
  echo "============== FALLÓ LA COMPILACIÓN =============="
  echo " Líneas con error (mándale esto a Claude):"
  grep -n -iE "error|failed|exception|not found|no such" "$PROY/build.log" | grep -viE "warning|-Werror" | tail -40
  echo " Log completo: $PROY/build.log"
  exit 1
fi

APK="$(ls -t bin/*.apk | head -1)"
echo ""
echo "============================================================"
echo " APK LISTO: $PROY/$APK"
echo "============================================================"

# --- 5. Si es WSL, copiar el APK al Escritorio de Windows ---
if grep -qi microsoft /proc/version 2>/dev/null; then
  WINUSER="$(cmd.exe /c 'echo %USERNAME%' 2>/dev/null | tr -d '\r')"
  DESK="/mnt/c/Users/$WINUSER/Desktop"
  if [ -d "$DESK" ]; then
    cp "$APK" "$DESK/"
    echo " Copia en el Escritorio de Windows: $DESK/$(basename "$APK")"
  fi
fi

# --- 6. Si hay un escáner conectado por USB con depuración, instalarlo ---
if adb devices | grep -qw device; then
  echo ">> Escáner detectado, instalando..."
  adb install -r "$APK" && echo " Instalado en el escáner."
else
  echo " (Para instalar directo: conecta el escáner con Depuración USB y corre: adb install -r $PROY/$APK)"
fi
