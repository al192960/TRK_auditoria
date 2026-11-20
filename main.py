# -*- coding: utf-8 -*-
"""
Auditoría de Carga - App Android (Kivy)

- Pantalla principal = el camión (posiciones grandes, se desliza para ver todas).
- Al tocar una posición: área + hasta 6 números de parte con cantidad (escáner).
  Más de 6 → se pide el nombre del responsable y se registra una incidencia.
- "Generar Excel y enviar" llena la plantilla original (HCARGA/COMPARAR) y la
  manda por correo. El archivo NO se guarda en el dispositivo.
- Menú oculto (tocar 5 veces el título): reporte de incidencias aparte,
  recuperar última troca enviada y modo dos trocas (estas dos con contraseña).
"""
import json
import os
import re
import tempfile
import time
import uuid
from datetime import datetime

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Line, Rectangle, RoundedRectangle
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.relativelayout import RelativeLayout
from kivy.uix.screenmanager import NoTransition, Screen, ScreenManager
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.utils import platform

import reporte

# ---------------------------------------------------------------- ajustes
CONTRASENA_INICIAL = "2468"     # contraseña de fábrica (se puede cambiar en el menú oculto)
TOQUES_MENU_OCULTO = 5          # toques seguidos al título para abrir el menú oculto
NUM_POS = 26
MAX_NORMAL = 6
CAT_MAS_DE_6 = "MAS DE 6 NUMEROS POR PALLET"
CATEGORIAS = [
    "MAS DE 6 NUMEROS POR PALLET",
    "CAMBIO DE TRK POR RETRAZO",
    "CANTIDAD ESCRITA A MANO",
    "ETIQUETAS MAL PUESTAS",
    "FALTO CAPTURAR",
    "FALTO DE MOVER MATERIAL",
    "MAL CAPTURADO",
    "MAL ETIQUETADO",
    "MAL ETIQUETADO Y CAPTURADO",
    "MATERIAL EQUIVOCADO",
    "MATERIAL MAL IDENTIFICADO",
    "OTRO",
]
TURNOS = ["TURNO 1", "TURNO 2", "TURNO 3"]

FONDO = (0.95, 0.96, 0.98, 1)
AZUL = (0.12, 0.31, 0.47, 1)
VERDE = (0.18, 0.62, 0.35, 1)
NARANJA = (0.93, 0.49, 0.13, 1)
GRIS = (0.80, 0.82, 0.86, 1)
ROJO = (0.78, 0.20, 0.20, 1)
MORADO = (0.40, 0.25, 0.60, 1)
OSCURO = (0.15, 0.17, 0.20, 1)
BLANCO = (1, 1, 1, 1)

Window.clearcolor = FONDO
Window.softinput_mode = "below_target"
AQUI = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.path.join(AQUI, "plantilla.xlsm")


# ---------------------------------------------------------------- widgets base
class Mayus(TextInput):
    """Campo que convierte todo a MAYÚSCULAS mientras se escribe."""

    def insert_text(self, substring, from_undo=False):
        return super().insert_text(substring.upper(), from_undo=from_undo)


def boton(texto, fondo=AZUL, accion=None, **kw):
    kw.setdefault("font_size", sp(16))
    kw.setdefault("bold", True)
    b = Button(text=texto, background_normal="", background_color=fondo, halign="center", **kw)
    b.bind(size=lambda w, s: setattr(w, "text_size", (s[0] - dp(6), None)))
    if accion:
        b.bind(on_release=lambda *_: accion())
    return b


def campo(hint, mayus=True, **kw):
    """Campo de una línea. mayus=False solo para números de parte, correos y contraseñas."""
    kw.setdefault("size_hint_y", None)
    kw.setdefault("height", dp(48))
    kw.setdefault("font_size", sp(17))
    clase = Mayus if mayus else TextInput
    return clase(hint_text=hint, multiline=False, write_tab=False,
                 text_validate_unfocus=False, padding=[dp(8), dp(12)], **kw)


def etiqueta(texto, **kw):
    kw.setdefault("color", OSCURO)
    kw.setdefault("font_size", sp(15))
    lb = Label(text=texto, **kw)
    lb.bind(size=lambda w, s: setattr(w, "text_size", (s[0], None)))
    return lb


def enfocar(widget):
    Clock.schedule_once(lambda *_: setattr(widget, "focus", True), 0.05)


def ventana(titulo, contenido, alto=dp(260), cerrable=True):
    return Popup(title=titulo, content=contenido, size_hint=(0.94, None), height=alto,
                 auto_dismiss=cerrable, title_size=sp(17))


def aviso(mensaje, titulo="Aviso"):
    caja = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(12))
    caja.add_widget(etiqueta(mensaje, color=BLANCO, halign="center", font_size=sp(16)))
    p = ventana(titulo, caja)
    caja.add_widget(boton("OK", accion=p.dismiss, size_hint_y=None, height=dp(50)))
    p.open()


def confirmar(mensaje, al_aceptar, titulo="Confirmar", si="Sí", no="Cancelar", al_cancelar=None):
    caja = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(12))
    caja.add_widget(etiqueta(mensaje, color=BLANCO, halign="center", font_size=sp(16)))
    fila = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
    p = ventana(titulo, caja, cerrable=False)

    def _si():
        p.dismiss()
        al_aceptar()

    def _no():
        p.dismiss()
        if al_cancelar:
            al_cancelar()

    fila.add_widget(boton(no, GRIS, _no, color=OSCURO))
    fila.add_widget(boton(si, ROJO, _si))
    caja.add_widget(fila)
    p.open()


def limpiar_cantidad(txt):
    """El escáner puede mandar 'Q100' o '100 PCS': nos quedamos con el número."""
    m = re.search(r"\d+(\.\d+)?", txt or "")
    if not m:
        return None
    v = float(m.group())
    return int(v) if v.is_integer() else v


def limpiar_parte(txt):
    # El número de parte se respeta tal cual lo manda el escáner (sin pasar a mayúsculas)
    return (txt or "").strip()


# ---------------------------------------------------------------- camión
class Camion(RelativeLayout):
    """Tractor + caja con las 26 posiciones en tamaño grande (va dentro de un ScrollView)."""
    CELDA_LARGO = dp(96)   # alto de cada posición (vertical) / ancho (horizontal)

    def __init__(self, app, **kw):
        super().__init__(size_hint=(None, None), **kw)
        self.app = app
        self.horizontal = False
        self.botones = {}
        self.fondo = Widget(size_hint=(None, None))
        self.add_widget(self.fondo)
        for n in range(1, NUM_POS + 1):
            b = Button(text=str(n), font_size=sp(28), bold=True, background_normal="",
                       background_color=GRIS, color=OSCURO, halign="center", valign="middle",
                       size_hint=(None, None))
            b.bind(on_release=lambda btn, n=n: app.abrir_posicion(n))
            b.bind(size=lambda w, s: setattr(w, "text_size", s))
            self.botones[n] = b
            self.add_widget(b)
        self.bind(size=self._dibujar)

    def ajustar(self, ancho, alto):
        """Calcula el tamaño según el espacio visible del ScrollView."""
        self.horizontal = ancho > alto * 1.2
        if self.horizontal:
            self.size = (dp(150) + 13 * self.CELDA_LARGO * 1.15 + dp(20), max(alto, dp(220)))
        else:
            self.size = (ancho, dp(150) + 13 * self.CELDA_LARGO + dp(20))
        self._dibujar()

    def _dibujar(self, *_):
        w, h = self.size
        pad = dp(8)
        self.fondo.size = self.size
        self.fondo.canvas.clear()
        with self.fondo.canvas:
            if self.horizontal:
                cab = dp(150)
                tx, ty, tw, th = cab + dp(12), pad, w - cab - dp(12) - pad, h - 2 * pad
                self._tractor_h(pad, ty, cab - pad, th)
            else:
                cab = dp(150)
                tw = min(w - 2 * pad, dp(520))
                tx = (w - tw) / 2
                ty, th = pad, h - cab - dp(12) - pad
                self._tractor_v(tx, h - cab, tw, cab - pad)
            Color(*OSCURO)
            Line(rectangle=(tx - dp(3), ty - dp(3), tw + dp(6), th + dp(6)), width=dp(2))
        g = dp(3)
        if self.horizontal:
            cw, ch = tw / 13, th / 2
            for i in range(13):
                x = tx + i * cw
                self._colocar(2 * i + 2, x + g, ty + ch + g, cw - 2 * g, ch - 2 * g)  # arriba pares
                self._colocar(2 * i + 1, x + g, ty + g, cw - 2 * g, ch - 2 * g)       # abajo nones
        else:
            cw, ch = tw / 2, th / 13
            for i in range(13):
                y = ty + th - (i + 1) * ch
                self._colocar(2 * i + 1, tx + g, y + g, cw - 2 * g, ch - 2 * g)       # izq nones
                self._colocar(2 * i + 2, tx + cw + g, y + g, cw - 2 * g, ch - 2 * g)  # der pares

    def _colocar(self, n, x, y, w, h):
        b = self.botones[n]
        b.pos, b.size = (x, y), (w, h)

    def _tractor_h(self, x, y, w, h):
        cy = y + h / 2
        ch = min(h * 0.7, dp(170))
        Color(*OSCURO)
        for yy in (cy + ch / 2, cy - ch / 2 - dp(10)):
            Rectangle(pos=(x + w * 0.12, yy), size=(w * 0.16, dp(10)))
            Rectangle(pos=(x + w * 0.62, yy), size=(w * 0.22, dp(10)))
        Color(0.30, 0.33, 0.38, 1)
        RoundedRectangle(pos=(x, cy - ch / 2), size=(w * 0.55, ch), radius=[dp(12), 0, 0, dp(12)])
        Color(0.75, 0.86, 0.95, 1)
        RoundedRectangle(pos=(x + w * 0.20, cy - ch * 0.38), size=(w * 0.12, ch * 0.76), radius=[dp(4)])
        Color(0.40, 0.43, 0.48, 1)
        Rectangle(pos=(x + w * 0.57, cy - ch * 0.33), size=(w * 0.30, ch * 0.66))
        Color(*OSCURO)
        Line(points=[x + w * 0.87, cy, x + w + dp(12), cy], width=dp(3))

    def _tractor_v(self, x, y, w, h):
        cx = x + w / 2
        cw = min(w * 0.7, dp(260))
        Color(*OSCURO)
        for xx in (cx - cw / 2 - dp(10), cx + cw / 2):
            Rectangle(pos=(xx, y + h * 0.72), size=(dp(10), h * 0.16))
            Rectangle(pos=(xx, y + h * 0.12), size=(dp(10), h * 0.22))
        Color(0.30, 0.33, 0.38, 1)
        RoundedRectangle(pos=(cx - cw / 2, y + h * 0.45), size=(cw, h * 0.55), radius=[dp(12), dp(12), 0, 0])
        Color(0.75, 0.86, 0.95, 1)
        RoundedRectangle(pos=(cx - cw * 0.38, y + h * 0.66), size=(cw * 0.76, h * 0.12), radius=[dp(4)])
        Color(0.40, 0.43, 0.48, 1)
        Rectangle(pos=(cx - cw * 0.33, y + h * 0.13), size=(cw * 0.66, h * 0.30))
        Color(*OSCURO)
        Line(points=[cx, y + h * 0.13, cx, y - dp(12)], width=dp(3))

    def refrescar(self):
        a = self.app.audit
        con_inc = {str(i.get("posicion")) for i in self.app.incidencias_troca()}
        for n, b in self.botones.items():
            d = a["posiciones"].get(str(n))
            lineas = len(d["lineas"]) if d else 0
            if str(n) in con_inc:
                b.background_color, b.color = NARANJA, BLANCO
            elif lineas:
                b.background_color, b.color = VERDE, BLANCO
            else:
                b.background_color, b.color = GRIS, OSCURO
            if lineas:
                area = d.get("area", "")
                b.text = f"{n}\n[size={int(sp(13))}]{area[:10]} · {lineas} pza[/size]"
                b.markup = True
            else:
                b.text = str(n)


# ---------------------------------------------------------------- pantalla principal
class PantallaPrincipal(Screen):
    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        self._toques = []
        raiz = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(6))

        barra = BoxLayout(size_hint_y=None, height=dp(46))
        with barra.canvas.before:
            Color(*AZUL)
            self._fondo = Rectangle()
        barra.bind(pos=lambda w, p: setattr(self._fondo, "pos", p),
                   size=lambda w, s: setattr(self._fondo, "size", s))
        # El título es un botón "invisible": 5 toques seguidos abren el menú oculto
        self.titulo = Button(text="AUDITORÍA DE CARGA", bold=True, font_size=sp(18),
                             background_normal="", background_down="", background_color=(0, 0, 0, 0))
        self.titulo.bind(on_release=self._toque_titulo)
        barra.add_widget(self.titulo)
        self.lbl_fecha = Label(text="", font_size=sp(13), size_hint_x=0.45)
        barra.add_widget(self.lbl_fecha)
        raiz.add_widget(barra)

        # pestañas para modo dos trocas (ocultas si está apagado)
        self.pestanas = BoxLayout(size_hint_y=None, height=0, spacing=dp(6), opacity=0)
        self.btn_t = [boton("TROCA 1", AZUL, lambda: app.cambiar_troca(0)),
                      boton("TROCA 2", GRIS, lambda: app.cambiar_troca(1), color=OSCURO)]
        for b in self.btn_t:
            self.pestanas.add_widget(b)
        raiz.add_widget(self.pestanas)

        datos = GridLayout(cols=3, size_hint_y=None, height=dp(48), spacing=dp(6))
        self.trk = campo("TRK / CAJA")
        self.auditor = campo("AUDITOR")
        self.turno = Spinner(text=TURNOS[0], values=TURNOS, size_hint_y=None, height=dp(48),
                             font_size=sp(16), background_normal="", background_color=AZUL)
        self.trk.bind(text=lambda w, t: app.set_dato("trk", t.strip()))
        self.auditor.bind(text=lambda w, t: app.set_dato("auditor", t.strip()))
        self.turno.bind(text=lambda w, t: app.set_dato("turno", t))
        self.trk.bind(on_text_validate=lambda *_: enfocar(self.auditor))
        for wdg in (self.trk, self.auditor, self.turno):
            datos.add_widget(wdg)
        raiz.add_widget(datos)

        self.scroll = ScrollView(bar_width=dp(8), scroll_type=["bars", "content"])
        self.camion = Camion(app)
        self.scroll.add_widget(self.camion)
        self.scroll.bind(size=self._ajustar_camion)
        raiz.add_widget(self.scroll)

        self.lbl_estado = etiqueta("", size_hint_y=None, height=dp(24), halign="center", font_size=sp(14))
        raiz.add_widget(self.lbl_estado)

        acciones = BoxLayout(size_hint_y=None, height=dp(56), spacing=dp(6))
        acciones.add_widget(boton("Incidencia", NARANJA, lambda: app.popup_incidencia()))
        acciones.add_widget(boton("Generar Excel y enviar", VERDE, app.popup_enviar))
        acciones.add_widget(boton("Nueva troca", ROJO, app.nueva_troca, size_hint_x=0.7))
        raiz.add_widget(acciones)
        self.add_widget(raiz)

    def _ajustar_camion(self, sv, size):
        self.camion.ajustar(*size)
        self.scroll.do_scroll_x = self.camion.width > size[0] + 1
        self.scroll.do_scroll_y = self.camion.height > size[1] + 1

    def _toque_titulo(self, *_):
        ahora = time.time()
        self._toques = [t for t in self._toques if ahora - t < 3] + [ahora]
        if len(self._toques) >= TOQUES_MENU_OCULTO:
            self._toques = []
            self.app.menu_oculto()

    def refrescar(self):
        app = self.app
        a = app.audit
        self.trk.text, self.auditor.text, self.turno.text = a["trk"], a["auditor"], a["turno"]
        self.lbl_fecha.text = datetime.fromisoformat(a["fecha"]).strftime("%d/%m/%Y %H:%M")
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        piezas = sum(len(d["lineas"]) for d in a["posiciones"].values())
        self.lbl_estado.text = (f"Posiciones: {capt}/{NUM_POS}  ·  Números de parte: {piezas}"
                                f"  ·  Incidencias: {len(app.incidencias_troca())}")
        dos = app.config_app.get("dos_trocas", False)
        self.pestanas.height, self.pestanas.opacity, self.pestanas.disabled = \
            (dp(46), 1, False) if dos else (0, 0, True)
        for i, b in enumerate(self.btn_t):
            t = app.estado["trocas"][i] if i < len(app.estado["trocas"]) else None
            nombre = f"TROCA {i + 1}" + (f": {t['trk']}" if t and t.get("trk") else "")
            b.text = nombre
            activa = i == app.estado["activa"]
            b.background_color, b.color = (AZUL, BLANCO) if activa else (GRIS, OSCURO)
        self.camion.refrescar()

    def on_pre_enter(self, *_):
        self.refrescar()


# ---------------------------------------------------------------- pantalla de posición
class PantallaPosicion(Screen):
    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        self.num = None
        self.filas = []
        self.extra_nombre = ""

        raiz = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(6))
        barra = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(6))
        barra.add_widget(boton("< Camión", GRIS, self.regresar, color=OSCURO, size_hint_x=0.35))
        self.titulo = Label(text="", bold=True, font_size=sp(22), color=AZUL)
        barra.add_widget(self.titulo)
        raiz.add_widget(barra)

        self.area = campo("ÁREA (teclear)")
        self.area.bind(on_text_validate=lambda *_: self.filas and enfocar(self.filas[0][0]))
        raiz.add_widget(self.area)

        enc = BoxLayout(size_hint_y=None, height=dp(26), spacing=dp(4))
        enc.add_widget(etiqueta("#", size_hint_x=None, width=dp(30), bold=True))
        enc.add_widget(etiqueta("No. de parte (escanear)", bold=True))
        enc.add_widget(etiqueta("Cantidad", size_hint_x=0.45, bold=True))
        enc.add_widget(Label(size_hint_x=None, width=dp(42)))
        raiz.add_widget(enc)

        self.lista = GridLayout(cols=1, spacing=dp(4), size_hint_y=None)
        self.lista.bind(minimum_height=self.lista.setter("height"))
        self.scroll = ScrollView(do_scroll_x=False)
        self.scroll.add_widget(self.lista)
        raiz.add_widget(self.scroll)

        self.lbl_extra = etiqueta("", size_hint_y=None, height=dp(22), color=NARANJA, font_size=sp(13))
        raiz.add_widget(self.lbl_extra)

        acciones = BoxLayout(size_hint_y=None, height=dp(56), spacing=dp(6))
        acciones.add_widget(boton("+ Agregar número", AZUL, self.agregar))
        acciones.add_widget(boton("Incidencia", NARANJA, lambda: app.popup_incidencia(posicion=self.num)))
        acciones.add_widget(boton("Guardar", VERDE, self.guardar))
        raiz.add_widget(acciones)
        self.add_widget(raiz)

    def _nueva_fila(self, parte="", cant=""):
        idx = len(self.filas) + 1
        fila = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(4))
        extra = idx > MAX_NORMAL
        fila.add_widget(etiqueta(str(idx), size_hint_x=None, width=dp(30), bold=True,
                                 color=NARANJA if extra else OSCURO))
        tp = campo("No. parte", mayus=False)          # el número de parte NO se pasa a mayúsculas
        tc = campo("Cant.", mayus=False, size_hint_x=0.45, input_type="number")
        tp.text, tc.text = parte, "" if cant in (None, "") else str(cant)
        if extra:
            tp.background_color = tc.background_color = (1, 0.93, 0.86, 1)
        tp.bind(on_text_validate=lambda *_: enfocar(tc))
        tc.bind(on_text_validate=lambda w: self._siguiente(w))
        fila.add_widget(tp)
        fila.add_widget(tc)
        borrar = boton("X", ROJO, size_hint_x=None, width=dp(42))
        borrar.bind(on_release=lambda *_: (setattr(tp, "text", ""), setattr(tc, "text", ""), enfocar(tp)))
        fila.add_widget(borrar)
        self.filas.append((tp, tc))
        self.lista.add_widget(fila)
        return tp

    def _siguiente(self, tc):
        v = limpiar_cantidad(tc.text)
        tc.text = "" if v is None else str(v)
        for i, (_, c) in enumerate(self.filas):
            if c is tc and i + 1 < len(self.filas):
                enfocar(self.filas[i + 1][0])
                self.scroll.scroll_to(self.filas[i + 1][0].parent, padding=dp(40))
                return
        tc.focus = False

    def cargar(self, num):
        self.num = num
        datos = self.app.audit["posiciones"].get(str(num), {"area": "", "lineas": []})
        self.titulo.text = f"Posición {num}"
        self.area.text = datos.get("area", "")
        self.extra_nombre = datos.get("extra_nombre", "")
        self.filas = []
        self.lista.clear_widgets()
        for ln in datos.get("lineas", []):
            self._nueva_fila(ln["parte"], ln["cant"])
        while len(self.filas) < MAX_NORMAL:
            self._nueva_fila()
        self._actualizar_extra()
        if not self.area.text:
            enfocar(self.area)
        else:
            vacia = next((p for p, _ in self.filas if not p.text), None)
            if vacia:
                enfocar(vacia)

    def _actualizar_extra(self):
        self.lbl_extra.text = (f"Más de 6 números autorizado — responsable: {self.extra_nombre}"
                               if self.extra_nombre else "")

    def _extras_otros(self):
        """Partes extra (>6) ya capturadas en las demás posiciones de esta troca."""
        return sum(max(0, len(d["lineas"]) - MAX_NORMAL)
                   for k, d in self.app.audit["posiciones"].items() if k != str(self.num))

    def agregar(self):
        if len(self.filas) < MAX_NORMAL:
            enfocar(self._nueva_fila())
            return
        extras_aqui = len(self.filas) - MAX_NORMAL
        if self._extras_otros() + extras_aqui >= reporte.MAX_EXTRA:
            aviso(f"El formato HCARGA solo tiene lugar para {reporte.MAX_EXTRA} números extra "
                  f"(bloques 27 y 28) en toda la troca.", "Sin espacio")
            return
        if self.extra_nombre:
            self._fila_extra()
            return
        self.app.popup_incidencia(posicion=self.num, categoria=CAT_MAS_DE_6,
                                  titulo=f"Posición {self.num}: más de 6 números de parte",
                                  tipo=f"PALLET EN POSICION {self.num} CON MAS DE 6 NUMEROS DE PARTE",
                                  al_guardar=self._autorizar_extra)

    def _fila_extra(self):
        tp = self._nueva_fila()
        enfocar(tp)
        Clock.schedule_once(lambda *_: self.scroll.scroll_to(tp.parent), 0.1)

    def _autorizar_extra(self, nombre):
        self.extra_nombre = nombre
        self._actualizar_extra()
        self.guardar(salir=False)
        self._fila_extra()

    def _leer(self):
        lineas, errores = [], []
        for i, (tp, tc) in enumerate(self.filas, start=1):
            parte, cant = limpiar_parte(tp.text), limpiar_cantidad(tc.text)
            if not parte and cant is None:
                continue
            if not parte:
                errores.append(f"Renglón {i}: falta el número de parte")
            elif cant is None:
                errores.append(f"Renglón {i}: falta la cantidad de {parte}")
            lineas.append({"parte": parte, "cant": cant})
        return lineas, errores

    def guardar(self, salir=True):
        lineas, errores = self._leer()
        if errores and salir:
            aviso("\n".join(errores[:6]), "Revisa la captura")
            return
        if lineas and not self.area.text.strip() and salir:
            aviso("Teclea el área de la posición.", "Falta el área")
            enfocar(self.area)
            return
        self.app.audit["posiciones"][str(self.num)] = {
            "area": self.area.text.strip().upper(),
            "lineas": [ln for ln in lineas if ln["parte"]],
            "extra_nombre": self.extra_nombre,
        }
        self.app.guardar_local()
        if salir:
            self.app.ir("principal")

    def regresar(self):
        lineas, _ = self._leer()
        guardado = self.app.audit["posiciones"].get(str(self.num), {}).get("lineas", [])
        if lineas != guardado:
            confirmar("Hay cambios sin guardar. ¿Salir sin guardar?", lambda: self.app.ir("principal"))
        else:
            self.app.ir("principal")


# ---------------------------------------------------------------- app
class AuditoriaApp(App):
    title = "Auditoría de Carga"

    def build(self):
        self.ruta_estado = os.path.join(self.user_data_dir, "estado.json")
        self.ruta_config = os.path.join(self.user_data_dir, "config.json")
        self.ruta_ultima = os.path.join(self.user_data_dir, "ultima_troca.json")
        self.config_app = {"correo": "", "correo_inc": "", "auditor": "", "turno": TURNOS[0],
                           "dos_trocas": False, "contrasena": CONTRASENA_INICIAL}
        self.config_app.update(self._leer_json(self.ruta_config) or {})
        self.estado = self._leer_json(self.ruta_estado) or {"trocas": [self._troca_vacia()],
                                                            "activa": 0, "incidencias": []}
        if not self.config_app["dos_trocas"]:
            self.estado["trocas"] = self.estado["trocas"][:1]
            self.estado["activa"] = 0

        self.sm = ScreenManager(transition=NoTransition())
        self.principal = PantallaPrincipal(self, name="principal")
        self.posicion = PantallaPosicion(self, name="posicion")
        self.sm.add_widget(self.principal)
        self.sm.add_widget(self.posicion)
        Window.bind(on_keyboard=self._tecla_atras)
        return self.sm

    # -- archivos temporales (el Excel nunca se queda en el dispositivo)
    def carpeta_temporal(self):
        if platform == "android":
            from android import mActivity  # noqa
            base = mActivity.getCacheDir().getAbsolutePath()
        else:
            base = tempfile.gettempdir()
        ruta = os.path.join(base, "reportes")
        os.makedirs(ruta, exist_ok=True)
        return ruta

    def borrar_temporales(self, mas_viejos_que=0):
        try:
            carpeta = self.carpeta_temporal()
            for f in os.listdir(carpeta):
                ruta = os.path.join(carpeta, f)
                if time.time() - os.path.getmtime(ruta) >= mas_viejos_que:
                    os.remove(ruta)
        except Exception:
            pass

    def on_start(self):
        self.borrar_temporales()

    def on_resume(self):
        # al regresar del correo se borra el archivo enviado
        Clock.schedule_once(lambda *_: self.borrar_temporales(mas_viejos_que=60), 2)

    def on_pause(self):
        return True

    # -- datos
    def _troca_vacia(self):
        return {"id": uuid.uuid4().hex[:10], "fecha": datetime.now().isoformat(timespec="seconds"),
                "trk": "", "auditor": self.config_app.get("auditor", ""),
                "turno": self.config_app.get("turno", TURNOS[0]), "posiciones": {}}

    @property
    def audit(self):
        return self.estado["trocas"][self.estado["activa"]]

    def incidencias_troca(self):
        return [i for i in self.estado["incidencias"] if i.get("troca_id") == self.audit["id"]]

    @staticmethod
    def _leer_json(ruta):
        try:
            with open(ruta, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None

    @staticmethod
    def _escribir_json(ruta, data):
        tmp = ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, ruta)

    def guardar_local(self):
        self._escribir_json(self.ruta_estado, self.estado)
        self._escribir_json(self.ruta_config, self.config_app)

    def set_dato(self, clave, valor):
        if clave in ("trk", "auditor"):
            valor = valor.upper()
        if self.audit.get(clave) != valor:
            self.audit[clave] = valor
            if clave in ("auditor", "turno"):
                self.config_app[clave] = valor
            self.guardar_local()

    # -- navegación
    def ir(self, pantalla):
        self.sm.current = pantalla

    def abrir_posicion(self, n):
        self.posicion.cargar(n)
        self.ir("posicion")

    def _tecla_atras(self, window, key, *args):
        if key == 27 and self.sm.current == "posicion":
            self.posicion.regresar()
            return True
        return False

    def cambiar_troca(self, i):
        if i >= len(self.estado["trocas"]):
            self.estado["trocas"].append(self._troca_vacia())
        self.estado["activa"] = i
        self.guardar_local()
        self.principal.refrescar()

    def _limpiar_troca_activa(self):
        self.estado["trocas"][self.estado["activa"]] = self._troca_vacia()
        self.guardar_local()
        self.principal.refrescar()
        enfocar(self.principal.trk)

    def nueva_troca(self):
        confirmar("Se borrará la captura de esta troca (asegúrate de haber enviado el Excel). ¿Continuar?",
                  self._limpiar_troca_activa, "Nueva troca")

    # -- incidencias
    def popup_incidencia(self, posicion="", categoria=None, titulo="Registrar incidencia",
                         tipo="", al_guardar=None):
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        nombre = campo("NOMBRE DE QUIEN CARGÓ (obligatorio)")
        cat = Spinner(text=categoria or CATEGORIAS[0], values=CATEGORIAS, size_hint_y=None, height=dp(48),
                      font_size=sp(15), background_normal="", background_color=AZUL)
        desc = Mayus(hint_text="DESCRIPCIÓN / COMENTARIO", text=tipo.upper(), font_size=sp(15),
                     size_hint_y=None, height=dp(90))
        pos = campo("Posición (opcional)", mayus=False, text=str(posicion or ""), input_filter="int")
        for w in (etiqueta("Nombre:", size_hint_y=None, height=dp(20), color=BLANCO), nombre,
                  etiqueta("Categoría:", size_hint_y=None, height=dp(20), color=BLANCO), cat, desc, pos):
            caja.add_widget(w)
        fila = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        p = ventana(titulo, caja, alto=dp(490), cerrable=False)

        def ok():
            if not nombre.text.strip():
                nombre.hint_text = "¡ESCRIBE EL NOMBRE!"
                enfocar(nombre)
                return
            quien = nombre.text.strip().upper()
            self.estado["incidencias"].append({
                "nombre": quien,
                "tipo": (desc.text.strip() or cat.text).upper(),
                "categoria": cat.text.upper(),
                "fecha": datetime.now().isoformat(timespec="seconds"),
                "posicion": pos.text.strip(),
                "trk": self.audit.get("trk", ""),
                "auditor": self.audit.get("auditor", ""),
                "troca_id": self.audit["id"],
            })
            self.guardar_local()
            p.dismiss()
            self.principal.refrescar()
            if al_guardar:
                al_guardar(quien)

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Guardar", NARANJA, ok))
        caja.add_widget(fila)
        nombre.bind(on_text_validate=lambda *_: enfocar(desc))
        p.open()
        enfocar(nombre)

    # -- envío del Excel de carga
    def popup_enviar(self):
        a = self.audit
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        caja.add_widget(etiqueta(
            f"TRK: {a['trk'] or '—'}   AUDITOR: {a['auditor'] or '—'}   {a['turno']}\n"
            f"Posiciones capturadas: {capt}/{NUM_POS}",
            size_hint_y=None, height=dp(56), color=BLANCO))
        correo = campo("Correo(s) destino, separados por coma", mayus=False,
                       text=self.config_app.get("correo", ""), input_type="mail")
        caja.add_widget(correo)
        fila = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        p = ventana("Enviar Excel de carga", caja, alto=dp(280))

        def enviar():
            if not a["trk"] or not a["auditor"]:
                aviso("Captura el TRK y el nombre del auditor en la pantalla principal.", "Faltan datos")
                return
            if not capt:
                aviso("No hay posiciones capturadas.", "Sin datos")
                return
            self.config_app["correo"] = correo.text.strip()
            self.guardar_local()
            p.dismiss()
            self.enviar_carga()

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Generar y enviar", VERDE, enviar))
        caja.add_widget(fila)
        p.open()

    def enviar_carga(self):
        a = self.audit
        ruta = os.path.join(self.carpeta_temporal(), reporte.nombre_carga(a))
        try:
            reporte.generar_carga(a, PLANTILLA, ruta)
        except Exception as e:
            aviso(f"No se pudo generar el Excel:\n{e}", "Error")
            return
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        asunto = f"Auditoría de carga TRK {a['trk']} - {datetime.now():%d/%m/%Y %H:%M}"
        cuerpo = (f"Se adjunta la auditoría de carga.\n\nTRK: {a['trk']}\nAuditor: {a['auditor']}\n"
                  f"{a['turno']}\nPosiciones capturadas: {capt}/{NUM_POS}\n")
        if not self._compartir(ruta, reporte.MIME_XLSM, self.config_app.get("correo", ""), asunto, cuerpo):
            return
        # respaldo de la última troca enviada (para "Recuperar última troca")
        self._escribir_json(self.ruta_ultima, a)
        Clock.schedule_once(lambda *_: confirmar(
            "¿Se envió el correo?\nSi ya se envió, se limpia esta troca para empezar la siguiente.",
            self._limpiar_troca_activa, "Troca enviada", si="Sí, nueva troca", no="No, conservar"), 1)

    def _compartir(self, ruta, mime, correos, asunto, cuerpo):
        para = [c.strip() for c in correos.split(",") if c.strip()]
        if platform == "android":
            try:
                compartir_android(ruta, mime, para, asunto, cuerpo)
                return True
            except Exception as e:
                aviso(f"No se pudo abrir el correo:\n{e}", "Error")
                self.borrar_temporales()
                return False
        aviso(f"(Modo PC) Archivo temporal generado:\n{ruta}", "Listo")
        return True

    # -- menú oculto
    def menu_oculto(self):
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        p = ventana("Menú de supervisor", caja, alto=dp(470))
        n_inc = len(self.estado["incidencias"])
        dos = self.config_app.get("dos_trocas", False)

        def cerrar_y(f):
            return lambda: (p.dismiss(), f())

        caja.add_widget(boton(f"Reporte de incidencias ({n_inc})", NARANJA, cerrar_y(self.popup_incidencias)))
        caja.add_widget(boton("Recuperar última troca enviada  [contraseña]", MORADO,
                              cerrar_y(lambda: self.pedir_contrasena(self.recuperar_ultima))))
        caja.add_widget(boton(f"Modo dos trocas: {'ENCENDIDO' if dos else 'APAGADO'}  [contraseña]", MORADO,
                              cerrar_y(lambda: self.pedir_contrasena(self.alternar_dos_trocas))))
        caja.add_widget(boton("Cambiar contraseña", GRIS,
                              cerrar_y(lambda: self.pedir_contrasena(self.cambiar_contrasena)), color=OSCURO))
        caja.add_widget(boton("Cerrar", AZUL, p.dismiss))
        p.open()

    def pedir_contrasena(self, al_ok):
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        clave = campo("Contraseña", mayus=False, password=True)
        caja.add_widget(clave)
        fila = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        p = ventana("Contraseña", caja, alto=dp(220), cerrable=False)

        def ok(*_):
            if clave.text == self.config_app.get("contrasena", CONTRASENA_INICIAL):
                p.dismiss()
                al_ok()
            else:
                clave.text = ""
                clave.hint_text = "Contraseña incorrecta"
                enfocar(clave)

        clave.bind(on_text_validate=ok)
        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Aceptar", AZUL, ok))
        caja.add_widget(fila)
        p.open()
        enfocar(clave)

    def cambiar_contrasena(self):
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        n1 = campo("Nueva contraseña", mayus=False, password=True)
        n2 = campo("Repetir nueva contraseña", mayus=False, password=True)
        caja.add_widget(n1)
        caja.add_widget(n2)
        fila = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        p = ventana("Cambiar contraseña", caja, alto=dp(280), cerrable=False)

        def ok():
            if len(n1.text) < 4 or n1.text != n2.text:
                n2.text = ""
                n2.hint_text = "No coinciden (mínimo 4 caracteres)"
                return
            self.config_app["contrasena"] = n1.text
            self.guardar_local()
            p.dismiss()
            aviso("Contraseña actualizada.")

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Guardar", VERDE, ok))
        caja.add_widget(fila)
        p.open()
        enfocar(n1)

    def recuperar_ultima(self):
        ultima = self._leer_json(self.ruta_ultima)
        if not ultima:
            aviso("No hay ninguna troca enviada guardada.")
            return

        def hacer():
            self.estado["trocas"][self.estado["activa"]] = ultima
            self.guardar_local()
            self.principal.refrescar()
            aviso(f"Se recuperó la troca {ultima.get('trk', '')}.", "Listo")

        tiene_datos = any(d["lineas"] for d in self.audit["posiciones"].values())
        if tiene_datos:
            confirmar(f"Se reemplazará la captura actual por la última troca enviada "
                      f"({ultima.get('trk', '')}). ¿Continuar?", hacer, "Recuperar")
        else:
            hacer()

    def alternar_dos_trocas(self):
        if not self.config_app.get("dos_trocas"):
            self.config_app["dos_trocas"] = True
            if len(self.estado["trocas"]) < 2:
                self.estado["trocas"].append(self._troca_vacia())
            self.guardar_local()
            self.principal.refrescar()
            aviso("Modo dos trocas ENCENDIDO.\nUsa las pestañas TROCA 1 / TROCA 2.")
            return

        def apagar():
            self.config_app["dos_trocas"] = False
            activa = self.estado["trocas"][self.estado["activa"]]
            self.estado["trocas"] = [activa]
            self.estado["activa"] = 0
            self.guardar_local()
            self.principal.refrescar()

        otra = self.estado["trocas"][1 - self.estado["activa"]] if len(self.estado["trocas"]) > 1 else None
        if otra and any(d["lineas"] for d in otra["posiciones"].values()):
            confirmar("La otra troca tiene captura y se borrará. ¿Apagar modo dos trocas?", apagar)
        else:
            apagar()

    # -- reporte de incidencias (aparte)
    def popup_incidencias(self):
        incs = self.estado["incidencias"]
        if not incs:
            aviso("No hay incidencias registradas.")
            return
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        caja.add_widget(etiqueta(f"Incidencias pendientes de enviar: {len(incs)}",
                                 size_hint_y=None, height=dp(30), color=BLANCO))
        correo = campo("Correo(s) para incidencias", mayus=False,
                       text=self.config_app.get("correo_inc") or self.config_app.get("correo", ""),
                       input_type="mail")
        caja.add_widget(correo)
        fila = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        p = ventana("Reporte de incidencias", caja, alto=dp(260))

        def enviar():
            self.config_app["correo_inc"] = correo.text.strip()
            self.guardar_local()
            p.dismiss()
            ruta = os.path.join(self.carpeta_temporal(), reporte.nombre_incidencias())
            try:
                reporte.generar_incidencias(incs, ruta)
            except Exception as e:
                aviso(f"No se pudo generar el reporte:\n{e}", "Error")
                return
            asunto = f"Reporte de incidencias de carga - {datetime.now():%d/%m/%Y %H:%M}"
            cuerpo = f"Se adjunta el reporte de incidencias ({len(incs)}).\n"
            if self._compartir(ruta, reporte.MIME_XLSX, correo.text, asunto, cuerpo):
                enviadas = len(incs)
                Clock.schedule_once(lambda *_: confirmar(
                    "¿Se envió el reporte de incidencias?\nSi ya se envió, se borran de la lista.",
                    lambda: self._borrar_incidencias(enviadas), "Incidencias", si="Sí, borrar",
                    no="No, conservar"), 1)

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Generar y enviar", NARANJA, enviar))
        caja.add_widget(fila)
        p.open()

    def _borrar_incidencias(self, n):
        del self.estado["incidencias"][:n]
        self.guardar_local()
        self.principal.refrescar()


# ---------------------------------------------------------------- Android
def compartir_android(ruta, mime, para, asunto, cuerpo):
    """Comparte el archivo desde la caché privada de la app (FileProvider): no queda en Descargas."""
    from jnius import autoclass, cast
    from android import mActivity  # noqa

    Intent = autoclass("android.content.Intent")
    JString = autoclass("java.lang.String")
    File = autoclass("java.io.File")
    FileProvider = autoclass("androidx.core.content.FileProvider")

    uri = FileProvider.getUriForFile(mActivity, mActivity.getPackageName() + ".fileprovider", File(ruta))
    intent = Intent(Intent.ACTION_SEND)
    intent.setType(mime)
    intent.putExtra(Intent.EXTRA_STREAM, cast("android.os.Parcelable", uri))
    intent.putExtra(Intent.EXTRA_SUBJECT, JString(asunto))
    intent.putExtra(Intent.EXTRA_TEXT, JString(cuerpo))
    if para:
        try:
            intent.putExtra(Intent.EXTRA_EMAIL, [JString(c) for c in para])
        except Exception:
            pass
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    mActivity.startActivity(Intent.createChooser(intent, JString("Enviar por correo")))


if __name__ == "__main__":
    AuditoriaApp().run()
