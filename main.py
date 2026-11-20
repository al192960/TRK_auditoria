# -*- coding: utf-8 -*-
"""
Auditoría de Carga - App Android (Kivy)

Pantalla principal = el camión con las 26 posiciones.
Al tocar una posición se captura el área y hasta 6 números de parte con su
cantidad (con escáner tipo teclado o a mano). Si se necesitan más de 6, se pide
el nombre del responsable y se registra una incidencia.
Al final se genera un Excel y se envía por correo.
"""
import json
import os
import re
from datetime import datetime

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Ellipse, Line, Rectangle, RoundedRectangle
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

NUM_POS = 26
MAX_NORMAL = 6
CAT_MAS_DE_6 = "Mas de 6 numeros por pallet"
CATEGORIAS = [
    "Mas de 6 numeros por pallet",
    "Cambio de TRK por retrazo",
    "Cantidad escrita a mano",
    "Etiquetas mal puestas",
    "Falto capturar",
    "Falto de mover material",
    "Mal capturado",
    "Mal etiquetado",
    "Mal etiquetado y capturado",
    "Material equivocado",
    "Material mal identificado",
    "Otro",
]
TURNOS = ["TURNO 1", "TURNO 2", "TURNO 3"]
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Colores
FONDO = (0.95, 0.96, 0.98, 1)
AZUL = (0.12, 0.31, 0.47, 1)
VERDE = (0.18, 0.62, 0.35, 1)
NARANJA = (0.93, 0.49, 0.13, 1)
GRIS = (0.80, 0.82, 0.86, 1)
ROJO = (0.78, 0.20, 0.20, 1)
OSCURO = (0.15, 0.17, 0.20, 1)

Window.clearcolor = FONDO
Window.softinput_mode = "below_target"


# ---------------------------------------------------------------- utilidades
def boton(texto, fondo=AZUL, accion=None, **kw):
    kw.setdefault("font_size", sp(16))
    kw.setdefault("bold", True)
    b = Button(text=texto, background_normal="", background_color=fondo, **kw)
    if accion:
        b.bind(on_release=lambda *_: accion())
    return b


def campo(hint, **kw):
    kw.setdefault("size_hint_y", None)
    kw.setdefault("height", dp(46))
    kw.setdefault("font_size", sp(17))
    return TextInput(hint_text=hint, multiline=False, write_tab=False,
                     text_validate_unfocus=False, padding=[dp(8), dp(12)], **kw)


def etiqueta(texto, **kw):
    kw.setdefault("color", OSCURO)
    kw.setdefault("font_size", sp(15))
    lb = Label(text=texto, **kw)
    lb.bind(size=lambda w, s: setattr(w, "text_size", (s[0], None)))
    return lb


def enfocar(widget):
    Clock.schedule_once(lambda *_: setattr(widget, "focus", True), 0.05)


def aviso(mensaje, titulo="Aviso"):
    caja = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(12))
    caja.add_widget(Label(text=mensaje, halign="center", font_size=sp(16),
                          text_size=(Window.width * 0.75, None)))
    p = Popup(title=titulo, content=caja, size_hint=(0.9, None), height=dp(240))
    caja.add_widget(boton("OK", accion=p.dismiss, size_hint_y=None, height=dp(48)))
    p.open()


def confirmar(mensaje, al_aceptar, titulo="Confirmar"):
    caja = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(12))
    caja.add_widget(Label(text=mensaje, halign="center", font_size=sp(16),
                          text_size=(Window.width * 0.75, None)))
    fila = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(10))
    p = Popup(title=titulo, content=caja, size_hint=(0.9, None), height=dp(240))

    def si():
        p.dismiss()
        al_aceptar()

    fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
    fila.add_widget(boton("Sí", ROJO, si))
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
    return (txt or "").strip().upper()


# ---------------------------------------------------------------- camión
class Camion(RelativeLayout):
    """Dibujo del tractor + caja con las 26 posiciones (se adapta a vertical u horizontal)."""

    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        self.botones = {}
        self.fondo = Widget(size_hint=(None, None))  # aquí se dibuja el tractor y la caja
        self.add_widget(self.fondo)
        for n in range(1, NUM_POS + 1):
            b = Button(text=str(n), font_size=sp(18), bold=True, background_normal="",
                       background_color=GRIS, color=OSCURO, halign="center", valign="middle",
                       size_hint=(None, None))
            b.bind(on_release=lambda btn, n=n: app.abrir_posicion(n))
            b.bind(size=lambda w, s: setattr(w, "text_size", s))
            self.botones[n] = b
            self.add_widget(b)
        self.bind(size=self._dibujar)

    def _dibujar(self, *_):
        w, h = self.size
        pad = dp(6)
        self.fondo.size = self.size
        self.fondo.canvas.clear()
        horizontal = w > h * 1.3
        with self.fondo.canvas:
            if horizontal:
                cab = w * 0.17
                tx, ty, tw, th = cab + dp(10), pad, w - cab - dp(10) - pad, h - 2 * pad
                th = min(th, tw / 13 * 2 * 1.6)
                ty = (h - th) / 2
                self._tractor_h(pad, ty, cab, th)
            else:
                cab = h * 0.13
                tx, tw = pad, w - 2 * pad
                tw = min(tw, (h - cab) / 13 * 2 * 1.9)
                tx = (w - tw) / 2
                ty, th = pad, h - cab - dp(10) - pad
                self._tractor_v(tx, h - cab - pad, tw, cab)
            # caja
            Color(*OSCURO)
            Line(rectangle=(tx - dp(3), ty - dp(3), tw + dp(6), th + dp(6)), width=dp(1.6))
        g = dp(2)
        if horizontal:
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
        b = self.botones.get(n)
        if b:
            b.pos, b.size = (x, y), (w, h)

    def _tractor_h(self, x, y, w, h):
        cy = y + h / 2
        ch = h * 0.78
        Color(*OSCURO)
        for yy in (cy + ch / 2, cy - ch / 2 - dp(8)):           # llantas
            Rectangle(pos=(x + w * 0.12, yy), size=(w * 0.16, dp(8)))
            Rectangle(pos=(x + w * 0.62, yy), size=(w * 0.22, dp(8)))
        Color(0.30, 0.33, 0.38, 1)
        RoundedRectangle(pos=(x, cy - ch / 2), size=(w * 0.55, ch), radius=[dp(10), 0, 0, dp(10)])
        Color(0.75, 0.86, 0.95, 1)
        RoundedRectangle(pos=(x + w * 0.20, cy - ch * 0.38), size=(w * 0.12, ch * 0.76), radius=[dp(4)])
        Color(0.40, 0.43, 0.48, 1)
        Rectangle(pos=(x + w * 0.57, cy - ch * 0.33), size=(w * 0.30, ch * 0.66))
        Color(*OSCURO)
        Line(points=[x + w * 0.87, cy, x + w + dp(8), cy], width=dp(3))

    def _tractor_v(self, x, y, w, h):
        cx = x + w / 2
        cw = w * 0.70
        Color(*OSCURO)
        for xx in (cx - cw / 2 - dp(8), cx + cw / 2):
            Rectangle(pos=(xx, y + h * 0.72), size=(dp(8), h * 0.16))
            Rectangle(pos=(xx, y + h * 0.12), size=(dp(8), h * 0.22))
        Color(0.30, 0.33, 0.38, 1)
        RoundedRectangle(pos=(cx - cw / 2, y + h * 0.45), size=(cw, h * 0.55), radius=[dp(10), dp(10), 0, 0])
        Color(0.75, 0.86, 0.95, 1)
        RoundedRectangle(pos=(cx - cw * 0.38, y + h * 0.66), size=(cw * 0.76, h * 0.12), radius=[dp(4)])
        Color(0.40, 0.43, 0.48, 1)
        Rectangle(pos=(cx - cw * 0.33, y + h * 0.13), size=(cw * 0.66, h * 0.30))
        Color(*OSCURO)
        Line(points=[cx, y + h * 0.13, cx, y - dp(8)], width=dp(3))

    def refrescar(self):
        pos = self.app.audit["posiciones"]
        con_inc = {str(i.get("posicion")) for i in self.app.audit["incidencias"]}
        for n, b in self.botones.items():
            d = pos.get(str(n))
            lineas = len(d["lineas"]) if d else 0
            if str(n) in con_inc:
                b.background_color, b.color = NARANJA, (1, 1, 1, 1)
            elif lineas:
                b.background_color, b.color = VERDE, (1, 1, 1, 1)
            else:
                b.background_color, b.color = GRIS, OSCURO
            b.text = f"{n} ({lineas})" if lineas else str(n)
            b.font_size = sp(15) if lineas else sp(18)


# ---------------------------------------------------------------- pantallas
class PantallaPrincipal(Screen):
    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        raiz = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(6))

        barra = BoxLayout(size_hint_y=None, height=dp(44))
        with barra.canvas.before:
            Color(*AZUL)
            self._fondo = Rectangle()
        barra.bind(pos=lambda w, p: setattr(self._fondo, "pos", p),
                   size=lambda w, s: setattr(self._fondo, "size", s))
        barra.add_widget(Label(text="AUDITORÍA DE CARGA", bold=True, font_size=sp(18)))
        self.lbl_fecha = Label(text="", font_size=sp(13), size_hint_x=0.5)
        barra.add_widget(self.lbl_fecha)
        raiz.add_widget(barra)

        datos = GridLayout(cols=3, size_hint_y=None, height=dp(46), spacing=dp(6))
        self.trk = campo("TRK / Caja")
        self.auditor = campo("Auditor")
        self.turno = Spinner(text=TURNOS[0], values=TURNOS, size_hint_y=None, height=dp(46),
                             font_size=sp(16), background_normal="", background_color=AZUL)
        self.trk.bind(text=lambda w, t: app.set_dato("trk", t.strip().upper()))
        self.auditor.bind(text=lambda w, t: app.set_dato("auditor", t.strip()))
        self.turno.bind(text=lambda w, t: app.set_dato("turno", t))
        self.trk.bind(on_text_validate=lambda *_: enfocar(self.auditor))
        for wdg in (self.trk, self.auditor, self.turno):
            datos.add_widget(wdg)
        raiz.add_widget(datos)

        self.camion = Camion(app)
        raiz.add_widget(self.camion)

        self.lbl_estado = etiqueta("", size_hint_y=None, height=dp(24), halign="center", font_size=sp(14))
        raiz.add_widget(self.lbl_estado)

        acciones = BoxLayout(size_hint_y=None, height=dp(54), spacing=dp(6))
        acciones.add_widget(boton("Incidencia", NARANJA, lambda: app.popup_incidencia()))
        acciones.add_widget(boton("Generar Excel\ny enviar", VERDE, app.popup_enviar, halign="center"))
        acciones.add_widget(boton("Nueva", ROJO, app.nueva_auditoria, size_hint_x=0.6))
        raiz.add_widget(acciones)
        self.add_widget(raiz)

    def refrescar(self):
        a = self.app.audit
        self.trk.text, self.auditor.text, self.turno.text = a["trk"], a["auditor"], a["turno"]
        self.lbl_fecha.text = datetime.fromisoformat(a["fecha"]).strftime("%d/%m/%Y %H:%M")
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        piezas = sum(len(d["lineas"]) for d in a["posiciones"].values())
        self.lbl_estado.text = (f"Posiciones: {capt}/{NUM_POS}   ·   Números de parte: {piezas}"
                                f"   ·   Incidencias: {len(a['incidencias'])}")
        self.camion.refrescar()

    def on_pre_enter(self, *_):
        self.refrescar()


class PantallaPosicion(Screen):
    def __init__(self, app, **kw):
        super().__init__(**kw)
        self.app = app
        self.num = None
        self.filas = []
        self.extra_nombre = ""

        raiz = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(6))
        barra = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(6))
        barra.add_widget(boton("< Camión", GRIS, self.regresar, color=OSCURO, size_hint_x=0.35))
        self.titulo = Label(text="", bold=True, font_size=sp(20), color=AZUL)
        barra.add_widget(self.titulo)
        raiz.add_widget(barra)

        self.area = campo("Área (teclear)")
        self.area.bind(on_text_validate=lambda *_: self.filas and enfocar(self.filas[0][0]))
        raiz.add_widget(self.area)

        enc = BoxLayout(size_hint_y=None, height=dp(26), spacing=dp(4))
        enc.add_widget(etiqueta("#", size_hint_x=None, width=dp(30), bold=True))
        enc.add_widget(etiqueta("No. de parte (escanear)", bold=True))
        enc.add_widget(etiqueta("Cantidad", size_hint_x=0.45, bold=True))
        enc.add_widget(Label(size_hint_x=None, width=dp(40)))
        raiz.add_widget(enc)

        self.lista = GridLayout(cols=1, spacing=dp(4), size_hint_y=None)
        self.lista.bind(minimum_height=self.lista.setter("height"))
        self.scroll = ScrollView(do_scroll_x=False)
        self.scroll.add_widget(self.lista)
        raiz.add_widget(self.scroll)

        self.lbl_extra = etiqueta("", size_hint_y=None, height=dp(22), color=NARANJA, font_size=sp(13))
        raiz.add_widget(self.lbl_extra)

        acciones = BoxLayout(size_hint_y=None, height=dp(54), spacing=dp(6))
        acciones.add_widget(boton("+ Agregar\nnúmero", AZUL, self.agregar, halign="center"))
        acciones.add_widget(boton("Incidencia", NARANJA,
                                  lambda: app.popup_incidencia(posicion=self.num)))
        acciones.add_widget(boton("Guardar", VERDE, self.guardar))
        raiz.add_widget(acciones)
        self.add_widget(raiz)

    # -- filas
    def _nueva_fila(self, parte="", cant=""):
        idx = len(self.filas) + 1
        fila = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(4))
        extra = idx > MAX_NORMAL
        fila.add_widget(etiqueta(str(idx), size_hint_x=None, width=dp(30), bold=True,
                                 color=NARANJA if extra else OSCURO))
        tp = campo("No. parte")
        tc = campo("Cant.", size_hint_x=0.45, input_type="number")
        tp.text, tc.text = parte, "" if cant in (None, "") else str(cant)
        if extra:
            tp.background_color = tc.background_color = (1, 0.93, 0.86, 1)
        tp.bind(on_text_validate=lambda *_: enfocar(tc))
        tc.bind(on_text_validate=lambda w: self._siguiente(w))
        fila.add_widget(tp)
        fila.add_widget(tc)
        borrar = boton("X", ROJO, size_hint_x=None, width=dp(40))
        borrar.bind(on_release=lambda *_: (setattr(tp, "text", ""), setattr(tc, "text", ""), enfocar(tp)))
        fila.add_widget(borrar)
        self.filas.append((tp, tc))
        self.lista.add_widget(fila)
        return tp

    def _siguiente(self, tc):
        tc.text = "" if limpiar_cantidad(tc.text) is None else str(limpiar_cantidad(tc.text))
        for i, (_, c) in enumerate(self.filas):
            if c is tc and i + 1 < len(self.filas):
                enfocar(self.filas[i + 1][0])
                self.scroll.scroll_to(self.filas[i + 1][0].parent, padding=dp(40))
                return
        tc.focus = False

    def cargar(self, pos):
        self.num = pos
        datos = self.app.audit["posiciones"].get(str(pos), {"area": "", "lineas": []})
        self.titulo.text = f"Posición {pos}"
        self.area.text = datos.get("area", "")
        self.extra_nombre = datos.get("extra_nombre", "")
        self.filas = []
        self.lista.clear_widgets()
        lineas = datos.get("lineas", [])
        for ln in lineas:
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

    def agregar(self):
        if len(self.filas) < MAX_NORMAL:
            enfocar(self._nueva_fila())
            return
        if self.extra_nombre:
            tp = self._nueva_fila()
            enfocar(tp)
            Clock.schedule_once(lambda *_: self.scroll.scroll_to(tp.parent), 0.1)
            return
        self.app.popup_incidencia(posicion=self.num, categoria=CAT_MAS_DE_6,
                                  titulo=f"Posición {self.num}: más de 6 números de parte",
                                  tipo=f"Pallet en posición {self.num} con más de 6 números de parte",
                                  al_guardar=self._autorizar_extra)

    def _autorizar_extra(self, nombre):
        self.extra_nombre = nombre
        self._actualizar_extra()
        self.guardar(salir=False)
        tp = self._nueva_fila()
        enfocar(tp)
        Clock.schedule_once(lambda *_: self.scroll.scroll_to(tp.parent), 0.1)

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
        self.ruta_datos = os.path.join(self.user_data_dir, "auditoria_actual.json")
        self.ruta_config = os.path.join(self.user_data_dir, "config.json")
        self.config_app = self._leer_json(self.ruta_config) or {"correo": "", "auditor": "", "turno": TURNOS[0]}
        self.audit = self._leer_json(self.ruta_datos) or self._audit_vacia()

        self.sm = ScreenManager(transition=NoTransition())
        self.principal = PantallaPrincipal(self, name="principal")
        self.posicion = PantallaPosicion(self, name="posicion")
        self.sm.add_widget(self.principal)
        self.sm.add_widget(self.posicion)
        Window.bind(on_keyboard=self._tecla_atras)
        return self.sm

    # -- datos
    def _audit_vacia(self):
        return {"fecha": datetime.now().isoformat(timespec="seconds"), "trk": "",
                "auditor": self.config_app.get("auditor", ""), "turno": self.config_app.get("turno", TURNOS[0]),
                "posiciones": {}, "incidencias": []}

    @staticmethod
    def _leer_json(ruta):
        try:
            with open(ruta, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None

    def guardar_local(self):
        for ruta, data in ((self.ruta_datos, self.audit), (self.ruta_config, self.config_app)):
            with open(ruta, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=1)

    def set_dato(self, clave, valor):
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
        if key == 27:  # botón atrás de Android
            if self.sm.current == "posicion":
                self.posicion.regresar()
                return True
        return False

    def nueva_auditoria(self):
        def hacer():
            self.audit = self._audit_vacia()
            self.guardar_local()
            self.principal.refrescar()
            enfocar(self.principal.trk)
        confirmar("Se borrará la captura actual (asegúrate de haber enviado el Excel). ¿Continuar?", hacer,
                  "Nueva auditoría")

    # -- incidencias
    def popup_incidencia(self, posicion="", categoria=None, titulo="Registrar incidencia",
                         tipo="", al_guardar=None):
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        nombre = campo("Nombre de quien cargó (obligatorio)")
        cat = Spinner(text=categoria or CATEGORIAS[0], values=CATEGORIAS, size_hint_y=None, height=dp(46),
                      font_size=sp(15), background_normal="", background_color=AZUL)
        desc = TextInput(hint_text="Descripción / comentario", text=tipo, font_size=sp(15),
                         size_hint_y=None, height=dp(90))
        pos = campo("Posición (opcional)", text=str(posicion or ""), input_filter="int")
        for w in (etiqueta("Nombre:", size_hint_y=None, height=dp(20), color=(1, 1, 1, 1)), nombre,
                  etiqueta("Categoría:", size_hint_y=None, height=dp(20), color=(1, 1, 1, 1)), cat,
                  desc, pos):
            caja.add_widget(w)
        fila = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(8))
        p = Popup(title=titulo, content=caja, size_hint=(0.95, None), height=dp(470), auto_dismiss=False)

        def ok():
            if not nombre.text.strip():
                nombre.hint_text = "¡Escribe el nombre!"
                enfocar(nombre)
                return
            self.audit["incidencias"].append({
                "nombre": nombre.text.strip().upper(),
                "tipo": desc.text.strip() or cat.text,
                "categoria": cat.text,
                "fecha": datetime.now().isoformat(timespec="seconds"),
                "posicion": pos.text.strip(),
            })
            self.guardar_local()
            p.dismiss()
            self.principal.refrescar()
            if al_guardar:
                al_guardar(nombre.text.strip().upper())

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Guardar", NARANJA, ok))
        caja.add_widget(fila)
        nombre.bind(on_text_validate=lambda *_: enfocar(desc))
        p.open()
        enfocar(nombre)

    # -- exportar / enviar
    def popup_enviar(self):
        a = self.audit
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        caja = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(8))
        caja.add_widget(etiqueta(
            f"TRK: {a['trk'] or '—'}   Auditor: {a['auditor'] or '—'}   {a['turno']}\n"
            f"Posiciones capturadas: {capt}/{NUM_POS}\nIncidencias: {len(a['incidencias'])}",
            size_hint_y=None, height=dp(70), color=(1, 1, 1, 1)))
        correo = campo("Correo(s) destino, separados por coma", text=self.config_app.get("correo", ""),
                       input_type="mail")
        caja.add_widget(correo)
        fila = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(8))
        p = Popup(title="Generar reporte", content=caja, size_hint=(0.95, None), height=dp(300))

        def enviar():
            if not a["trk"] or not a["auditor"]:
                aviso("Captura el TRK y el nombre del auditor en la pantalla principal.", "Faltan datos")
                return
            self.config_app["correo"] = correo.text.strip()
            self.guardar_local()
            p.dismiss()
            self.exportar_y_enviar()

        fila.add_widget(boton("Cancelar", GRIS, p.dismiss, color=OSCURO))
        fila.add_widget(boton("Generar y enviar", VERDE, enviar))
        caja.add_widget(fila)
        p.open()

    def exportar_y_enviar(self):
        a = self.audit
        nombre = reporte.nombre_archivo(a)
        carpeta = os.path.join(self.user_data_dir, "reportes")
        os.makedirs(carpeta, exist_ok=True)
        ruta = os.path.join(carpeta, nombre)
        try:
            reporte.generar_excel(a, ruta)
        except Exception as e:
            aviso(f"No se pudo generar el Excel:\n{e}", "Error")
            return
        para = [c.strip() for c in self.config_app.get("correo", "").split(",") if c.strip()]
        capt = sum(1 for d in a["posiciones"].values() if d["lineas"])
        asunto = f"Auditoría de carga TRK {a['trk']} - {datetime.now():%d/%m/%Y %H:%M}"
        cuerpo = (f"Se adjunta la auditoría de carga.\n\nTRK: {a['trk']}\nAuditor: {a['auditor']}\n"
                  f"{a['turno']}\nPosiciones capturadas: {capt}/{NUM_POS}\n"
                  f"Incidencias: {len(a['incidencias'])}\n")
        if platform == "android":
            try:
                compartir_android(ruta, nombre, para, asunto, cuerpo)
            except Exception as e:
                aviso(f"El Excel se generó pero no se pudo abrir el correo:\n{e}\n\nArchivo: {ruta}", "Error")
        else:
            aviso(f"Excel generado:\n{ruta}", "Listo")


# ---------------------------------------------------------------- Android
def compartir_android(ruta, nombre, para, asunto, cuerpo):
    """Copia el Excel a Descargas/AuditoriaCarga y abre el selector para enviarlo por correo."""
    from jnius import autoclass, cast
    from android import mActivity  # noqa

    Intent = autoclass("android.content.Intent")
    JString = autoclass("java.lang.String")
    Build = autoclass("android.os.Build$VERSION")

    uri = None
    if Build.SDK_INT >= 29:
        # MediaStore: queda una copia en Descargas y el content:// se puede compartir
        ContentValues = autoclass("android.content.ContentValues")
        Downloads = autoclass("android.provider.MediaStore$Downloads")
        resolver = mActivity.getContentResolver()
        cv = ContentValues()
        cv.put("_display_name", nombre)
        cv.put("mime_type", MIME_XLSX)
        cv.put("relative_path", "Download/AuditoriaCarga")
        uri = resolver.insert(Downloads.EXTERNAL_CONTENT_URI, cv)
        salida = resolver.openOutputStream(uri)
        with open(ruta, "rb") as fh:
            salida.write(bytearray(fh.read()))
        salida.flush()
        salida.close()
    else:
        import shutil
        from android.permissions import Permission, request_permissions
        request_permissions([Permission.WRITE_EXTERNAL_STORAGE])
        Environment = autoclass("android.os.Environment")
        destino = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS).getAbsolutePath()
        os.makedirs(os.path.join(destino, "AuditoriaCarga"), exist_ok=True)
        ruta2 = os.path.join(destino, "AuditoriaCarga", nombre)
        shutil.copy(ruta, ruta2)
        StrictMode = autoclass("android.os.StrictMode")
        StrictMode.setVmPolicy(autoclass("android.os.StrictMode$VmPolicy$Builder")().build())
        uri = autoclass("android.net.Uri").fromFile(autoclass("java.io.File")(ruta2))

    intent = Intent(Intent.ACTION_SEND)
    intent.setType(MIME_XLSX)
    intent.putExtra(Intent.EXTRA_STREAM, cast("android.os.Parcelable", uri))
    intent.putExtra(Intent.EXTRA_SUBJECT, JString(asunto))
    intent.putExtra(Intent.EXTRA_TEXT, JString(cuerpo))
    if para:
        try:
            intent.putExtra(Intent.EXTRA_EMAIL, [JString(c) for c in para])
        except Exception:
            pass
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    mActivity.startActivity(Intent.createChooser(intent, JString("Enviar reporte por correo")))


if __name__ == "__main__":
    AuditoriaApp().run()
