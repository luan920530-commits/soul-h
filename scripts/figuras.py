# -*- coding: utf-8 -*-
"""Figuras del informe (matplotlib). Composición con ajuste automático de texto."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import os
import textwrap

AZUL = "#1F3355"
AZUL_M = "#2E4A6B"
CELESTE = "#DCE4EE"
CELESTE2 = "#EDF1F6"
VINO = "#8C1C13"
VERDE = "#1E5B3F"
AMBAR = "#8A6A17"
GRIS = "#444B55"
GRIS_C = "#6B727C"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 8.5,
    "text.color": GRIS,
    "axes.labelcolor": GRIS,
    "xtick.color": GRIS,
    "ytick.color": GRIS,
})

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")


class Lienzo:
    """Lienzo en coordenadas 0-1 con cajas de texto autoajustadas."""

    def __init__(self, w_in, h_in):
        self.fig = plt.figure(figsize=(w_in, h_in))
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, 1)
        self.ax.set_ylim(0, 1)
        self.ax.axis("off")
        self.w_in, self.h_in = w_in, h_in

    def _wrap(self, texto, ancho_frac, fs):
        chars = max(6, int(ancho_frac * self.w_in * 72 * 0.85 / (0.50 * fs)))
        salida = []
        for linea in texto.split("\n"):
            if not linea.strip():
                salida.append("")
            else:
                salida.extend(textwrap.wrap(linea, chars) or [""])
        return salida

    def caja(self, x, y, w, h, texto, fc=CELESTE2, ec=AZUL, tc=AZUL, fs=8.0,
             bold=False, lw=1.0, radio=0.012, it=False):
        fs_max, fit = fs, None
        while fs_max >= 5.3:
            lineas = self._wrap(texto, w, fs_max)
            alto = len(lineas) * fs_max * 1.55 / 72
            if alto <= h * self.h_in * 0.90:
                fit = (fs_max, lineas)
                break
            fs_max -= 0.3
        if fit is None:
            fit = (5.3, self._wrap(texto, w, 5.3))
        fs_f, lineas = fit
        self.ax.add_patch(FancyBboxPatch((x, y), w, h,
                                         boxstyle="round,pad=0.004,rounding_size=%s" % radio,
                                         linewidth=lw, facecolor=fc, edgecolor=ec,
                                         mutation_aspect=self.w_in / self.h_in))
        self.ax.text(x + w / 2, y + h / 2, "\n".join(lineas), ha="center", va="center",
                     fontsize=fs_f, color=tc, linespacing=1.5,
                     fontweight="bold" if bold else "normal",
                     style="italic" if it else "normal")

    def rombo(self, cx, cy, hw, hh, texto, fc="#FBF6E6", ec=AMBAR, tc=AMBAR, fs=7.4):
        self.ax.add_patch(plt.Polygon([[cx, cy + hh], [cx + hw, cy], [cx, cy - hh], [cx - hw, cy]],
                                      closed=True, facecolor=fc, edgecolor=ec, lw=1.0))
        self.ax.text(cx, cy, texto, ha="center", va="center", fontsize=fs,
                     color=tc, fontweight="bold", linespacing=1.45)

    def flecha(self, x1, y1, x2, y2, color=AZUL_M, lw=1.1, ls="-"):
        self.ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                                          mutation_scale=11, linewidth=lw, color=color,
                                          linestyle=ls, shrinkA=0, shrinkB=0))

    def etiqueta(self, x, y, texto, color=VINO, fs=7.0, ha="left"):
        self.ax.text(x, y, texto, fontsize=fs, color=color, fontweight="bold", ha=ha, va="center")

    def texto(self, x, y, t, fs=7.0, color=GRIS_C, it=True, bold=False, ha="left"):
        self.ax.text(x, y, t, fontsize=fs, color=color, ha=ha, va="center",
                     style="italic" if it else "normal", fontweight="bold" if bold else "normal",
                     linespacing=1.4)

    def guardar(self, ruta, dpi=230):
        self.fig.savefig(ruta, dpi=dpi, facecolor="white")
        plt.close(self.fig)


# ---------------------------------------------------------------- Figura 1
def fig_flujo_ibc(ruta):
    L = Lienzo(7.0, 4.75)
    L.texto(0.012, 0.972, "A. TRABAJADOR DEPENDIENTE", fs=8.4, color=VINO, it=False, bold=True)
    L.caja(0.012, 0.740, 0.285, 0.190,
           "Factores que SÍ integran salario (art. 127 CST): básico · horas extras y recargos · comisiones · % sobre ventas · viáticos permanentes · primas habituales",
           fc=CELESTE2, fs=7.4)
    L.caja(0.012, 0.480, 0.285, 0.190,
           "No integran el IBC: auxilio de transporte (art. 128 CST) · viáticos accidentales · dotación · propinas · primas extralegales",
           fc="#FBF6E6", ec=AMBAR, tc=AMBAR, fs=7.4)
    L.caja(0.335, 0.760, 0.295, 0.150, "TOTAL DEVENGADO\n(art. 127 CST)\nfactores salariales del mes",
           fc=CELESTE, fs=8.4, bold=True)
    L.caja(0.335, 0.510, 0.295, 0.130, "SE EXCLUYEN los pagos no salariales",
           fc="#FFFFFF", ec="#B7C0CC", tc=GRIS, fs=7.6)
    L.caja(0.668, 0.720, 0.318, 0.255,
           "IBC\n\nPiso: 1 SMMLV\nTecho: 25 SMMLV\n(art. 18, Ley 100/1993 y art. 3, Dec. 1772/1994)",
           fc=CELESTE2, ec=VINO, tc=VINO, fs=7.5, lw=1.3)
    L.flecha(0.299, 0.835, 0.333, 0.835)
    L.flecha(0.299, 0.575, 0.333, 0.575, color=AMBAR, ls=(0, (3, 2)))
    L.flecha(0.632, 0.835, 0.666, 0.835)

    L.texto(0.012, 0.408, "B. CONTRATISTA INDEPENDIENTE (prestación de servicios personales)",
            fs=8.4, color=VINO, it=False, bold=True)
    L.caja(0.012, 0.185, 0.285, 0.195,
           "Valor mensualizado del contrato SIN IVA (el retardo en el pago no altera la base)",
           fs=7.5)
    L.caja(0.335, 0.185, 0.295, 0.195,
           "× 40%\n\nart. 89, Ley 100/1993 y art. 23, Dec. 1703/2002: el 60% restante son costos",
           fs=7.5, bold=True)
    L.caja(0.668, 0.160, 0.318, 0.248,
           "IBC\n\nPiso: 1 SMMLV · Techo: 25 SMMLV\n\nCotiza el 100%: salud 12,5% · pensión 16% · ARL según clase",
           fc=CELESTE2, ec=VINO, tc=VINO, fs=7.4, lw=1.3)
    L.flecha(0.299, 0.283, 0.333, 0.283)
    L.flecha(0.632, 0.283, 0.666, 0.283)

    L.texto(0.012, 0.115, "Salario integral (art. 132 CST): el IBC es el 70% del total pactado, entre 1 y 25 SMMLV.")
    L.texto(0.012, 0.062, "El IBC excluye los pagos no salariales (auxilio de transporte); estos sí se suman para cesantías y prima.")
    L.texto(0.012, 0.016, "Fuente: elaboración propia (Ley 100 de 1993; Decreto 1703 de 2002; CST, arts. 127, 128 y 132).",
            fs=6.4)
    L.guardar(ruta)


# ---------------------------------------------------------------- Figura 2
def fig_aportes(ruta):
    fig, ax = plt.subplots(figsize=(7.0, 3.05))
    fig.subplots_adjust(left=0.285, right=0.985, top=0.86, bottom=0.20)
    conceptos = ["Salud", "Pensión", "ARL (clase I)", "Caja de\ncompensación", "SENA", "ICBF"]
    # Caso EXTINTORES S C M: persona natural con 3 trabajadores → exoneración del art. 114-1 E.T.
    empleador = [0.0, 12.0, 0.522, 4.0, 0.0, 0.0]
    empleado = [4.0, 4.0, 0.0, 0.0, 0.0, 0.0]
    y = np.arange(len(conceptos))[::-1]

    ax.barh(y, [-v for v in empleador], color=AZUL, height=0.56, label="Empleador")
    ax.barh(y, empleado, color="#B98A4B", height=0.56, label="Trabajador")
    ax.axvline(0, color=GRIS, lw=0.9)

    for yi, v in zip(y, empleador):
        if not v:
            ax.text(-0.18, yi, "exonerado", ha="right", va="center", fontsize=6.8,
                    color=VINO, style="italic")
            continue
        txt = "0,522%" if v == 0.522 else (("%.2f%%" % v).replace(".", ",") if v != int(v)
                                           else "%d,00%%" % v)
        ax.text(-v - 0.18, yi, txt, ha="right", va="center", fontsize=7.2,
                color=AZUL, fontweight="bold")
    for yi, v in zip(y, empleado):
        if v:
            ax.text(v + 0.18, yi, "4,00%", ha="left", va="center", fontsize=7.2, color="#7A571F")
        else:
            ax.text(0.18, yi, "—", ha="left", va="center", fontsize=7.2, color=GRIS)

    ax.set_yticks(y)
    ax.set_yticklabels(conceptos, fontsize=7.6)
    ax.set_xlim(-16.2, 5.6)
    ax.set_xlabel("% del IBC", fontsize=7.2, labelpad=2)
    ax.set_xticks([-15, -12, -9, -6, -3, 0, 3])
    ax.set_xticklabels(["15", "12", "9", "6", "3", "0", "3"], fontsize=7)
    ax.set_ylim(-0.62, 5.55)
    ax.tick_params(axis="y", length=0)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="x", color="#E6EAEF", lw=0.7)
    ax.set_axisbelow(True)
    ax.legend(loc="lower right", fontsize=6.6, frameon=False, ncol=2,
              bbox_to_anchor=(1.0, -0.30))
    ax.set_title("Cotización mensual al SGSSS en EXTINTORES S C M con la exoneración del art. 114-1 E.T.",
                 fontsize=8.2, color=AZUL, fontweight="bold", loc="left", pad=8)

    fig.text(0.005, -0.035,
             "Nota: la exoneración del art. 114-1 del E.T. (SENA, ICBF y salud) no aplica a personas naturales que "
             "empleen menos de dos trabajadores; EXTINTORES S C M tiene tres y sí le aplica. Los aportes a pensión "
             "y a ARL nunca están exonerados.",
             fontsize=6.0, color=VINO, style="italic")

    fig.savefig(ruta, dpi=230, facecolor="white", bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- Figura 3
def fig_mapa_pensional(ruta, pm=2, rais=1):
    fig = plt.figure(figsize=(7.0, 2.85))
    ax1 = fig.add_axes([0.015, 0.05, 0.30, 0.82])
    ax2 = fig.add_axes([0.515, 0.12, 0.47, 0.70])
    vals = [pm, rais]
    cols = [AZUL, "#B98A4B"]
    ax1.pie(vals, colors=cols, startangle=90, counterclock=False,
            wedgeprops=dict(width=0.34, edgecolor="white", linewidth=1.6))
    ax1.text(0, 0.12, "%d" % (pm + rais), ha="center", va="center", fontsize=16,
             color=AZUL, fontweight="bold")
    ax1.text(0, -0.20, "afiliados", ha="center", va="center", fontsize=6.8, color=GRIS)
    ax1.set_aspect("equal")
    ax1.set_title("Planta de personal", fontsize=8, color=AZUL, fontweight="bold", pad=2)

    cats = ["Prima Media\n(Colpensiones)", "Ahorro Individual\n(AFP)"]
    counts = [pm, rais]
    y = [1, 0]
    ax2.barh(y, counts, color=cols, height=0.40)
    for yi, c in zip(y, counts):
        ax2.text(c + 0.15, yi, "%d  ·  %.0f%%" % (c, 100.0 * c / (pm + rais)), va="center",
                 fontsize=7.6, color=GRIS, fontweight="bold")
    ax2.set_yticks(y)
    ax2.set_yticklabels(cats, fontsize=7.4)
    ax2.set_xlim(0, max(counts) * 1.55)
    ax2.set_ylim(-0.55, 1.55)
    ax2.set_xticks([])
    ax2.tick_params(axis="y", length=0)
    for s in ("top", "right", "bottom", "left"):
        ax2.spines[s].set_visible(False)
    ax2.set_title("Distribución por régimen pensional",
                  fontsize=8, color=AZUL, fontweight="bold", loc="left", pad=6)

    fig.text(0.02, 0.012, "Fuente: elaboración propia con base en la planta de personal y los comprobantes de nómina "
                          "y planilla PILA del periodo analizado.",
             fontsize=6.2, color=GRIS_C, style="italic")
    fig.savefig(ruta, dpi=230, facecolor="white", bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------- Figura 4
def fig_ruta_at(ruta):
    L = Lienzo(7.0, 5.55)
    L.caja(0.235, 0.925, 0.53, 0.062, "Accidente o enfermedad con ocasión del trabajo",
           fc=CELESTE, bold=True, fs=8.0)
    L.flecha(0.50, 0.923, 0.50, 0.893)

    L.rombo(0.50, 0.842, 0.225, 0.050, "¿Origen laboral acreditado?")
    L.flecha(0.272, 0.842, 0.208, 0.842)
    L.etiqueta(0.240, 0.862, "NO")
    L.caja(0.005, 0.782, 0.200, 0.115,
           "Enfermedad general\nEPS: 66,67% (días 3 a 90) y 50% (días 91 a 180)\nEmpleador: días 1 y 2",
           fc="#FBEDEA", ec=VINO, tc=VINO, fs=6.9)

    L.flecha(0.50, 0.790, 0.50, 0.762)
    L.etiqueta(0.512, 0.776, "SÍ", color=VERDE)
    L.caja(0.150, 0.660, 0.700, 0.098,
           "ARL · prestaciones asistenciales (art. 5, Dec. 1295/1994) y económicas: subsidio por incapacidad temporal (100% del SBC), indemnización por IPP (art. 38) o pensión de invalidez",
           fc=CELESTE2, fs=7.2)

    L.flecha(0.50, 0.658, 0.50, 0.630)
    L.rombo(0.50, 0.575, 0.265, 0.055, "¿Falla del SG-SST imputable al empleador?")
    L.flecha(0.233, 0.575, 0.172, 0.575)
    L.etiqueta(0.200, 0.596, "NO", color=GRIS)
    L.caja(0.005, 0.512, 0.164, 0.126,
           "Reparación tarifada exclusiva: la ARL asume el riesgo y el empleador no responde",
           fc="#FFFFFF", ec="#B7C0CC", tc=GRIS, fs=6.9)

    L.flecha(0.50, 0.518, 0.50, 0.485)
    L.etiqueta(0.512, 0.500, "SÍ")
    L.caja(0.130, 0.375, 0.740, 0.100,
           "Art. 216 CST · indemnización total y ordinaria de perjuicios (culpa leve, art. 63 C.C.; carga probatoria del trabajador)",
           fc="#FBEDEA", ec=VINO, tc=VINO, fs=7.4, bold=True)

    L.flecha(0.50, 0.373, 0.50, 0.342)
    L.caja(0.115, 0.242, 0.770, 0.092,
           "Del monto se descuentan las prestaciones en dinero pagadas por la ARL (inciso final del art. 216 CST) · prescripción: 3 años (art. 151 CPTSS)",
           fc="#FFFFFF", ec="#B7C0CC", tc=GRIS, fs=7.2)

    L.texto(0.012, 0.185, "La culpa patronal no se presume: el trabajador debe probar el daño, la omisión de los deberes de")
    L.texto(0.012, 0.140, "seguridad y el nexo causal. Acreditada la omisión, la carga se traslada al empleador (diligencia).")
    L.texto(0.012, 0.062, "Fuente: elaboración propia (Decreto 1295 de 1994; Ley 1562 de 2012; Decreto 2943 de 2013; art. 216 CST; CPTSS).",
            fs=6.4)
    L.guardar(ruta)


# ---------------------------------------------------------------- Figura 5
def fig_reparacion(ruta, tarifada, plena, neto):
    fig, ax = plt.subplots(figsize=(7.0, 2.55))
    fig.subplots_adjust(left=0.05, right=0.985, top=0.80, bottom=0.24)
    labels = ["Reparación tarifada\nARL (sin culpa)", "Reparación plena\nart. 216 CST",
              "Mayor valor a cargo\ndel empleador"]
    vals = [tarifada, plena, neto]
    cols = [AZUL_M, VINO, AMBAR]
    x = np.arange(3)
    ax.bar(x, vals, color=cols, width=0.52)
    for xi, v in zip(x, vals):
        ax.text(xi, v * 1.015, "$ %s" % f"{v:,.0f}".replace(",", "."), ha="center",
                va="bottom", fontsize=7.4, color=GRIS, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=7.3)
    ax.set_yticks([])
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.set_ylim(0, max(vals) * 1.16)
    ax.set_title("Caso hipotético AT-2026-014: escenarios de reparación (COP)",
                 fontsize=8.2, color=AZUL, fontweight="bold", loc="left", pad=8)
    ax.text(0.0, -0.26, "Cifras ilustrativas de la liquidación del Eje II, numeral 2.6; no constituyen tasación pericial.",
            transform=ax.transAxes, fontsize=6.3, color=GRIS_C, style="italic")
    fig.savefig(ruta, dpi=230, facecolor="white", bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    fig_flujo_ibc(os.path.join(OUT, "f1_ibc.png"))
    fig_aportes(os.path.join(OUT, "f2_aportes.png"))
    fig_mapa_pensional(os.path.join(OUT, "f3_mapa_pensional.png"))
    fig_ruta_at(os.path.join(OUT, "f4_ruta_at.png"))
    fig_reparacion(os.path.join(OUT, "f5_reparacion.png"), 9044000, 27284000, 18240000)
    print("figuras en", OUT)
