#!/usr/bin/env python3
"""
Ventas mensuales por sucursal (2025 y 2026) desde la API de FUDO.

Uso:
    python3 ventas_mensuales.py              # rápido: salta el historial anterior a 2025
    python3 ventas_mensuales.py --completo   # lee todas las páginas desde la primera

Usa las credenciales de config.py (las mismas de sheets.py y dashboard.py).
Muestra las tablas en pantalla, guarda ventas_mensuales.xlsx (Excel; no hace
falta instalar nada) junto a este archivo y, en Mac, lo abre al terminar.

Criterio (el mismo del dashboard): ventas no canceladas, total de la venta tal
como lo guarda FUDO (con IVA) y fecha de cierre en hora de Chile.
"""

import argparse
import csv
import os
import re
import subprocess
import sys
import time
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from xml.sax.saxutils import escape

import requests

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("America/Santiago")          # maneja solo el horario de verano
    TZ_EXACTA = True
except Exception:
    TZ = timezone(timedelta(hours=-4))
    TZ_EXACTA = False

ANIOS = (2025, 2026)
PAGINA = 500
IVA = 1.19
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


class ErrorFudo(Exception):
    pass


# ── Fechas ───────────────────────────────────────────────────

_FECHA = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})(?:\.\d+)?\s*(Z|[+-]\d{2}:?\d{2})?$")


def a_chile(texto):
    """Fecha ISO 8601 de FUDO (normalmente en UTC) -> datetime en hora de Chile."""
    m = _FECHA.match(str(texto).strip())
    if not m:
        return None
    y, mo, d, h, mi, s = (int(x) for x in m.groups()[:6])
    z = m.group(7)
    if z in (None, "Z"):
        zona = timezone.utc
    else:
        signo = 1 if z[0] == "+" else -1
        zona = timezone(signo * timedelta(hours=int(z[1:3]), minutes=int(z[-2:])))
    try:
        return datetime(y, mo, d, h, mi, s, tzinfo=zona).astimezone(TZ)
    except ValueError:
        return None


def fecha_venta(v):
    a = v.get("attributes") or {}
    for k in ("closedAt", "createdAt", "openedAt"):
        if a.get(k):
            d = a_chile(a[k])
            if d is not None:
                return d
    return None


def fechas_validas(items):
    return [d for d in (fecha_venta(v) for v in items) if d is not None]


def ahora():
    return datetime.now(TZ)


def rango_anios():
    """'2025-2026' (o '2026' si es un solo año)."""
    return f"{ANIOS[0]}-{ANIOS[-1]}" if len(ANIOS) > 1 else str(ANIOS[0])


def texto_anios():
    """'2025 y 2026' (o '2026' si es un solo año)."""
    if len(ANIOS) == 1:
        return str(ANIOS[0])
    return ", ".join(str(a) for a in ANIOS[:-1]) + f" y {ANIOS[-1]}"


# ── Configuración y acceso a FUDO ────────────────────────────

def cargar_config():
    aqui = os.path.dirname(os.path.abspath(__file__))
    for carpeta in (aqui, os.getcwd(),
                    os.path.expanduser("~/Downloads/schwencke_dashboard-2")):
        if os.path.exists(os.path.join(carpeta, "config.py")):
            sys.path.insert(0, carpeta)
            import config
            return (config.SUCURSALES,
                    getattr(config, "FUDO_AUTH_URL", "https://auth.fu.do/api"),
                    getattr(config, "FUDO_BASE_URL", "https://api.fu.do/v1alpha1"))
    sys.exit("No encontré config.py. Copia este archivo a la carpeta "
             "schwencke_dashboard-2 (donde está config.py) y vuelve a ejecutarlo.")


class Fudo:
    """Cliente mínimo: una autenticación por sucursal y reintentos ante fallas."""

    def __init__(self, auth_url, base_url, suc, http=None):
        self.auth_url, self.base_url, self.suc = auth_url, base_url, suc
        self.http = http or requests               # p. ej. requests.Session()
        self.token = None

    def autenticar(self):
        for intento in range(5):
            r = self.http.post(
                self.auth_url,
                json={"apiKey": self.suc["api_key"],
                      "apiSecret": self.suc["api_secret"]},
                headers={"Content-Type": "application/json"}, timeout=30)
            if r.status_code == 429:           # demasiadas autenticaciones seguidas
                time.sleep(20 * (intento + 1))
                continue
            if r.status_code in (401, 403):
                raise ErrorFudo("FUDO rechazó las credenciales "
                                "(revisa la API key y el API secret de esta sucursal)")
            r.raise_for_status()
            self.token = r.json()["token"]
            return
        raise ErrorFudo("FUDO limitó las autenticaciones (429); "
                        "espera unos minutos y vuelve a intentarlo")

    def pagina(self, numero):
        """Ventas de una página de /sales ([] si la página no existe)."""
        reautenticado = False
        ultimo = ""
        for intento in range(6):
            try:
                r = self.http.get(
                    self.base_url + "/sales",
                    headers={"Authorization": "Bearer " + self.token},
                    params={"page[size]": PAGINA, "page[number]": numero},
                    timeout=90)
            except requests.exceptions.RequestException as e:
                ultimo = str(e)
            else:
                if r.status_code == 401 and not reautenticado:
                    self.autenticar()          # el token duró menos de lo esperado
                    reautenticado = True
                    continue
                if r.status_code in (400, 404, 416, 422):
                    return []                  # página fuera de rango
                if r.status_code == 429 or r.status_code >= 500:
                    ultimo = f"HTTP {r.status_code}"
                else:
                    r.raise_for_status()
                    return r.json().get("data") or []
            time.sleep(min(60, 2 ** intento * 2))
        raise ErrorFudo(f"página {numero}: {ultimo}. FUDO no respondió tras "
                        f"varios intentos; vuelve a intentarlo en unos minutos")


# ── Búsqueda de páginas ──────────────────────────────────────

def ultima_pagina(cli):
    """Número de la última página con ventas (0 si no hay ninguna)."""
    if not cli.pagina(1):
        return 0
    lo, paso = 1, 2
    while paso < 1_000_000 and cli.pagina(paso):
        lo, paso = paso, paso * 2
    hi = paso                                  # página vacía
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if cli.pagina(mid):
            lo = mid
        else:
            hi = mid
    return lo


def pagina_inicio(cli, ultima, inicio):
    """Página desde la que hay que leer para no perder ventas >= inicio.

    /sales viene en orden ascendente: busca (búsqueda binaria) la primera
    página cuyas fechas ya llegan a `inicio` y retrocede un margen de seguridad.
    """
    lo, hi, hallada = 1, ultima, ultima
    while lo <= hi:
        mid = (lo + hi) // 2
        fechas = fechas_validas(cli.pagina(mid))
        if not fechas:
            return 1                           # no se puede decidir: leer todo
        if max(fechas) < inicio:
            lo = mid + 1
        else:
            hallada, hi = mid, mid - 1
    pag = max(1, hallada - 2)
    while pag > 1:                             # verificar que el margen alcanzó
        fechas = fechas_validas(cli.pagina(pag))
        if fechas and min(fechas) >= inicio:
            pag = max(1, pag - 3)
        else:
            break
    return pag


# ── Lectura y acumulación ────────────────────────────────────

def leer_sucursal(cli, nombre, completo=False, progreso=None):
    """Suma ventas y tickets por mes (hora de Chile) de una sucursal.

    `progreso` (opcional) es una función que recibe un diccionario con el
    avance; si no se entrega, se imprime una línea de avance en pantalla.
    """
    t0 = time.time()
    inicio = datetime(ANIOS[0], 1, 1, tzinfo=TZ)
    r = {"meses": {}, "leidas": 0, "canceladas": 0, "sin_fecha": 0,
         "total_invalido": 0, "primera": None, "ultima": None,
         "paginas": 0, "saltadas": 0, "total_paginas": 0, "segundos": 0}
    if progreso:
        progreso({"etapa": "buscando"})
    ultima = ultima_pagina(cli)
    r["total_paginas"] = ultima
    if ultima == 0:
        return r

    f1 = fechas_validas(cli.pagina(1))
    fn = fechas_validas(cli.pagina(ultima))
    ascendente = bool(f1 and fn and min(f1) <= max(fn))
    pagina = 1
    if not completo and ascendente and ultima > 4:
        pagina = pagina_inicio(cli, ultima, inicio)
    r["saltadas"] = pagina - 1
    desde = pagina
    if progreso:
        progreso({"etapa": "leyendo", "pagina": pagina, "total": max(ultima, pagina),
                  "desde": desde, "leidas": 0})

    meses = defaultdict(lambda: [0.0, 0])      # "AAAA-MM" -> [ventas, tickets]
    vistos = set()
    while True:
        items = cli.pagina(pagina)
        if not items:
            break
        r["paginas"] += 1
        for v in items:
            vid = v.get("id")
            if vid is not None:
                if vid in vistos:
                    continue
                vistos.add(vid)
            r["leidas"] += 1
            a = v.get("attributes") or {}
            if "cancel" in str(a.get("saleState") or "").lower():
                r["canceladas"] += 1
                continue
            d = fecha_venta(v)
            if d is None:
                r["sin_fecha"] += 1
                continue
            if r["ultima"] is None or d > r["ultima"]:
                r["ultima"] = d
            if d.year in ANIOS:
                try:
                    total = float(a.get("total") or 0)
                except (TypeError, ValueError):
                    total = 0.0
                    r["total_invalido"] += 1
                if r["primera"] is None or d < r["primera"]:
                    r["primera"] = d
                m = meses[d.strftime("%Y-%m")]
                m[0] += total
                m[1] += 1
        if progreso:
            progreso({"etapa": "leyendo", "pagina": pagina, "total": max(ultima, pagina),
                      "desde": desde, "leidas": r["leidas"]})
        else:
            print(f"\r  {nombre:<14} página {pagina}/{max(ultima, pagina)} · "
                  f"{r['leidas']:,} ventas leídas".replace(",", "."),
                  end="", flush=True)
        pagina += 1
    if not progreso:
        print()
    r["meses"] = dict(meses)
    r["segundos"] = time.time() - t0
    return r


# ── Armado de tablas ─────────────────────────────────────────

def mes_actual():
    a = ahora()
    return f"{a.year}-{a.month:02d}"


def lista_meses():
    a = ahora()
    return [f"{y}-{m:02d}" for y in ANIOS for m in range(1, 13)
            if (y, m) <= (a.year, a.month)]


def etiqueta(ym):
    return f"{MESES[int(ym[5:]) - 1]} {ym[:4]}"


def valor(r, ym, idx):
    if r is None:
        return None
    return r["meses"].get(ym, (0, 0))[idx]


def construir_filas(nombres, datos, meses, idx, factor=1.0):
    """Filas (es_total, etiqueta, valores por sucursal, total) con subtotal anual.

    Cada mes se redondea a entero antes de sumar, así lo que se ve cuadra.
    """
    filas = []
    acum = [0] * len(nombres)
    for i, ym in enumerate(meses):
        vals = []
        for n in nombres:
            v = valor(datos[n], ym, idx)
            vals.append(None if v is None else int(round(v * factor)))
        filas.append((False, etiqueta(ym) + ("*" if ym == mes_actual() else ""),
                      vals, total_fila(vals)))
        acum = [None if (a is None or v is None) else a + v
                for a, v in zip(acum, vals)]
        if i == len(meses) - 1 or meses[i + 1][:4] != ym[:4]:
            filas.append((True, f"TOTAL {ym[:4]}", acum, total_fila(acum)))
            acum = [0] * len(nombres)
    return filas


def total_fila(vals):
    return None if any(v is None for v in vals) else sum(vals)


def clp(v):
    return "$" + f"{v:,}".replace(",", ".")


def num(v):
    return f"{v:,}".replace(",", ".")


def celda(v, fmt):
    if v is None:
        return "ERROR"
    return "—" if v == 0 else fmt(v)


def imprimir(titulo, nombres, filas, fmt):
    ancho = 15
    linea = "-" * (11 + ancho * (len(nombres) + 1))
    print(f"\n{titulo}")
    print("Mes".ljust(11) + "".join(n[:ancho - 1].rjust(ancho) for n in nombres)
          + "Total".rjust(ancho))
    print(linea)
    for es_total, etiq, vals, tot in filas:
        if es_total:
            print(linea)
        print(etiq.ljust(11) + "".join(celda(v, fmt).rjust(ancho) for v in vals)
              + celda(tot, fmt).rjust(ancho))
        if es_total:
            print(linea)


# ── Archivos de salida ───────────────────────────────────────

def guardar_csv(carpeta, nombres, datos, meses):
    ruta = os.path.join(carpeta, "ventas_mensuales.csv")
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Mes", "Sucursal", "Ventas con IVA", "Ventas netas", "Tickets"])
        for ym in meses:
            for n in nombres:
                if datos[n] is None:
                    continue
                v, t = valor(datos[n], ym, 0), valor(datos[n], ym, 1)
                w.writerow([ym, n, round(v), round(v / IVA), t])
    return ruta


def _col(n):
    """Número de columna -> letra (1 = A, 27 = AA)."""
    letras = ""
    while n:
        n, resto = divmod(n - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _attr(texto):
    return escape(texto, {'"': "&quot;"})


# Posición de cada estilo dentro de <cellXfs> de _ESTILOS
S_NORMAL, S_ENCAB, S_NUM, S_TOT_TXT, S_TOT_NUM, S_NOTA = range(6)

_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_ESTILOS = (
    _XML + '<styleSheet xmlns="' + _NS + '">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0;-#,##0;&quot;–&quot;"/></numFmts>'
    '<fonts count="4">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>'
    '</fonts>'
    '<fills count="4">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDEEAF1"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="6">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" '
    'applyAlignment="1"><alignment horizontal="center"/></xf>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="164" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>')


def _celda(col, fila, valor, estilo):
    ref = f"{_col(col)}{fila}"
    if isinstance(valor, str):
        return (f'<c r="{ref}" s="{estilo}" t="inlineStr">'
                f'<is><t>{escape(valor)}</t></is></c>')
    return f'<c r="{ref}" s="{estilo}"><v>{valor}</v></c>'


def _xml_hoja(encabezado, filas, nota, seleccionada):
    """XML de una hoja. `filas` = [(es_total, [etiqueta, valor, valor, ...])]."""
    ncol = len(encabezado)
    xml_filas = ['<row r="1">' + "".join(
        _celda(i + 1, 1, t, S_ENCAB) for i, t in enumerate(encabezado)) + "</row>"]
    fila = 1
    for es_total, celdas in filas:
        fila += 1
        partes = []
        for i, v in enumerate(celdas):
            if i == 0:
                estilo = S_TOT_TXT if es_total else S_NORMAL
            else:
                estilo = S_TOT_NUM if es_total else S_NUM
            partes.append(_celda(i + 1, fila, v, estilo))
        xml_filas.append(f'<row r="{fila}">' + "".join(partes) + "</row>")
    fila += 2
    xml_filas.append(f'<row r="{fila}">' + _celda(1, fila, nota, S_NOTA) + "</row>")
    sel = 'tabSelected="1" ' if seleccionada else ""
    datos = "".join(xml_filas)
    return (
        _XML + '<worksheet xmlns="' + _NS + '">'
        f'<dimension ref="A1:{_col(ncol)}{fila}"/>'
        f'<sheetViews><sheetView {sel}workbookViewId="0">'
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
        '<selection pane="topRight"/><selection pane="bottomLeft"/>'
        '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="14" customWidth="1"/>'
        f'<col min="2" max="{ncol}" width="17" customWidth="1"/></cols>'
        f'<sheetData>{datos}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>')


def escribir_xlsx(ruta, hojas):
    """Escribe un .xlsx usando solo la librería estándar (no hace falta openpyxl).

    hojas = [{"nombre", "encabezado": [...], "filas": [(es_total, [celdas])], "nota"}]
    """
    n = len(hojas)
    tipo_hoja = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    contenido = (
        _XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{tipo_hoja}"/>'
                  for i in range(n))
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.styles+xml"/>'
        '</Types>')
    raiz = (
        _XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_REL}/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    libro = (
        _XML + '<workbook xmlns="' + _NS + '" xmlns:r="' + _REL + '">'
        '<bookViews><workbookView/></bookViews><sheets>'
        + "".join(f'<sheet name="{_attr(h["nombre"])}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                  for i, h in enumerate(hojas))
        + '</sheets></workbook>')
    rels_libro = (
        _XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i + 1}" Type="{_REL}/worksheet" '
                  f'Target="worksheets/sheet{i + 1}.xml"/>' for i in range(n))
        + f'<Relationship Id="rId{n + 1}" Type="{_REL}/styles" Target="styles.xml"/>'
        '</Relationships>')
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", contenido)
        z.writestr("_rels/.rels", raiz)
        z.writestr("xl/workbook.xml", libro)
        z.writestr("xl/_rels/workbook.xml.rels", rels_libro)
        z.writestr("xl/styles.xml", _ESTILOS)
        for i, h in enumerate(hojas):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml",
                       _xml_hoja(h["encabezado"], h["filas"], h["nota"], i == 0))


HOJAS = [("Ventas con IVA", 0, 1.0, "Total FUDO con IVA"),
         ("Ventas netas", 0, 1 / IVA, "Total FUDO dividido por 1,19"),
         ("Tickets", 1, 1.0, "Cantidad de ventas")]


def construir_hojas(nombres, datos, meses):
    """Las tres hojas del Excel (con IVA, netas, tickets) listas para escribir."""
    hojas = []
    for titulo, idx, factor, nota in HOJAS:
        filas = []
        for es_total, etiq, vals, tot in construir_filas(nombres, datos, meses, idx, factor):
            filas.append((es_total, [etiq] + ["ERROR" if v is None else v for v in vals]
                          + ["n/d" if tot is None else tot]))
        hojas.append({"nombre": titulo, "encabezado": ["Mes"] + list(nombres) + ["Total"],
                      "filas": filas,
                      "nota": f"{nota} · no canceladas · hora de Chile · * mes en curso (parcial)"})
    return hojas


def guardar(carpeta, nombres, datos, meses):
    """Guarda ventas_mensuales.xlsx (si no se puede, un CSV) y devuelve la ruta."""
    hojas = construir_hojas(nombres, datos, meses)
    ruta = os.path.join(carpeta, "ventas_mensuales.xlsx")
    try:
        escribir_xlsx(ruta, hojas)
        return ruta
    except Exception as e:
        print(f"\nNo pude guardar el Excel ({e}); dejo un CSV en su lugar.")
        return guardar_csv(carpeta, nombres, datos, meses)


def abrir(ruta):
    """En Mac abre el archivo con la aplicación predeterminada (Excel, Numbers…)."""
    if sys.platform == "darwin":
        try:
            subprocess.run(["open", ruta], timeout=15)
        except Exception:
            pass


# ── Programa principal ───────────────────────────────────────

def info_lectura(r, pista="--completo"):
    """(líneas informativas, avisos) de una sucursal ya leída."""
    if r["leidas"] == 0:
        return [], ["FUDO no devolvió ventas para esta sucursal"]
    p = r["primera"].strftime("%d/%m/%Y") if r["primera"] else "—"
    u = r["ultima"].strftime("%d/%m/%Y") if r["ultima"] else "—"
    info = [f"primera venta {rango_anios()}: {p} · última venta: {u}",
            f"{num(r['leidas'])} ventas leídas "
            f"({num(r['canceladas'])} canceladas excluidas) · {r['paginas']} páginas"
            + (f", saltó {r['saltadas']} anteriores a {ANIOS[0]}" if r["saltadas"] else "")
            + f" · {int(r['segundos'])} s"]
    avisos = []
    if r["sin_fecha"]:
        avisos.append(f"{r['sin_fecha']} ventas sin fecha (no se sumaron)")
    if r["total_invalido"]:
        avisos.append(f"{r['total_invalido']} ventas con total ilegible (se contaron como $0)")
    if r["ultima"] and (ahora() - r["ultima"]).days > 2:
        avisos.append(f"la última venta leída es de hace {(ahora() - r['ultima']).days} días: "
                      f"si la sucursal sigue operando puede faltar información "
                      f"(prueba con {pista})")
    return info, avisos


def resumen_cobertura(nombres, datos):
    print("\nCobertura (qué se leyó de cada sucursal):")
    for nom in nombres:
        r = datos[nom]
        if r is None:
            print(f"  {nom:<14} ERROR: no se pudo leer; vuelve a ejecutar el script")
            continue
        info, avisos = info_lectura(r)
        for i, linea in enumerate(info + ["⚠ " + a for a in avisos]):
            print(f"  {nom if i == 0 else '':<14} {linea}")


def main():
    ap = argparse.ArgumentParser(description="Ventas mensuales por sucursal desde FUDO")
    ap.add_argument("--completo", action="store_true",
                    help="leer todas las páginas desde la primera (más lento)")
    args = ap.parse_args()

    sucursales, auth_url, base_url = cargar_config()
    carpeta = os.path.dirname(os.path.abspath(__file__))
    print(f"Ventas mensuales por sucursal — {texto_anios()}")
    if not TZ_EXACTA:
        print("Aviso: no pude cargar la zona horaria de Chile; uso UTC-4 fijo "
              "(solo afecta ventas de medianoche en horario de verano).")
    print("Leyendo ventas desde FUDO (puede tardar varios minutos).\n")

    nombres, datos = [], {}
    for suc in sucursales:
        nombre = suc.get("nombre", "?")
        nombres.append(nombre)
        try:
            cli = Fudo(auth_url, base_url, suc)
            cli.autenticar()
            datos[nombre] = leer_sucursal(cli, nombre, args.completo)
        except Exception as e:
            print(f"\n  ✗ {nombre}: {e}")
            datos[nombre] = None

    if all(v is None for v in datos.values()):
        sys.exit("No se pudo leer ninguna sucursal. Revisa las credenciales "
                 "en config.py y tu conexión a internet.")

    meses = lista_meses()
    imprimir("VENTAS (total FUDO, con IVA)", nombres,
             construir_filas(nombres, datos, meses, 0), clp)
    imprimir("TICKETS (ventas no canceladas)", nombres,
             construir_filas(nombres, datos, meses, 1), num)
    print("\n* mes en curso (parcial)")

    resumen_cobertura(nombres, datos)

    ruta = guardar(carpeta, nombres, datos, meses)
    print(f"\nArchivo guardado: {ruta}")
    if ruta.endswith(".xlsx"):
        abrir(ruta)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nCancelado.")
