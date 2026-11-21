#!/usr/bin/env python3
"""
Ventas mensuales por sucursal (2025 y 2026) desde la API de FUDO.

Uso:
    python3 ventas_mensuales.py                  # rápido: salta el historial anterior a 2025
    python3 ventas_mensuales.py --completo       # lee todas las páginas desde la primera
    python3 ventas_mensuales.py --sin-productos  # solo ventas mensuales (sin ítems; más rápido)

Usa las credenciales de config.py (las mismas de sheets.py y dashboard.py).
Muestra las tablas en pantalla, guarda ventas_mensuales.xlsx (Excel; no hace
falta instalar nada) junto a este archivo y, en Mac, lo abre al terminar.
Además de las ventas mensuales trae las ventas por producto y sucursal de cada año
(hojas «Productos AAAA» y «Detalle productos»).

Criterio (el mismo del dashboard): ventas no canceladas, total de la venta tal
como lo guarda FUDO (con IVA) y fecha de cierre en hora de Chile. Por producto:
cantidad y precio de los ítems no anulados de esas mismas ventas.
"""

import argparse
import csv
import math
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
# Subítems (modificadores o componentes de combos): "incluir" los suma como cualquier otro
# ítem de la venta; "excluir" cuenta solo los ítems principales. Se cambia con la variable
# de entorno PRODUCTOS_SUBITEMS (en Railway: servicio → Variables).
SUBITEMS = os.environ.get("PRODUCTOS_SUBITEMS", "incluir").strip().lower() != "excluir"
# Precio de los ítems: por defecto se toma como precio unitario (ventas = precio × cantidad).
# Con PRODUCTOS_PRECIO=linea se toma como el total de la línea (ventas = precio). Solo hace
# falta si el «cuadre» con el total de las ventas queda muy por encima de 100 %.
PRECIO_LINEA = os.environ.get("PRODUCTOS_PRECIO", "unitario").strip().lower() in ("linea", "línea", "total")
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

    def _pedir(self, ruta, params, que, timeout=90):
        """GET con reintentos. Devuelve el JSON, o None si la página no existe."""
        reautenticado = False
        ultimo = ""
        for intento in range(6):
            try:
                r = self.http.get(
                    self.base_url + ruta,
                    headers={"Authorization": "Bearer " + self.token},
                    params=params, timeout=timeout)
            except requests.exceptions.RequestException as e:
                ultimo = str(e)
            else:
                if r.status_code == 401 and not reautenticado:
                    self.autenticar()          # el token duró menos de lo esperado
                    reautenticado = True
                    continue
                if r.status_code in (400, 404, 416, 422):
                    return None                # página fuera de rango
                if r.status_code == 429 or r.status_code >= 500:
                    ultimo = f"HTTP {r.status_code}"
                else:
                    r.raise_for_status()
                    return r.json()
            time.sleep(min(60, 2 ** intento * 2))
        raise ErrorFudo(f"{que}: {ultimo}. FUDO no respondió tras "
                        f"varios intentos; vuelve a intentarlo en unos minutos")

    def pagina(self, numero):
        """Ventas de una página de /sales ([] si la página no existe)."""
        j = self._pedir("/sales", {"page[size]": PAGINA, "page[number]": numero},
                        f"página {numero}")
        return (j or {}).get("data") or []

    def pagina_con_items(self, numero):
        """(ventas, {id: ítem}) de una página de /sales pidiendo también sus ítems."""
        j = self._pedir("/sales", {"page[size]": PAGINA, "page[number]": numero,
                                   "include": "items"}, f"página {numero}", timeout=180)
        j = j or {}
        items = {str(x["id"]): x for x in (j.get("included") or [])
                 if isinstance(x, dict) and x.get("type") == "Item" and x.get("id") is not None}
        return j.get("data") or [], items

    def productos(self):
        """{id de producto: nombre} leído de /products."""
        nombres, vistos = {}, set()
        for numero in range(1, 201):
            j = self._pedir("/products", {"page[size]": PAGINA, "page[number]": numero},
                            f"productos, página {numero}")
            nuevos = [p for p in (j or {}).get("data") or []
                      if isinstance(p, dict) and p.get("id") is not None
                      and str(p["id"]) not in vistos]
            if not nuevos:
                break
            for p in nuevos:
                vistos.add(str(p["id"]))
                a = p.get("attributes")
                nombre = " ".join(str(a.get("name") or "").split()) if isinstance(a, dict) else ""
                if nombre:
                    nombres[str(p["id"])] = nombre
        return nombres


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

def _num(x):
    """float(x) si es un número finito (acepta texto numérico); si no, None."""
    if x is None or isinstance(x, bool):
        return None
    try:
        n = float(x)
    except (TypeError, ValueError):
        return None
    return n if math.isfinite(n) else None


def _rel(obj, nombre):
    rels = obj.get("relationships") if isinstance(obj, dict) else None
    return rels.get(nombre) if isinstance(rels, dict) else None


def _ids(rel):
    """Ids de una relación JSON:API ({"data": [{"id": ...}]}); [] si no hay."""
    data = rel.get("data") if isinstance(rel, dict) else None
    if not isinstance(data, list):
        return []
    return [str(x["id"]) for x in data if isinstance(x, dict) and x.get("id") is not None]


def lineas_de_venta(v, items):
    """([(ítem, es_subítem)], faltantes) de una venta.

    Parte de los ítems de la venta (relationships.items) y suma los subítems que cuelgan de
    ellos (relationships.subitems), aunque FUDO no los liste en la venta. `items` es el
    {id: ítem} que FUDO entregó junto con la página. `faltantes` cuenta los ids que la venta
    o un ítem nombran pero FUDO no entregó.
    """
    ids = list(dict.fromkeys(_ids(_rel(v, "items"))))
    hijos = []
    for i in ids:
        if i in items:
            hijos += [h for h in _ids(_rel(items[i], "subitems")) if h not in ids and h not in hijos]
    marcados = set(hijos)
    for i in ids:                              # un subítem también puede venir listado en la venta
        if i in items:
            marcados.update(_ids(_rel(items[i], "subitems")))
    todos = ids + hijos
    lineas = [(items[i], i in marcados) for i in todos if i in items]
    return lineas, len(todos) - len(lineas)


def sumar_items(r, ym, v, mapa):
    """Suma los ítems de una venta válida en r["prod"][mes][id_producto] = [unidades, ventas]."""
    lineas, faltan = lineas_de_venta(v, mapa)
    r["items_faltantes"] += faltan
    if not lineas and not faltan:
        r["ventas_sin_items"] += 1
        return
    acum = r["prod"].setdefault(ym, {})
    for it, es_sub in lineas:
        a = it.get("attributes")
        a = a if isinstance(a, dict) else {}
        if a.get("canceled"):
            r["items_anulados"] += 1
            continue
        if es_sub and not SUBITEMS:
            r["subitems_excluidos"] += 1
            continue
        cantidad = _num(a.get("quantity"))
        if cantidad is None:
            cantidad = 1.0                     # sin cantidad: una unidad
        precio = _num(a.get("price"))
        if precio is None:
            r["items_sin_precio"] += 1
            precio = 0.0
        prod = _rel(it, "product")
        prod = prod.get("data") if isinstance(prod, dict) else None
        pid = str(prod["id"]) if isinstance(prod, dict) and prod.get("id") is not None else ""
        if not pid:
            r["items_sin_producto"] += 1
        monto = precio if PRECIO_LINEA else cantidad * precio
        fila = acum.setdefault(pid, [0.0, 0.0])
        fila[0] += cantidad
        fila[1] += monto
        r["items"] += 1
        if es_sub:
            r["subitems"] += 1
            r["subitems_ventas"] += monto


def leer_sucursal(cli, nombre, completo=False, progreso=None, productos=False):
    """Suma ventas y tickets por mes (hora de Chile) de una sucursal.

    Con productos=True lee además los ítems de cada venta (include=items) y acumula
    unidades y ventas por producto y mes en r["prod"]; los nombres vienen de /products.

    `progreso` (opcional) es una función que recibe un diccionario con el
    avance; si no se entrega, se imprime una línea de avance en pantalla.
    """
    t0 = time.time()
    inicio = datetime(ANIOS[0], 1, 1, tzinfo=TZ)
    r = {"meses": {}, "leidas": 0, "canceladas": 0, "sin_fecha": 0,
         "total_invalido": 0, "primera": None, "ultima": None,
         "paginas": 0, "saltadas": 0, "total_paginas": 0, "segundos": 0,
         "prod_activo": bool(productos), "prod": {}, "nombres_prod": {}, "nombres_error": "",
         "items": 0, "items_anulados": 0, "items_sin_precio": 0, "items_sin_producto": 0,
         "items_faltantes": 0, "ventas_sin_items": 0, "subitems": 0,
         "subitems_excluidos": 0, "subitems_ventas": 0.0}
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
    if productos:
        if progreso:
            progreso({"etapa": "productos"})
        try:                                   # sin nombres el reporte sale igual (con «Producto #id»)
            r["nombres_prod"] = cli.productos()
        except (ErrorFudo, requests.exceptions.RequestException) as e:
            r["nombres_error"] = str(e) if isinstance(e, ErrorFudo) else type(e).__name__
    if progreso:
        progreso({"etapa": "leyendo", "pagina": pagina, "total": max(ultima, pagina),
                  "desde": desde, "leidas": 0})

    meses = defaultdict(lambda: [0.0, 0])      # "AAAA-MM" -> [ventas, tickets]
    vistos = set()
    while True:
        if productos:
            items, mapa = cli.pagina_con_items(pagina)
        else:
            items, mapa = cli.pagina(pagina), {}
        if not items:
            if pagina <= ultima:               # FUDO ya había entregado esta página: no es el final
                raise ErrorFudo(
                    f"página {pagina}: FUDO no la entregó"
                    + (" al pedir las ventas con sus ítems. Si se repite, genera el reporte "
                       "sin productos (opción --sin-productos o casilla «Incluir ventas por "
                       "producto»)" if productos else "; vuelve a intentarlo en unos minutos"))
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
                total = _num(a.get("total") or 0)
                if total is None:
                    total = 0.0
                    r["total_invalido"] += 1
                if r["primera"] is None or d < r["primera"]:
                    r["primera"] = d
                ym = d.strftime("%Y-%m")
                m = meses[ym]
                m[0] += total
                m[1] += 1
                if productos:
                    sumar_items(r, ym, v, mapa)
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


# ── Productos ────────────────────────────────────────────────

def pct(parte):
    """0.1234 -> '12,3 %'."""
    return f"{parte * 100:.1f}".replace(".", ",") + " %"


def fmt_unid(u):
    """Unidades: enteras con punto de miles; las fraccionadas (kilos) con dos decimales."""
    if float(u).is_integer():
        return num(int(u))
    return f"{u:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _unid(u):
    u = round(u, 2)
    return int(u) if float(u).is_integer() else u


def unir_productos(nombres, datos):
    """Une los productos de todas las sucursales por nombre (sin distinguir mayúsculas ni espacios).

    Devuelve ({(mes, clave, sucursal): [unidades, ventas]}, {clave: nombre para mostrar}).
    Un producto sin nombre en FUDO (por ejemplo, uno eliminado) se muestra como
    «Producto #id (sucursal)» y no se junta con ningún otro.
    """
    tabla, mostrar = {}, {}
    for suc in nombres:
        r = datos.get(suc)
        if not r or not r.get("prod"):
            continue
        conocidos = r.get("nombres_prod") or {}
        for ym, prods in r["prod"].items():
            for pid, (u, v) in prods.items():
                if not pid:
                    clave, etiq = "\0sin-producto", "(sin producto)"
                elif pid in conocidos:
                    etiq = conocidos[pid]
                    clave = " ".join(etiq.split()).casefold()
                else:
                    clave, etiq = f"\0{suc}\0{pid}", f"Producto #{pid} ({suc})"
                mostrar.setdefault(clave, etiq)
                fila = tabla.setdefault((ym, clave, suc), [0.0, 0.0])
                fila[0] += u
                fila[1] += v
    return tabla, mostrar


def filas_productos(nombres, datos, anio):
    """Productos de un año, unidos entre sucursales y ordenados por ventas (de mayor a menor).

    {"filas": [(producto, por_suc, unid, ventas, neto, parte)],
     "total": (por_suc, unid, ventas, neto), "falla": bool}
    con por_suc = [(unidades, ventas) | None] (None = la sucursal dio ERROR). Si alguna
    sucursal falló, los totales son None. Lo que se ve cuadra: cada total suma las celdas
    ya redondeadas (ventas en pesos enteros).
    """
    tabla, mostrar = unir_productos(nombres, datos)
    pos = {n: j for j, n in enumerate(nombres)}
    ok = [datos.get(n) is not None for n in nombres]
    falla = not all(ok)
    acum = {}
    for (ym, clave, suc), (u, v) in tabla.items():
        if ym.startswith(f"{anio}-"):
            fila = acum.setdefault(clave, [[0.0, 0.0] for _ in nombres])
            fila[pos[suc]][0] += u
            fila[pos[suc]][1] += v
    crudas = []
    for clave, celdas in acum.items():
        por_suc = [(_unid(u), int(round(v))) if ok[j] else None
                   for j, (u, v) in enumerate(celdas)]
        buenas = [c for c in por_suc if c is not None]
        if any(c[0] or c[1] for c in buenas):          # sin movimiento: no se muestra
            crudas.append((mostrar[clave], por_suc, _unid(sum(c[0] for c in buenas)),
                           sum(c[1] for c in buenas)))
    crudas.sort(key=lambda f: (-f[3], f[0].casefold()))
    tot_v = sum(f[3] for f in crudas)
    tot_u = _unid(sum(f[2] for f in crudas))
    tot_suc = [None if not ok[j] else (_unid(sum(f[1][j][0] for f in crudas)),
                                       sum(f[1][j][1] for f in crudas))
               for j in range(len(nombres))]
    filas = []
    for nombre, por_suc, u, v in crudas:
        if falla:
            filas.append((nombre, por_suc, None, None, None, None))
        else:
            filas.append((nombre, por_suc, u, v, int(round(v / IVA)), (v / tot_v) if tot_v else 0.0))
    tot_neto = sum(f[4] for f in filas) if not falla else None
    return {"filas": filas, "falla": falla,
            "total": (tot_suc, None if falla else tot_u, None if falla else tot_v, tot_neto)}


def filas_detalle(nombres, datos):
    """[(mes, sucursal, producto, unidades, ventas)] ordenadas por mes, sucursal y ventas."""
    tabla, mostrar = unir_productos(nombres, datos)
    pos = {n: j for j, n in enumerate(nombres)}
    filas = []
    for (ym, clave, suc), (u, v) in tabla.items():
        u2, v2 = _unid(u), int(round(v))
        if u2 or v2:
            filas.append((ym, suc, mostrar[clave], u2, v2))
    filas.sort(key=lambda f: (f[0], pos[f[1]], -f[4], f[2].casefold()))
    return filas


def cuadre_sucursal(r, prefijo=""):
    """(ventas según los ítems, ventas según el total de cada venta); None si no hay productos."""
    if not r or not r.get("prod_activo"):
        return None
    items = sum(v for ym, d in r["prod"].items() if ym.startswith(prefijo)
                for _, v in d.values())
    ventas = sum(m[0] for ym, m in r["meses"].items() if ym.startswith(prefijo))
    return items, ventas


def cuadre_total(nombres, datos, anio):
    """Lo mismo sumando todas las sucursales; None si alguna dio ERROR o no hay productos."""
    partes = [cuadre_sucursal(datos.get(n), f"{anio}-") for n in nombres]
    if not partes or any(p is None for p in partes):
        return None
    return sum(p[0] for p in partes), sum(p[1] for p in partes)


def texto_cuadre(c):
    items, ventas = c
    if ventas <= 0:
        return f"ítems {clp(int(round(items)))} · ventas $0"
    return (f"ítems {clp(int(round(items)))} vs ventas {clp(int(round(ventas)))} "
            f"({pct(items / ventas)})")


def anios_con_datos():
    """Años del reporte que ya empezaron (los futuros no tienen ventas)."""
    return sorted({ym[:4] for ym in lista_meses()})


def imprimir_productos(nombres, datos, top=10):
    for anio in anios_con_datos():
        d = filas_productos(nombres, datos, anio)
        if not d["filas"]:
            continue
        print(f"\nPRODUCTOS MÁS VENDIDOS {anio} (top {top} por ventas con IVA; "
              f"el Excel trae los {len(d['filas'])})")
        print("Producto".ljust(34) + "Unidades".rjust(12) + "Ventas".rjust(16))
        print("-" * 62)
        for nombre, _, u, v, _, _ in d["filas"][:top]:
            print(nombre[:33].ljust(34) + ("n/d" if u is None else fmt_unid(u)).rjust(12)
                  + ("n/d" if v is None else clp(v)).rjust(16))
        c = cuadre_total(nombres, datos, anio)
        if c:
            print(f"Cuadre con las ventas {anio}: {texto_cuadre(c)}")


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
S_DEC, S_TOT_DEC, S_PCT, S_TOT_PCT = range(6, 10)    # decimales y porcentaje

_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_ESTILOS = (
    _XML + '<styleSheet xmlns="' + _NS + '">'
    '<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0;-#,##0;&quot;–&quot;"/>'
    '<numFmt numFmtId="165" formatCode="#,##0.00;-#,##0.00;&quot;–&quot;"/>'
    '<numFmt numFmtId="166" formatCode="0.0%"/></numFmts>'
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
    '<cellXfs count="10">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" '
    'applyAlignment="1"><alignment horizontal="center"/></xf>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="164" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="165" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyAlignment="1"><alignment horizontal="right"/></xf>'
    '<xf numFmtId="166" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" '
    'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>')


def _celda(col, fila, valor, estilo):
    ref = f"{_col(col)}{fila}"
    if isinstance(valor, str):
        return (f'<c r="{ref}" s="{estilo}" t="inlineStr">'
                f'<is><t>{escape(valor)}</t></is></c>')
    return f'<c r="{ref}" s="{estilo}"><v>{valor}</v></c>'


def _xml_hoja(encabezado, filas, nota, seleccionada, anchos=None, congelar_col=True):
    """XML de una hoja. `filas` = [(es_total, [etiqueta, valor, valor, ...])].

    Un valor puede ser (valor, estilo) para fijar su formato. `nota` es un texto o una lista
    de textos (una línea cada uno). `anchos` = ancho de cada columna (por defecto 14 y 17).
    """
    ncol = len(encabezado)
    xml_filas = ['<row r="1">' + "".join(
        _celda(i + 1, 1, t, S_ENCAB) for i, t in enumerate(encabezado)) + "</row>"]
    fila = 1
    for es_total, celdas in filas:
        fila += 1
        partes = []
        for i, v in enumerate(celdas):
            if isinstance(v, tuple):
                v, estilo = v
            elif i == 0:
                estilo = S_TOT_TXT if es_total else S_NORMAL
            else:
                estilo = S_TOT_NUM if es_total else S_NUM
            partes.append(_celda(i + 1, fila, v, estilo))
        xml_filas.append(f'<row r="{fila}">' + "".join(partes) + "</row>")
    fila += 1                                  # una fila en blanco antes de las notas
    for linea in ([nota] if isinstance(nota, str) else list(nota)):
        fila += 1
        xml_filas.append(f'<row r="{fila}">' + _celda(1, fila, linea, S_NOTA) + "</row>")
    sel = 'tabSelected="1" ' if seleccionada else ""
    datos = "".join(xml_filas)
    if congelar_col:
        panel = ('<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
                 '<selection pane="topRight"/><selection pane="bottomLeft"/>'
                 '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>')
    else:
        panel = ('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                 '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
    if anchos:
        cols = "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
                       for i, w in enumerate(anchos))
    else:
        cols = ('<col min="1" max="1" width="14" customWidth="1"/>'
                f'<col min="2" max="{ncol}" width="17" customWidth="1"/>')
    return (
        _XML + '<worksheet xmlns="' + _NS + '">'
        f'<dimension ref="A1:{_col(ncol)}{fila}"/>'
        f'<sheetViews><sheetView {sel}workbookViewId="0">{panel}</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{cols}</cols>'
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
                       _xml_hoja(h["encabezado"], h["filas"], h["nota"], i == 0,
                                 h.get("anchos"), h.get("congelar_col", True)))


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


def _c_unid(u, total=False):
    """Celda de unidades: entera con miles; fraccionada (kilos) con dos decimales."""
    return (u, S_TOT_DEC if total else S_DEC) if isinstance(u, float) else u


def construir_hojas_productos(nombres, datos):
    """Hojas de productos: una por año (producto × sucursal) y el detalle mes a mes."""
    caidas = [n for n in nombres if datos.get(n) is None]
    aviso = f"No se pudo leer: {', '.join(caidas)} (columnas en ERROR)" if caidas else None
    cab_suc = [t for n in nombres for t in (f"{n} unid.", f"{n} ventas")]
    hojas = []
    for anio in anios_con_datos():
        d = filas_productos(nombres, datos, anio)
        tot_suc, tot_u, tot_v, tot_n = d["total"]

        def fila(etiqueta, por_suc, u, v, n, parte, total):
            celdas = [etiqueta]
            for par in por_suc:
                celdas += ["ERROR", "ERROR"] if par is None else [_c_unid(par[0], total), par[1]]
            if u is None:
                return celdas + ["n/d"] * 4
            return celdas + [_c_unid(u, total), v, n, (parte, S_TOT_PCT if total else S_PCT)]

        filas = [(True, fila(f"TOTAL {anio} · {len(d['filas'])} productos", tot_suc, tot_u,
                             tot_v, tot_n, 1.0 if tot_v else 0.0, True))]
        filas += [(False, fila(nom, ps, u, v, n, p, False)) for nom, ps, u, v, n, p in d["filas"]]
        nota = ["Ventas con IVA = precio × cantidad de los ítems no anulados de las ventas no "
                "canceladas · neto = ÷ 1,19 · hora de Chile · * mes en curso (parcial)"]
        c = cuadre_total(nombres, datos, anio)
        if c:
            nota.append(f"Cuadre con la hoja «Ventas con IVA» de {anio}: {texto_cuadre(c)}. "
                        "La diferencia puede venir de descuentos, redondeos u otros ajustes "
                        "que FUDO aplica a nivel de venta.")
        if aviso:
            nota.append(aviso)
        hojas.append({"nombre": f"Productos {anio}",
                      "encabezado": ["Producto"] + cab_suc + ["Total unid.", "Total ventas",
                                                              "Total neto", "% ventas"],
                      "filas": filas, "nota": nota,
                      "anchos": [38] + [max(12, len(t) + 3) for t in cab_suc] + [15, 16, 16, 11]})
    filas = []
    for ym, suc, nombre, u, v in filas_detalle(nombres, datos):
        filas.append((False, [ym[:4], (ym, S_NORMAL), (suc, S_NORMAL), (nombre, S_NORMAL),
                              _c_unid(u), v, int(round(v / IVA))]))
    nota = ["Una fila por mes, sucursal y producto con movimiento · pensada para tablas "
            "dinámicas · ventas netas = ventas con IVA ÷ 1,19 · hora de Chile"]
    if aviso:
        nota.append(aviso.replace("(columnas en ERROR)", "(sin filas en esta hoja)"))
    hojas.append({"nombre": "Detalle productos",
                  "encabezado": ["Año", "Mes", "Sucursal", "Producto", "Unidades",
                                 "Ventas con IVA", "Ventas netas"],
                  "filas": filas, "nota": nota, "congelar_col": False,
                  "anchos": [8, 10, 18, 38, 12, 16, 16]})
    return hojas


def hojas_del_reporte(nombres, datos, meses):
    """Todas las hojas del Excel: las tres mensuales y, si se leyeron, las de productos."""
    hojas = construir_hojas(nombres, datos, meses)
    if any(r and r.get("prod_activo") for r in datos.values()):
        hojas += construir_hojas_productos(nombres, datos)
    return hojas


def guardar_csv_productos(carpeta, nombres, datos):
    ruta = os.path.join(carpeta, "ventas_por_producto.csv")
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Año", "Mes", "Sucursal", "Producto", "Unidades",
                    "Ventas con IVA", "Ventas netas"])
        for ym, suc, nombre, u, v in filas_detalle(nombres, datos):
            w.writerow([ym[:4], ym, suc, nombre, str(u).replace(".", ","), v,
                        int(round(v / IVA))])
    return ruta


def guardar(carpeta, nombres, datos, meses):
    """Guarda ventas_mensuales.xlsx (si no se puede, un CSV) y devuelve la ruta."""
    hojas = hojas_del_reporte(nombres, datos, meses)
    ruta = os.path.join(carpeta, "ventas_mensuales.xlsx")
    try:
        escribir_xlsx(ruta, hojas)
        return ruta
    except Exception as e:
        print(f"\nNo pude guardar el Excel ({e}); dejo un CSV en su lugar.")
        if any(r and r.get("prod_activo") for r in datos.values()):
            print(f"Productos: {guardar_csv_productos(carpeta, nombres, datos)}")
        return guardar_csv(carpeta, nombres, datos, meses)


def abrir(ruta):
    """En Mac abre el archivo con la aplicación predeterminada (Excel, Numbers…)."""
    if sys.platform == "darwin":
        try:
            subprocess.run(["open", ruta], timeout=15)
        except Exception:
            pass


# ── Programa principal ───────────────────────────────────────

def info_productos(r):
    """(líneas informativas, avisos) de los productos leídos en una sucursal."""
    g = r.get
    tickets = sum(m[1] for m in r["meses"].values())
    sin_items = g("ventas_sin_items", 0)
    distintos = len({pid for d in g("prod", {}).values() for pid in d})
    linea = (f"productos: {num(g('items', 0))} líneas de ítems en {num(tickets - sin_items)} "
             f"ventas · {num(distintos)} productos distintos")
    if g("items_anulados", 0):
        linea += f" · {num(g('items_anulados', 0))} ítems anulados excluidos"
    if 0 < sin_items < tickets:
        linea += f" · {num(sin_items)} ventas sin ítems"
    info, avisos = [linea], []
    if g("subitems", 0):
        info.append(f"incluye {num(g('subitems', 0))} subítems (modificadores o componentes de "
                    f"combos) por {clp(int(round(g('subitems_ventas', 0.0))))}")
    if g("subitems_excluidos", 0):
        info.append(f"{num(g('subitems_excluidos', 0))} subítems excluidos "
                    "(PRODUCTOS_SUBITEMS=excluir)")
    if PRECIO_LINEA:
        info.append("precio de cada ítem tomado como total de la línea (PRODUCTOS_PRECIO=linea)")
    c = cuadre_sucursal(r)
    if c and c[1] > 0:
        info.append("cuadre con el total de las ventas: " + texto_cuadre(c))
        if abs(c[0] / c[1] - 1) > 0.10:
            avisos.append("los ítems difieren más de un 10 % del total de las ventas; posibles "
                          "causas: descuentos, subítems (PRODUCTOS_SUBITEMS=excluir), precio "
                          "unitario o de línea (PRODUCTOS_PRECIO=linea), precios sin IVA o ítems "
                          "que FUDO no entregó")
    if g("nombres_error", ""):
        avisos.append(f"no se pudieron leer los nombres de los productos ({r['nombres_error']}); "
                      "aparecen como «Producto #id»")
    if tickets and sin_items == tickets:
        avisos.append("FUDO no entregó ítems en ninguna venta; revisa los permisos de la API key")
    if g("items_faltantes", 0):
        avisos.append(f"{num(g('items_faltantes', 0))} ítems que FUDO nombra pero no entregó "
                      "(no se sumaron)")
    if g("items_sin_precio", 0):
        avisos.append(f"{num(g('items_sin_precio', 0))} ítems sin precio legible "
                      "(se contaron como $0)")
    if g("items_sin_producto", 0):
        avisos.append(f"{num(g('items_sin_producto', 0))} ítems sin producto asociado "
                      "(aparecen como «(sin producto)»)")
    return info, avisos


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
    if r.get("prod_activo"):
        info_p, avisos_p = info_productos(r)
        info += info_p
        avisos += avisos_p
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
    ap.add_argument("--sin-productos", action="store_true",
                    help="no leer las ventas por producto (mucho más rápido)")
    args = ap.parse_args()
    productos = not args.sin_productos

    sucursales, auth_url, base_url = cargar_config()
    carpeta = os.path.dirname(os.path.abspath(__file__))
    print(f"Ventas mensuales por sucursal — {texto_anios()}")
    if not TZ_EXACTA:
        print("Aviso: no pude cargar la zona horaria de Chile; uso UTC-4 fijo "
              "(solo afecta ventas de medianoche en horario de verano).")
    print("Leyendo ventas" + (" y productos" if productos else "")
          + " desde FUDO (puede tardar varios minutos).\n")

    nombres, datos = [], {}
    for suc in sucursales:
        nombre = suc.get("nombre", "?")
        nombres.append(nombre)
        try:
            cli = Fudo(auth_url, base_url, suc)
            cli.autenticar()
            datos[nombre] = leer_sucursal(cli, nombre, args.completo, productos=productos)
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
    if productos:
        imprimir_productos(nombres, datos)

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
