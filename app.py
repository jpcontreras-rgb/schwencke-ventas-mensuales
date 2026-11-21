#!/usr/bin/env python3
"""
Ventas mensuales por sucursal (FUDO) — servicio web para Railway.

Página privada (protegida con clave) con un botón «Generar»: lee las ventas de
cada sucursal desde la API de FUDO, muestra las tablas y entrega el Excel.

Es un servicio APARTE del dashboard: no lo modifica, no comparte memoria con él
y no escribe nada en FUDO (solo lee /sales).

Variables de entorno (en Railway → Variables):
    SUC1_NOMBRE, SUC1_APIKEY, SUC1_APISECRET     sucursal 1 (igual que el dashboard)
    SUC2_…, SUC3_…                                 una tanda por sucursal
    SUCn_COLOR                                     opcional, p. ej. #D85A30
    DASHBOARD_PASSWORD                             clave de acceso (obligatoria)
    ANIOS                                          opcional, por defecto 2025,2026
    SECRET_KEY                                     opcional (si falta se deriva de la clave)
    PORT                                           lo define Railway
"""

import hashlib
import hmac
import io
import ipaddress
import os
import re
import secrets
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from html import escape

import requests
from flask import Flask, Response, g, jsonify, redirect, request, send_file, session

import ventas_mensuales as vm

FUDO_AUTH = os.environ.get("FUDO_AUTH_URL", "https://auth.fu.do/api")
FUDO_BASE = os.environ.get("FUDO_BASE_URL", "https://api.fu.do/v1alpha1")
PASSWORD = os.environ.get("DASHBOARD_PASSWORD", "")
EN_RAILWAY = any(k.startswith("RAILWAY_") for k in os.environ)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def log(texto):
    try:
        print(f"[{vm.ahora():%d/%m %H:%M:%S}] {texto}", flush=True)
    except Exception:                                        # un log nunca debe tumbar la lectura
        pass


# ── Configuración ────────────────────────────────────────────

def leer_anios():
    """ANIOS="2025,2026" -> (2025, 2026). Si falta o es inválido, usa el valor por defecto."""
    texto = os.environ.get("ANIOS", "").strip()
    if not texto:
        return vm.ANIOS
    anios = sorted({int(x) for x in re.findall(r"\d{4}", texto) if 2000 <= int(x) <= 2100})
    if not anios:
        print(f"Aviso: ANIOS='{texto}' no es válido; uso {vm.texto_anios()}", flush=True)
        return vm.ANIOS
    return tuple(anios)


vm.ANIOS = leer_anios()

HEX = re.compile(r"^#[0-9a-fA-F]{6}$")
PALETA = ["#D85A30", "#1D9E75", "#7F77DD", "#C99A2E", "#3B82C4", "#B45A8C"]
COLORES_POR_NOMBRE = (("viña", "#D85A30"), ("vina", "#D85A30"),
                      ("parque", "#1D9E75"), ("mut", "#7F77DD"))


def leer_sucursales():
    """SUC1_NOMBRE/_APIKEY/_APISECRET/_COLOR, SUC2_…, igual que el dashboard."""
    sucs, usados, i = [], set(), 1
    while True:
        nombre = (os.environ.get(f"SUC{i}_NOMBRE") or "").strip()
        if not nombre:
            break
        if nombre in usados:
            nombre = f"{nombre} ({i})"
        usados.add(nombre)
        color = (os.environ.get(f"SUC{i}_COLOR") or "").strip()
        if not HEX.match(color):
            bajo = nombre.lower()
            color = next((c for k, c in COLORES_POR_NOMBRE if k in bajo),
                         PALETA[(i - 1) % len(PALETA)])
        sucs.append({"n": i, "nombre": nombre, "color": color,
                     "api_key": (os.environ.get(f"SUC{i}_APIKEY") or "").strip(),
                     "api_secret": (os.environ.get(f"SUC{i}_APISECRET") or "").strip()})
        i += 1
    return sucs


SUCURSALES = leer_sucursales()


# ── Estado de la lectura (en memoria, un solo proceso) ───────

LOCK = threading.Lock()
JOB = {"id": 0, "fase": "libre", "completo": False, "inicio": 0.0, "fin": 0.0,
       "mensaje": "", "sucs": []}                  # fase: libre | leyendo | listo | error
ULTIMO = {"resultado": None}                       # último resultado bueno


def nueva_suc(s):
    return {"nombre": s["nombre"], "etapa": "pendiente", "pagina": 0, "total": 0,
            "desde": 1, "hechas": 0, "leidas": 0, "segundos": 0.0, "error": ""}


def dur(seg):
    m, s = divmod(int(seg), 60)
    return f"{m} min {s:02d} s" if m else f"{s} s"


def limpiar(texto):
    """Mensaje de error sin credenciales y de largo razonable."""
    texto = str(texto)
    secretos = sorted({v for s in SUCURSALES for v in (s["api_key"], s["api_secret"])
                       if len(v) >= 4}, key=len, reverse=True)      # los más largos primero
    for secreto in secretos:
        texto = texto.replace(secreto, "***")
    return texto[:300]


def texto_suc(s, completo=False):
    e = s["etapa"]
    if e == "pendiente":
        return "En espera"
    if e == "autenticando":
        return "Conectando con FUDO…"
    if e == "buscando":
        return ("Contando las páginas de ventas…" if completo
                else f"Buscando dónde empiezan las ventas de {vm.ANIOS[0]}…")
    if e == "leyendo":
        return f"Página {s['pagina']} de {s['total']} · {vm.num(s['leidas'])} ventas leídas"
    if e == "listo":
        return f"Listo · {vm.num(s['leidas'])} ventas leídas en {dur(s['segundos'])}"
    return s["error"] or "Error"


def pct_suc(s):
    """Avance 0-100, o None si no se puede medir (barra animada)."""
    e = s["etapa"]
    if e == "pendiente":
        return 0
    if e in ("autenticando", "buscando"):
        return None
    if e == "leyendo":
        return min(99.0, round(100.0 * s["hechas"] / max(1, s["total"] - s["desde"] + 1), 1))
    return 100


def instantanea():
    with LOCK:
        j = {k: v for k, v in JOB.items() if k != "sucs"}
        sucs = [dict(s) for s in JOB["sucs"]]
    fin = j["fin"] or time.time()
    return {"fase": j["fase"], "completo": j["completo"], "mensaje": j["mensaje"],
            "transcurrido": int(fin - j["inicio"]) if j["inicio"] else 0,
            "sucs": [{"nombre": s["nombre"], "etapa": s["etapa"], "pct": pct_suc(s),
                      "texto": texto_suc(s, j["completo"]), "error": s["error"]} for s in sucs]}


def terminar_con_error(token, mensaje):
    with LOCK:
        if JOB["id"] == token:
            JOB.update(fase="error", mensaje=mensaje, fin=time.time())
    log(f"lectura fallida: {mensaje}")


def leer_una(token, i, suc, completo):
    """Lee una sucursal (en su propio hilo y con su propia sesión HTTP)."""
    nombre = suc["nombre"]

    def progreso(info):
        with LOCK:
            if JOB["id"] != token:
                return
            s = JOB["sucs"][i]
            previa = s["etapa"]
            s["etapa"] = info.get("etapa", previa)
            if s["etapa"] == "leyendo":
                s["hechas"] = s["hechas"] + 1 if previa == "leyendo" else 0
                s["pagina"], s["total"] = info["pagina"], info["total"]
                s["desde"], s["leidas"] = info["desde"], info["leidas"]

    try:
        if not suc["api_key"] or not suc["api_secret"]:
            raise vm.ErrorFudo(f"faltan las variables SUC{suc['n']}_APIKEY y/o "
                               f"SUC{suc['n']}_APISECRET en Railway")
        with requests.Session() as http:
            cli = vm.Fudo(FUDO_AUTH, FUDO_BASE, suc, http=http)
            progreso({"etapa": "autenticando"})
            cli.autenticar()
            r = vm.leer_sucursal(cli, nombre, completo, progreso=progreso)
        with LOCK:
            if JOB["id"] == token:
                JOB["sucs"][i].update(etapa="listo", leidas=r["leidas"],
                                      segundos=r["segundos"])
        log(f"{nombre}: {r['leidas']} ventas leídas en {dur(r['segundos'])}")
        return nombre, r
    except Exception as e:                                   # noqa: BLE001
        msg = str(e) if isinstance(e, vm.ErrorFudo) else f"{type(e).__name__}: {e}"
        msg = limpiar(msg)
        with LOCK:
            if JOB["id"] == token:
                JOB["sucs"][i].update(etapa="error", error=msg)
        log(f"{nombre}: ERROR — {msg}")
        return nombre, None


def ejecutar(token, completo):
    t0 = time.time()
    log(f"lectura iniciada ({'completa' if completo else 'rápida'}, "
        f"{len(SUCURSALES)} sucursales, {vm.texto_anios()})")
    try:
        with ThreadPoolExecutor(max_workers=max(1, len(SUCURSALES))) as pool:
            futuros = [pool.submit(leer_una, token, i, s, completo)
                       for i, s in enumerate(SUCURSALES)]
            resultados = [f.result() for f in futuros]
        nombres = [s["nombre"] for s in SUCURSALES]
        datos = dict(resultados)
        with LOCK:
            errores = {s["nombre"]: s["error"] for s in JOB["sucs"] if s["etapa"] == "error"}
        if all(v is None for v in datos.values()):
            terminar_con_error(token, "No se pudo leer ninguna sucursal. "
                                      "Revisa el detalle de cada una.")
            return
        meses = vm.lista_meses()
        buf = io.BytesIO()
        vm.escribir_xlsx(buf, vm.construir_hojas(nombres, datos, meses))
        ahora = vm.ahora()
        resultado = {
            "xlsx": buf.getvalue(),
            "archivo": f"ventas_mensuales_{vm.rango_anios()}_{ahora:%Y-%m-%d}.xlsx",
            "generado": ahora.strftime("%d/%m/%Y %H:%M"),
            "duracion": time.time() - t0,
            "anios": vm.texto_anios(),
            "errores": errores,
            "tablas": [tabla_html(nombres, datos, meses, 0, 1.0, vm.clp),
                       tabla_html(nombres, datos, meses, 0, 1 / vm.IVA, vm.clp),
                       tabla_html(nombres, datos, meses, 1, 1.0, vm.num)],
            "cobertura": cobertura_html(nombres, datos, errores),
        }
        with LOCK:
            if JOB["id"] != token:
                return
            ULTIMO["resultado"] = resultado
            JOB.update(fase="listo", mensaje="", fin=time.time())
        log(f"reporte listo en {dur(time.time() - t0)} ({len(resultado['xlsx'])} bytes)")
    except Exception as e:                                   # noqa: BLE001
        terminar_con_error(token, limpiar(f"{type(e).__name__}: {e}"))


def iniciar(completo):
    """Arranca una lectura; devuelve False si ya hay una en curso."""
    with LOCK:
        if JOB["fase"] == "leyendo" and time.time() - JOB["inicio"] < 2 * 3600:
            return False
        JOB["id"] += 1
        token = JOB["id"]
        JOB.update(fase="leyendo", completo=completo, inicio=time.time(), fin=0.0,
                   mensaje="", sucs=[nueva_suc(s) for s in SUCURSALES])
    threading.Thread(target=ejecutar, args=(token, completo), daemon=True).start()
    return True


# ── Tablas en HTML ───────────────────────────────────────────

def celda_html(v, fmt, total=False):
    if v is None:
        return f'<td class="num mal">{"n/d" if total else "ERROR"}</td>'
    if v == 0:
        return '<td class="num cero">—</td>'
    return f'<td class="num">{escape(fmt(v))}</td>'


def tabla_html(nombres, datos, meses, idx, factor, fmt):
    filas = vm.construir_filas(nombres, datos, meses, idx, factor)
    out = ['<table><thead><tr><th class="mes">Mes</th>']
    for i, n in enumerate(nombres):
        out.append(f'<th class="s{i}"><span class="dot"></span>{escape(n)}</th>')
    out.append("<th>Total</th></tr></thead><tbody>")
    for es_total, etiq, vals, tot in filas:
        out.append('<tr class="total">' if es_total else "<tr>")
        out.append(f'<th scope="row" class="mes">{escape(etiq)}</th>')
        out.extend(celda_html(v, fmt) for v in vals)
        out.append(celda_html(tot, fmt, total=True))
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out)


def cobertura_html(nombres, datos, errores):
    hay_avisos, bloques = False, []
    for i, n in enumerate(nombres):
        r = datos[n]
        if r is None:
            hay_avisos = True
            lineas = [("mal", f"✗ {errores.get(n) or 'no se pudo leer'} — "
                              "esta sucursal aparece como ERROR en las tablas y en el Excel")]
        else:
            info, avisos = vm.info_lectura(r, pista="la opción «Leer todo el historial»")
            hay_avisos = hay_avisos or bool(avisos)
            lineas = [("", t) for t in info] + [("aviso", "⚠ " + a) for a in avisos]
        items = "".join(f'<li class="{c}">{escape(t)}</li>' for c, t in lineas)
        bloques.append(f'<div class="cobsuc s{i}"><strong><span class="dot"></span>'
                       f"{escape(n)}</strong><ul>{items}</ul></div>")
    abierto = " open" if hay_avisos else ""
    return (f'<details class="cob"{abierto}><summary>Cobertura: qué se leyó de cada '
            f'sucursal{" · hay avisos" if hay_avisos else ""}</summary>{"".join(bloques)}</details>')


# ── Aplicación web ───────────────────────────────────────────

app = Flask(__name__)
app.config.update(
    SECRET_KEY=(os.environ.get("SECRET_KEY")
                or (hashlib.sha256(("ventas-mensuales|" + PASSWORD).encode()).hexdigest()
                    if PASSWORD else secrets.token_hex(32))),
    SESSION_COOKIE_NAME="vm_sesion",
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=EN_RAILWAY,
    PERMANENT_SESSION_LIFETIME=12 * 3600,
    MAX_CONTENT_LENGTH=16 * 1024,
)

# Límite de intentos fallidos de clave: por IP y, además, global (por si falsean la IP)
VENTANA, MAX_IP, MAX_GLOBAL = 600, 5, 30
FALLOS_IP, FALLOS_TODOS, LOCK_FALLOS = {}, [], threading.Lock()


def ip_cliente():
    """IP del cliente. Waitress deja en REMOTE_ADDR el último valor de X-Forwarded-For (lo que
    vio el proxy de Railway) o, sin proxy, la IP directa. Si no es una IP válida, un solo grupo."""
    try:
        return str(ipaddress.ip_address(request.remote_addr or ""))
    except ValueError:
        return "desconocida"


def bloqueado(ip):
    t = time.time()
    with LOCK_FALLOS:
        FALLOS_TODOS[:] = [x for x in FALLOS_TODOS if t - x < VENTANA]
        if len(FALLOS_IP) > 500:
            for k in [k for k, v in FALLOS_IP.items() if all(t - x >= VENTANA for x in v)]:
                del FALLOS_IP[k]
        propios = [x for x in FALLOS_IP.get(ip, []) if t - x < VENTANA]
        if propios:
            FALLOS_IP[ip] = propios
        else:
            FALLOS_IP.pop(ip, None)
        return len(propios) >= MAX_IP or len(FALLOS_TODOS) >= MAX_GLOBAL


def registrar_fallo(ip):
    t = time.time()
    with LOCK_FALLOS:
        FALLOS_IP.setdefault(ip, []).append(t)
        FALLOS_TODOS.append(t)


def csrf_ok():
    enviado = request.headers.get("X-CSRF-Token") or request.form.get("csrf") or ""
    esperado = session.get("csrf") or ""
    return bool(esperado) and hmac.compare_digest(enviado.encode(), esperado.encode())


@app.before_request
def antes():
    g.nonce = secrets.token_urlsafe(16)
    if request.path == "/healthz":
        return None
    if not PASSWORD:
        return pagina_sin_clave()
    return None


@app.after_request
def cabeceras(resp):
    n = getattr(g, "nonce", "")
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Content-Security-Policy"] = (
        f"default-src 'none'; script-src 'nonce-{n}'; style-src 'nonce-{n}'; "
        "connect-src 'self'; img-src data:; form-action 'self'; base-uri 'none'; "
        "frame-ancestors 'none'")
    if EN_RAILWAY:
        resp.headers["Strict-Transport-Security"] = "max-age=15552000"
    return resp


def autenticado():
    return bool(session.get("ok"))


@app.get("/healthz")
def healthz():
    return Response("ok", mimetype="text/plain")


@app.route("/login", methods=["GET", "POST"])
def login():
    if autenticado():
        return redirect("/")
    error = ""
    if request.method == "POST":
        ip = ip_cliente()
        if bloqueado(ip):
            error = "Demasiados intentos fallidos. Espera unos minutos e inténtalo de nuevo."
            return Response(pagina_login(error), status=429, mimetype="text/html")
        clave = request.form.get("clave", "")
        if hmac.compare_digest(clave.encode(), PASSWORD.encode()):
            with LOCK_FALLOS:
                FALLOS_IP.pop(ip, None)
            session.clear()
            session.permanent = True
            session["ok"] = True
            session["csrf"] = secrets.token_urlsafe(24)
            return redirect("/")
        registrar_fallo(ip)
        log(f"clave incorrecta desde {ip}")
        error = "Clave incorrecta."
        return Response(pagina_login(error), status=401, mimetype="text/html")
    return Response(pagina_login(error), mimetype="text/html")


@app.post("/logout")
def logout():
    if autenticado() and csrf_ok():
        session.clear()
    return redirect("/login")


@app.get("/")
def inicio():
    if not autenticado():
        return redirect("/login")
    return Response(pagina_principal(instantanea(), ULTIMO["resultado"]), mimetype="text/html")


@app.get("/estado")
def estado():
    if not autenticado():
        return jsonify(error="sesión vencida"), 401
    return jsonify(instantanea())


@app.post("/generar")
def generar():
    if not autenticado():
        return jsonify(error="sesión vencida"), 401
    if not csrf_ok():
        return jsonify(error="solicitud no válida; recarga la página"), 403
    if not SUCURSALES:
        return jsonify(error="No hay sucursales configuradas (variables SUC1_NOMBRE, "
                             "SUC1_APIKEY y SUC1_APISECRET)"), 400
    completo = bool((request.get_json(silent=True) or {}).get("completo"))
    return jsonify(ok=True, nueva=iniciar(completo))


@app.get("/descargar.xlsx")
def descargar():
    if not autenticado():
        return redirect("/login")
    r = ULTIMO["resultado"]
    if not r:
        return Response("Todavía no hay un reporte generado.", status=404, mimetype="text/plain")
    return send_file(io.BytesIO(r["xlsx"]), mimetype=XLSX, as_attachment=True,
                     download_name=r["archivo"])


# ── HTML ─────────────────────────────────────────────────────

FAVICON = ("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E"
           "%3Crect width='32' height='32' rx='7' fill='%23C24A1F'/%3E"
           "%3Crect x='7' y='16' width='5' height='9' rx='1' fill='white'/%3E"
           "%3Crect x='14' y='11' width='5' height='14' rx='1' fill='white'/%3E"
           "%3Crect x='21' y='6' width='5' height='19' rx='1' fill='white'/%3E%3C/svg%3E")

CSS = """
[hidden]{display:none!important}
:root{color-scheme:light dark;--bg:#f4f3ef;--card:#fff;--line:#e8e7e2;--ink:#1a1a1a;--mute:#6b6a65;
--acc:#C24A1F;--acch:#A83E18;--track:#eceae3;--head:#1F3864;--tot:#E4EDF4;--ok:#3B6D11;--okbg:#EAF3DE;
--warn:#7a5200;--warnbg:#FFF3D1;--bad:#A32D2D;--badbg:#FDE8E8}
@media (prefers-color-scheme:dark){:root{--bg:#151514;--card:#1e1e1c;--line:#33322e;--ink:#ecebe6;
--mute:#a09f98;--track:#2c2b27;--tot:#243038;--ok:#a5d676;--okbg:#222d18;--warn:#f1c76e;
--warnbg:#352b10;--bad:#f2a0a0;--badbg:#382020}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 -apple-system,BlinkMacSystemFont,
"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:0 16px}
.top{border-bottom:1px solid var(--line);background:var(--card)}
.top .wrap{display:flex;align-items:center;justify-content:space-between;gap:12px;padding-top:14px;padding-bottom:14px}
h1{font-size:18px;margin:0;font-weight:650}
h2{font-size:15px;margin:0;font-weight:650}
.sub,.mute{color:var(--mute)}
.sub{margin:2px 0 0;font-size:12.5px}
main.wrap{padding-top:18px;padding-bottom:40px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}
.fila{display:flex;flex-wrap:wrap;align-items:center;gap:14px}
.cab{display:flex;flex-wrap:wrap;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:12px}
.cab p{margin:2px 0 0;font-size:12.5px}
.ayuda{margin:12px 0 0;font-size:12.5px;color:var(--mute)}
.btn{display:inline-block;border:0;border-radius:8px;background:var(--acc);color:#fff;font-weight:600;
font-size:14px;font-family:inherit;padding:11px 18px;cursor:pointer;text-decoration:none;line-height:1}
.btn:hover:not(:disabled){background:var(--acch)}
.btn:disabled{opacity:.55;cursor:default}
.btn:focus-visible,.link:focus-visible,.tabs button:focus-visible,summary:focus-visible,input:focus-visible{
outline:2px solid var(--ink);outline-offset:2px}
.link{background:none;border:0;color:var(--mute);font:inherit;cursor:pointer;text-decoration:underline;padding:4px}
.chk{display:flex;align-items:center;gap:8px;cursor:pointer}
.chk input{width:16px;height:16px;accent-color:var(--acc)}
.msg{margin:12px 0 0;padding:9px 12px;border-radius:8px;background:var(--badbg);color:var(--bad);font-size:13px}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:var(--c,#999);margin-right:7px}
.suc{padding:10px 0;border-top:1px solid var(--line)}
.suc .nom{font-weight:600;margin-bottom:6px}
.pista{height:7px;background:var(--track);border-radius:4px;overflow:hidden}
.barra{height:100%;width:0;background:var(--c,#999);border-radius:4px;transition:width .5s ease}
.suc.indet .barra{animation:vaiven 1.4s ease-in-out infinite alternate}
@keyframes vaiven{from{margin-left:0}to{margin-left:70%}}
.suc .txt{margin-top:6px;font-size:12.5px;color:var(--mute)}
.suc.fallo .txt{color:var(--bad)}
.suc.fallo .barra{background:var(--bad)}
.suc.listo .txt{color:var(--ok)}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.tabs button{border:1px solid var(--line);background:transparent;color:var(--mute);border-radius:20px;
padding:6px 14px;font:inherit;font-size:13px;cursor:pointer}
.tabs button.on{background:var(--ink);color:var(--card);border-color:var(--ink)}
.scroll{overflow-x:auto;border:1px solid var(--line);border-radius:8px}
table{border-collapse:collapse;width:100%;font-size:13px;font-variant-numeric:tabular-nums}
th,td{padding:7px 12px;white-space:nowrap}
thead th{background:var(--head);color:#fff;font-weight:600;text-align:right}
thead th .dot{box-shadow:0 0 0 1.5px #fff}
tbody tr+tr th,tbody tr+tr td{border-top:1px solid var(--line)}
th.mes{text-align:left;position:sticky;left:0;background:var(--card);font-weight:500}
thead th.mes{background:var(--head);font-weight:600}
td.num{text-align:right}
td.cero{color:var(--mute)}
td.mal{color:var(--bad);font-weight:600}
tr.total th,tr.total td{background:var(--tot);font-weight:700}
.nota{margin:10px 0 0;font-size:12px}
.aviso{margin:0 0 12px;padding:9px 12px;border-radius:8px;background:var(--warnbg);color:var(--warn);font-size:13px}
details.cob{margin-top:14px;border-top:1px solid var(--line);padding-top:12px}
details.cob summary{cursor:pointer;font-weight:600}
.cobsuc{margin-top:10px}
.cobsuc ul{margin:4px 0 0;padding-left:26px;color:var(--mute);font-size:12.5px}
.cobsuc li.aviso{background:none;color:var(--warn);padding:0;margin:0;font-size:12.5px}
.cobsuc li.mal{color:var(--bad)}
.malo{border-color:var(--bad)}
.malo ul{margin:8px 0 0;padding-left:18px}
.login{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px}
.caja{width:100%;max-width:340px;padding:28px}
.caja h1{margin-bottom:2px}
.caja label{display:block;margin:18px 0 6px;font-weight:600;font-size:13px}
.caja input[type=password]{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:8px;
background:var(--bg);color:var(--ink);font:inherit}
.caja .btn{width:100%;margin-top:14px;text-align:center;padding:12px}
.caja .msg{margin:14px 0 0}
@media (max-width:560px){th,td{padding:7px 9px}.top .wrap{align-items:flex-start}}
@media (prefers-reduced-motion:reduce){.barra{transition:none}.suc.indet .barra{animation:none;width:40%}}
"""

JS = """
(function () {
  "use strict";
  var cuerpo = document.body;
  var csrf = cuerpo.getAttribute("data-csrf") || "";
  var boton = document.getElementById("generar");
  var completo = document.getElementById("completo");
  var panel = document.getElementById("progreso");
  var reloj = document.getElementById("reloj");
  var msg = document.getElementById("msg");
  var tarjetas = document.querySelectorAll("#progreso .suc");
  var ocultables = document.querySelectorAll("[data-oculta-al-leer]");
  var ya = false;

  function paraCada(lista, fn) { Array.prototype.forEach.call(lista, fn); }

  function dur(s) {
    var m = Math.floor(s / 60), r = s % 60;
    return m > 0 ? m + " min " + (r < 10 ? "0" : "") + r + " s" : r + " s";
  }

  function pintar(e) {
    reloj.textContent = dur(e.transcurrido || 0);
    paraCada(e.sucs || [], function (s, i) {
      var t = tarjetas[i];
      if (!t) { return; }
      t.querySelector(".txt").textContent = s.texto;
      t.classList.toggle("indet", s.pct === null);
      t.classList.toggle("listo", s.etapa === "listo");
      t.classList.toggle("fallo", s.etapa === "error");
      t.querySelector(".barra").style.width = (s.pct === null ? 30 : s.pct) + "%";
    });
  }

  function modoLeyendo() {
    panel.hidden = false;
    paraCada(ocultables, function (el) { el.hidden = true; });
    boton.disabled = true;
    boton.textContent = "Leyendo…";
    msg.hidden = true;
  }

  function consultar() {
    fetch("/estado", { cache: "no-store", credentials: "same-origin" })
      .then(function (r) {
        if (r.status === 401) { ya = true; location.href = "/login"; return null; }
        return r.json();
      })
      .then(function (e) {
        if (!e) { return; }
        pintar(e);
        if (e.fase === "leyendo") { setTimeout(consultar, 1500); }
        else { location.reload(); }
      })
      .catch(function () { if (!ya) { setTimeout(consultar, 4000); } });
  }

  function falla(texto) {
    boton.disabled = false;
    msg.textContent = texto;
    msg.hidden = false;
  }

  boton.addEventListener("click", function () {
    boton.disabled = true;
    fetch("/generar", {
      method: "POST", credentials: "same-origin",
      headers: { "Content-Type": "application/json", "X-CSRF-Token": csrf },
      body: JSON.stringify({ completo: !!(completo && completo.checked) })
    })
      .then(function (r) {
        if (r.status === 401) { ya = true; location.href = "/login"; return null; }
        return r.json().then(function (j) { return { ok: r.ok, json: j }; });
      })
      .then(function (x) {
        if (!x) { return; }
        if (x.ok) { modoLeyendo(); consultar(); }
        else { falla(x.json.error || "No se pudo iniciar la lectura."); }
      })
      .catch(function () { falla("No se pudo conectar con el servidor. Inténtalo de nuevo."); });
  });

  var pestanas = document.querySelectorAll("[data-tab]");
  var paneles = document.querySelectorAll("[data-panel]");
  paraCada(pestanas, function (b) {
    b.addEventListener("click", function () {
      var k = b.getAttribute("data-tab");
      paraCada(pestanas, function (x) {
        x.classList.toggle("on", x === b);
        x.setAttribute("aria-selected", x === b ? "true" : "false");
      });
      paraCada(paneles, function (p) { p.hidden = p.getAttribute("data-panel") !== k; });
    });
  });

  if (cuerpo.getAttribute("data-fase") === "leyendo") { modoLeyendo(); consultar(); }
})();
"""


def pagina(titulo, cuerpo, atributos="", con_js=False):
    n = g.nonce
    script = f'<script nonce="{n}">{JS}</script>' if con_js else ""
    return ('<!doctype html>\n<html lang="es"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1">'
            '<meta name="robots" content="noindex, nofollow">'
            '<meta name="color-scheme" content="light dark">'
            f'<title>{escape(titulo)}</title><link rel="icon" href="{FAVICON}">'
            f'<style nonce="{n}">{CSS}{css_sucursales()}</style></head>'
            f"<body{atributos}>{cuerpo}{script}</body></html>")


def css_sucursales():
    return "".join(f".s{i}{{--c:{s['color']}}}" for i, s in enumerate(SUCURSALES))


def pagina_sin_clave():
    cuerpo = ('<main class="login"><div class="card caja"><h1>Falta la clave de acceso</h1>'
              '<p class="mute">Define la variable <strong>DASHBOARD_PASSWORD</strong> en '
              "Railway (servicio → Variables) y el servicio se reiniciará solo. "
              "Mientras no exista, esta página permanece cerrada.</p></div></main>")
    return Response(pagina("Ventas mensuales", cuerpo), status=503, mimetype="text/html")


def pagina_login(error):
    aviso = f'<p class="msg" role="alert">{escape(error)}</p>' if error else ""
    cuerpo = ('<main class="login"><form class="card caja" method="post" action="/login">'
              '<h1>Ventas mensuales</h1><p class="mute">Emporio Schwencke</p>'
              f'{aviso}<label for="clave">Clave de acceso</label>'
              '<input id="clave" name="clave" type="password" autocomplete="current-password" '
              'autofocus required><button class="btn" type="submit">Entrar</button></form></main>')
    return pagina("Ingresar · Ventas mensuales", cuerpo)


def tarjetas_progreso(snap):
    out = []
    for i, s in enumerate(snap["sucs"] or [{"nombre": x["nombre"], "texto": "En espera"}
                                           for x in SUCURSALES]):
        clase = {"listo": " listo", "error": " fallo"}.get(s.get("etapa"), "")
        out.append(f'<div class="suc s{i}{clase}"><div class="nom"><span class="dot"></span>'
                   f'{escape(s["nombre"])}</div><div class="pista"><div class="barra"></div></div>'
                   f'<div class="txt" aria-live="polite">{escape(s["texto"])}</div></div>')
    return "".join(out)


def bloque_resultado(res):
    if not res:
        return ""
    errores = res["errores"]
    aviso = ""
    if errores:
        nombres = ", ".join(errores)
        aviso = (f'<p class="aviso"><strong>No se pudo leer: {escape(nombres)}.</strong> '
                 "Su columna aparece como ERROR en las tablas y en el Excel; "
                 "vuelve a generar el reporte para completarla.</p>")
    pestanas = "".join(
        f'<button type="button" role="tab" data-tab="{i}" aria-selected="{"true" if i == 0 else "false"}"'
        f'{" class=on" if i == 0 else ""}>{escape(t)}</button>'
        for i, t in enumerate(["Ventas con IVA", "Ventas netas", "Tickets"]))
    paneles = "".join(
        f'<div data-panel="{i}"{"" if i == 0 else " hidden"}><div class="scroll">{t}</div></div>'
        for i, t in enumerate(res["tablas"]))
    return (
        '<section class="card" id="resultados" data-oculta-al-leer>'
        '<div class="cab"><div><h2>Resultado · ' + escape(res["anios"]) + '</h2>'
        f'<p class="mute">Generado el {escape(res["generado"])} · tardó {escape(dur(res["duracion"]))}</p></div>'
        '<a class="btn" href="/descargar.xlsx">Descargar Excel</a></div>'
        f'{aviso}<div class="tabs" role="tablist">{pestanas}</div>{paneles}'
        '<p class="nota mute">* mes en curso (parcial) · ventas no canceladas · hora de Chile · '
        "netas = total ÷ 1,19</p>" + res["cobertura"] + "</section>")


def pagina_principal(snap, res):
    fase = snap["fase"]
    leyendo = fase == "leyendo"
    if not SUCURSALES:
        aviso_cfg = ('<p class="msg">No encontré sucursales configuradas. Agrega en Railway las '
                     "variables SUC1_NOMBRE, SUC1_APIKEY y SUC1_APISECRET (y SUC2_…, SUC3_…).</p>")
    else:
        aviso_cfg = ""
    if not vm.TZ_EXACTA:
        aviso_cfg += ('<p class="msg">No pude cargar la zona horaria de Chile; uso UTC-4 fijo '
                      "(solo afecta ventas de medianoche en horario de verano).</p>")
    deshabilitar = " disabled" if (leyendo or not SUCURSALES) else ""
    etiqueta = "Leyendo…" if leyendo else ("Generar de nuevo" if res else "Generar")
    control = (
        '<section class="card"><div class="fila">'
        f'<button id="generar" class="btn" type="button"{deshabilitar}>{etiqueta}</button>'
        '<label class="chk"><input type="checkbox" id="completo"> '
        "Leer todo el historial (más lento)</label></div>"
        '<p class="ayuda">Lee las ventas de cada sucursal directamente desde FUDO. Tarda varios '
        "minutos; puedes cerrar la página y volver más tarde: la lectura sigue en segundo plano "
        "y el resultado queda aquí mientras el servicio siga activo. Por defecto se salta el "
        "historial anterior a " + str(vm.ANIOS[0]) + "; marca la opción solo si la cobertura avisa "
        'que falta información.</p><p class="msg" id="msg" role="alert" hidden></p>'
        + aviso_cfg + "</section>")
    oculto = "" if leyendo else " hidden"
    progreso = ('<section class="card" id="progreso"' + oculto + ">"
                '<div class="cab"><h2>Leyendo ventas desde FUDO…</h2>'
                '<span class="mute" id="reloj"></span></div>' + tarjetas_progreso(snap) + "</section>")
    error = ""
    if fase == "error":
        filas = "".join(f'<li>{escape(s["nombre"])}: {escape(s["error"] or "sin detalle")}</li>'
                        for s in snap["sucs"] if s["etapa"] == "error")
        error = ('<section class="card malo" data-oculta-al-leer><h2>No se pudo generar el reporte</h2>'
                 f'<p>{escape(snap["mensaje"])}</p>' + (f"<ul>{filas}</ul>" if filas else "") +
                 ("<p class=\"mute\">Se muestra abajo el último reporte bueno.</p>" if res else "")
                 + "</section>")
    resultado = "" if leyendo else bloque_resultado(res)
    cuerpo = (
        '<header class="top"><div class="wrap"><div><h1>Ventas mensuales por sucursal</h1>'
        f'<p class="sub">Emporio Schwencke · {escape(vm.texto_anios())} · datos de FUDO · hora de Chile</p></div>'
        f'<form method="post" action="/logout"><input type="hidden" name="csrf" value="{escape(session.get("csrf", ""))}">'
        '<button class="link" type="submit">Salir</button></form></div></header>'
        '<main class="wrap"><noscript><p class="msg">Esta página necesita JavaScript.</p></noscript>'
        + control + progreso + error + resultado + "</main>")
    atributos = f' data-csrf="{escape(session.get("csrf", ""))}" data-fase="{escape(fase)}"'
    return pagina("Ventas mensuales por sucursal", cuerpo, atributos, con_js=True)


# ── Arranque ─────────────────────────────────────────────────

def main():
    port = int(os.environ.get("PORT", "8080"))
    nombres = ", ".join(s["nombre"] for s in SUCURSALES) or "ninguna"
    print(f"Ventas mensuales · sucursales: {nombres} · años: {vm.texto_anios()} · "
          f"clave {'definida' if PASSWORD else 'NO definida (página cerrada)'} · puerto {port}",
          flush=True)
    if not SUCURSALES:
        print("Aviso: no hay sucursales (faltan SUC1_NOMBRE, SUC1_APIKEY, SUC1_APISECRET).",
              flush=True)
    try:
        from waitress import serve
    except ImportError:                                     # sin waitress: servidor de Flask
        app.run(host="0.0.0.0", port=port, threaded=True)
    else:
        # detrás del proxy de Railway: toma la IP real del cliente desde X-Forwarded-For
        serve(app, host="0.0.0.0", port=port, threads=8, trusted_proxy="*",
              trusted_proxy_headers={"x-forwarded-for"}, trusted_proxy_count=1)


if __name__ == "__main__":
    main()
