# Ventas mensuales por sucursal (FUDO)

Servicio web **aparte** para Railway: una página privada con un botón **Generar** que lee las ventas
de FUDO de cada sucursal y entrega las ventas mensuales de 2025 y 2026 en pantalla y en Excel
(hojas: ventas con IVA, ventas netas y tickets).

- No toca el dashboard ni sus variables: es otro repositorio y otro servicio de Railway.
- Solo **lee** de FUDO (`/sales`); no escribe nada.
- Criterio igual al del dashboard: ventas no canceladas, total de FUDO (con IVA), mes según la fecha de
  cierre en hora de Chile. Netas = total ÷ 1,19.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | La página web: ingreso con clave, botón Generar, avance por sucursal y descarga del Excel |
| `ventas_mensuales.py` | La lectura de FUDO y el armado del Excel (también se puede usar solo, en la Mac) |
| `requirements.txt` | Librerías (Flask, requests, waitress, tzdata) |
| `Procfile`, `railway.json` | Le indican a Railway cómo arrancar: `python app.py` |

## Puesta en marcha

1. **GitHub**: crea un repositorio **privado nuevo** (por ejemplo `schwencke-ventas-mensuales`) y sube
   estos archivos (*Add file → Upload files*). No subas claves ni un `config.py`.
2. **Railway**: *New Project → Deploy from GitHub repo* y elige el repositorio nuevo.
   Si no aparece, dale acceso en *Configure GitHub App*.
3. En el servicio → **Variables** → *RAW Editor*: copia desde el servicio actual **solo** estas líneas
   (`SUC1_…`, `SUC2_…`, `SUC3_…` y `DASHBOARD_PASSWORD`) y pégalas en el nuevo:

   ```
   SUC1_NOMBRE / SUC1_APIKEY / SUC1_APISECRET
   SUC2_NOMBRE / SUC2_APIKEY / SUC2_APISECRET
   SUC3_NOMBRE / SUC3_APIKEY / SUC3_APISECRET
   DASHBOARD_PASSWORD
   ```

4. *Settings → Networking → Generate Domain* para obtener la dirección.
5. Abre la dirección, ingresa con la clave, pulsa **Generar**, espera (se ve el avance por sucursal) y
   usa **Descargar Excel**.

## Variables

| Variable | Obligatoria | Para qué |
|---|---|---|
| `SUCn_NOMBRE`, `SUCn_APIKEY`, `SUCn_APISECRET` | sí | Una tanda por sucursal (n = 1, 2, 3…), igual que el dashboard |
| `DASHBOARD_PASSWORD` | sí | Clave de ingreso. Si falta, la página queda cerrada (no hay clave por defecto) |
| `SUCn_COLOR` | no | Color de la sucursal, p. ej. `#D85A30` |
| `ANIOS` | no | Años a leer; por defecto `2025,2026` (p. ej. `ANIOS=2026`) |
| `SECRET_KEY` | no | Firma de la sesión; si falta se deriva de la clave |
| `PORT` | no | Lo define Railway |

## Cómo funciona

- Cada sucursal se lee en paralelo, con una sola autenticación por lectura.
- Por defecto se salta el historial anterior al primer año (búsqueda binaria de páginas). Si la sección
  **Cobertura** avisa que falta información, marca **Leer todo el historial** y genera de nuevo.
- Puedes cerrar la página durante la lectura: sigue en segundo plano y el resultado queda disponible
  mientras el servicio no se reinicie (el resultado vive en memoria; nada se guarda en disco).
- Si una sucursal falla (credenciales, FUDO caído), su columna sale como `ERROR` y las demás quedan intactas.
- Tras 5 intentos fallidos de clave desde una misma IP, el ingreso se bloquea 10 minutos.

## Si algo falla

| Mensaje | Qué hacer |
|---|---|
| «Falta la clave de acceso» | Define `DASHBOARD_PASSWORD` en Variables |
| «No encontré sucursales configuradas» | Revisa que existan `SUC1_NOMBRE`, `SUC1_APIKEY` y `SUC1_APISECRET` |
| «FUDO rechazó las credenciales» | Revisa `SUCn_APIKEY` y `SUCn_APISECRET` de esa sucursal |
| «FUDO limitó las autenticaciones (429)» | Espera unos minutos y vuelve a generar |
| «la última venta leída es de hace N días» | Marca *Leer todo el historial* y genera de nuevo |
| La página no abre | En Railway → *Deployments* → *View logs*; revisa que el último deploy diga *Success* |

## Probarlo en tu computador (opcional)

```bash
pip install -r requirements.txt
export DASHBOARD_PASSWORD="una-clave"
export SUC1_NOMBRE="Viña del Mar" SUC1_APIKEY="…" SUC1_APISECRET="…"   # y SUC2_…, SUC3_…
python app.py        # http://localhost:8080
```

`ventas_mensuales.py` también funciona solo (usa un `config.py` con `SUCURSALES`):
`python3 ventas_mensuales.py` o `python3 ventas_mensuales.py --completo`.
