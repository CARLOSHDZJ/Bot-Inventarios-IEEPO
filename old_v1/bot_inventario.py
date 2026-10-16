import os
import sys
import time
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# -----------------------------------------------------------------------------
# CONFIGURACIÓN Y CONSTANTES
# -----------------------------------------------------------------------------
load_dotenv()

# Credenciales de inicio de sesión
USUARIO = os.getenv("IEEPO_USER", "TU_USUARIO_AQUI")
PASSWORD = os.getenv("IEEPO_PASS", "TU_PASSWORD_AQUI")

# URLs y Rutas de trabajo
URL_LOGIN = "https://portal-inventarios.ejemplo.gob.mx/#login"
EXCEL_PATH = Path("centros.xlsx")

# Organización por carpetas con fecha del día (ej. Descargas_Inventario/2026-08-12)
FECHA_HOY = datetime.now().strftime("%Y-%m-%d")
CARPETA_BASE_DESCARGAS = Path("Descargas_Inventario")
CARPETA_DESCARGAS = CARPETA_BASE_DESCARGAS / FECHA_HOY
CARPETA_ERRORES = CARPETA_DESCARGAS / "errores"

# Parámetros de resiliencia
TIMEOUT_BASE = 30000       # 30 segundos de espera máxima por elementos ExtJS
ESPERA_CORTA = 2000        # Pausa de 2s entre transiciones
MAX_REINTENTOS_POR_CT = 2  # Reintentos por centro de trabajo si falla la red

# Crear carpetas si no existen
CARPETA_DESCARGAS.mkdir(parents=True, exist_ok=True)
CARPETA_ERRORES.mkdir(parents=True, exist_ok=True)

# Configuración del registro de eventos (Logging)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(CARPETA_DESCARGAS / "ejecucion_bot.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------
def cargar_centros_trabajo(ruta_excel: Path) -> list:
    """Lee el archivo de Excel y extrae la lista de Centros de Trabajo."""
    if not ruta_excel.exists():
        logger.error(f"El archivo Excel '{ruta_excel}' no existe.")
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel.resolve()}")

    df = pd.read_excel(ruta_excel)
    if df.empty:
        logger.warning("El archivo de Excel está vacío.")
        return []

    columnas_posibles = ["centro_trabajo", "centro", "cct", "area", "centro de trabajo"]
    columna_encontrada = None

    for col in df.columns:
        if str(col).strip().lower() in columnas_posibles:
            columna_encontrada = col
            break

    if not columna_encontrada:
        columna_encontrada = df.columns[0]
        logger.info(f"Usando la primera columna del Excel: '{columna_encontrada}'")
    else:
        logger.info(f"Columna identificada en Excel: '{columna_encontrada}'")

    centros = df[columna_encontrada].dropna().astype(str).str.strip().tolist()
    logger.info(f"Se cargaron {len(centros)} centros de trabajo para procesar.")
    return centros


def limpiar_nombre_archivo(cadena: str) -> str:
    """Reemplaza caracteres no válidos para guardar archivos en el sistema operativo."""
    import re
    return re.sub(r'[\\/*?:"<>|]', "_", cadena)


def ya_fue_descargado(centro: str, carpeta_destino: Path) -> bool:
    """Comprueba si el PDF del centro de trabajo ya fue descargado previamente."""
    nombre_limpio = limpiar_nombre_archivo(centro.strip().upper())
    
    # 1. Comprobar si existe <CENTRO>.pdf exacto en la carpeta del día o base
    if (carpeta_destino / f"{nombre_limpio}.pdf").exists():
        return True
    if (CARPETA_BASE_DESCARGAS / f"{nombre_limpio}.pdf").exists():
        return True

    # 2. Retrocompatibilidad con archivos antiguos (ej. Constancia_20dtv0001a_...)
    patron = f"*{nombre_limpio}*.pdf"
    coincidencias = list(carpeta_destino.glob(patron)) + list(CARPETA_BASE_DESCARGAS.glob(patron))
    for f in coincidencias:
        if f.is_file() and f.stat().st_size > 1000:
            return True
            
    return False


# -----------------------------------------------------------------------------
# BOT DE NAVEGACIÓN Y DESCARGA AUTOMATIZADA
# -----------------------------------------------------------------------------
def ejecutar_bot():
    """Flujo principal del bot con visión de supervisión y resiliencia mejorada."""
    logger.info("=== INICIANDO BOT DE DESCARGA DE REPORTES IEEPO (ExtJS PRO) ===")

    try:
        centros = cargar_centros_trabajo(EXCEL_PATH)
        if not centros:
            logger.error("No hay centros de trabajo para procesar. Abortando.")
            return
        
        logger.info(f"📁 CARPETA DE ALMACENAMIENTO:\n   {CARPETA_DESCARGAS.resolve()}")
    except Exception as e:
        logger.critical(f"Error al leer el archivo Excel: {e}")
        return

    ruta_absoluta_descargas = CARPETA_DESCARGAS.resolve()

    with sync_playwright() as p:
        # EJECUCIÓN SILENCIOSA EN SEGUNDO PLANO (headless=True) - Solo visible desde la consola CMD
        logger.info("Lanzando navegador Chromium en modo silencioso (segundo plano, headless=True)...")
        browser = p.chromium.launch(
            headless=True,
            downloads_path=str(ruta_absoluta_descargas),
            args=[
                "--disable-pdf-extension", # Forzar descarga directa de PDFs
                "--no-sandbox"
            ]
        )
        
        context = browser.new_context(
            no_viewport=True,
            accept_downloads=True
        )
        
        page = context.new_page()
        page.set_default_timeout(TIMEOUT_BASE)

        # ---------------------------------------------------------------------
        # PASO A: INICIO DE SESIÓN
        # ---------------------------------------------------------------------
        try:
            logger.info(f"Navegando a URL: {URL_LOGIN}")
            page.goto(URL_LOGIN, wait_until="networkidle")

            logger.info("Ingresando credenciales...")
            selector_usuario = page.locator(
                "input[name='username'], input[name='usuario'], input[name='login'], input[type='text']"
            ).first
            selector_usuario.wait_for(state="visible", timeout=TIMEOUT_BASE)
            selector_usuario.fill(USUARIO)

            selector_pass = page.locator(
                "input[name='password'], input[name='contrasena'], input[name='clave'], input[type='password']"
            ).first
            selector_pass.wait_for(state="visible", timeout=TIMEOUT_BASE)
            selector_pass.fill(PASSWORD)

            logger.info("Enviando formulario de inicio de sesión...")
            btn_login = page.locator(
                ".x-button:has-text('Login'), button:has-text('Login'), .x-btn:has-text('Login'), div:has-text('Login'), text='Login', button:has-text('Iniciar'), input[type='submit']"
            ).first
            
            try:
                btn_login.wait_for(state="visible", timeout=5000)
                btn_login.click()
            except Exception:
                selector_pass.press("Enter")

            page.wait_for_load_state("networkidle")
            time.sleep(ESPERA_CORTA / 1000)
            logger.info("Inicio de sesión completado correctamente.")

        except PlaywrightTimeoutError:
            logger.error("Tiempo de espera agotado durante el Inicio de Sesión.")
            page.screenshot(path=CARPETA_ERRORES / "error_login_timeout.png")
            browser.close()
            return
        except Exception as e:
            logger.error(f"Error inesperado en Inicio de Sesión: {e}")
            page.screenshot(path=CARPETA_ERRORES / "error_login_general.png")
            browser.close()
            return

        # ---------------------------------------------------------------------
        # PASO B: NAVEGACIÓN EN MENÚS (MOBILIARIO -> CONSTANCIA DE ÁREA)
        # ---------------------------------------------------------------------
        try:
            import re
            logger.info("Navegando por menú 'Mobiliario'...")
            menu_mobiliario = page.locator("text='Mobiliario'").first
            menu_mobiliario.wait_for(state="visible", timeout=TIMEOUT_BASE)
            menu_mobiliario.hover()
            time.sleep(1)

            locator_sub = page.get_by_text(re.compile(r"Constancia\s+de\s+[AÁ]rea", re.IGNORECASE))

            clic_exitoso = False
            for intento in range(1, 4):
                try:
                    if not locator_sub.first.is_visible(timeout=3000):
                        menu_mobiliario.hover()
                        time.sleep(1)
                    
                    opcion_sub = locator_sub.first
                    opcion_sub.hover()
                    opcion_sub.click(force=True)
                    clic_exitoso = True
                    break
                except Exception:
                    time.sleep(1)

            if not clic_exitoso:
                locator_sub.first.evaluate("el => el.click()")

            header_vista = page.get_by_text(re.compile(r"Constancia\s+de\s+[AÁ]rea\s+Resguardante", re.IGNORECASE)).first
            header_vista.wait_for(state="visible", timeout=TIMEOUT_BASE)
            logger.info("Formulario 'Constancia de Área Resguardante' cargado.")
            page.mouse.move(0, 0)
            time.sleep(ESPERA_CORTA / 1000)

        except Exception as e:
            logger.error(f"Error al navegar por el menú: {e}")
            page.screenshot(path=CARPETA_ERRORES / "error_navegacion_menu.png")
            browser.close()
            return

        # ---------------------------------------------------------------------
        # PASO C: BUCLE DE DESCARGA CON OMISIÓN Y REINTENTOS
        # ---------------------------------------------------------------------
        exitosos = 0
        omitidos = 0
        fallidos = 0
        resumen_procesamiento = []

        logger.info(f"Iniciando procesamiento secuencial de {len(centros)} centros de trabajo...")

        for idx, centro in enumerate(centros, start=1):
            logger.info(f"\n==================================================")
            logger.info(f"--- [{idx}/{len(centros)}] Procesando Centro de Trabajo: {centro} ---")
            logger.info(f"==================================================")

            # MEJORA: Verificar si ya fue descargado previamente
            if ya_fue_descargado(centro, CARPETA_DESCARGAS):
                logger.info(f"⏩ [OMITIDO] El reporte para '{centro}' ya existe en la carpeta. Saltando...")
                omitidos += 1
                resumen_procesamiento.append({"centro": centro, "estado": "OMITIDO (YA EXISTE)", "archivo": "Previamente descargado"})
                continue

            # MEJORA: Sistema de reintentos por centro de trabajo
            ct_procesado_exito = False

            for intento_ct in range(1, MAX_REINTENTOS_POR_CT + 1):
                if intento_ct > 1:
                    logger.warning(f"🔄 Reintentando centro '{centro}' (Intento {intento_ct}/{MAX_REINTENTOS_POR_CT})...")

                try:
                    # 1. Localizar y limpiar campo Área
                    campo_area = page.locator(
                        "input[role='combobox'], input[name='area'], label:has-text('Área') + div input, label:has-text('Área*') + div input, input.x-form-text"
                    ).first
                    
                    campo_area.wait_for(state="visible", timeout=15000)
                    campo_area.click()
                    campo_area.fill("")
                    page.keyboard.press("Control+A")
                    page.keyboard.press("Backspace")
                    time.sleep(0.5)

                    # 2. Escribir clave y seleccionar de lista desplegable
                    logger.info(f"1. Escribiendo la clave '{centro}'...")
                    campo_area.type(centro, delay=90)
                    time.sleep(1.5)

                    logger.info(f"2. Seleccionando de la lista desplegable...")
                    item_desplegable = page.locator(
                        f".x-boundlist-item:has-text('{centro}'), .x-list-item:has-text('{centro}'), div.x-boundlist-item:has-text('{centro}')"
                    ).first

                    if item_desplegable.is_visible(timeout=4000):
                        item_desplegable.click(force=True)
                    else:
                        opcion_alt = page.get_by_text(re.compile(re.escape(centro), re.IGNORECASE)).last
                        if opcion_alt.is_visible(timeout=3000):
                            opcion_alt.click(force=True)
                        else:
                            page.keyboard.press("ArrowDown")
                            page.keyboard.press("Enter")

                    time.sleep(2)  # Carga de datos del resguardante

                    # 3. Botón Reporte
                    logger.info("3. Buscando el botón REPORTE...")
                    btn_reporte = page.locator(
                        "a.x-btn:has-text('Reporte'), button:has-text('Reporte'), span.x-btn-inner:has-text('Reporte'), .x-btn:has-text('Reporte')"
                    ).first
                    btn_reporte.wait_for(state="visible", timeout=10000)

                    # Nombre del PDF únicamente con la clave del CT en MAYÚSCULAS (ej. 20DTV0001A.pdf)
                    nombre_pdf_limpio = limpiar_nombre_archivo(centro.strip().upper())
                    nombre_archivo = f"{nombre_pdf_limpio}.pdf"
                    ruta_guardado = CARPETA_DESCARGAS / nombre_archivo

                    # 4 y 5. Clic en Reporte, capturar descarga nativa (headless) o pestaña emergente (headful)
                    logger.info("4. Solicitando reporte...")
                    download_exitoso = False

                    try:
                        # Prioridad 1: Captura de descarga nativa en modo headless (instantáneo)
                        with page.expect_download(timeout=10000) as download_info:
                            btn_reporte.click(force=True)

                        download = download_info.value
                        download.save_as(ruta_guardado)
                        download_exitoso = True
                        logger.info(f"   -> [OK] Descarga nativa completada e identificada como: {nombre_archivo}")

                    except PlaywrightTimeoutError:
                        # Prioridad 2: Si no fue descarga directa, comprobar pestaña emergente
                        logger.info("   -> Comprobando si se abrió pestaña emergente...")
                        pages_actuales = context.pages
                        if len(pages_actuales) > 1:
                            new_tab = pages_actuales[-1]
                            new_tab.wait_for_load_state("load", timeout=10000)
                            tab_url = new_tab.url
                            logger.info(f"   -> URL de pestaña detectada: {tab_url}")

                            res = new_tab.request.get(tab_url)
                            if res.ok and len(res.body()) > 500:
                                ruta_guardado.write_bytes(res.body())
                                download_exitoso = True
                                logger.info(f"   -> [OK] PDF extraído de pestaña y guardado como: {nombre_archivo}")

                            new_tab.close()

                    if download_exitoso:
                        logger.info(f"[ÉXITO] ✓ Reporte para '{centro}' guardado en:\n        {ruta_guardado}")
                        exitosos += 1
                        resumen_procesamiento.append({"centro": centro, "estado": "ÉXITO", "archivo": nombre_archivo})
                        ct_procesado_exito = True
                        break
                    else:
                        raise Exception("No se pudo obtener la descarga del archivo PDF.")

                except PlaywrightTimeoutError:
                    logger.warning(f"⚠️ Timeout en intento {intento_ct} para '{centro}'.")
                    if intento_ct == MAX_REINTENTOS_POR_CT:
                        foto_error = CARPETA_ERRORES / f"timeout_{limpiar_nombre_archivo(centro)}.png"
                        page.screenshot(path=foto_error)
                        fallidos += 1
                        resumen_procesamiento.append({"centro": centro, "estado": "TIMEOUT", "archivo": None})

                except Exception as ex:
                    logger.warning(f"⚠️ Error en intento {intento_ct} para '{centro}': {ex}")
                    if intento_ct == MAX_REINTENTOS_POR_CT:
                        foto_error = CARPETA_ERRORES / f"error_{limpiar_nombre_archivo(centro)}.png"
                        page.screenshot(path=foto_error)
                        fallidos += 1
                        resumen_procesamiento.append({"centro": centro, "estado": f"ERROR: {str(ex)}", "archivo": None})

                finally:
                    # Limpiar campo Área para el siguiente intento/centro
                    try:
                        campo_area = page.locator(
                            "input[role='combobox'], input[name='area'], label:has-text('Área') + div input, input.x-form-text"
                        ).first
                        if campo_area.is_visible():
                            campo_area.click()
                            campo_area.fill("")
                            page.keyboard.press("Control+A")
                            page.keyboard.press("Backspace")
                    except Exception:
                        pass

        # ---------------------------------------------------------------------
        # PASO D: RESUMEN Y APERTURA DE CARPETA
        # ---------------------------------------------------------------------
        logger.info("\n==================================================")
        logger.info("=== RESUMEN FINAL DE LA EJECUCIÓN ===")
        logger.info(f"Total de centros en Excel : {len(centros)}")
        logger.info(f"Descargas Exitosas        : {exitosos}")
        logger.info(f"Omitidos (Ya Existían)    : {omitidos}")
        logger.info(f"Descargas Fallidas        : {fallidos}")
        logger.info("==================================================")

        df_resumen = pd.DataFrame(resumen_procesamiento)
        df_resumen.to_excel(CARPETA_DESCARGAS / "resumen_ejecucion.xlsx", index=False)
        logger.info(f"Reporte de ejecución guardado en: {CARPETA_DESCARGAS / 'resumen_ejecucion.xlsx'}")

        # Abrir automáticamente el Explorador de Windows en la carpeta del día
        try:
            logger.info("Abriendo la carpeta de descargas en el Explorador de Windows...")
            os.startfile(CARPETA_DESCARGAS.resolve())
        except Exception:
            pass

        time.sleep(2)
        browser.close()


if __name__ == "__main__":
    ejecutar_bot()
