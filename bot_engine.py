import os
import sys
import time
import re
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

def limpiar_nombre_archivo(cadena: str) -> str:
    """Limpia caracteres no válidos para nombres de archivo en Windows."""
    return re.sub(r'[\\/*?:"<>|]', "_", str(cadena).strip())

def generar_plantilla_excel(ruta_excel: Path, sobreescribir: bool = False) -> tuple[bool, str]:
    """Genera un archivo centros.xlsx con ejemplos de Claves de Centro de Trabajo."""
    try:
        ruta = Path(ruta_excel)
        if ruta.exists() and not sobreescribir:
            return True, f"El archivo ya existe en: {ruta.resolve()}"

        ruta.parent.mkdir(parents=True, exist_ok=True)
        datos = {
            "centro_trabajo": [
                "20DTV0001A",
                "20DPR0002B",
                "20DJN0003C",
                "20DES0004D"
            ]
        }
        df = pd.DataFrame(datos)
        df.to_excel(ruta, index=False)
        return True, f"Plantilla creada exitosamente en:\n{ruta.resolve()}"
    except Exception as e:
        return False, f"Error al generar plantilla Excel: {e}"

def leer_centros_excel(ruta_excel: Path) -> tuple[list[str], str]:
    """Lee el archivo de Excel y extrae la lista de Centros de Trabajo."""
    ruta = Path(ruta_excel)
    if not ruta.exists():
        return [], f"No se encontró el archivo: {ruta.resolve()}"

    try:
        df = pd.read_excel(ruta)
        if df.empty:
            return [], "El archivo de Excel está vacío."

        columnas_posibles = ["centro_trabajo", "centro", "cct", "area", "centro de trabajo", "clave"]
        columna_encontrada = None

        for col in df.columns:
            if str(col).strip().lower() in columnas_posibles:
                columna_encontrada = col
                break

        if not columna_encontrada:
            columna_encontrada = df.columns[0]

        centros_raw = df[columna_encontrada].dropna().astype(str).str.strip().tolist()
        # Filtrar valores vacíos o nulos
        centros = [c for c in centros_raw if c and c.lower() != "nan" and len(c) >= 3]
        return centros, f"Se cargaron {len(centros)} centros (columna '{columna_encontrada}')."
    except Exception as e:
        return [], f"Error al leer el archivo Excel: {e}"

def ya_fue_descargado(centro: str, carpeta_destino: Path) -> bool:
    """Comprueba si el PDF del centro de trabajo ya fue descargado previamente."""
    nombre_limpio = limpiar_nombre_archivo(centro.strip().upper())
    
    # 1. Comprobar <CENTRO>.pdf exacto
    archivo_directo = carpeta_destino / f"{nombre_limpio}.pdf"
    if archivo_directo.exists() and archivo_directo.stat().st_size > 1000:
        return True

    # 2. Coincidencia por patrón (retrocompatibilidad)
    patron = f"*{nombre_limpio}*.pdf"
    for f in carpeta_destino.glob(patron):
        if f.is_file() and f.stat().st_size > 1000:
            return True

    return False

def test_connection_and_credentials(url: str, user: str, password: str, log_cb=None) -> tuple[bool, str]:
    """Prueba la conexión al portal IEEPO e intenta autenticar con las credenciales dadas."""
    def log(lvl, msg):
        if log_cb:
            log_cb(lvl, msg)
        else:
            print(f"[{lvl}] {msg}")

    log("INFO", f"Iniciando prueba de conexión con: {url}")
    if not user or not password:
        err = "El usuario o la contraseña están vacíos. Por favor ingresa ambos."
        log("ERROR", err)
        return False, err

    try:
        with sync_playwright() as p:
            log("INFO", "Iniciando navegador Chromium para prueba rápida...")
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-pdf-extension", "--no-sandbox"]
            )
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()
            page.set_default_timeout(25000)

            log("INFO", f"Conectando a {url}...")
            try:
                page.goto(url, wait_until="networkidle", timeout=25000)
            except Exception as e:
                # Si falla networkidle, intentar con domcontentloaded
                log("WARNING", f"networkidle demoró, reintentando con carga DOM: {e}")
                page.goto(url, wait_until="domcontentloaded", timeout=25000)

            log("INFO", "Esperando campos de inicio de sesión de ExtJS...")
            page.wait_for_selector("input[name='username'], input[name='usuario'], input[type='text']", timeout=15000)
            time.sleep(1)

            user_input = page.locator("input[name='username'], input[name='usuario'], input[type='text']").first
            pass_input = page.locator("input[name='password'], input[name='contrasena'], input[type='password']").first

            user_input.fill(user)
            pass_input.fill(password)
            log("INFO", f"Credenciales ingresadas (Usuario: {user}). Enviando formulario...")

            btn_login = page.locator(".x-btn-inner:has-text('Login'), button:has-text('Login'), span:has-text('Login'), input[type='submit']").first
            if btn_login.is_visible():
                btn_login.click()
            else:
                pass_input.press("Enter")

            # Esperar respuesta
            time.sleep(3)
            current_url = page.url
            log("INFO", f"URL posterior al login: {current_url}")

            # Comprobar si hay ventanas emergentes de error de ExtJS
            error_box = page.locator(".x-message-box, .x-window-header-text:has-text('Error')")
            if error_box.count() > 0 and error_box.first.is_visible():
                err_text = page.locator(".x-window-text").first.inner_text() if page.locator(".x-window-text").count() > 0 else "Credenciales incorrectas"
                log("ERROR", f"El portal devolvió un mensaje de error: {err_text}")
                browser.close()
                return False, f"Fallo de autenticación: {err_text}"

            # Verificar si cargó la pantalla principal (#home o menú Mobiliario)
            if "#home" in current_url or page.locator("text='Mobiliario'").count() > 0:
                log("SUCCESS", "¡Conexión y Autenticación EXITOSAS! Las credenciales son válidas y el sistema respondió correctamente.")
                browser.close()
                return True, "¡Conexión y autenticación exitosas en el portal IEEPO!"
            
            # Si no está en #home pero tampoco hay error explícito, esperar 2 segundos más
            time.sleep(2)
            if page.locator("text='Mobiliario'").count() > 0 or "#home" in page.url:
                log("SUCCESS", "¡Conexión y Autenticación EXITOSAS!")
                browser.close()
                return True, "¡Conexión y autenticación exitosas!"

            browser.close()
            log("WARNING", "El login finalizó pero no se detectó la pantalla principal #home.")
            return False, "No se detectó confirmación de inicio de sesión exitoso."

    except PlaywrightTimeoutError:
        err = "Tiempo de espera agotado al intentar conectar con el portal IEEPO."
        log("ERROR", err)
        return False, err
    except Exception as ex:
        err = f"Error durante la prueba de conexión: {ex}"
        log("ERROR", err)
        return False, err

def ejecutar_proceso_descargas(
    config: dict,
    log_cb,
    progress_cb,
    item_status_cb,
    stop_event=None
) -> dict:
    """
    Ejecuta el ciclo completo de descarga de inventarios para todos los CTs.
    No cierra la app ante errores, reporta todo en tiempo real.
    """
    def log(lvl, msg):
        if log_cb:
            log_cb(lvl, msg)
        else:
            print(f"[{lvl}] {msg}")

    url = config.get("url", "https://portal-inventarios.ejemplo.gob.mx/#login")
    user = config.get("user", "")
    password = config.get("password", "")
    carpeta_destino = Path(config.get("download_dir", "Descargas_Inventario")).resolve()
    ruta_excel = Path(config.get("excel_path", "centros.xlsx")).resolve()
    modo_headless = config.get("headless", True)
    omitir_existentes = config.get("skip_existing", True)
    timeout_base = int(config.get("timeout_seconds", 30)) * 1000
    max_reintentos = int(config.get("retries", 2))

    carpeta_destino.mkdir(parents=True, exist_ok=True)
    carpeta_errores = carpeta_destino / "errores"
    carpeta_errores.mkdir(parents=True, exist_ok=True)

    # Configurar archivo log en la carpeta de descarga
    file_handler = logging.FileHandler(carpeta_destino / "ejecucion_bot.log", encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    file_logger = logging.getLogger("IEEPOBotLogger")
    file_logger.setLevel(logging.INFO)
    file_logger.addHandler(file_handler)

    log("INFO", "==================================================")
    log("INFO", "🚀 INICIANDO BOT DE INVENTARIOS IEEPO v2.0 PRO")
    log("INFO", f"📁 Carpeta de destino: {carpeta_destino}")
    log("INFO", f"📄 Archivo de centros: {ruta_excel}")
    log("INFO", f"🌐 Modo navegador: {'Segundo plano (Silencioso)' if modo_headless else 'Visible (Supervisión)'}")
    log("INFO", "==================================================")

    # 1. Cargar centros
    centros, msg = leer_centros_excel(ruta_excel)
    log("INFO", msg)
    if not centros:
        log("ERROR", "No se encontraron Claves de Centro de Trabajo para procesar.")
        return {"total": 0, "exitosos": 0, "omitidos": 0, "fallidos": 0, "cancelado": False}

    total_centros = len(centros)
    exitosos = 0
    omitidos = 0
    fallidos = 0
    resumen_procesamiento = []

    if progress_cb:
        progress_cb(0, total_centros, 0, f"Iniciando navegador (0/{total_centros})...")

    # 2. Inicializar Playwright
    try:
        with sync_playwright() as p:
            log("INFO", "Lanzando navegador Chromium...")
            args_launch = ["--disable-pdf-extension", "--no-sandbox"]
            if not modo_headless:
                args_launch.append("--start-maximized")

            browser = p.chromium.launch(
                headless=modo_headless,
                downloads_path=str(carpeta_destino),
                args=args_launch
            )

            context = browser.new_context(
                no_viewport=True if not modo_headless else False,
                accept_downloads=True,
                ignore_https_errors=True
            )

            page = context.new_page()
            page.set_default_timeout(timeout_base)

            # -----------------------------------------------------------------
            # LOGIN
            # -----------------------------------------------------------------
            if stop_event and stop_event.is_set():
                log("WARNING", "Proceso cancelado por el usuario antes de iniciar sesión.")
                browser.close()
                return {"total": total_centros, "exitosos": 0, "omitidos": 0, "fallidos": 0, "cancelado": True}

            log("INFO", f"Navegando a {url}...")
            try:
                page.goto(url, wait_until="networkidle", timeout=timeout_base)
            except Exception:
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_base)

            log("INFO", "Ingresando usuario y contraseña...")
            user_input = page.locator("input[name='username'], input[name='usuario'], input[type='text']").first
            user_input.wait_for(state="visible", timeout=timeout_base)
            user_input.fill(user)

            pass_input = page.locator("input[name='password'], input[name='contrasena'], input[type='password']").first
            pass_input.wait_for(state="visible", timeout=timeout_base)
            pass_input.fill(password)

            log("INFO", "Enviando inicio de sesión...")
            btn_login = page.locator(".x-btn-inner:has-text('Login'), button:has-text('Login'), span:has-text('Login'), input[type='submit']").first
            if btn_login.is_visible():
                btn_login.click()
            else:
                pass_input.press("Enter")

            time.sleep(3)
            page.wait_for_load_state("networkidle")
            log("SUCCESS", "Sesión iniciada con éxito.")

            # -----------------------------------------------------------------
            # NAVEGACIÓN EN MENÚ: Mobiliario -> Constancia de Área
            # -----------------------------------------------------------------
            if stop_event and stop_event.is_set():
                log("WARNING", "Proceso cancelado por el usuario.")
                browser.close()
                return {"total": total_centros, "exitosos": 0, "omitidos": 0, "fallidos": 0, "cancelado": True}

            log("INFO", "Abriendo menú 'Mobiliario' -> 'Constancia de Área'...")
            menu_mob = page.locator("text='Mobiliario'").first
            menu_mob.wait_for(state="visible", timeout=timeout_base)
            menu_mob.click()
            time.sleep(1)

            # Seleccionar 'Constancia de Área'
            item_constancia = page.locator("text=/Constancia de .?rea/i").first
            item_constancia.wait_for(state="visible", timeout=10000)
            item_constancia.click()
            time.sleep(2)

            # Mover el ratón a la esquina para disipar cualquier tooltip/overlay residual
            page.mouse.move(0, 0)
            time.sleep(1)

            log("SUCCESS", "Formulario 'Constancia de Área' cargado correctamente.")

            # -----------------------------------------------------------------
            # CICLO DE DESCARGAS
            # -----------------------------------------------------------------
            for idx, centro in enumerate(centros, start=1):
                if stop_event and stop_event.is_set():
                    log("WARNING", f"Proceso detenido por el usuario en el centro #{idx} ({centro}).")
                    break

                pct = int((idx - 1) / total_centros * 100)
                if progress_cb:
                    progress_cb(idx - 1, total_centros, pct, f"Procesando [{idx}/{total_centros}]: {centro}")

                log("INFO", f"--------------------------------------------------")
                log("INFO", f"[{idx}/{total_centros}] Procesando Centro: {centro}")

                # Verificar si ya existe
                if omitir_existentes and ya_fue_descargado(centro, carpeta_destino):
                    nombre_archivo = f"{limpiar_nombre_archivo(centro.upper())}.pdf"
                    log("INFO", f"⏩ [OMITIDO] El archivo '{nombre_archivo}' ya existe en la carpeta.")
                    omitidos += 1
                    resumen_procesamiento.append({"centro": centro.upper(), "estado": "OMITIDO (YA EXISTE)", "archivo": nombre_archivo})
                    if item_status_cb:
                        item_status_cb(idx - 1, centro.upper(), "OMITIDO", str(carpeta_destino / nombre_archivo), "Ya existía previamente")
                    continue

                ct_exitoso = False
                nombre_pdf_limpio = limpiar_nombre_archivo(centro.strip().upper())
                nombre_archivo = f"{nombre_pdf_limpio}.pdf"
                ruta_guardado = carpeta_destino / nombre_archivo

                for intento in range(1, max_reintentos + 1):
                    if stop_event and stop_event.is_set():
                        break

                    if intento > 1:
                        log("WARNING", f"🔄 Reintentando '{centro}' (Intento {intento}/{max_reintentos})...")

                    try:
                        # Asegurar que el mouse no tape ningún elemento
                        page.mouse.move(0, 0)

                        # 1. Campo Área
                        campo_area = page.locator("input[name='idarea'], input[role='combobox'], label:has-text('Área') + div input, label:has-text('Area') + div input").first
                        campo_area.wait_for(state="visible", timeout=15000)

                        # Bucle de reintentos para asegurar que el valor quede correctamente seleccionado
                        valor_correcto = False
                        for intento_sel in range(3):
                            campo_area.click(force=True)
                            page.keyboard.press("Control+A")
                            page.keyboard.press("Backspace")
                            time.sleep(1.0) # Dar tiempo a que reaccione al borrado

                            # 2. Escribir clave de CT
                            campo_area.type(centro, delay=100)
                            
                            # 3. Seleccionar de lista desplegable (Espera estricta al elemento exacto)
                            item_desplegable = page.locator(f".x-boundlist-item:has-text('{centro.upper()}'), .x-boundlist-item:has-text('{centro}')").first
                            
                            try:
                                # Esperar a que la red traiga el resultado y se muestre en la lista
                                item_desplegable.wait_for(state="visible", timeout=6000)
                                
                                # Click directo (force=True para evitar problemas de headless)
                                item_desplegable.click(force=True)
                                time.sleep(1.0)
                                
                                # Presionar Tab para obligar a ExtJS a disparar el evento "blur" y guardar el valor interno
                                page.keyboard.press("Tab")
                                time.sleep(1.0)

                                # Verificar si el valor actual coincide
                                valor_actual = campo_area.input_value()
                                if valor_actual and centro.upper() in valor_actual.upper():
                                    valor_correcto = True
                                    log("INFO", f"Centro '{centro}' seleccionado correctamente en el desplegable.")
                                    break
                                else:
                                    log("WARNING", f"Se hizo click pero el valor actual '{valor_actual}' no coincide con '{centro}'. Reintentando...")
                            except Exception:
                                log("WARNING", f"La lista desplegable no cargó el centro '{centro}'. Reintentando...")

                        if not valor_correcto:
                            log("WARNING", f"No se pudo fijar de manera segura el valor '{centro}' después de 3 intentos.")

                        # Pausa final de seguridad para que ExtJS asimile el cambio antes de descargar
                        time.sleep(3.0)

                        # 4. Localizar botón Reporte
                        btn_reporte = page.locator(
                            "a.x-btn:has-text('Reporte'), button:has-text('Reporte'), span.x-btn-inner:has-text('Reporte'), .x-btn:has-text('Reporte')"
                        ).first
                        btn_reporte.wait_for(state="visible", timeout=10000)

                        # 5. Solicitar reporte y gestionar descarga
                        download_ok = False
                        
                        try:
                            # En modo headless Chromium convierte la navegación del PDF en una descarga nativa
                            with page.expect_download(timeout=15000) as download_info:
                                btn_reporte.click(force=True)
                            
                            download = download_info.value
                            download.save_as(ruta_guardado)
                            download_ok = True
                            time.sleep(1.0) # Breve pausa tras guardar para evitar pisar archivos
                            
                            
                        except PlaywrightTimeoutError:
                            # Fallback si por alguna razón abre un popup o no descarga
                            log("WARNING", "No se detectó evento de descarga nativa. Verificando páginas actuales...")
                            pages_actuales = context.pages
                            if len(pages_actuales) > 1:
                                new_tab = pages_actuales[-1]
                                tab_url = new_tab.url
                                if "pdf" in tab_url.lower():
                                    res = new_tab.request.get(tab_url)
                                    if res.ok and len(res.body()) > 500:
                                        ruta_guardado.write_bytes(res.body())
                                        download_ok = True
                                new_tab.close()
                        
                        if download_ok and ruta_guardado.exists() and ruta_guardado.stat().st_size > 500:
                            log("SUCCESS", f"✓ [ÉXITO] Descargado: {nombre_archivo} ({ruta_guardado.stat().st_size // 1024} KB)")
                            exitosos += 1
                            ct_exitoso = True
                            resumen_procesamiento.append({"centro": centro.upper(), "estado": "ÉXITO", "archivo": nombre_archivo})
                            if item_status_cb:
                                item_status_cb(idx - 1, centro.upper(), "ÉXITO", str(ruta_guardado), "Descargado correctamente")
                            break
                        else:
                            raise Exception("No se generó el archivo PDF o está incompleto.")

                    except PlaywrightTimeoutError:
                        log("WARNING", f"Timeout en intento {intento} para '{centro}'.")
                        if intento == max_reintentos:
                            foto_error = carpeta_errores / f"timeout_{nombre_pdf_limpio}.png"
                            try:
                                page.screenshot(path=foto_error)
                            except Exception:
                                pass
                            fallidos += 1
                            resumen_procesamiento.append({"centro": centro.upper(), "estado": "TIMEOUT", "archivo": None})
                            if item_status_cb:
                                item_status_cb(idx - 1, centro.upper(), "ERROR", "", "Tiempo de espera agotado")

                    except Exception as ex:
                        log("WARNING", f"Error en intento {intento} para '{centro}': {ex}")
                        if intento == max_reintentos:
                            foto_error = carpeta_errores / f"error_{nombre_pdf_limpio}.png"
                            try:
                                page.screenshot(path=foto_error)
                            except Exception:
                                pass
                            fallidos += 1
                            resumen_procesamiento.append({"centro": centro.upper(), "estado": f"ERROR: {str(ex)}", "archivo": None})
                            if item_status_cb:
                                item_status_cb(idx - 1, centro.upper(), "ERROR", "", str(ex)[:40])

                    finally:
                        # Limpiar campo para la siguiente iteración
                        try:
                            campo_area = page.locator("input[name='idarea'], input[role='combobox']").first
                            if campo_area.is_visible():
                                campo_area.click(force=True)
                                campo_area.fill("")
                        except Exception:
                            pass

                time.sleep(0.8)

            # Limpiar archivos temporales de descarga de Chromium
            for f in carpeta_destino.glob("*"):
                if f.is_file() and not f.name.endswith(".pdf") and not f.name.endswith(".xlsx") and not f.name.endswith(".log") and not f.name.endswith(".json") and not f.name.endswith(".env"):
                    try:
                        f.unlink()
                    except Exception:
                        pass

            # Generar Excel de resumen
            df_resumen = pd.DataFrame(resumen_procesamiento)
            ruta_resumen = carpeta_destino / "resumen_ejecucion.xlsx"
            df_resumen.to_excel(ruta_resumen, index=False)
            log("INFO", f"Reporte de resumen guardado en: {ruta_resumen}")

            browser.close()

    except Exception as e:
        log("ERROR", f"Error crítico en el proceso de descarga: {e}")

    final_pct = 100
    cancelado = stop_event.is_set() if stop_event else False
    estado_final_txt = "Proceso cancelado por el usuario" if cancelado else "Proceso de descarga completado"
    if progress_cb:
        progress_cb(total_centros, total_centros, final_pct, estado_final_txt)

    log("INFO", "==================================================")
    log("INFO", f"=== RESUMEN FINAL: {estado_final_txt.upper()} ===")
    log("INFO", f"Total de centros en Excel : {total_centros}")
    log("INFO", f"Descargas Exitosas        : {exitosos}")
    log("INFO", f"Omitidos (Ya Existían)    : {omitidos}")
    log("INFO", f"Descargas Fallidas        : {fallidos}")
    log("INFO", "==================================================")

    return {
        "total": total_centros,
        "exitosos": exitosos,
        "omitidos": omitidos,
        "fallidos": fallidos,
        "cancelado": cancelado,
        "carpeta": str(carpeta_destino),
        "resumen_excel": str(carpeta_destino / "resumen_ejecucion.xlsx")
    }
