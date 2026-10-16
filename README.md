# Bot Inventarios IEEPO v2.0 Pro

![Bot IEEPO](https://img.shields.io/badge/Versi%C3%B3n-2.0%20Pro-blue) ![Python](https://img.shields.io/badge/Python-3.12-yellow) ![Playwright](https://img.shields.io/badge/Automatizaci%C3%B3n-Playwright-green)

Una aplicación de escritorio con interfaz gráfica (GUI) diseñada para automatizar la descarga masiva de **Constancias de Área Resguardante (PDFs)** desde el portal de control de inventarios del Instituto Estatal de Educación Pública de Oaxaca (IEEPO).

## 🚀 Características Principales

*   **Interfaz Gráfica Amigable:** Desarrollada con `tkinter`, permite a los usuarios configurar credenciales, rutas de descarga y ver el progreso en tiempo real sin tocar la consola.
*   **Automatización Robusta:** Utiliza `Playwright` para navegar por el sistema web basado en ExtJS del IEEPO, gestionando inicios de sesión, esperas dinámicas, menús desplegables complejos y ventanas emergentes de forma silenciosa (modo *headless*).
*   **Procesamiento por Lotes:** Lee las Claves de Centro de Trabajo (CT) directamente desde un archivo de Excel (`centros.xlsx`) y procesa cada una secuencialmente.
*   **Reportes en Tiempo Real:** Muestra una tabla visual con el estado de cada descarga (Éxito, Error) y genera automáticamente un reporte Excel final (`resumen_ejecucion.xlsx`) con los resultados.
*   **Gestión Segura:** Guarda las configuraciones de ruta y encripta/protege las contraseñas en un archivo `.env` local que no se sube a internet.

## 🛠️ ¿Cómo funciona?

1.  **Carga de Datos:** El usuario selecciona un archivo Excel que contenga una columna llamada `centro_trabajo` con las claves a procesar.
2.  **Configuración:** Ingresa su usuario y contraseña del IEEPO, y elige en qué carpeta guardar los PDFs.
3.  **Ejecución:** Al iniciar, el Bot abre una sesión privada, se autentica en el portal, navega al menú de *Mobiliario > Constancia de Área*, busca cada clave, simula las interacciones de teclado/ratón necesarias y captura el PDF nativo generado por el sistema.
4.  **Finalización:** Todos los PDFs quedan guardados con el nombre del CT (ej. `20DCC0001X.pdf`) en la carpeta elegida.

## 💻 Uso (Ejecutable de Windows)

El proyecto incluye un empaquetado `Bot_Inventarios_IEEPO.exe` que contiene todas las dependencias necesarias (incluyendo los navegadores). Solo basta con abrir el ejecutable para empezar a utilizar la herramienta.
