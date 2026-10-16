import os
import json
from pathlib import Path
from dotenv import load_dotenv

CONFIG_FILE = "config.json"
DEFAULT_URL = "https://portal-inventarios.ejemplo.gob.mx/#login"

def get_base_dir() -> Path:
    """Obtiene el directorio base de la aplicación."""
    import sys
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent.resolve()

def get_config_path() -> Path:
    return get_base_dir() / CONFIG_FILE

def get_env_path() -> Path:
    return get_base_dir() / ".env"

def load_config() -> dict:
    """Carga la configuración desde config.json y complementa con .env si existen."""
    base_dir = get_base_dir()
    default_download_dir = str(base_dir / "Descargas_Inventario")
    default_excel_path = str(base_dir / "centros.xlsx")

    config = {
        "url": DEFAULT_URL,
        "user": "",
        "password": "",
        "download_dir": default_download_dir,
        "excel_path": default_excel_path,
        "headless": True,
        "skip_existing": True,
        "auto_open_folder": True,
        "timeout_seconds": 30,
        "retries": 2
    }

    # Intentar cargar de .env
    env_file = get_env_path()
    if env_file.exists():
        load_dotenv(env_file)
        env_user = os.getenv("IEEPO_USER", "").strip()
        env_pass = os.getenv("IEEPO_PASS", "").strip()
        if env_user and env_user != "TU_USUARIO_AQUI":
            config["user"] = env_user
        if env_pass and env_pass != "TU_PASSWORD_AQUI":
            config["password"] = env_pass

    # Intentar cargar de config.json
    cfg_file = get_config_path()
    if cfg_file.exists():
        try:
            with open(cfg_file, "r", encoding="utf-8") as f:
                saved_cfg = json.load(f)
                config.update(saved_cfg)
        except Exception:
            pass

    # Asegurar que los directorios por defecto sean rutas válidas
    if not config.get("download_dir"):
        config["download_dir"] = default_download_dir
    if not config.get("excel_path"):
        config["excel_path"] = default_excel_path

    return config

def save_config(config_data: dict) -> bool:
    """Guarda la configuración en config.json y actualiza .env."""
    try:
        cfg_file = get_config_path()
        with open(cfg_file, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)

        # Actualizar .env para mantener compatibilidad
        env_file = get_env_path()
        user = config_data.get("user", "")
        pwd = config_data.get("password", "")
        with open(env_file, "w", encoding="utf-8") as f:
            f.write("# Credenciales para el sistema de Inventarios IEEPO\n")
            f.write(f"IEEPO_USER={user}\n")
            f.write(f"IEEPO_PASS={pwd}\n")

        return True
    except Exception as e:
        print(f"Error al guardar la configuración: {e}")
        return False
