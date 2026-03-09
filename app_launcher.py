# app_launcher.py (OPÇÃO A - simples e direta, com mensagens de arranque)
import os
import sys
import time
import warnings

# (Opcional) console mais limpo
warnings.filterwarnings("ignore")

# Configurações opcionais de UX do arranque
SHOW_START_MESSAGE = True
INITIAL_PAUSE_SECONDS = 0.6  # pequena pausa para a mensagem aparecer antes do Streamlit

def run():
    # Working dir correto (PyInstaller onefile/onedir)
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    os.chdir(base_dir)

    # Caminho do seu app
    script_path = os.path.join(base_dir, "main.py")

    # SSL em Windows (mitiga aviso)
    try:
        import certifi  # type: ignore
        os.environ["SSL_CERT_FILE"] = certifi.where()
    except Exception:
        pass

    # Produção (sem Node dev server)
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "false"  # abrir navegador

    # Mensagem amigável no console antes do Streamlit assumir
    if SHOW_START_MESSAGE:
        print("\nIniciando a aplicação…")
        print("Aguarde enquanto o servidor prepara a interface (porta 8501).")
        print("Dica: a primeira inicialização pode levar alguns segundos.\n")
        # pequena pausa apenas para que o usuário leia a mensagem
        if INITIAL_PAUSE_SECONDS > 0:
            time.sleep(INITIAL_PAUSE_SECONDS)

    # Executa Streamlit no mesmo processo
    try:
        from streamlit.web import cli as stcli
    except Exception:
        import streamlit.web.bootstrap as stcli  # fallback

    # Importante: evitar flags na argv quando já setamos ENV VARs
    sys.argv = ["streamlit", "run", script_path]

    # A partir daqui, o Streamlit assume o controle do processo/console
    sys.exit(stcli.main())


if __name__ == "__main__":
    run()