import os
import sys
import subprocess
import webbrowser
from threading import Timer
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env (se existir)
load_dotenv()

from app import create_app

app = create_app()

def free_port(port=5000):
    try:
        current_pid = str(os.getpid())
        cmd = f'Get-NetTCPConnection -LocalPort {port} -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess'
        res = subprocess.run(['powershell', '-Command', cmd], capture_output=True, text=True)
        pids = set(res.stdout.strip().split())
        for pid_str in pids:
            if pid_str and pid_str != current_pid and pid_str != '0' and pid_str.isdigit() and int(pid_str) > 0:
                subprocess.run(['powershell', '-Command', f'Stop-Process -Id {pid_str} -Force -ErrorAction SilentlyContinue'], capture_output=True)
    except Exception as e:
        print(f"Aviso ao liberar porta {port}: {e}")

def open_browser():
    try:
        webbrowser.open('http://127.0.0.1:5000/')
    except Exception as e:
        print(f"Não foi possível abrir o navegador automaticamente: {e}")

def run_server():
    host = os.environ.get("SERVER_HOST", "127.0.0.1")
    port = int(os.environ.get("SERVER_PORT", 5000))
    threads = int(os.environ.get("SERVER_THREADS", 8))
    app_env = os.environ.get("APP_ENV", "development").upper()
    pg_db = os.environ.get("PG_DB", "vendas_db")
    open_browser_flag = os.environ.get("OPEN_BROWSER_ON_START", "True").lower() in ("true", "1", "t")

    free_port(port)
    print("=" * 68)
    print(f"  Sistema Vendas - Gestao Comercial e PDV (Flask / {app_env})")
    print(f"  Servidor WSGI de Alta Performance: Waitress ({threads} threads)")
    print(f"  Base de Dados: PostgreSQL ({pg_db})")
    print(f"  Endereco Web: http://{host}:{port}/")
    print("=" * 68)
    print("  [STATUS]: Servidor INICIADO com sucesso e PRONTO para uso!")
    print("  (O terminal permanece nesta tela mantendo o servidor web ativo)")
    print("=" * 68)

    if open_browser_flag and host in ("127.0.0.1", "localhost"):
        Timer(0.8, lambda: webbrowser.open(f'http://127.0.0.1:{port}/')).start()

    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=threads)
    except Exception as e:
        debug_mode = os.environ.get("DEBUG", "False").lower() in ("true", "1", "t")
        print(f"Iniciando via Flask em fallback: {e}")
        app.run(host=host, port=port, debug=debug_mode, use_reloader=False, threaded=True)

if __name__ == '__main__':
    run_server()
