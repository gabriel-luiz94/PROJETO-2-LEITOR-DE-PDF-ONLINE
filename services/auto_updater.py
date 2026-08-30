"""
services/auto_updater.py — Serviço de detecção e download de atualizações (Desktop).
"""
import os
import sys
import json
import tempfile
import threading
import urllib.request
import subprocess
from config import APP_VERSION, APP_MODE, SERVER_URL, IS_FROZEN, logger

# URL de onde baixar as informações de release (pode ser o servidor central ou um repositório Git)
# Se estiver usando o modo Híbrido, o servidor central (Fly.io/Railway) fornecerá essa rota.
UPDATE_CHECK_URL = f"{SERVER_URL}/api/update/check"


def check_for_updates() -> dict:
    """Verifica se há uma nova versão disponível."""
    if not SERVER_URL:
        return {"has_update": False}
        
    try:
        req = urllib.request.Request(UPDATE_CHECK_URL, headers={'User-Agent': 'LeitorPro-Updater'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            
        latest_version = data.get("latest_version")
        download_url = data.get("download_url")
        release_notes = data.get("release_notes", "")
        
        # Lógica semântica ultra simples (ex: 2.0.1 > 2.0.0)
        if latest_version and latest_version != APP_VERSION:
            return {
                "has_update": True,
                "latest_version": latest_version,
                "download_url": download_url,
                "release_notes": release_notes
            }
    except Exception as e:
        logger.warning(f"Erro ao verificar atualizações: {e}")
        
    return {"has_update": False}


def download_and_install_update(download_url: str):
    """
    Baixa o novo executável e inicia o script de substituição.
    (Só funciona corretamente se IS_FROZEN for True).
    """
    if not IS_FROZEN:
        logger.info("Auto-update ignorado: Executando em modo de desenvolvimento.")
        return False
        
    try:
        temp_dir = tempfile.gettempdir()
        installer_path = os.path.join(temp_dir, "Leitor_PDF_Pro_Update.exe")
        
        logger.info(f"Baixando atualização de {download_url}...")
        
        # Download do arquivo
        urllib.request.urlretrieve(download_url, installer_path)
        
        # Caminho do executável atual
        current_exe = sys.executable
        
        # Cria um script .bat temporário para fazer a substituição (pois o .exe atual está em uso)
        bat_script = f'''@echo off
echo Aguardando o fechamento do aplicativo...
timeout /t 3 /nobreak > NUL
move /Y "{installer_path}" "{current_exe}"
echo Atualizacao concluida. Reiniciando...
start "" "{current_exe}"
del "%~f0"
'''
        bat_path = os.path.join(temp_dir, "update_leitor.bat")
        with open(bat_path, "w") as f:
            f.write(bat_script)
            
        logger.info("Atualização baixada. O aplicativo será reiniciado.")
        
        # Executa o .bat desvinculado do processo atual
        subprocess.Popen([bat_path], creationflags=subprocess.CREATE_NEW_CONSOLE)
        
        # Encerra o app atual para o .bat poder substituir o arquivo
        os._exit(0)
        
    except Exception as e:
        logger.error(f"Erro ao baixar e instalar atualização: {e}")
        raise e
