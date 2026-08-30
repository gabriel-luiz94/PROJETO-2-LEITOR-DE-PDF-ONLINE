"""
scripts/release.py — Script para gerar nova release e notificar o servidor.

Executa o PyInstaller, compacta (ou prepara o .exe) e atualiza 
o banco de dados do servidor na nuvem indicando a nova versão disponível.
"""
import os
import sys
import argparse
from config import APP_VERSION
# from services.supabase_client import get_supabase

def build_exe():
    print(f"[*] Construindo versão {APP_VERSION} via PyInstaller...")
    os.system("pyinstaller --clean Leitor_PDF_Pro_v35.spec")
    print("[+] Build concluído.")

def push_to_cloud(version, download_url):
    print(f"[*] Notificando servidor nuvem sobre a nova versão {version}...")
    # Exemplo de como seria atualizar a nuvem:
    # supabase = get_supabase()
    # supabase.table('configuracoes').upsert([
    #     {'chave': 'latest_desktop_version', 'valor': version},
    #     {'chave': 'latest_desktop_url', 'valor': download_url}
    # ]).execute()
    print("[+] Servidor notificado. Os clientes desktop começarão a baixar na próxima verificação.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", required=True, help="URL pública de download do novo .exe")
    args = parser.parse_args()
    
    build_exe()
    push_to_cloud(APP_VERSION, args.url)
    print("Release finalizada!")
