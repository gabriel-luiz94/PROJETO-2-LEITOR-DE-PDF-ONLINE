"""
scripts/release.py — Prepara uma nova versão do executável (TASK-127).

Uso:  python scripts/release.py 2.1.0

Atualiza APP_VERSION em config.py, faz o commit e cria a tag v2.1.0. Depois é só enviar:  git push origin main --tags
O GitHub Actions (.github/workflows/release-exe.yml) compila o .exe no Windows e publica a Release; os programas instalados
oferecem a atualização na próxima vez que abrirem (services/auto_updater.py).
"""
import os
import re
import subprocess
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _git(*args):
    return subprocess.run(["git", *args], cwd=RAIZ, check=True, capture_output=True, text=True).stdout.strip()


def main(argv) -> int:
    if len(argv) != 2 or not re.fullmatch(r"\d+\.\d+\.\d+", argv[1]):
        print("Uso: python scripts/release.py MAIOR.MENOR.PATCH   (ex.: 2.1.0)")
        return 2
    versao = argv[1]
    if _git("status", "--porcelain"):
        print("Há alterações não commitadas; faça o commit antes de criar a versão.")
        return 1
    if f"v{versao}" in _git("tag", "--list", f"v{versao}").split():
        print(f"A tag v{versao} já existe.")
        return 1
    caminho = os.path.join(RAIZ, "config.py")
    with open(caminho, encoding="utf-8") as f:
        texto = f.read()
    novo, n = re.subn(r'^APP_VERSION = ".*"$', f'APP_VERSION = "{versao}"', texto, count=1, flags=re.M)
    if n != 1:
        print("Não achei a linha APP_VERSION em config.py.")
        return 1
    with open(caminho, "w", encoding="utf-8", newline="") as f:
        f.write(novo)
    _git("add", "config.py")
    _git("commit", "-m", f"Versão {versao}")
    _git("tag", "-a", f"v{versao}", "-m", f"Versão {versao}")
    print(f"Versão {versao} pronta (commit + tag v{versao}). Para publicar:  git push origin main --tags")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
