"""
services/inicio_janela.py — Abertura da janela do programa desktop (TASK-128).

O servidor (uvicorn) só aceita conexões depois de terminar a inicialização (teste de rede, banco, nuvem…). Antes, a janela abria logo e mostrava
"recusou a conectar" até o usuário atualizar. Agora a janela abre com uma tela de "Iniciando…" e só carrega o login quando a porta responde.
"""
import socket
import time

TELA_INICIAL = """<!doctype html><html lang="pt-br"><head><meta charset="utf-8"><title>ReadGrid</title>
<style>html,body{height:100%;margin:0;background:#0d1117;color:#e6edf3;font-family:Segoe UI,system-ui,sans-serif}
.c{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px}
.s{width:38px;height:38px;border:4px solid #30363d;border-top-color:#d97706;border-radius:50%;animation:g 0.9s linear infinite}
@keyframes g{to{transform:rotate(360deg)}}</style></head>
<body><div class="c"><div class="s"></div><div>Iniciando o ReadGrid…<div style="color:#6e7681;font-size:12px;margin-top:6px">Desenvolvido por Gabriel Sales</div></div></div></body></html>"""

TELA_ERRO = """<!doctype html><html lang="pt-br"><head><meta charset="utf-8"><title>ReadGrid</title>
<style>html,body{height:100%;margin:0;background:#0d1117;color:#e6edf3;font-family:Segoe UI,system-ui,sans-serif}
.c{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;text-align:center;padding:20px}</style></head>
<body><div class="c"><div style="font-size:20px">Não foi possível iniciar o servidor do programa.</div>
<div style="color:#8b949e">Feche e abra o programa de novo. Se continuar, veja se outro programa está usando a porta.</div></div></body></html>"""


def esperar_servidor(host: str, porta: int, limite: float = 90.0, intervalo: float = 0.2) -> bool:
    """True assim que `host:porta` aceitar conexão; False se passar `limite` segundos."""
    fim = time.monotonic() + limite
    while time.monotonic() < fim:
        try:
            with socket.create_connection((host, porta), timeout=1.0):
                return True
        except OSError:
            time.sleep(intervalo)
    return False
