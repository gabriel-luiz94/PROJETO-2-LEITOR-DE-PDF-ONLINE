"""
services/plugin_instalador.py — o programa principal traz o plugin do nanoCAD (TASK-131).

O executável (build.bat) leva as DLLs de plugin_nanocad/appload/nanoCAD26 e nanoCAD5 (cada uma com versao.json). Ao abrir o .exe,
elas são instaladas numa pasta fixa do usuário:

    %APPDATA%\\OrcamentoNanoCAD\\plugin\\nanoCAD26\\OrcamentoNanoCAD.dll
    %APPDATA%\\OrcamentoNanoCAD\\plugin\\nanoCAD5\\OrcamentoNanoCAD.dll

Só instala quando a versão do programa é MAIS NOVA que a instalada (o próprio plugin também se atualiza sozinho e grava o versao.json
ao lado da DLL; o programa nunca volta uma versão). Com o nanoCAD aberto a DLL está em uso: o Windows deixa RENOMEAR um arquivo em
uso, então a atual vira .anterior e a nova entra no lugar (vale ao reabrir o nanoCAD) — o mesmo jeito do plugin (Nucleo/Atualizacao.cs).
No nanoCAD, uma vez: APPLOAD → "Inicialização" (Startup Suite) → adicionar o arquivo da pasta acima.
"""
import json
import os
import shutil
import sys
import threading

from config import logger

ALVOS = {"nanocad26": ("nanoCAD26", "nanoCAD 26 (.NET 6+)"), "nanocad5": ("nanoCAD5", "nanoCAD 5 (.NET Framework 4)")}
ARQUIVO_DLL = "OrcamentoNanoCAD.dll"


def pasta_destino() -> str:
    """Pasta do plugin instalado (variável PLUGIN_NANOCAD_DESTINO troca, para testes)."""
    p = os.environ.get("PLUGIN_NANOCAD_DESTINO")
    if p:
        return p
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "OrcamentoNanoCAD", "plugin")


def _pasta_origem() -> str:
    from routers.plugin_cad import PASTA_APPLOAD
    return PASTA_APPLOAD


def _versao(v) -> tuple:
    partes = []
    for x in str(v or "").strip().lstrip("vV").split("-")[0].split("."):
        try:
            partes.append(int(x))
        except ValueError:
            partes.append(0)
    while partes and partes[-1] == 0:
        partes.pop()
    return tuple(partes)


def _ler_versao(pasta: str):
    try:
        with open(os.path.join(pasta, "versao.json"), encoding="utf-8") as f:
            return str(json.load(f).get("versao") or "") or None
    except (OSError, ValueError):
        return None


def situacao() -> list:
    """Por alvo: versão que o programa traz, versão instalada, caminho da DLL para o APPLOAD."""
    r = []
    for alvo, (pasta, nome) in ALVOS.items():
        origem = os.path.join(_pasta_origem(), pasta)
        destino = os.path.join(pasta_destino(), pasta)
        dll = os.path.join(destino, ARQUIVO_DLL)
        v_prog = _ler_versao(origem) if os.path.isfile(os.path.join(origem, ARQUIVO_DLL)) else None
        v_inst = _ler_versao(destino) if os.path.isfile(dll) else None
        r.append({"alvo": alvo, "nome": nome, "versao_programa": v_prog, "versao_instalada": v_inst,
                  "desatualizado": bool(v_prog) and os.path.isfile(dll) and _versao(v_inst) < _versao(v_prog),
                  "caminho": dll, "pasta": destino, "instalado": os.path.isfile(dll)})
    return r


def _trocar(origem_dll: str, dll: str):
    """Grava .nova e troca (a em uso vira .anterior). Nunca deixa a pasta sem DLL."""
    nova, anterior = dll + ".nova", dll + ".anterior"
    shutil.copyfile(origem_dll, nova)
    if os.path.exists(dll):
        if os.path.exists(anterior):
            os.remove(anterior)
        os.rename(dll, anterior)
    try:
        os.rename(nova, dll)
    except OSError:
        if not os.path.exists(dll) and os.path.exists(anterior):
            os.rename(anterior, dll)
        raise


def instalar(forcar: bool = False) -> list:
    """Instala/atualiza cada alvo que o programa traz. Devolve o resultado por alvo (instalado | em dia | erro | ausente)."""
    res = []
    for alvo, (pasta, nome) in ALVOS.items():
        origem = os.path.join(_pasta_origem(), pasta)
        origem_dll = os.path.join(origem, ARQUIVO_DLL)
        destino = os.path.join(pasta_destino(), pasta)
        dll = os.path.join(destino, ARQUIVO_DLL)
        item = {"alvo": alvo, "nome": nome, "caminho": dll}
        v_prog = _ler_versao(origem)
        if not os.path.isfile(origem_dll) or not v_prog:
            res.append(dict(item, resultado="ausente", detalhe="o programa não traz o plugin para este nanoCAD"))
            continue
        v_inst = _ler_versao(destino) if os.path.isfile(dll) else None
        if v_inst and not forcar and _versao(v_inst) >= _versao(v_prog):
            res.append(dict(item, resultado="em dia", versao=v_inst))
            continue
        try:
            os.makedirs(destino, exist_ok=True)
            _trocar(origem_dll, dll)
            with open(os.path.join(destino, "versao.json"), "w", encoding="utf-8") as f:
                json.dump({"versao": v_prog, "origem": "programa principal"}, f, ensure_ascii=False, indent=2)
            res.append(dict(item, resultado="instalado", versao=v_prog, anterior=v_inst))
            logger.info(f"Plugin {nome} {v_prog} instalado em {dll}")
        except OSError as e:
            res.append(dict(item, resultado="erro", detalhe=str(e)))
            logger.warning(f"Plugin {nome}: instalação falhou: {e}")
    return res


def instalar_em_segundo_plano():
    """Chamado ao abrir o .exe (só o desktop empacotado): instala sem atrasar a abertura."""
    if not getattr(sys, "frozen", False):
        return
    threading.Thread(target=lambda: _seguro(instalar), daemon=True, name="instalar-plugin-nanocad").start()


def _seguro(f):
    try:
        f()
    except Exception as e:  # noqa: BLE001 — instalar o plugin nunca impede o programa de abrir
        logger.warning(f"Instalação do plugin do nanoCAD falhou: {e}")
