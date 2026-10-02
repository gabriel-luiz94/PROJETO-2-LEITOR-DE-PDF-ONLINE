"""
scripts/autonomo_arquivos_teste.py — Arquivos e situações de teste para o modo autônomo (TASK-031, roteiro de testes no Windows).

  python scripts\\autonomo_arquivos_teste.py gerar  <pasta_amostras>                 cria os arquivos de teste (NÃO mexe na pasta de entrada)
  python scripts\\autonomo_arquivos_teste.py copiar-devagar <arquivo> <destino> [--segundos 20]   copia aos poucos (simula cópia lenta/rede)
  python scripts\\autonomo_arquivos_teste.py copiar-e-travar <origem> <destino> [--segundos 60]   copia e MANTÉM o destino ABERTO (no Windows,
                                                                                ninguém consegue mover/apagar até soltar)

Os DXF usam os textos do projeto de teste (cor vermelha, como os do programa). Os resultados esperados estão no roteiro
(docs/ROTEIRO_TESTE_WINDOWS_MODO_AUTONOMO.md).
"""
import argparse
import os
import sys
import time

import ezdxf


def gravar_dxf(caminho, textos):
    doc = ezdxf.new()
    msp = doc.modelspace()
    for i, t in enumerate(textos):
        msp.add_text(t, dxfattribs={"color": 1, "insert": (i * 10, 0), "height": 2})
    doc.saveas(caminho)


def gerar(pasta):
    os.makedirs(pasta, exist_ok=True)
    arquivos = {
        "01_normal.dxf": ["1-U4", "2-CFU", "DT11/300", "CAA 2 ABC 35 m", "TROCAR 2 RS M AC", "1-N1"],
        "02_com_exclusao.dxf": ["1-U3 1-U4", "1-U3 1-CFU", "2-CFU"],
        "03_com_exclusao_b.dxf": ["1-U3 1-U4", "1-U4"],
        "04_sem_nada_a_excluir.dxf": ["1-U4", "2-CFU"],
    }
    for nome, textos in arquivos.items():
        gravar_dxf(os.path.join(pasta, nome), textos)
    with open(os.path.join(pasta, "05_quebrado.dxf"), "w", encoding="utf-8") as f:
        f.write("isto não é um DXF de verdade")
    with open(os.path.join(pasta, "06_nota.txt"), "w", encoding="utf-8") as f:
        f.write("arquivo de tipo não suportado")
    gravar_dxf(os.path.join(pasta, "07_grande.dxf"), [f"{i % 9 + 1}-U4" for i in range(8000)])
    gravar_dxf(os.path.join(pasta, "08_nome com espaço e acentuação ção.dxf"), ["1-U4", "2-CFU"])
    print(f"Arquivos criados em {os.path.abspath(pasta)}:")
    for n in sorted(os.listdir(pasta)):
        print(f"  {n}  ({os.path.getsize(os.path.join(pasta, n)) // 1024} KB)")
    print("\nOBS.: o arquivo duplicado (T7) é só uma CÓPIA de um arquivo já processado — copie 01_normal.dxf de novo com outro nome.")


def copiar_devagar(origem, destino, segundos):
    tamanho = os.path.getsize(origem)
    bloco = max(1, tamanho // max(1, segundos * 2))
    with open(origem, "rb") as o, open(destino, "wb") as d:
        feitos = 0
        while True:
            parte = o.read(bloco)
            if not parte:
                break
            d.write(parte)
            d.flush()
            feitos += len(parte)
            print(f"\r  copiando… {feitos * 100 // tamanho}%", end="", flush=True)
            time.sleep(0.5)
    print("\nCópia concluída.")


def travar(arquivo, segundos):
    """Mantém o arquivo aberto. No Windows, um arquivo aberto pelo Python NÃO pode ser movido nem apagado por outro programa
    (o `open` do Python não compartilha a permissão de exclusão) — é exatamente a situação de um arquivo ainda em uso pelo CAD/cópia."""
    f = open(arquivo, "rb")
    try:
        print(f"Arquivo mantido ABERTO por {segundos}s: {arquivo}\nNo Windows, tente movê-lo/apagá-lo no Explorer agora (deve recusar). Ctrl+C solta antes.")
        time.sleep(segundos)
    except KeyboardInterrupt:
        pass
    finally:
        f.close()
        print("Arquivo liberado.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    g = sub.add_parser("gerar"); g.add_argument("pasta")
    c = sub.add_parser("copiar-devagar"); c.add_argument("origem"); c.add_argument("destino"); c.add_argument("--segundos", type=int, default=20)
    t = sub.add_parser("copiar-e-travar"); t.add_argument("origem"); t.add_argument("destino"); t.add_argument("--segundos", type=int, default=60)
    a = ap.parse_args()
    if a.cmd == "gerar":
        gerar(a.pasta)
    elif a.cmd == "copiar-devagar":
        copiar_devagar(a.origem, a.destino, a.segundos)
    else:
        import shutil
        shutil.copyfile(a.origem, a.destino)
        travar(a.destino, a.segundos)
