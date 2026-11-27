import json
import struct
import subprocess
import sys
from pathlib import Path

import ezdxf
from ezdxf.entities import factory
from ezdxf.lldxf.extendedtags import ExtendedTags
from ezdxf.proxygraphic import ProxyGraphicTypes

from services.cadsuite_proxy import main


def inserir_proxy(modelo, tipo, raio=None):
    grafico = ""
    if raio is not None:
        dados = bytes(8) + struct.pack(
            "<II7d", 64, ProxyGraphicTypes.CIRCLE, 1, 2, 3, raio, 0, 0, 1
        )
        grafico = f"160\n{len(dados)}\n310\n{dados.hex()}\n"
    comum = f"0\n{tipo}\n100\nAcDbEntity\n8\n0\n"
    if tipo == "ACAD_PROXY_ENTITY":
        tags = comum + "100\nAcDbProxyEntity\n90\n498\n91\n500\n" + grafico
    else:
        tags = comum + grafico + "100\nSPDSNotePosition\n"
    modelo.add_entity(factory.load(ExtendedTags.from_text(tags)))


def test_worker_converte_os_dois_tipos_e_preserva_o_dxf_de_entrada(tmp_path):
    doc = ezdxf.new("R2010")
    modelo = doc.modelspace()
    modelo.add_line((0, 0), (1, 2))
    inserir_proxy(modelo, "ACAD_PROXY_ENTITY", raio=5)
    inserir_proxy(modelo, "SPDSNOTEPOSITION", raio=7)
    inserir_proxy(modelo, "ACAD_PROXY_ENTITY")
    entrada = tmp_path / "entrada.dxf"
    doc.saveas(entrada)
    original = entrada.read_bytes()

    app = Path(__file__).resolve().parents[1] / "app.py"
    subprocess.run(
        [sys.executable, str(app), "--trabalhador-cadsuite-proxy", str(tmp_path)],
        check=True,
        timeout=15,
    )

    relatorio = json.loads((tmp_path / "resultado.json").read_text(encoding="utf-8"))
    assert relatorio["convertidos"] == 2
    assert len(relatorio["avisos"]) == 1
    assert "sem representação gráfica" in relatorio["avisos"][0]
    convertido = ezdxf.readfile(tmp_path / "saida.dxf").modelspace()
    assert len(convertido.query("LINE")) == 1
    assert len(convertido.query("ACAD_PROXY_ENTITY")) == 1
    circulos = convertido.query("CIRCLE")
    assert [circulo.dxf.radius for circulo in circulos] == [5, 7]
    assert all(tuple(circulo.dxf.center) == (1, 2, 3) for circulo in circulos)
    assert entrada.read_bytes() == original


def test_worker_sem_geometria_recuperavel_nao_publica_dxf(tmp_path):
    doc = ezdxf.new("R2010")
    inserir_proxy(doc.modelspace(), "ACAD_PROXY_ENTITY")
    doc.saveas(tmp_path / "entrada.dxf")

    assert main([str(tmp_path)]) == 0
    relatorio = json.loads((tmp_path / "resultado.json").read_text(encoding="utf-8"))
    assert relatorio["convertidos"] == 0
    assert len(relatorio["avisos"]) == 1
    assert not (tmp_path / "saida.dxf").exists()
