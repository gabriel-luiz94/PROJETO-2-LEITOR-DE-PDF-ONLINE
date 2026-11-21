"""TASK-216 — as bibliotecas de formatos de arquivo não entram na carga inicial."""
import subprocess
import sys
from pathlib import Path


def test_app_importa_sem_carregar_dependencias_de_pdf_dxf_e_excel():
    resultado = subprocess.run(
        [sys.executable, "-c", """
import sys
import tests.conftest
from services.supabase_client import get_supabase
assert get_supabase() is None
import app

formatos = {'pymupdf', 'fitz', 'ezdxf', 'openpyxl'}
carregados = formatos.intersection(sys.modules)
assert not carregados, f'Bibliotecas carregadas na abertura: {carregados}'
"""],
        cwd=Path(__file__).resolve().parents[1],
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert resultado.returncode == 0, resultado.stderr
