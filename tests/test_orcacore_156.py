"""TASK-156 — o programa passa a se chamar OrçaCore (o logo/ícone não muda). Arquivos e pastas usam OrcaCore (sem acento)."""
import glob
import hashlib
import os

import config
import pytest

from services import auto_updater as au
from services import janela_desktop as jd

ESTATICOS = config.STATIC_DIR


def ler(*partes):
    return open(os.path.join(*partes), encoding="utf-8").read()


def test_nome_nas_telas_e_o_icone_nao_mudou():
    paginas = glob.glob(os.path.join(ESTATICOS, "*.html"))
    assert len(paginas) >= 8
    for p in paginas:
        html = ler(p)
        assert "GridR</title>" not in html and "— GridR" not in html, p
        assert "/static/gridr.svg" in html and "/static/gridr.ico" in html, p                  # o ícone é o mesmo
    assert "<title>Login — OrçaCore</title>" in ler(ESTATICOS, "login.html")
    marca = ler(ESTATICOS, "marca.js")
    assert "const NOME = 'OrçaCore';" in marca and "Orça<span" in marca and "/static/gridr.svg" in marca
    assert "window.OrcaCore = window.GridR = window.GridReader" in marca                       # nomes antigos continuam existindo no JS
    assert hashlib.sha256(open(os.path.join(ESTATICOS, "gridr.svg"), "rb").read()).hexdigest()   # (o desenho do logo não é tocado: só o aria-label)
    assert "#d97706" in ler(ESTATICOS, "gridr.svg")


def test_nome_na_janela_no_build_e_no_servidor():
    app = ler(config.BASE_DIR, "app.py")
    assert 'title="OrçaCore",' in app and 'title=f"OrçaCore v{APP_VERSION}"' in app
    assert "OrçaCore" in ler(config.BASE_DIR, "services", "inicio_janela.py")
    bat = ler(config.BASE_DIR, "build.bat")
    assert '--name "OrcaCore"' in bat and "dist\\OrcaCore.exe" in bat and '--icon "static\\gridr.ico"' in bat
    assert "OrçaCore" in jd.LOCALIZACAO["global.quitConfirmation"]


def test_plugin_do_nanocad_continua_reconhecendo_o_programa():
    src = ler(config.BASE_DIR, "routers", "plugin_cad.py")
    assert '{"app": "GridReader", "nome": "OrçaCore"' in src


def test_perfil_da_janela_migra_do_nome_anterior(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    antiga = tmp_path / "GridR" / "janela"
    antiga.mkdir(parents=True)
    (antiga / "Login Data").write_text("senha")
    p = jd.pasta_perfil()
    assert os.path.normpath(str(p)) == os.path.normpath(str(tmp_path / "OrcaCore" / "janela")) and (tmp_path / "OrcaCore" / "janela" / "Login Data").read_text() == "senha"
    assert not antiga.exists()


def release(*nomes):
    base = au._url_permitida()
    return {"tag_name": "v9.9.9", "assets": [{"name": n, "browser_download_url": base + n} for n in nomes]}


def test_atualizador_aceita_o_nome_novo_e_os_antigos():
    novo = au._interpretar(release("OrcaCore.exe", "OrcaCore.exe.sha256", "GridR.exe", "GridR.exe.sha256"))
    assert novo["has_update"] and novo["download_url"].endswith("/OrcaCore.exe") and novo["sha256_url"].endswith("/OrcaCore.exe.sha256")
    so_gridr = au._interpretar(release("GridR.exe", "GridR.exe.sha256"))
    assert so_gridr["has_update"] and so_gridr["download_url"].endswith("/GridR.exe")
    so_legado = au._interpretar(release("Leitor_Projetos.exe", "Leitor_Projetos.exe.sha256"))
    assert so_legado["has_update"] and so_legado["download_url"].endswith("/Leitor_Projetos.exe")
    assert au._interpretar(release("OrcaCore.exe"))["has_update"] is False                     # sem o SHA-256 não oferece
    assert au.NOME_EXE == "OrcaCore.exe" and au.NOMES_EXE_ANTIGOS == ("GridR.exe", "Leitor_Projetos.exe")


def test_release_publica_o_nome_novo_e_as_copias_antigas():
    fluxo = ler(config.BASE_DIR, ".github", "workflows", "release-exe.yml")
    for nome in ("OrcaCore.exe", "GridR.exe", "Leitor_Projetos.exe"):
        assert f"dist/{nome} dist/{nome}.sha256" in fluxo, nome
    assert "Copy-Item dist\\OrcaCore.exe dist\\GridR.exe" in fluxo and "Copy-Item dist\\OrcaCore.exe dist\\Leitor_Projetos.exe" in fluxo
    assert '--title "OrçaCore ${GITHUB_REF_NAME}"' in fluxo
