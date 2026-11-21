"""TASK-213 — regressão do Resumo no Chromium, se Playwright e Chromium estiverem disponíveis."""
import json
import shutil
import socket
import threading
import time

import pytest
import uvicorn

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import programa_local


def test_importacao_comparacao_e_janelas_auxiliares(tmp_path, monkeypatch):
    playwright = pytest.importorskip("playwright.sync_api")
    chromium = shutil.which("chromium") or shutil.which("chromium-browser")
    if not chromium:
        pytest.skip("Chromium indisponível.")

    import app as appmod

    monkeypatch.setattr(config, "APP_MODE", "desktop")
    database.init_db()
    ativo = {"entidade": "CABO", "operacao": "I", "ativo": "CAA2 ABC 35 m", "_geo": [[0, 0], [100, 0]]}
    poste = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": "U4", "_poste": [0, 0]}
    obra = {"formato": "leitor-obra", "versao": 1, "nome": "Atual", "projeto": "027",
            "dados": {"cabos": [ativo], "outros": [poste]}}
    caminho = tmp_path / "Atual.obra"
    caminho.write_text(json.dumps(obra), encoding="utf-8")
    pendente = programa_local.guardar_projeto("", str(caminho), "027")
    comparada = tmp_path / "Comparada.obra.json"
    comparada.write_text(json.dumps({**obra, "nome": "Comparada", "dados": {
        "cabos": [{**ativo, "ativo": "CAA2 ABC 42 m"}], "outros": [poste]}}), encoding="utf-8")

    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        porta = sock.getsockname()[1]
        servidor = uvicorn.Server(uvicorn.Config(appmod.app, lifespan="off", log_level="error"))
        thread = threading.Thread(target=servidor.run, kwargs={"sockets": [sock]}, daemon=True)
        thread.start()
        try:
            for _ in range(200):
                if servidor.started:
                    break
                time.sleep(0.025)
            assert servidor.started
            with playwright.sync_playwright() as pw:
                navegador = pw.chromium.launch(executable_path=chromium, args=["--no-sandbox"])
                pagina = navegador.new_page(viewport={"width": 960, "height": 510})
                pagina.set_default_timeout(8000)
                pagina.route("https://**/*", lambda rota: rota.abort())
                token = create_jwt_token("teste-local", "teste@example.test", "operador")
                pagina.add_init_script(f"""
                    localStorage.setItem('auth_token', {json.dumps(token)});
                    if (location.pathname === '/resumo') {{
                        localStorage.setItem('projeto_selecionado_codigo', '027');
                        localStorage.setItem('resumo_layout_v1', '{{"outros":{{"visivel":true,"recolhida":true}}}}');
                    }}
                    window.pywebview = {{api: {{projeto_criar() {{}}}}}};
                """)
                erros = []
                pagina.on("pageerror", lambda erro: erros.append(str(erro)))
                pagina.goto(f"http://127.0.0.1:{porta}/resumo?plugin_pendente={pendente}", wait_until="domcontentloaded")
                pagina.wait_for_function("VinculacaoHost.outros().length === 1 && !document.querySelector('#section-outros').classList.contains('recolhida')")
                altura = "() => document.querySelector('#section-outros .table-wrap').clientHeight"
                assert pagina.evaluate(altura) > 60
                pagina.set_viewport_size({"width": 1490, "height": 920})
                assert pagina.evaluate(altura) > 400

                pagina.locator("#input-importar-json").set_input_files(str(comparada))
                pagina.locator("[data-acao=comparar]").click()
                dialogo = pagina.get_by_role("dialog", name="Comparar obras")
                dialogo.get_by_role("button", name="Cabos").click()
                pagina.locator("#section-cabos .btn-recolher").evaluate("(b) => b.click()")
                dialogo.get_by_role("button", name="Ir à linha 1").click()
                assert pagina.evaluate("document.activeElement.closest('#body-cabos tr').dataset.index") == "0"
                dialogo.get_by_role("button", name="Voltar à comparação").click()
                assert dialogo.get_by_text("+7 m", exact=True).is_visible()
                dialogo.get_by_role("button", name="Fechar", exact=True).click()

                pagina.locator("#btn-toggle-config").click()
                pagina.locator("#btn-tabela-orcamento").click()
                base = pagina.get_by_role("region", name="Base de Orçamento")
                pagina.locator("[data-modal-id=vincular]").click()
                pagina.get_by_role("button", name="Abrir em janela").click()
                vinculo = pagina.get_by_role("region", name="Vinculação")
                vinculo.get_by_role("button", name="Minimizar", exact=True).click()
                pagina.locator("[data-modal-id=vincular]").click()
                assert vinculo.locator("#vz-corpo").is_visible()

                pagina.route("**/api/validacao/planilhas", lambda rota: rota.fulfill(json={"achados": []}))
                pagina.route("**/api/orcamento/calcular", lambda rota: rota.fulfill(json={
                    "resultado": [{"operacao": "I", "mdo": "", "codigo": "4321", "desc_codigo": "Item", "total": 1}]}))
                resultado = pagina.get_by_role("region", name="Resultado do orçamento")
                continuar = pagina.get_by_role("button", name="Continuar mesmo assim", exact=True)
                for _ in range(2):
                    with pagina.expect_response(lambda r: r.url.endswith("/api/orcamento/calcular")):
                        pagina.locator("#btn-montar-orcamento").click()
                        # Este projeto exige passar pelo aviso local de vinculação antes de calcular.
                        continuar.click()
                    assert resultado.count() == 1
                    assert base.is_visible() and vinculo.is_visible()
                resultado.locator("iframe").evaluate("(f) => f.contentWindow.voltarOuFechar()")
                resultado.wait_for(state="detached")
                assert base.is_visible() and vinculo.is_visible()
                assert not erros
                conferir_largura_da_obra_restaurada(navegador, porta, token, ativo, poste)
                navegador.close()
        finally:
            servidor.should_exit = True
            thread.join(timeout=5)


def conferir_largura_da_obra_restaurada(navegador, porta, token, ativo, poste):
    # TASK-214: restauração em uma nova página. As linhas abaixo da área visível
    # continuam pendentes, como no caso de 801 linhas da Paraíba.
    dados_grandes = {"cabos": [ativo] * 357, "outros": [poste] * 801}
    restaurada = navegador.new_page(viewport={"width": 960, "height": 650})
    restaurada.set_default_timeout(8000)
    restaurada.route("https://**/*", lambda rota: rota.abort())
    restaurada.add_init_script(f"""
        localStorage.setItem('auth_token', {json.dumps(token)});
        localStorage.setItem('projeto_selecionado_codigo', '027');
        localStorage.setItem('resumo_ocultar_entidade', '1');
        localStorage.setItem('resumo_ocultar_cabo_extra', '1');
        const sessao = document.cookie.split('; ').find(c => c.startsWith('gridr_sessao='));
        if (sessao) localStorage.setItem('gridr_sessao', sessao.split('=')[1]);
        localStorage.setItem('gridr_projeto', JSON.stringify({{caminho: 'Grande.obra', projeto: '027'}}));
        window.pywebview = {{api: {{
            projeto_criar() {{}},
            projeto_assinatura: async () => ({{ok: true, existe: true, assinatura: '1'}}),
            projeto_abrir: async () => ({{ok: true, caminho: 'Grande.obra', assinatura: '1',
                projeto: {{nome: 'Grande', projeto: '027', dados: {json.dumps(dados_grandes)}}}}})
        }}}};
    """)
    restaurada.goto(f"http://127.0.0.1:{porta}/resumo", wait_until="domcontentloaded")
    restaurada.wait_for_function("VinculacaoHost.outros().length === 801")
    larguras = """id => {
        const tabela = document.getElementById(id);
        const borda = tabela.getBoundingClientRect();
        return {
            pendentes: tabela.querySelectorAll('.lazy-row').length,
            larguraDesocupada: tabela.parentElement.clientWidth - borda.width,
            sobra: borda.right - tabela.tHead.rows[0].lastElementChild.getBoundingClientRect().right
        };
    }"""

    def conferir_larguras():
        for tabela in ("table-outros", "table-cabos"):
            medidas = restaurada.evaluate(larguras, tabela)
            assert medidas["pendentes"] > 0
            assert abs(medidas["sobra"]) < 3, (tabela, medidas)
            assert abs(medidas["larguraDesocupada"]) < 3, (tabela, medidas)

    conferir_larguras()
    restaurada.set_viewport_size({"width": 1872, "height": 860})
    conferir_larguras()
    restaurada.locator("#btn-toggle-entidade").click()
    restaurada.get_by_text("Mostrar a coluna Entidade", exact=True).click()
    restaurada.locator("#btn-toggle-cabo-extra").click()
    conferir_larguras()
    restaurada.locator("#btn-toggle-entidade").click()
    restaurada.get_by_text("Agrupar por Entidade", exact=True).click()
    separador = restaurada.locator("#body-outros .grupo-sep td").first
    assert separador.get_attribute("colspan") == "5"
    restaurada.locator("#btn-toggle-entidade").click()
    restaurada.get_by_text("Ocultar a coluna Entidade", exact=True).click()
    assert separador.get_attribute("colspan") == "4"
    restaurada.locator("#btn-toggle-cabo-extra").click()
    conferir_larguras()
