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
                pagina.context.add_init_script(f"""
                    localStorage.setItem('auth_token', {json.dumps(token)});
                    if (location.pathname === '/resumo') {{
                        localStorage.setItem('projeto_selecionado_codigo', '027');
                        localStorage.setItem('resumo_layout_v1', '{{"outros":{{"visivel":true,"recolhida":true}}}}');
                    }}
                    const auxiliares = new Map();
                    window.pywebview = {{api: {{
                        projeto_criar() {{}},
                        projeto_gravar: async (caminho, dados) => {{
                            window.ultimaGravacao = JSON.parse(dados);
                            return {{ok: true, assinatura: 'gravada'}};
                        }},
                        janela_auxiliar_abrir: async (tipo, id) => {{
                            const w = window.open('/static/janela_auxiliar.html#' + tipo + '=' + id, '_blank');
                            auxiliares.set(id, w);
                            return {{ok: !!w}};
                        }},
                        janela_auxiliar_fechar: async id => auxiliares.get(id)?.close(),
                        salvar_arquivo: async () => ({{ok: true}})
                    }}}};
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
                base.get_by_role("button", name="Minimizar", exact=True).click()
                assert not base.is_visible()
                assert pagina.get_by_role("button", name="Maximizar Base de Orçamento").is_visible()
                pagina.locator("[data-modal-id=vincular]").click()
                pagina.get_by_role("button", name="Abrir em janela").click()
                vinculo = pagina.get_by_role("region", name="Vinculação")
                vinculo.get_by_role("button", name="Minimizar", exact=True).click()
                pagina.locator("[data-modal-id=vincular]").click()
                assert vinculo.locator("#vz-corpo").is_visible()

                pagina.route("**/api/validacao/planilhas", lambda rota: rota.fulfill(json={"achados": []}))
                contador = iter((4321, 4322))
                pagina.context.route("**/api/orcamento/calcular", lambda rota: rota.fulfill(json={
                    "resultado": [{"operacao": "I", "mdo": "", "codigo": str(next(contador)), "desc_codigo": "Item", "total": 1}]}))
                resultado = pagina.locator('section[aria-label="Resultado do orçamento"]')
                continuar = pagina.get_by_role("button", name="Continuar mesmo assim", exact=True)

                def montar():
                    pagina.locator("#btn-montar-orcamento").click()
                    continuar.click()

                montar()
                largura = resultado.bounding_box()["width"]
                assert largura == pagina.viewport_size["width"]
                linha = resultado.frame_locator("iframe").locator("#resultado-orcamento-tbody tr")
                linha.get_by_text("4321", exact=True).wait_for()
                resultado.get_by_role("button", name="Minimizar", exact=True).click()
                assert not resultado.is_visible()
                pagina.get_by_role("button", name="Maximizar Resultado", exact=True).click()
                assert resultado.is_visible()
                with pagina.expect_popup() as popup:
                    resultado.get_by_role("button", name="Abrir em nova janela").click()
                nativa = popup.value
                nativa.set_default_timeout(8000)
                nativa.on("pageerror", lambda erro: erros.append(str(erro)))
                quadro_nativo = nativa.frame_locator("iframe")
                linha_nativa = quadro_nativo.locator("#resultado-orcamento-tbody tr")
                linha_nativa.get_by_text("4321", exact=True).wait_for()
                resultado.wait_for(state="hidden")
                assert vinculo.is_visible()
                montar()
                linha_nativa.get_by_text("4322", exact=True).wait_for()
                linha_nativa.locator("td").nth(3).fill("Ajuste <literal>")
                pagina.evaluate("GridRProjeto.salvar()")
                assert pagina.evaluate("ultimaGravacao.resultado.linhas[0].desc_codigo") == "Ajuste <literal>"
                with nativa.expect_event("close"):
                    nativa.get_by_role("button", name="Retornar à visualização principal").click()
                linha.get_by_text("Ajuste <literal>", exact=True).wait_for()
                assert resultado.is_visible()
                resultado.locator("iframe").evaluate("(f) => f.contentWindow.voltarOuFechar()")
                resultado.wait_for(state="detached")
                assert vinculo.is_visible()
                assert pagina.get_by_role("button", name="Maximizar Base de Orçamento").is_visible()
                # A Base usa o mesmo protocolo e mantém a edição ainda não gravada.
                pagina.get_by_role("button", name="Maximizar Base de Orçamento").click()
                pagina.context.route("**/api/orcamento/dados?genericas=true", lambda rota: rota.fulfill(json={
                    "dados": [{"ativo": "CABO", "codigo": "6789", "desc_ativo": "Base"}]}))
                base.frame_locator("iframe").get_by_role("button", name="Comuns a todos os projetos").click()
                campo_base = base.frame_locator("iframe").locator("#orcamento-tbody .cell-input").first
                campo_base.fill("BASE EDITADA")
                with pagina.expect_popup() as popup_base:
                    base.get_by_role("button", name="Abrir em nova janela").click()
                nativa_base = popup_base.value
                nativa_base.set_default_timeout(8000)
                nativa_base.on("pageerror", lambda erro: erros.append(str(erro)))
                campo_nativo = nativa_base.frame_locator("iframe").locator("#orcamento-tbody .cell-input").first
                playwright.expect(campo_nativo).to_have_value("BASE EDITADA")
                base.wait_for(state="hidden")
                campo_nativo.fill("BASE NO OUTRO MONITOR")
                with nativa_base.expect_event("close"):
                    nativa_base.get_by_role("button", name="Retornar à visualização principal").click()
                playwright.expect(campo_base).to_have_value("BASE NO OUTRO MONITOR")
                base.get_by_role("button", name="Fechar", exact=True).click()
                base.wait_for(state="detached")
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
