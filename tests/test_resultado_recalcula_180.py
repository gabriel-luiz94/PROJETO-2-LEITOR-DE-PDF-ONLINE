"""TASK-180: Montar Orçamento sempre recalcula; o resultado gravado no projeto só volta ao abrir o projeto na tela do Resultado."""
import os
import re

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _ler(*p):
    with open(os.path.join(RAIZ, *p), encoding="utf-8") as f:
        return f.read()


def test_montar_nao_restaura_resultado_gravado():
    html = _ler("static", "resultado_orcamento.html")
    assert "GridRProjeto.carregarResultado" not in html          # a restauração por assinatura saiu
    assert "/api/orcamento/calcular" in html
    assert "as edições manuais do resultado anterior foram substituídas" in html


def test_abrir_projeto_mostra_resultado_gravado():
    html = _ler("static", "resultado_orcamento.html")
    assert "function mostrarResultadoDoProjeto" in html
    assert "GridRProjeto.aoAbrirProjeto(mostrarResultadoDoProjeto)" in html
    assert "Este projeto não tem resultado gravado" in html
    assert "GridRProjeto.pararAcompanharResultado()" in html


def test_projeto_js_expoe_ligacoes():
    js = _ler("static", "projeto.js")
    for nome in ("resultadoEditado", "pararAcompanharResultado", "aoAbrirProjeto", "__notificarAbertura",
                 "__marcarEdicaoResultado", "__limparEdicaoResultado", "__registrarAoAbrir"):
        assert re.search(r"\b" + nome + r"\b", js), nome
    # Mesclar/Sobrescrever não notificam: só o modo Abrir
    assert "if (modo === 'abrir') { resultadoEditado = false; notificarAbertura(pr); }" in js


def test_formato_do_obra_nao_muda():
    js = _ler("static", "projeto.js")
    assert "enviar({ resultado: entregar() })" in js and "assinatura, linhas: coletar()" in js
