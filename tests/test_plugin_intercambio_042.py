import re
from pathlib import Path


PLUGIN = Path(__file__).resolve().parents[1] / "plugin_nanocad" / "src" / "Plugin"


def test_acoes_do_menu_correspondem_aos_comandos_publicos_do_orcacore():
    inicio = (PLUGIN / "Comandos.cs").read_text(encoding="utf-8")
    menu = (PLUGIN / "TelaIntercambio.cs").read_text(encoding="utf-8")
    comandos = (PLUGIN / "ComandosDesenho.cs").read_text(encoding="utf-8")
    assert 'Escolher(Escolha.Intercambio)' in (PLUGIN / "Telas.cs").read_text(encoding="utf-8")
    assert inicio.index("if (t.Escolha == Escolha.Intercambio)") < inicio.index("RegrasFrescas(t.Regras, t.Projeto)")

    acoes_do_menu = re.findall(r'Adicionar\(menu, \d+, "[^"]+", "([A-Z_]+)"\);', menu)
    acoes_publicas = re.findall(r'\[CommandMethod\("([A-Z_]+)"', comandos)
    assert len(acoes_do_menu) == len(acoes_publicas) == 8
    assert set(acoes_do_menu) == set(acoes_publicas)
    assert all(acao.startswith("ORC_") for acao in acoes_publicas)
