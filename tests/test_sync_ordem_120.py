"""TASK-120: a sincronização da base mestra mantém a ORDEM da planilha (pede as linhas ao Supabase ordenadas por id, paginado)."""
import database
from services import sync_service as ss


class _Tabela:
    """Supabase de mentira que, SEM order(), devolve as linhas embaralhadas (ordem física do Postgres) e com order('id') devolve por id."""
    def __init__(self, linhas):
        self.linhas, self.ordenar, self.faixa = linhas, False, None

    def select(self, *_): return self
    def like(self, *_): return self
    def order(self, coluna): self.ordenar = coluna; return self
    def range(self, ini, fim): self.faixa = (ini, fim); return self

    def execute(self):
        class R: pass
        linhas = sorted(self.linhas, key=lambda r: r["id"]) if self.ordenar == "id" else sorted(self.linhas, key=lambda r: (r["id"] * 7919) % 1000)
        if self.faixa:
            linhas = linhas[self.faixa[0]:self.faixa[1] + 1]
        r = R(); r.data = linhas
        return r


def test_sync_da_base_mestra_preserva_a_ordem_por_id(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    linhas = [{"id": i, "ativo": f"A{i}" if i % 3 == 1 else "", "codigo": f"c{i:04d}", "componente": "X"} for i in range(1, 2501)]   # 2500 linhas = 3 páginas de 1000
    class Sb:
        def table(self, _): return _Tabela(linhas)
    monkeypatch.setattr(ss, "get_supabase", lambda: Sb())
    monkeypatch.setattr(ss, "process_queue", lambda: None, raising=False)
    assert ss.sync_tabela_master() is True
    local = [r["codigo"] for r in ss.get_merged_orcamento(None)]
    assert local == [f"c{i:04d}" for i in range(1, 2501)]               # mesma ordem e sem repetir/pular linha entre páginas


def test_sync_pede_a_ordem_ao_supabase():
    import inspect
    fonte = inspect.getsource(ss.sync_tabela_master)
    assert '.order("id")' in fonte
