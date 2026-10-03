"""
services/autonomo/pipeline.py — Pipeline autônomo SEM TELA e SEM IA (TASK-031, fase B).

ler → processar (leitor) → montar tabelas → ramais → validar → ajustar → [confirmação de exclusões] → Totalizadora → orçamento → obra + pasta.

Decisões do usuário (TASK-031): ajustes que NÃO excluem são aplicados sozinhos; os que EXCLUEM linhas ou itens (`excluir_linhas`,
`remover_ativo`) deixam o arquivo "aguardando_confirmacao" — nada é excluído nem orçado até alguém responder (`confirmar`, "Sim" ou
"Sim para todos" = todas as pendências do arquivo; `rejeitar`). Erro de validação sem correção: salva "com_pendencias" com relatório.
O backend guarda as tabelas ORIGINAIS e o diff e reaplica do zero a cada decisão/reversão (ids posicionais não se desencontram).
Nenhuma chamada à IA aqui: nada importa routers/ai_chat, validacao_ia ou correcao_ia (há teste).
"""
import hashlib
import json
import os
import time
from dataclasses import dataclass, field
from datetime import datetime

from services.ajustes_planilhas import ajustar, validar_receitas
from services.autonomo import execucoes, saida
from services.autonomo.aplicar import aplicar_operacoes
from services.autonomo.leitor_js import obter_leitor
from services.autonomo.montagem import montar_tabelas
from services.autonomo.ramais import linhas_ramais
from services.autonomo.totalizadora import REGRA_PADRAO, montar_totalizadora, normalizar_para_json, payload_calculo
from services.autonomo.vinculacao import avaliar_vinculacao, itens_vinculo_para_totalizadora, tentar_vincular_automaticamente
from services.orcamento_calc import processar_calculo
from services.regras_dominio import avaliar
from services.validacao_planilhas import resumir, validar_planilhas

ACOES_DESTRUTIVAS = {"excluir_linhas", "remover_ativo"}
LIMITE_ACOES_AUTONOMO = 300   # a cadeia de todos os ajustes habilitados (nas rotas/tela o limite é 50); projetos reais passam de 50
LIMITE_BYTES = 200 * 1024 * 1024
EXTENSOES = (".dxf", ".pdf")


class ErroPipeline(Exception):
    """Falha esperada e explicável (vira status 'erro' com a mensagem, sem derrubar a fila)."""


@dataclass
class Contexto:
    """Tudo que o pipeline precisa do banco — injetável nos testes."""
    projeto_codigo: str
    projeto_nome: str
    user_id: str
    regras_proc: list = field(default_factory=list)
    regras_cls: list = field(default_factory=list)
    regras_conversao: list = field(default_factory=list)
    base_orcamento: list = field(default_factory=list)
    regras_dominio: list = field(default_factory=list)
    grupos: dict = field(default_factory=dict)
    receitas: list = field(default_factory=list)      # ajustes efetivos habilitados do projeto (ativa e não ocultos)
    regras_vinculacao: list = field(default_factory=list)  # TASK-036
    leitor: object = None


def carregar_contexto(projeto_codigo: str, user_id: str) -> Contexto:
    """Monta o contexto a partir do banco (mesmas fontes que a tela usa)."""
    import database
    from routers.regras import get_regras_conversao
    from routers.regras_leitor import _get_regras
    from routers.regras_vinculacao import _get_regras as _get_regras_vinculacao
    from routers.validacao_ajustes import ajustes_efetivos
    from routers.validacao_regras import regras_efetivas
    from services.sync_service import get_merged_orcamento

    conn = database.get_row_connection()
    try:
        r = conn.execute("SELECT nome FROM projetos WHERE codigo = ?", (projeto_codigo,)).fetchone()
    finally:
        conn.close()
    if not r:
        raise ErroPipeline(f"Projeto '{projeto_codigo}' não cadastrado.")
    regras_dominio, grupos, _ = regras_efetivas(projeto_codigo)
    receitas = [x for x in ajustes_efetivos(projeto_codigo) if x.get("ativa") and not x.get("oculta")]
    conversao = (get_regras_conversao(projeto_codigo) or {}).get("regras") or [dict(REGRA_PADRAO)]   # como a tela sem regras salvas
    return Contexto(projeto_codigo=projeto_codigo, projeto_nome=r["nome"], user_id=user_id,
                    regras_proc=_get_regras("processamento", projeto_codigo), regras_cls=_get_regras("classificacao", projeto_codigo),
                    regras_conversao=conversao, base_orcamento=get_merged_orcamento(user_id),
                    regras_dominio=regras_dominio, grupos=grupos, receitas=receitas,
                    regras_vinculacao=_get_regras_vinculacao(projeto_codigo))


# ── utilidades ──────────────────────────────────────────────────────────────
def hash_arquivo(caminho: str) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


def ler_arquivo(caminho: str) -> list:
    """Extrai os itens {pagina, texto, cor, layer} do DXF/PDF (as mesmas funções de POST /upload)."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext not in EXTENSOES:
        raise ErroPipeline(f"Tipo de arquivo não suportado: '{ext}' (use .dxf ou .pdf).")
    if os.path.getsize(caminho) > LIMITE_BYTES:
        raise ErroPipeline("Arquivo grande demais (limite de 200 MB).")
    try:
        if ext == ".dxf":
            import ezdxf
            from services.dxf_service import extract_dxf_content
            return extract_dxf_content(ezdxf.readfile(caminho))
        import pymupdf
        from services.pdf_service import extract_pdf_content
        with pymupdf.open(caminho) as doc:
            return extract_pdf_content(doc)
    except ErroPipeline:
        raise
    except Exception as e:  # noqa: BLE001 — arquivo corrompido não pode derrubar a fila
        raise ErroPipeline(f"Não foi possível ler o arquivo: {e}") from e


def _destrutiva(op: dict, acoes: list) -> bool:
    return op["op"] == "excluir" or any(acoes[k]["acao"] in ACOES_DESTRUTIVAS for k in op.get("acoes", []) if 0 <= k < len(acoes))


def _frase_op(op: dict) -> str:
    v = lambda x: f"{x.get('operacao') or '·'} {x.get('ativo') or '(vazio)'}"  # noqa: E731
    if op["op"] == "excluir":
        return f"Excluir {op['linha_id']}: {v(op['antes'])}"
    if op["op"] == "editar":
        return f"Editar {op['linha_id']}: {v(op['antes'])} → {v(op['depois'])}"
    if op["op"] == "inserir":
        return f"Inserir ({op['tabela']}): {v(op['depois'])}"
    return f"Reordenar {op['tabela']}"


class _Etapas:
    def __init__(self):
        self.lista = []

    def __call__(self, nome):
        etapas = self

        class _Ctx:
            def __enter__(self):
                self.t = time.time()
                self.item = {"etapa": nome, "status": "ok", "ms": 0}
                etapas.lista.append(self.item)
                return self.item

            def __exit__(self, tipo, valor, tb):
                self.item["ms"] = round((time.time() - self.t) * 1000)
                if tipo is not None:
                    self.item["status"] = "erro"
                    self.item["detalhe"] = str(valor)
                return False
        return _Ctx()


def _validar(ctx: Contexto, cabos: list, outros: list) -> list:
    return validar_planilhas(cabos, outros) + avaliar(ctx.regras_dominio, cabos, outros, ctx.grupos)


def _salvar_obra(obra_id: str, nome: str, projeto: str, user_id: str, dados: dict) -> None:
    """Mesma gravação de POST /api/obras (SQLite + Supabase quando disponível)."""
    import database
    from services.supabase_client import get_supabase
    registro = {"id": obra_id, "nome": nome, "data": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"), "dados_json": json.dumps(dados, ensure_ascii=False),
                "user_id": user_id, "projeto": projeto}
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("obras").upsert(registro).execute()
        except Exception:  # noqa: BLE001 — como a rota: a nuvem é best-effort, o local é a fonte
            pass
    conn = database.get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id, projeto) VALUES (?, ?, ?, ?, ?, ?)",
                     (registro["id"], registro["nome"], registro["data"], registro["dados_json"], registro["user_id"], registro["projeto"]))
        conn.commit()
    finally:
        conn.close()


# ── fluxo ───────────────────────────────────────────────────────────────────
def processar_arquivo(caminho: str, projeto_codigo: str, user_id: str, pasta_saida: str, ctx: Contexto = None, forcar: bool = False) -> dict:
    """Processa UM arquivo. Devolve a execução (dict). Nunca levanta por falha do arquivo: o erro vira status 'erro'."""
    if not os.path.isfile(caminho):
        raise FileNotFoundError(caminho)
    h = hash_arquivo(caminho)
    if not forcar:
        anterior = execucoes.buscar_por_hash(h, projeto_codigo)
        if anterior:
            return {**anterior, "duplicado": True}
    exec_id = execucoes.criar(os.path.basename(caminho), h, projeto_codigo, user_id)
    execucoes.atualizar(exec_id, arquivo_caminho=os.path.abspath(caminho))
    pasta = saida.pasta_da_execucao(pasta_saida, projeto_codigo, caminho, exec_id)
    etapas = _Etapas()
    try:
        ctx = ctx or carregar_contexto(projeto_codigo, user_id)
        leitor = ctx.leitor or obter_leitor()
        ctx.leitor = leitor
        with etapas("ler") as e:
            itens = ler_arquivo(caminho)
            e["detalhe"] = {"itens": len(itens)}
            if not itens:
                raise ErroPipeline("O arquivo não tem texto para processar.")
        if not ctx.regras_proc and not ctx.regras_cls:
            raise ErroPipeline(f"O projeto '{projeto_codigo}' não tem regras do leitor cadastradas.")
        with etapas("montar") as e:
            tabelas = montar_tabelas(itens, ctx.regras_proc, ctx.regras_cls, leitor)
            e["detalhe"] = {k: len(v) for k, v in tabelas.items()}
        with etapas("ramais") as e:      # decisão do usuário: os ramais ENTRAM no orçamento autônomo (como o botão "Adicionar" do modal RAMAIS)
            novas = linhas_ramais([r["texto"] for r in tabelas.get("ramais", [])])
            tabelas["outros"] = tabelas["outros"] + novas
            e["detalhe"] = {"itens_ramais": len(tabelas.get("ramais", [])), "linhas_geradas": len(novas)}
        with etapas("validar_antes") as e:
            antes = _validar(ctx, tabelas["cabos"], tabelas["outros"])
            e["detalhe"] = resumir(antes)
        with etapas("ajustar") as e:
            acoes, ignorados = [], []
            for r in ctx.receitas:
                erros = validar_receitas([r], ctx.grupos)
                if erros:
                    ignorados.append({"ajuste": r.get("id"), "erros": erros})
                else:
                    acoes += r.get("acoes") or []
            if len(acoes) > LIMITE_ACOES_AUTONOMO:
                raise ErroPipeline(f"Ajustes habilitados somam {len(acoes)} ações (máx. {LIMITE_ACOES_AUTONOMO} em cadeia no modo autônomo).")
            try:
                diff = ajustar(acoes, tabelas["cabos"], tabelas["outros"], ctx.grupos, LIMITE_ACOES_AUTONOMO) if acoes else \
                    {"operacoes": [], "descartadas": [], "avisos": [], "resumo": {}}
            except ValueError as ve:
                raise ErroPipeline(f"Os ajustes não puderam ser calculados: {ve.args[0] if ve.args else ve}") from ve
            ops = diff["operacoes"]
            pendentes = [i for i, op in enumerate(ops) if _destrutiva(op, acoes)]
            e["detalhe"] = {"operacoes": len(ops), "pendentes_de_confirmacao": len(pendentes), "descartadas": len(diff["descartadas"]),
                            "ajustes_ignorados": ignorados}
        execucoes.atualizar(exec_id, pasta_saida=pasta, originais=tabelas,
                            diff={"operacoes": ops, "acoes": acoes, "descartadas": diff["descartadas"], "avisos": diff["avisos"],
                                  "frases": [_frase_op(op) for op in ops]},
                            decisoes={"pendentes": pendentes, "confirmadas": [], "rejeitadas": []},
                            relatorio={"etapas": etapas.lista, "validacao_antes": resumir(antes), "ajustes_ignorados": ignorados})
        if pendentes:
            _gravar_pendencias(exec_id, pasta)
            execucoes.atualizar(exec_id, status="aguardando_confirmacao",
                                mensagem=f"{len(pendentes)} exclusão(ões) aguardando confirmação.")
        else:
            _finalizar(exec_id, ctx, pasta)
    except Exception as e:  # noqa: BLE001
        msg = str(e) if isinstance(e, ErroPipeline) else f"Erro inesperado: {type(e).__name__}"
        execucoes.atualizar(exec_id, status="erro", mensagem=msg,
                            relatorio={"etapas": etapas.lista, "erro": msg})
        if not isinstance(e, ErroPipeline):
            import logging
            logging.getLogger(__name__).exception("Falha inesperada no modo autônomo (execução %s)", exec_id)
    return execucoes.buscar(exec_id)


def _gravar_pendencias(exec_id: str, pasta: str) -> None:
    ex = execucoes.buscar(exec_id)
    ops = ex["diff"]["operacoes"]
    lista = [{"indice": i, "descricao": ex["diff"]["frases"][i], "operacao": ops[i]} for i in ex["decisoes"]["pendentes"]]
    saida.gravar(pasta, tabelas=ex["originais"], pendencias=lista, relatorio={"status": "aguardando_confirmacao", **(ex["relatorio"] or {})})


def _finalizar(exec_id: str, ctx: Contexto, pasta: str, reverter: bool = False) -> None:
    """Aplica o que foi decidido sobre as tabelas ORIGINAIS, gera Totalizadora/orçamento, salva obra e pasta e fecha a execução."""
    ex = execucoes.buscar(exec_id)
    ops, dec = ex["diff"]["operacoes"], ex["decisoes"]
    etapas = _Etapas()
    etapas.lista = list((ex["relatorio"] or {}).get("etapas", []))
    if reverter:
        escolhidas = set()
    else:
        escolhidas = {i for i in range(len(ops)) if i not in set(dec["pendentes"])} | set(dec["confirmadas"])
    with etapas("aplicar") as e:
        r = aplicar_operacoes(ex["originais"]["cabos"], ex["originais"]["outros"], ops, escolhidas, ctx.regras_cls, ctx.leitor or obter_leitor())
        cabos, outros = r["cabos"], r["outros"]
        e["detalhe"] = {"aplicadas": r["aplicadas"], "ignoradas": r["ignoradas"], "rejeitadas": len(dec["rejeitadas"]) if not reverter else 0}
    with etapas("vincular") as e:
        # TASK-036: auto-link por coordenada sobre as tabelas FINAIS (depois do diff aplicado) — nunca antes, para não
        # correr o risco de `vinculoEstruturas` (índices em `outros`) ficar desalinhado por uma exclusão do ajuste.
        e["detalhe"] = tentar_vincular_automaticamente(cabos, outros, ctx.regras_vinculacao)
    with etapas("validar_depois") as e:
        achados_vinculo = avaliar_vinculacao(cabos, outros, ctx.regras_vinculacao)
        depois = _validar(ctx, cabos, outros) + achados_vinculo
        e["detalhe"] = resumir(depois)
    with etapas("totalizadora") as e:
        itens_vinculo = itens_vinculo_para_totalizadora(cabos, outros)
        tot_bruta = montar_totalizadora(cabos, outros, ctx.regras_conversao, ctx.base_orcamento, ctx.projeto_nome, itens_vinculo=itens_vinculo)
        payload = payload_calculo(tot_bruta)
        totalizadora = normalizar_para_json(tot_bruta)
        e["detalhe"] = {"linhas": len(totalizadora)}
    with etapas("orcamento") as e:
        orcamento = processar_calculo(payload["cabos"], payload["outros"], ctx.projeto_nome, ctx.base_orcamento)
        e["detalhe"] = {"linhas": len(orcamento["resultado"]), "nao_encontrados": len(orcamento["nao_encontrados"])}
    obra_id = f"auto_{exec_id}"
    with etapas("salvar") as e:
        nome = f"[Autônomo] {os.path.splitext(ex['arquivo'])[0]}" + (" (original)" if reverter else "")
        snapshot = {"cabos": {"bodyId": "body-cabos", "data": cabos}, "outros": {"bodyId": "body-outros", "data": outros},
                    "totalizadora": {"bodyId": "body-totalizadora", "data": totalizadora},
                    "regras": {"bodyId": "body-regras", "data": ctx.regras_conversao},
                    "autonomo": {"execucao_id": exec_id, "arquivo": ex["arquivo"], "revertida": reverter}}
        _salvar_obra(obra_id, nome, ex["projeto_codigo"], ex["user_id"], snapshot)
        ex_diff = ex["diff"]
        relatorio = {
            "execucao_id": exec_id, "arquivo": ex["arquivo"], "projeto": ex["projeto_codigo"], "etapas": etapas.lista,
            "validacao_antes": (ex["relatorio"] or {}).get("validacao_antes"), "validacao_depois": resumir(depois),
            "achados_restantes": depois, "ajustes_ignorados": (ex["relatorio"] or {}).get("ajustes_ignorados", []),
            "ajustes": {"aplicados": [ex_diff["frases"][i] for i in sorted(escolhidas)], "rejeitados": [ex_diff["frases"][i] for i in dec["rejeitadas"]]
                        if not reverter else [], "descartados": ex_diff["descartadas"], "avisos": ex_diff["avisos"],
                        "ignorados_por_linha_alterada": r["ignoradas_detalhe"]},
            "orcamento": {"linhas": len(orcamento["resultado"]), "nao_encontrados": sorted(orcamento["nao_encontrados"])},
            "obra_id": obra_id,
        }
    pendencia = bool(resumir(depois).get("erro")) or bool(ex_diff["descartadas"]) or bool(r["ignoradas"]) \
        or bool(dec["rejeitadas"] and not reverter) or bool(orcamento["nao_encontrados"])
    status = "revertida" if reverter else ("com_pendencias" if pendencia else "ok")
    relatorio["status"] = status
    saida.gravar(pasta, tabelas={"cabos": cabos, "outros": outros, "ramais": ex["originais"].get("ramais", [])}, totalizadora=totalizadora,
                 orcamento=orcamento, relatorio=relatorio, pendencias=[])
    execucoes.atualizar(exec_id, status=status, obra_id=obra_id, relatorio=relatorio,
                        mensagem={"ok": "Concluída.", "com_pendencias": "Concluída com pendências (veja o relatório).",
                                  "revertida": "Revertida: obra com as tabelas originais."}[status])


def _contexto_da(ex: dict, ctx: Contexto = None) -> Contexto:
    return ctx or carregar_contexto(ex["projeto_codigo"], ex["user_id"])


def _exige(exec_id: str, status: tuple) -> dict:
    ex = execucoes.buscar(exec_id)
    if not ex:
        raise ErroPipeline("Execução não encontrada.")
    if ex["status"] not in status:
        raise ErroPipeline(f"Esta execução está '{ex['status']}'; ação disponível só para: {', '.join(status)}.")
    return ex


def _decidir(exec_id: str, indices, confirmar: bool, ctx: Contexto = None) -> dict:
    ex = _exige(exec_id, ("aguardando_confirmacao",))
    dec = ex["decisoes"]
    pendentes = [i for i in dec["pendentes"] if i not in dec["confirmadas"] and i not in dec["rejeitadas"]]
    alvo = pendentes if indices is None else list(dict.fromkeys(int(i) for i in indices))
    invalidos = [i for i in alvo if i not in pendentes]
    if invalidos:
        raise ErroPipeline(f"Índice(s) sem pendência de confirmação: {invalidos}.")
    dec["confirmadas" if confirmar else "rejeitadas"] += alvo
    execucoes.atualizar(exec_id, decisoes=dec)
    restantes = [i for i in dec["pendentes"] if i not in dec["confirmadas"] and i not in dec["rejeitadas"]]
    if restantes:
        execucoes.atualizar(exec_id, mensagem=f"{len(restantes)} exclusão(ões) ainda aguardando confirmação.")
        _gravar_pendencias(exec_id, ex["pasta_saida"])
    else:
        try:
            _finalizar(exec_id, _contexto_da(ex, ctx), ex["pasta_saida"])
        except Exception as e:  # noqa: BLE001
            execucoes.atualizar(exec_id, status="erro", mensagem=f"Falha ao concluir após a decisão: {e}")
    return execucoes.buscar(exec_id)


def confirmar(exec_id: str, indices=None, ctx: Contexto = None) -> dict:
    """'Sim' (indices) ou 'Sim para todos' (indices=None: todas as exclusões pendentes DESTE arquivo). Termina o pipeline quando nada mais pende."""
    return _decidir(exec_id, indices, True, ctx)


def rejeitar(exec_id: str, indices=None, ctx: Contexto = None) -> dict:
    """'Não': as exclusões indicadas (ou todas) NÃO são aplicadas; o arquivo termina marcado 'com_pendencias'."""
    return _decidir(exec_id, indices, False, ctx)


def reverter(exec_id: str, ctx: Contexto = None) -> dict:
    """Refaz obra, orçamento e pasta com as tabelas ORIGINAIS (nenhum ajuste). Só para execuções concluídas."""
    ex = _exige(exec_id, ("ok", "com_pendencias"))
    _finalizar(exec_id, _contexto_da(ex, ctx), ex["pasta_saida"], reverter=True)
    return execucoes.buscar(exec_id)
