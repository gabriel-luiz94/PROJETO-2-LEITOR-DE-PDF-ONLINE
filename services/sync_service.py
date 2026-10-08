"""
services/sync_service.py — Lógica de Sincronização Offline-First (Supabase <-> SQLite)
"""
from services.supabase_client import get_supabase
from services.offline_queue import process_queue
from database import get_connection, get_row_connection
from config import logger


def sync_tabela_master():
    """
    Baixa a tabela mestre (Admin) do Supabase e atualiza o SQLite local.
    Suporta paginação para baixar mais de 1000 registros do Supabase.
    """
    supabase = get_supabase()
    if not supabase:
        logger.info("Sync Master cancelado: Cliente Supabase não configurado/offline.")
        return False

    try:
        process_queue()

        # Busca todas as páginas do Supabase (1000 por lote)
        master_rows = []
        page_size = 1000
        start = 0
        while True:
            res = supabase.table("tabela_orcamento_master").select("*").range(start, start + page_size - 1).execute()
            data = res.data or []
            master_rows.extend(data)
            if len(data) < page_size:
                break
            start += page_size

        conn = get_connection()
        cursor = conn.cursor()
        
        cursor.execute("BEGIN TRANSACTION")
        cursor.execute("DELETE FROM tabela_orcamento_master")
        
        for row in master_rows:
            cursor.execute('''
                INSERT INTO tabela_orcamento_master 
                (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                row.get("ativo", ""),
                row.get("desc_ativo", ""),
                row.get("componente", ""),
                row.get("projeto", ""),
                row.get("mdo", ""),
                row.get("codigo", ""),
                row.get("desc_codigo", ""),
                row.get("fator_i", 0.0),
                row.get("fator_r", 0.0),
                row.get("filtro", ""),
                row.get("origem", "")
            ))
            
        cursor.execute("COMMIT")
        conn.close()
        logger.info(f"Sync Master concluído. {len(master_rows)} registros atualizados localmente.")
        return True

    except Exception as e:
        logger.error(f"Erro ao sincronizar tabela master: {e}")
        try:
            conn = get_connection()
            conn.cursor().execute("ROLLBACK")
            conn.close()
        except:
            pass
        return False


def normalizar_projeto(valor) -> str:
    """TASK-052: mesma normalização usada por `processar_calculo` pra comparar nomes de projeto."""
    return (valor or "").strip().upper()


def validar_linhas_do_projeto(dados: list, projeto: str) -> list:
    """TASK-052: pra uma importação/gravação escopada a um projeto (Importar CSV Master,
    Sincronizar Tudo, Salvar Alterações quando chamados com `projeto`), confere que toda linha com
    `projeto` preenchido bate com o projeto selecionado (normalizado) — linha com `projeto` vazio
    recebe o projeto selecionado automaticamente (conveniência: uma planilha exportada da própria
    visão por projeto já vem com o campo preenchido, mas uma digitada à mão pode vir em branco).
    Devolve a lista com o campo `projeto` já resolvido; lança `ValueError` com a primeira linha
    divergente encontrada, se houver — quem chama decide como reportar (não grava nada antes)."""
    projeto_norm = normalizar_projeto(projeto)
    resolvidas = []
    for linha in dados:
        projeto_linha = normalizar_projeto(linha.get("projeto"))
        if projeto_linha and projeto_linha != projeto_norm:
            raise ValueError(
                f"Linha com ativo '{linha.get('ativo', '')}' código '{linha.get('codigo', '')}' "
                f"pertence ao projeto '{linha.get('projeto')}', diferente do projeto selecionado "
                f"'{projeto}'."
            )
        resolvidas.append({**linha, "projeto": projeto_norm})
    return resolvidas


def filtrar_linhas_por_categoria(linhas: list, *, projeto: str = None, genericas: bool = False,
                                  nao_reconhecido: bool = False, projetos_validos=None) -> list:
    """TASK-052: recorte da base pra visão por projeto em `/orcamento` — `projeto` (igualdade
    exata normalizada), `genericas` (`projeto` vazio) ou `nao_reconhecido` (`projeto` preenchido
    mas fora de `projetos_validos`). No máximo um critério é aplicado por chamada; sem nenhum,
    devolve `linhas` sem filtrar (comportamento usado hoje por quem não pedir nenhuma categoria)."""
    if genericas:
        return [r for r in linhas if not normalizar_projeto(r.get("projeto"))]
    if nao_reconhecido:
        validos = {normalizar_projeto(p) for p in (projetos_validos or [])}
        return [r for r in linhas
                if normalizar_projeto(r.get("projeto")) and normalizar_projeto(r.get("projeto")) not in validos]
    if projeto:
        projeto_norm = normalizar_projeto(projeto)
        return [r for r in linhas if normalizar_projeto(r.get("projeto")) == projeto_norm]
    return linhas


def _linhas_divididas(linha: dict) -> list:
    """TASK-052: divide uma linha cujo `projeto` combina N nomes separados por "/" (ex.
    "PARAIBA/PARAIBANOVO") em N dicionários, cada um com um projeto só e os demais campos
    idênticos. Linha sem "/" (ou com um nome só) devolve lista vazia (nada a dividir)."""
    projetos = [p.strip() for p in (linha.get("projeto") or "").split("/") if p.strip()]
    if len(projetos) < 2:
        return []
    campos = {k: v for k, v in linha.items() if k not in ("id", "updated_at")}
    return [{**campos, "projeto": p} for p in projetos]


def migrar_projetos_compartilhados() -> dict:
    """TASK-052: divide toda linha de `tabela_orcamento_master` cujo `projeto` combina mais de um
    nome separado por "/" numa linha exclusiva por projeto da lista, com os mesmos valores dos
    demais campos — sem perda de dado, só duplicação em cópias independentes. Idempotente: sem
    nenhuma linha com "/", não faz nada.

    Local e Supabase são migrados cada um a partir da sua própria consulta: os `id`s de uma base
    não correspondem aos da outra (`sync_tabela_master` nunca preserva `id` ao copiar do Supabase
    pro SQLite, cada um tem seu próprio autoincremento/identity independente), então tentar casar
    uma linha local com sua "equivalente" na nuvem pelo `id` seria errado — a divisão roda duas
    vezes, cada uma só contra a sua própria fonte.
    """
    resultado = {"linhas_divididas_local": 0, "linhas_criadas_local": 0,
                 "linhas_divididas_nuvem": 0, "linhas_criadas_nuvem": 0}

    # ── Local (SQLite) ──
    conn = get_row_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tabela_orcamento_master WHERE projeto LIKE '%/%'")
    linhas_local = [dict(r) for r in cursor.fetchall()]
    conn.close()

    if linhas_local:
        conn = get_connection()
        cursor = conn.cursor()
        for linha in linhas_local:
            divididas = _linhas_divididas(linha)
            if not divididas:
                continue
            cursor.execute("UPDATE tabela_orcamento_master SET projeto = ? WHERE id = ?",
                           (divididas[0]["projeto"], linha["id"]))
            for nova in divididas[1:]:
                cursor.execute('''
                    INSERT INTO tabela_orcamento_master
                    (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    nova.get("ativo"), nova.get("desc_ativo"), nova.get("componente"), nova.get("projeto"),
                    nova.get("mdo"), nova.get("codigo"), nova.get("desc_codigo"),
                    nova.get("fator_i"), nova.get("fator_r"), nova.get("filtro"), nova.get("origem"),
                ))
                resultado["linhas_criadas_local"] += 1
            resultado["linhas_divididas_local"] += 1
        conn.commit()
        conn.close()

    # ── Supabase ──
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("tabela_orcamento_master").select("*").like("projeto", "%/%").execute()
            linhas_nuvem = res.data or []
            for linha in linhas_nuvem:
                divididas = _linhas_divididas(linha)
                if not divididas:
                    continue
                supabase.table("tabela_orcamento_master").update(
                    {"projeto": divididas[0]["projeto"]}).eq("id", linha["id"]).execute()
                for nova in divididas[1:]:
                    nova_sem_id = {k: v for k, v in nova.items() if k != "id"}
                    supabase.table("tabela_orcamento_master").insert(nova_sem_id).execute()
                    resultado["linhas_criadas_nuvem"] += 1
                resultado["linhas_divididas_nuvem"] += 1
        except Exception as e:
            logger.error(f"Erro ao migrar projetos compartilhados no Supabase: {e}")
            raise RuntimeError(f"Dados locais migrados, mas falhou ao sincronizar com o Supabase: {e}")

    logger.info(
        f"Migração de projetos compartilhados concluída: local "
        f"{resultado['linhas_divididas_local']} divididas/{resultado['linhas_criadas_local']} criadas; "
        f"nuvem {resultado['linhas_divididas_nuvem']} divididas/{resultado['linhas_criadas_nuvem']} criadas."
    )
    return resultado


def get_merged_orcamento(user_id: str = None) -> list[dict]:
    """
    Retorna a tabela de orçamento priorizando as edições locais do usuário.
    Se o usuário não tiver salvo edições locais, retorna a tabela Master do Supabase.
    Se a Master estiver vazia, retorna o seed inicial (user_id IS NULL).
    """
    conn = get_row_connection()
    cursor = conn.cursor()
    
    # 1. Tabela local com as modificações do próprio usuário
    if user_id:
        cursor.execute("SELECT * FROM tabela_orcamento WHERE user_id = ?", (user_id,))
        user_rows = [dict(r) for r in cursor.fetchall()]
        if user_rows:
            conn.close()
            return user_rows

    # 2. Tabela Master oficial sincronizada do Supabase
    cursor.execute("SELECT * FROM tabela_orcamento_master")
    master_rows = [dict(r) for r in cursor.fetchall()]
    if master_rows and len(master_rows) > 0:
        conn.close()
        return master_rows
    
    # 3. Fallback genérico (Seed sem user_id)
    cursor.execute("SELECT * FROM tabela_orcamento WHERE user_id IS NULL")
    seed_rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return seed_rows
