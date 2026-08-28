import sqlite3
from services.supabase_client import get_supabase
from config import DB_PATH

def migrate_data():
    supabase = get_supabase()
    if not supabase:
        print("Erro: Supabase não conectado.")
        return

    # Conectar ao banco local
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Buscar os dados locais
    cursor.execute("SELECT * FROM tabela_orcamento")
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        print("A tabela local está vazia. Não há nada para migrar.")
        return

    print(f"Encontrados {len(rows)} registros locais. Preparando para enviar...")

    records_to_insert = []
    for row in rows:
        records_to_insert.append({
            "ativo": row["ativo"],
            "desc_ativo": row["desc_ativo"],
            "componente": row["componente"],
            "projeto": row["projeto"],
            "mdo": row["mdo"],
            "codigo": row["codigo"],
            "desc_codigo": row["desc_codigo"],
            "fator_i": row["fator_i"],
            "fator_r": row["fator_r"],
            "filtro": row["filtro"]
        })

    # Inserir no Supabase (em lotes para não sobrecarregar)
    batch_size = 500
    try:
        # Primeiro, limpa a tabela master na nuvem
        supabase.table("tabela_orcamento_master").delete().neq("id", 0).execute()
        print("Tabela mestre na nuvem limpa. Iniciando uploads...")

        for i in range(0, len(records_to_insert), batch_size):
            batch = records_to_insert[i:i + batch_size]
            supabase.table("tabela_orcamento_master").insert(batch).execute()
            print(f"Enviados {i + len(batch)} de {len(records_to_insert)} registros...")

        print("Migração concluída com sucesso!")
    except Exception as e:
        print(f"Erro durante a migração: {e}")

if __name__ == "__main__":
    migrate_data()
