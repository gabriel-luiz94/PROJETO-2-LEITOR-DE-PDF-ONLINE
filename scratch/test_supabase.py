from services.sync_service import sync_tabela_master
from services.supabase_client import get_supabase

client = get_supabase()
if client:
    print("Conexão Supabase OK!")
    res = sync_tabela_master()
    if res:
        print("Sincronização rodou sem erros (Tabela vazia ou populada)!")
    else:
        print("Erro na sincronização.")
else:
    print("Falha ao inicializar o cliente Supabase.")
