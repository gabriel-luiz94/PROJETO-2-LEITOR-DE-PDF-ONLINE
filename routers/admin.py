"""
routers/admin.py — Rotas exclusivas para o Administrador.
"""
import io
import csv
from fastapi import APIRouter, UploadFile, File, Request, HTTPException
from fastapi.responses import JSONResponse
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/admin", tags=["admin"])


@router.post("/upload-master")
async def upload_master_csv(request: Request, file: UploadFile = File(...)):
    # 1. Valida se o usuário é admin
    auth_header = request.headers.get("Authorization")
    if not auth_header or not auth_header.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Não autorizado")
    
    token = auth_header.split(" ")[1]
    
    supabase = get_supabase()
    if not supabase:
        raise HTTPException(status_code=503, detail="Supabase não configurado.")

    try:
        # Pega o usuário logado no Supabase a partir do token
        user_res = supabase.auth.get_user(token)
        if not user_res or not user_res.user:
            raise HTTPException(status_code=401, detail="Sessão inválida.")
            
        # (Futuro) Verificar na tabela profiles se o user é admin.
        # Por enquanto, assumimos que quem chega aqui com token válido e sabe usar a tela pode subir.
        # Em produção restrita, você checaria: if user_res.user.email != 'seuemail@admin.com': raise ...
    except Exception as e:
        raise HTTPException(status_code=401, detail="Sessão inválida ou erro de auth.")

    # 2. Processa o CSV
    try:
        contents = await file.read()
        text = contents.decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(text), delimiter=";")

        if not reader.fieldnames or len(reader.fieldnames) < 2:
            reader = csv.DictReader(io.StringIO(text), delimiter=",")

        records_to_insert = []
        for row in reader:
            row_upper = {k.strip().upper(): v for k, v in row.items() if k}
            ativo = row_upper.get("ATIVO", "")
            codigo = row_upper.get("CODIGO", "")
            if not ativo and not codigo:
                continue

            try:
                fator_i = float(row_upper.get("FATOR I", "0").replace(",", "."))
            except ValueError:
                fator_i = 0.0

            try:
                fator_r = float(row_upper.get("FATOR R", "0").replace(",", "."))
            except ValueError:
                fator_r = 0.0

            records_to_insert.append({
                "ativo": ativo,
                "desc_ativo": row_upper.get("DESC ATIVO", ""),
                "componente": row_upper.get("COMPONENTE", ""),
                "projeto": row_upper.get("PROJETO", ""),
                "mdo": row_upper.get("MDO", ""),
                "codigo": row_upper.get("CODIGO", ""),
                "desc_codigo": row_upper.get("DESC CODIGO", ""),
                "fator_i": fator_i,
                "fator_r": fator_r,
                "filtro": row_upper.get("FILTRO", "")
            })

        # 3. Limpa a tabela e sobe os novos pro Supabase
        # IMPORTANTE: Supabase REST API não tem um TRUNCATE simples que retorne tudo,
        # geralmente deletamos onde id > 0 ou usando rpc.
        # Vamos usar um DELETE com filtro.
        supabase.table("tabela_orcamento_master").delete().neq("id", 0).execute()
        
        # Insere em lotes se for muito grande
        batch_size = 500
        for i in range(0, len(records_to_insert), batch_size):
            batch = records_to_insert[i:i + batch_size]
            supabase.table("tabela_orcamento_master").insert(batch).execute()

        return {"status": "ok", "message": f"{len(records_to_insert)} registros atualizados na Nuvem."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@router.post("/add-row")
async def add_master_row(request: Request):
    # 1. Valida se o usuário é admin
    auth_header = request.headers.get("Authorization")
    if not auth_header or not auth_header.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Não autorizado")
    
    token = auth_header.split(" ")[1]
    
    supabase = get_supabase()
    if not supabase:
        raise HTTPException(status_code=503, detail="Supabase não configurado.")

    try:
        user_res = supabase.auth.get_user(token)
        if not user_res or not user_res.user:
            raise HTTPException(status_code=401, detail="Sessão inválida.")
    except Exception as e:
        raise HTTPException(status_code=401, detail="Sessão inválida ou erro de auth.")

    data = await request.json()

    # Prepara o payload
    row_data = {
        "ativo": data.get("ativo", ""),
        "desc_ativo": data.get("desc_ativo", ""),
        "componente": data.get("componente", ""),
        "projeto": data.get("projeto", ""),
        "mdo": data.get("mdo", ""),
        "codigo": data.get("codigo", ""),
        "desc_codigo": data.get("desc_codigo", ""),
        "fator_i": data.get("fator_i", 0.0),
        "fator_r": data.get("fator_r", 0.0),
        "filtro": data.get("filtro", "")
    }

    try:
        # Insere a nova linha no Supabase
        supabase.table("tabela_orcamento_master").insert(row_data).execute()
        return {"status": "ok", "message": "Linha inserida com sucesso."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
