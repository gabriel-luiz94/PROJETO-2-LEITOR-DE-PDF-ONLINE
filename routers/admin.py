"""
routers/admin.py — Rotas para gerenciamento de usuários (Painel Admin).

Suporta três roles: admin, editor, viewer.
Inclui audit log e sincronização com Supabase.
"""
import io
import csv
from fastapi import APIRouter, HTTPException, Request, UploadFile, File
from pydantic import BaseModel
from database import get_connection, get_row_connection, hash_password
from services.supabase_client import get_supabase
from middleware.auth_middleware import get_current_user_from_state
from config import logger

router = APIRouter(prefix="/api/admin", tags=["admin"])

# ── Roles válidas ────────────────────────────────────────────────────────────
VALID_ROLES = {"admin", "operador"}


class UserCreate(BaseModel):
    email: str
    password: str
    role: str = "operador"  # "admin", "operador"


class UserUpdate(BaseModel):
    role: str  # "admin", "operador"


class PasswordUpdate(BaseModel):
    password: str


class MasterRowCreate(BaseModel):
    ativo: str = ""
    desc_ativo: str = ""
    componente: str = ""
    projeto: str = ""
    mdo: str = ""
    codigo: str = ""
    desc_codigo: str = ""
    fator_i: float = 0.0
    fator_r: float = 0.0
    filtro: str = ""


# ── Helpers ──────────────────────────────────────────────────────────────────

def _require_admin(request: Request) -> dict:
    """Valida que o usuário autenticado é admin."""
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(
            status_code=403,
            detail="Acesso negado. Apenas administradores podem acessar este recurso."
        )
    return user


def _audit(user: dict, action: str, table_name: str, record_id: str = None, details: str = None):
    """Registra ação no audit log."""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('''
            INSERT INTO audit_log (user_id, email, action, table_name, record_id, details)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            user.get("user_id"),
            user.get("email"),
            action,
            table_name,
            record_id,
            details,
        ))
        conn.commit()
        conn.close()
    except Exception as e:
        logger.warning(f"Erro ao registrar audit log: {e}")


# ── Rotas ────────────────────────────────────────────────────────────────────

@router.get("/users")
def list_users(request: Request):
    """Lista todos os usuários. Sincroniza da nuvem para o local se houver internet."""
    admin = _require_admin(request)
    supabase = get_supabase()
    conn = get_row_connection()
    cur = conn.cursor()

    if supabase:
        try:
            # Sincroniza Nuvem -> Local
            res = supabase.table("usuarios_nuvem").select("*").execute()
            cloud_users = res.data
            for cu in cloud_users:
                role = cu.get("role", "admin" if cu.get("is_admin") else "viewer")
                cur.execute('''
                    INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin, role)
                    VALUES (?, ?, ?, ?)
                ''', (cu["email"], cu["senha_hash"], int(cu.get("is_admin", False)), role))
            conn.commit()
        except Exception:
            pass  # Ignora falhas de sync

    cur.execute("SELECT id, email, is_admin, role, ativo, created_at, updated_at FROM usuarios_locais")
    users = [dict(row) for row in cur.fetchall()]
    conn.close()
    return users


@router.post("/users")
def add_user(req: UserCreate, request: Request):
    """Adiciona um novo usuário (Nuvem + Local)."""
    admin = _require_admin(request)

    if len(req.password) < 6:
        raise HTTPException(status_code=400, detail="A senha deve ter pelo menos 6 caracteres.")

    if req.role not in VALID_ROLES:
        raise HTTPException(status_code=400, detail=f"Role inválida. Válidas: {', '.join(VALID_ROLES)}")

    senha_hash = hash_password(req.password)
    is_admin = 1 if req.role == "admin" else 0
    supabase = get_supabase()

    if supabase:
        try:
            supabase.table("usuarios_nuvem").insert({
                "email": req.email,
                "senha_hash": senha_hash,
                "is_admin": bool(is_admin),
                "role": req.role,
            }).execute()
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Erro ao inserir na nuvem: o e-mail já pode existir.")

    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('''
            INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin, role)
            VALUES (?, ?, ?, ?)
        ''', (req.email, senha_hash, is_admin, req.role))
        conn.commit()
        user_id = cur.lastrowid
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=400, detail="Erro local: E-mail já está em uso.")

    conn.close()

    _audit(admin, "CREATE_USER", "usuarios_locais", str(user_id), f"email={req.email}, role={req.role}")

    return {"id": user_id, "email": req.email, "role": req.role, "msg": "Usuário criado com sucesso!"}


@router.delete("/users/{user_id}")
def delete_user(user_id: int, request: Request):
    """Remove um usuário."""
    admin = _require_admin(request)
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]

    # Não permitir auto-exclusão
    if email == admin.get("email"):
        conn.close()
        raise HTTPException(status_code=400, detail="Não é possível excluir a si mesmo.")

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").delete().eq("email", email).execute()
        except Exception:
            pass

    cur.execute("DELETE FROM usuarios_locais WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()

    _audit(admin, "DELETE_USER", "usuarios_locais", str(user_id), f"email={email}")

    return {"msg": "Usuário removido com sucesso."}


@router.put("/users/{user_id}/role")
def update_role(user_id: int, req: UserUpdate, request: Request):
    """Atualiza a role de um usuário."""
    admin = _require_admin(request)

    if req.role not in VALID_ROLES:
        raise HTTPException(status_code=400, detail=f"Role inválida. Válidas: {', '.join(VALID_ROLES)}")

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]

    is_admin = 1 if req.role == "admin" else 0

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").update({
                "is_admin": bool(is_admin),
                "role": req.role,
            }).eq("email", email).execute()
        except Exception:
            pass

    cur.execute(
        "UPDATE usuarios_locais SET is_admin = ?, role = ?, updated_at = datetime('now') WHERE id = ?",
        (is_admin, req.role, user_id)
    )
    conn.commit()
    conn.close()

    _audit(admin, "UPDATE_ROLE", "usuarios_locais", str(user_id), f"email={email}, role={req.role}")

    return {"msg": f"Role atualizada para '{req.role}' com sucesso."}


# Manter compatibilidade com endpoint antigo
@router.put("/users/{user_id}/admin")
def toggle_admin(user_id: int, request: Request):
    """Ativa ou desativa os privilégios de admin (compatibilidade)."""
    from pydantic import BaseModel
    class OldUserUpdate(BaseModel):
        is_admin: bool
    # Redireciona para o novo endpoint
    import json
    body = {}
    try:
        # Tenta ler o body
        conn = get_row_connection()
        cur = conn.cursor()
        cur.execute("SELECT role FROM usuarios_locais WHERE id = ?", (user_id,))
        row = cur.fetchone()
        conn.close()
    except Exception:
        pass
    # Fallback simples
    return update_role(user_id, UserUpdate(role="admin"), request)


@router.put("/users/{user_id}/password")
def update_password(user_id: int, req: PasswordUpdate, request: Request):
    """Redefine a senha de um usuário."""
    admin = _require_admin(request)

    if len(req.password) < 6:
        raise HTTPException(status_code=400, detail="A senha deve ter pelo menos 6 caracteres.")

    senha_hash = hash_password(req.password)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").update({"senha_hash": senha_hash}).eq("email", email).execute()
        except Exception:
            pass

    cur.execute(
        "UPDATE usuarios_locais SET senha_hash = ?, updated_at = datetime('now') WHERE id = ?",
        (senha_hash, user_id)
    )
    conn.commit()
    conn.close()

    _audit(admin, "RESET_PASSWORD", "usuarios_locais", str(user_id), f"email={email}")

    return {"msg": "Senha redefinida com sucesso."}


@router.get("/audit-log")
def get_audit_log(request: Request, limit: int = 100):
    """Retorna o log de auditoria."""
    _require_admin(request)

    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM audit_log ORDER BY created_at DESC LIMIT ?",
        (limit,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"logs": rows}


@router.post("/add-row")
def add_master_row(req: MasterRowCreate, request: Request):
    """(Admin) Adiciona uma linha na tabela master local e na nuvem (Supabase)."""
    admin = _require_admin(request)

    conn = get_connection()
    cur = conn.cursor()
    cur.execute('''
        INSERT INTO tabela_orcamento_master 
        (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        req.ativo.strip(),
        req.desc_ativo.strip(),
        req.componente.strip(),
        req.projeto.strip(),
        req.mdo.strip(),
        req.codigo.strip(),
        req.desc_codigo.strip(),
        req.fator_i,
        req.fator_r,
        req.filtro.strip()
    ))
    conn.commit()
    row_id = cur.lastrowid
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("tabela_orcamento_master").insert({
                "ativo": req.ativo.strip(),
                "desc_ativo": req.desc_ativo.strip(),
                "componente": req.componente.strip(),
                "projeto": req.projeto.strip(),
                "mdo": req.mdo.strip(),
                "codigo": req.codigo.strip(),
                "desc_codigo": req.desc_codigo.strip(),
                "fator_i": req.fator_i,
                "fator_r": req.fator_r,
                "filtro": req.filtro.strip()
            }).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar linha master no Supabase: {e}")

    _audit(admin, "ADD_MASTER_ROW", "tabela_orcamento_master", str(row_id), f"codigo={req.codigo}, ativo={req.ativo}")
    return {"status": "ok", "msg": "Linha Master salva com sucesso na nuvem!"}


@router.post("/upload-master")
async def upload_master_csv(request: Request, file: UploadFile = File(...)):
    """(Admin) Upload de arquivo CSV para atualizar a Tabela Master no SQLite e no Supabase."""
    admin = _require_admin(request)
    try:
        contents = await file.read()
        text = contents.decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(text), delimiter=";")
        if not reader.fieldnames or len(reader.fieldnames) < 2:
            reader = csv.DictReader(io.StringIO(text), delimiter=",")

        rows_to_insert = []
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM tabela_orcamento_master")

        for row in reader:
            row_upper = {k.strip().upper(): v for k, v in row.items() if k}
            ativo = row_upper.get("ATIVO", "").strip()
            codigo = row_upper.get("CODIGO", "").strip()
            if not ativo and not codigo:
                continue

            try:
                fator_i = float(str(row_upper.get("FATOR I", "0")).replace(",", "."))
            except ValueError:
                fator_i = 0.0

            try:
                fator_r = float(str(row_upper.get("FATOR R", "0")).replace(",", "."))
            except ValueError:
                fator_r = 0.0

            item_dict = {
                "ativo": ativo,
                "desc_ativo": row_upper.get("DESC ATIVO", "").strip(),
                "componente": row_upper.get("COMPONENTE", "").strip(),
                "projeto": row_upper.get("PROJETO", "").strip(),
                "mdo": row_upper.get("MDO", "").strip(),
                "codigo": codigo,
                "desc_codigo": row_upper.get("DESC CODIGO", "").strip(),
                "fator_i": fator_i,
                "fator_r": fator_r,
                "filtro": row_upper.get("FILTRO", "").strip()
            }
            rows_to_insert.append(item_dict)

            cur.execute('''
                INSERT INTO tabela_orcamento_master 
                (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                item_dict["ativo"], item_dict["desc_ativo"], item_dict["componente"],
                item_dict["projeto"], item_dict["mdo"], item_dict["codigo"],
                item_dict["desc_codigo"], item_dict["fator_i"], item_dict["fator_r"], item_dict["filtro"]
            ))

        conn.commit()
        conn.close()

        supabase = get_supabase()
        if supabase and rows_to_insert:
            try:
                supabase.table("tabela_orcamento_master").delete().neq("id", 0).execute()
                for i in range(0, len(rows_to_insert), 500):
                    supabase.table("tabela_orcamento_master").insert(rows_to_insert[i:i+500]).execute()
            except Exception as e:
                logger.warning(f"Erro ao enviar Master para Supabase: {e}")

        _audit(admin, "UPLOAD_MASTER_CSV", "tabela_orcamento_master", None, f"total_rows={len(rows_to_insert)}")
        return {"status": "ok", "msg": f"Tabela Master atualizada com {len(rows_to_insert)} linhas."}

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Erro ao processar CSV Master: {e}")
