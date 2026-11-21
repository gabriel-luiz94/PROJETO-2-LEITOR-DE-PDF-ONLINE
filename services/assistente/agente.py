"""
services/assistente/agente.py (TASK-191) — O laço do assistente: o modelo pede ferramentas, nós executamos e devolvemos, até ele responder em texto.

Mensagens NEUTRAS (independem do provedor):
    {"role": "user", "texto": str}
    {"role": "assistant", "texto": str, "chamadas": [{"id", "nome", "args"}], "bruto": <resposta original do SDK, opcional>}
    {"role": "tool", "resultados": [{"id", "nome", "resultado": dict}]}
Cada Adaptador converte isso para o formato do seu SDK (Gemini / OpenAI / Claude) e lê a resposta de volta num `Turno`. O laço (`rodar`) é um gerador de
eventos {"tipo": ...} que a rota manda ao navegador em NDJSON: ferramenta, texto, plano, acao_ui, erro, fim.
"""
import asyncio
import json
import time
from dataclasses import dataclass, field
from typing import Any, AsyncIterator, Optional

from services.assistente import ferramentas

MAX_PASSOS = 8                      # idas e vindas com ferramentas por pergunta
GATILHOS_FALLBACK = ("429", "quota", "exhausted", "not found", "404", "unavailable", "invalid_argument", "multiturn", "overloaded", "503")


@dataclass
class Chamada:
    id: str
    nome: str
    args: dict


@dataclass
class Turno:
    texto: str = ""
    chamadas: list = field(default_factory=list)
    bruto: Any = None


def _resumo_args(args: dict) -> str:
    """Uma linha curta do que a ferramenta recebeu, para o navegador mostrar «buscando TR125…»."""
    if not isinstance(args, dict):
        return ""
    for k in ("texto", "ativo", "nome", "obra_id", "tabela"):
        if args.get(k):
            return str(args[k])[:60]
    if isinstance(args.get("alteracoes"), list):
        return f"{len(args['alteracoes'])} alteração(ões)"
    return ""


# Tempo máximo de UMA chamada ao provedor: sem isto, um provedor que trava deixa o chat em «Gerando» para sempre.
TIMEOUT_CHAMADA = 45.0
TIMEOUT_COM_RESERVA = 20.0          # com reserva configurado o principal tem menos tempo antes de passar a vez
PRAZO_TOTAL = 120.0                 # teto da pergunta inteira (várias idas e vindas com ferramentas)


class ErroTempo(RuntimeError):
    """O provedor não respondeu no tempo: passa ao próximo modelo SEM repetir (repetir custaria outro tempo inteiro)."""


# Falhas de momento (provedor sobrecarregado, rede): vale tentar de novo o MESMO modelo antes de trocar; cota esgotada/modelo inexistente trocam direto.
TRANSITORIOS = ("503", "502", "504", "unavailable", "overloaded", "high demand", "timeout", "timed out", "connection")
ESPERAS = (1.5, 4.0)                # segundos antes da 2ª e da 3ª tentativa no mesmo modelo


def eh_transitorio(e) -> bool:
    return any(k in str(e).lower() for k in TRANSITORIOS)


class ErroProvedores(RuntimeError):
    """Principal e reserva falharam: `mensagem_pronta` já é o texto para o usuário."""
    def __init__(self, mensagem_pronta: str):
        super().__init__(mensagem_pronta)
        self.mensagem_pronta = mensagem_pronta


def mensagem_amigavel(e) -> str:
    """O erro do provedor em português quando é de momento; senão o texto original (cortado)."""
    if isinstance(e, ErroProvedores):
        return e.mensagem_pronta
    if isinstance(e, ErroTempo):
        return "O provedor de IA demorou demais para responder. Tente de novo, escolha outro modelo ou configure um provedor reserva em «Configuração de IA»."
    if eh_transitorio(e):
        return ("O provedor de IA está sobrecarregado agora (erro 503/indisponível). Tentei de novo e troquei de modelo, sem sucesso. "
                "Aguarde um instante e repita, ou mude o provedor/modelo em «Configuração de IA».")
    return str(e)[:400]


async def _com_tempo(chamar, modelo, limite):
    try:
        return await asyncio.wait_for(chamar(modelo), limite)
    except (asyncio.TimeoutError, TimeoutError):
        raise ErroTempo(f"sem resposta em {limite:.0f} s ({modelo})")


async def com_fallback(modelos: list, chamar, esperas=None, timeout=None):
    """Para cada modelo: até 3 tentativas se a falha for de momento (503…); cota/indisponibilidade/tempo esgotado passam ao próximo modelo; os demais erros sobem.
    `esperas` (segundos entre tentativas) e `timeout` (por chamada) trocam o padrão: com provedor reserva o principal desiste mais cedo."""
    esperas = ESPERAS if esperas is None else esperas
    limite = TIMEOUT_CHAMADA if timeout is None else timeout
    ultimo = None
    for m in modelos:
        for tentativa in range(len(esperas) + 1):
            try:
                return await _com_tempo(chamar, m, limite)
            except ErroTempo as e:
                ultimo = e
                break
            except Exception as e:
                if eh_transitorio(e) and tentativa < len(esperas):
                    await asyncio.sleep(esperas[tentativa])
                    continue
                if any(k in str(e).lower() for k in GATILHOS_FALLBACK):
                    ultimo = e
                    break
                raise
    if isinstance(ultimo, ErroTempo):
        raise ultimo
    raise RuntimeError(str(ultimo) if ultimo else "Nenhum modelo disponível.")


# ── adaptadores ───────────────────────────────────────────────────────────────────────────────────────
class AdaptadorBase:
    modelos: list = []
    esperas = None                  # None = ESPERAS padrão
    timeout = None                  # None = TIMEOUT_CHAMADA

    async def turno(self, system: str, mensagens: list, specs: list) -> Turno:   # pragma: no cover - interface
        raise NotImplementedError

    def pegar_avisos(self) -> list:
        """Avisos para o usuário gerados desde a última leitura (ex.: «usando o provedor reserva»); lê e limpa."""
        return []


# ── provedor reserva ──────────────────────────────────────────────────────────────────────────────────
PAUSA_S = 90                        # depois de uma falha, o principal fica de lado este tempo (as próximas perguntas vão direto ao reserva)
ESPERAS_COM_RESERVA = (1.0,)        # com reserva configurado o principal desiste cedo: 2 tentativas por modelo
_pausa = {}                         # chave do principal → instante (monotônico) em que volta a ser tentado


def _agora() -> float:
    return time.monotonic()


def limpar_pausas():
    _pausa.clear()


class AdaptadorComReserva(AdaptadorBase):
    """Tenta o principal; se falhar (qualquer erro), segue com o reserva NA MESMA pergunta e avisa. Falha do principal pausa-o por PAUSA_S segundos
    (não repetir a espera de cada mensagem). Se os dois falharem, o erro traz os dois motivos."""

    def __init__(self, principal: AdaptadorBase, reserva: AdaptadorBase, chave: str, rotulo_reserva: str = "reserva"):
        self.principal, self.reserva, self.chave, self.rotulo_reserva = principal, reserva, chave, rotulo_reserva
        principal.esperas = ESPERAS_COM_RESERVA
        principal.timeout = TIMEOUT_COM_RESERVA
        self.no_reserva, self.erro_principal, self._avisos = False, None, []

    def _ir_para_reserva(self, motivo: str):
        self.no_reserva = True
        self._avisos.append(f"Usando o provedor reserva ({self.rotulo_reserva}): {motivo}")

    async def turno(self, system, mensagens, specs) -> Turno:
        if not self.no_reserva and _pausa.get(self.chave, 0) > _agora():
            self._ir_para_reserva("o principal falhou há pouco e está em pausa")
        if not self.no_reserva:
            try:
                t = await self.principal.turno(system, mensagens, specs)
                _pausa.pop(self.chave, None)
                return t
            except Exception as e:
                self.erro_principal = e
                _pausa[self.chave] = _agora() + PAUSA_S
                self._ir_para_reserva(mensagem_amigavel(e)[:120])
        try:
            return await self.reserva.turno(system, mensagens, specs)
        except Exception as e2:
            if self.erro_principal is None:
                raise
            raise ErroProvedores(f"Nenhum provedor respondeu. Principal: {mensagem_amigavel(self.erro_principal)[:160]} Reserva: {mensagem_amigavel(e2)[:160]}")

    def pegar_avisos(self) -> list:
        a, self._avisos = self._avisos, []
        return a


class AdaptadorGemini(AdaptadorBase):
    def __init__(self, client, modelos):
        self.client, self.modelos = client, modelos

    @staticmethod
    def converter(mensagens: list) -> list:
        from google.genai import types
        saida = []
        for m in mensagens:
            if m["role"] == "user":
                saida.append(types.Content(role="user", parts=[types.Part(text=m["texto"])]))
            elif m["role"] == "assistant":
                if m.get("bruto") is not None:                      # preserva as "assinaturas de pensamento" que o Gemini exige de volta
                    saida.append(m["bruto"])
                else:
                    parts = ([types.Part(text=m["texto"])] if m.get("texto") else []) + [types.Part.from_function_call(name=c["nome"], args=c["args"]) for c in m.get("chamadas", [])]
                    saida.append(types.Content(role="model", parts=parts or [types.Part(text="")]))
            else:
                saida.append(types.Content(role="user", parts=[types.Part.from_function_response(name=r["nome"], response={"resultado": r["resultado"]}) for r in m["resultados"]]))
        return saida

    async def turno(self, system, mensagens, specs) -> Turno:
        from google.genai import types
        decl = [types.FunctionDeclaration(name=s["name"], description=s["description"], parameters_json_schema=s["parameters"]) for s in specs]
        def cfg_de(modelo, pensar_pouco):
            extra = {}
            if pensar_pouco and str(modelo).startswith("gemini-3"):          # Gemini 3 pensa muito por padrão: «low» corta segundos de cada volta
                extra["thinking_config"] = types.ThinkingConfig(thinking_level="low")
            return types.GenerateContentConfig(system_instruction=system, tools=[types.Tool(function_declarations=decl)] if decl else None,
                                               automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True), **extra)
        contents = self.converter(mensagens)

        async def chamar(modelo):
            try:
                return await self.client.aio.models.generate_content(model=modelo, contents=contents, config=cfg_de(modelo, True))
            except Exception as e:
                if "thinking" not in str(e).lower():
                    raise
                return await self.client.aio.models.generate_content(model=modelo, contents=contents, config=cfg_de(modelo, False))      # esta versão não aceita o nível de pensamento
        resp = await com_fallback(self.modelos, chamar, self.esperas, self.timeout)
        cand = resp.candidates[0] if getattr(resp, "candidates", None) else None
        conteudo = getattr(cand, "content", None)
        texto, chamadas = [], []
        for i, p in enumerate(getattr(conteudo, "parts", None) or []):
            if getattr(p, "function_call", None):
                fc = p.function_call
                chamadas.append(Chamada(getattr(fc, "id", None) or f"g{i}", fc.name, dict(fc.args or {})))
            elif getattr(p, "text", None) and not getattr(p, "thought", False):
                texto.append(p.text)
        return Turno("".join(texto), chamadas, conteudo)


class AdaptadorOpenAI(AdaptadorBase):
    def __init__(self, client, modelos):
        self.client, self.modelos = client, modelos

    @staticmethod
    def converter(system: str, mensagens: list) -> list:
        saida = [{"role": "system", "content": system}]
        for m in mensagens:
            if m["role"] == "user":
                saida.append({"role": "user", "content": m["texto"]})
            elif m["role"] == "assistant":
                msg = {"role": "assistant", "content": m.get("texto") or None}
                if m.get("chamadas"):
                    msg["tool_calls"] = [{"id": c["id"], "type": "function", "function": {"name": c["nome"], "arguments": json.dumps(c["args"], ensure_ascii=False)}} for c in m["chamadas"]]
                saida.append(msg)
            else:
                saida += [{"role": "tool", "tool_call_id": r["id"], "content": json.dumps(r["resultado"], ensure_ascii=False)} for r in m["resultados"]]
        return saida

    async def turno(self, system, mensagens, specs) -> Turno:
        tools = [{"type": "function", "function": {"name": s["name"], "description": s["description"], "parameters": s["parameters"]}} for s in specs]
        msgs = self.converter(system, mensagens)

        async def chamar(modelo):
            extra = {"tools": tools} if tools else {}
            return await self.client.chat.completions.create(model=modelo, messages=msgs, **extra)
        resp = await com_fallback(self.modelos, chamar, self.esperas, self.timeout)
        msg = resp.choices[0].message
        chamadas = []
        for tc in getattr(msg, "tool_calls", None) or []:
            try:
                args = json.loads(tc.function.arguments or "{}")
            except ValueError:
                args = {}
            chamadas.append(Chamada(tc.id, tc.function.name, args if isinstance(args, dict) else {}))
        return Turno(msg.content or "", chamadas)


class AdaptadorClaude(AdaptadorBase):
    def __init__(self, client, modelos):
        self.client, self.modelos = client, modelos

    @staticmethod
    def converter(mensagens: list) -> list:
        saida = []
        for m in mensagens:
            if m["role"] == "user":
                saida.append({"role": "user", "content": m["texto"]})
            elif m["role"] == "assistant":
                blocos = ([{"type": "text", "text": m["texto"]}] if m.get("texto") else []) + [{"type": "tool_use", "id": c["id"], "name": c["nome"], "input": c["args"]} for c in m.get("chamadas", [])]
                saida.append({"role": "assistant", "content": blocos or [{"type": "text", "text": "…"}]})
            else:
                saida.append({"role": "user", "content": [{"type": "tool_result", "tool_use_id": r["id"], "content": json.dumps(r["resultado"], ensure_ascii=False)} for r in m["resultados"]]})
        return saida

    async def turno(self, system, mensagens, specs) -> Turno:
        tools = [{"name": s["name"], "description": s["description"], "input_schema": s["parameters"]} for s in specs]
        msgs = self.converter(mensagens)

        async def chamar(modelo):
            extra = {"tools": tools} if tools else {}
            return await self.client.messages.create(model=modelo, max_tokens=4096, system=system, messages=msgs, **extra)
        resp = await com_fallback(self.modelos, chamar, self.esperas, self.timeout)
        texto, chamadas = [], []
        for b in resp.content:
            if getattr(b, "type", "") == "text":
                texto.append(b.text)
            elif getattr(b, "type", "") == "tool_use":
                chamadas.append(Chamada(b.id, b.name, dict(b.input or {})))
        return Turno("".join(texto), chamadas)


def criar_adaptador(provider: str, api_key: str, modelos: list, base_url: str = "") -> AdaptadorBase:
    if provider == "gemini":
        from google import genai
        return AdaptadorGemini(genai.Client(api_key=api_key), modelos)
    if provider == "openai":
        from openai import AsyncOpenAI
        return AdaptadorOpenAI(AsyncOpenAI(api_key=api_key, base_url=base_url or "https://api.openai.com/v1"), modelos)
    from anthropic import AsyncAnthropic
    return AdaptadorClaude(AsyncAnthropic(api_key=api_key), modelos)


# ── o laço ────────────────────────────────────────────────────────────────────────────────────────────
def historico_neutro(history: list) -> list:
    """O `chatHistory` do navegador ([{role: user|model, parts: [{text}]}]) → mensagens neutras; só texto, alternando começando por «user»."""
    saida = []
    for h in history or []:
        parts = h.get("parts") if isinstance(h, dict) else None
        texto = (parts[0].get("text", "") if parts and isinstance(parts[0], dict) else str(parts[0]) if parts else "") or ""
        papel = "assistant" if h.get("role") == "model" else "user"
        if not texto.strip():
            continue
        if saida and saida[-1]["role"] == papel:                    # dois seguidos do mesmo lado (ação sem resposta): junta
            saida[-1]["texto"] += "\n" + texto
            continue
        saida.append({"role": papel, "texto": texto} if papel == "user" else {"role": papel, "texto": texto, "chamadas": []})
    while saida and saida[0]["role"] != "user":
        saida.pop(0)
    saida = saida[-20:]
    while saida and saida[0]["role"] != "user":
        saida.pop(0)
    return saida


# Ferramentas que ENCERRAM a pergunta: depois delas não vale outra volta ao modelo só para «resumir» (o cartão do plano já mostra tudo).
TERMINAIS = {"propor_alteracoes", "propor_obra", "ajustar_tela"}


def _terminal_ok(chamadas, resultados) -> bool:
    """Só chamadas terminais, todas aceitas, sem rejeição nem aviso (aviso/rejeição = o modelo deve corrigir numa nova volta)."""
    if not chamadas or any(c.nome not in TERMINAIS for c in chamadas):
        return False
    for r in resultados:
        res = r["resultado"]
        if not isinstance(res, dict) or res.get("erro") or res.get("rejeitadas") or res.get("avisos") or res.get("aviso"):
            return False
    return True


def _resumo_local(ctx, chamadas) -> str:
    """O texto que o modelo escreveria depois de um plano limpo, sem gastar outra chamada."""
    nomes = {c.nome for c in chamadas}
    if "propor_obra" in nomes:
        return "Confira o cartão e continue para aplicar a obra."
    if "propor_alteracoes" in nomes and ctx.plano:
        n = len(ctx.plano)
        return (f"Proponho {n} alteração(ões)" + (f": {ctx.resumo_plano}" if ctx.resumo_plano else "") + ". Confira o cartão e clique em Aplicar.")
    return "Pronto."


async def rodar(adaptador: AdaptadorBase, system: str, historico: list, prompt: str, ctx: ferramentas.Contexto) -> AsyncIterator[dict]:
    """Gera os eventos da conversa. O último é sempre {"tipo": "fim"}; falhas viram {"tipo": "erro"} (nunca exceção para a rota)."""
    mensagens = list(historico) + [{"role": "user", "texto": prompt}]
    specs = ferramentas.especificacoes()
    texto_final = ""
    inicio = _agora()
    try:
        for _ in range(MAX_PASSOS):
            if _agora() - inicio > PRAZO_TOTAL:
                texto_final = "Esta pergunta passou do tempo máximo. Diga o que falta e continuo daqui."
                break
            # O prazo global também limita a chamada em andamento, não só o início do passo.
            restante = PRAZO_TOTAL - (_agora() - inicio)
            if restante <= 0:
                raise ErroTempo("prazo total da pergunta esgotado")
            try:
                turno = await asyncio.wait_for(
                    adaptador.turno(system, mensagens, specs), timeout=restante
                )
            except asyncio.TimeoutError as exc:
                raise ErroTempo("prazo total da pergunta esgotado") from exc
            if not turno.chamadas and not turno.texto.strip():
                raise RuntimeError(
                    "O modelo retornou uma resposta vazia. Verifique o modelo configurado "
                    "e tente novamente."
                )
            for aviso in adaptador.pegar_avisos():
                yield {"tipo": "aviso", "texto": aviso}
            if not turno.chamadas:
                texto_final = turno.texto
                break
            if turno.texto.strip():
                yield {"tipo": "texto", "texto": turno.texto}                # o que o modelo disse antes de chamar ferramentas
            mensagens.append({"role": "assistant", "texto": turno.texto, "bruto": turno.bruto,
                              "chamadas": [{"id": c.id, "nome": c.nome, "args": c.args} for c in turno.chamadas]})
            resultados = []
            for c in turno.chamadas:
                yield {"tipo": "ferramenta", "nome": c.nome, "detalhe": _resumo_args(c.args)}
                # em outra thread: cálculo/base grandes não travam o servidor (nem as outras telas)
                restante = PRAZO_TOTAL - (_agora() - inicio)
                if restante <= 0:
                    raise ErroTempo("prazo total esgotado durante a execução da ferramenta")
                try:
                    resultado = await asyncio.wait_for(
                        asyncio.to_thread(ferramentas.executar, ctx, c.nome, c.args),
                        timeout=restante,
                    )
                except asyncio.TimeoutError as exc:
                    raise ErroTempo("prazo total esgotado durante a execução da ferramenta") from exc
                resultados.append({"id": c.id, "nome": c.nome, "resultado": resultado})
            mensagens.append({"role": "tool", "resultados": resultados})
            if _terminal_ok(turno.chamadas, resultados):
                texto_final = "" if turno.texto.strip() else _resumo_local(ctx, turno.chamadas)
                break
        else:
            texto_final = "Parei por excesso de passos nesta pergunta. Diga o que falta e continuo daqui."
    except Exception as e:
        for aviso in adaptador.pegar_avisos():
            yield {"tipo": "aviso", "texto": aviso}
        yield {"tipo": "erro", "mensagem": mensagem_amigavel(e)}
        yield {"tipo": "fim"}
        return
    if texto_final.strip():
        yield {"tipo": "texto", "texto": texto_final}
    for a in ctx.acoes_ui:
        yield {"tipo": "acao_ui", **a}
    if ctx.plano:
        yield {"tipo": "plano", "resumo": ctx.resumo_plano, "projeto": ctx.projeto_codigo, "itens": ctx.plano}
    yield {"tipo": "fim"}
