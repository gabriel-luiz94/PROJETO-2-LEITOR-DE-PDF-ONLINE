"""Testes de regressão do laço do agente; não requerem credenciais Gemini."""
import asyncio
import unittest
from unittest.mock import patch

from services.assistente import agente, ferramentas


class AdaptadorFalso(agente.AdaptadorBase):
    def __init__(self, turno=None, atraso=0):
        self.resposta = turno or agente.Turno(texto="ok")
        self.atraso = atraso

    async def turno(self, system, mensagens, specs):
        await asyncio.sleep(self.atraso)
        return self.resposta


async def eventos(adaptador):
    ctx = ferramentas.Contexto(projeto_codigo="TESTE")
    return [e async for e in agente.rodar(adaptador, "", [], "olá", ctx)]


class TestePrazoAgente(unittest.IsolatedAsyncioTestCase):
    async def test_resposta_normal_termina(self):
        saida = await eventos(AdaptadorFalso())
        self.assertEqual(saida[-1]["tipo"], "fim")
        self.assertTrue(any(e.get("texto") == "ok" for e in saida))

    async def test_resposta_vazia_exibe_erro_e_fim(self):
        saida = await eventos(AdaptadorFalso(agente.Turno(texto=" ")))
        self.assertEqual([e["tipo"] for e in saida[-2:]], ["erro", "fim"])
        self.assertIn("vazia", saida[-2]["mensagem"])

    async def test_provedor_lento_respeita_prazo_global(self):
        with patch.object(agente, "PRAZO_TOTAL", 0.02):
            saida = await eventos(AdaptadorFalso(atraso=0.2))
        self.assertEqual([e["tipo"] for e in saida[-2:]], ["erro", "fim"])


if __name__ == "__main__":
    unittest.main()
