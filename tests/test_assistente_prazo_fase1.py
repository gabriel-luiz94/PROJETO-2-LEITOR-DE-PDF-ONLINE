"""Testes de regressão do laço do agente; não requerem credenciais Gemini."""
import asyncio
import time
from threading import Event
import unittest
from unittest.mock import patch

from services.assistente import agente, ferramentas
from services.prazo import verificar_prazo


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


    async def test_prazo_interrompe_ferramenta_cooperativa(self):
        iniciou = Event()
        parou = Event()
        concluiu = Event()

        def ferramenta_lenta(*_):
            iniciou.set()
            try:
                for _ in range(100):
                    verificar_prazo()
                    time.sleep(0.005)
                concluiu.set()
            finally:
                parou.set()

        adaptador = AdaptadorFalso(agente.Turno(chamadas=[agente.Chamada("1", "totalizar", {})]))
        with patch.object(agente, "PRAZO_TOTAL", 0.15), patch.object(ferramentas, "executar", ferramenta_lenta):
            saida = await eventos(adaptador)
        self.assertEqual([e["tipo"] for e in saida], ["ferramenta", "erro", "fim"])
        self.assertIn("tempo máximo", saida[1]["mensagem"])
        self.assertTrue(iniciou.is_set() and await asyncio.to_thread(parou.wait, 0.5))
        self.assertFalse(concluiu.is_set())


if __name__ == "__main__":
    unittest.main()
