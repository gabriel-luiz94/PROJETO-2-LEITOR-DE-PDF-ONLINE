"""services/janelas_auxiliares.py — visualizações nativas do mesmo processo desktop."""
import re
import threading

from config import logger
from services.arquivos_desktop import ApiArquivos
from services.projeto_desktop import ApiProjeto

TITULOS = {"resultado": "Resultado", "base": "Base de Orçamento", "mapa": "Mapa"}          # TASK-241: o Mapa (Resumo em modo janela)


class ApiVisualizacao(ApiArquivos):
    """Ponte da janela secundária, com os mesmos diálogos nativos de arquivo."""

    def __init__(self):
        super().__init__()
        self._fechando = False

    def _fechar(self):
        if self._fechando:
            return
        self._fechando = True
        try:
            self._janela.destroy()
        except Exception:
            self._fechando = False
            raise

    def vincular(self, janela):
        super().vincular(janela)

        def antes_de_fechar():
            if self._fechando:
                return True

            def solicitar():
                try:
                    janela.evaluate_js("window.fecharVisualizacaoNativa?.()")
                except Exception as erro:
                    logger.debug("Não foi possível solicitar fechamento: %s", erro)

            threading.Thread(target=solicitar, daemon=True).start()
            return False

        janela.events.closing += antes_de_fechar


class ApiProjetoComJanelas(ApiProjeto):
    def __init__(self, origem, criar_janela):
        super().__init__()
        self._origem = origem
        self._criar_janela = criar_janela
        self._auxiliares = {}
        self._trava = threading.Lock()
        self._encerrando = False

    def vincular(self, janela):
        super().vincular(janela)
        janela.events.closed += self._encerrar_auxiliares

    def _encerrar_auxiliares(self):
        with self._trava:
            self._encerrando = True
            canais = list(self._auxiliares)
        for canal in canais:
            try:
                self.janela_auxiliar_fechar(canal)
            except Exception as erro:
                logger.warning("Falha ao fechar visualização auxiliar: %s", erro)

    def janela_auxiliar_abrir(self, tipo, canal):
        if not isinstance(tipo, str) or tipo not in TITULOS or not isinstance(canal, str) or not re.fullmatch(r"[0-9a-f]{32}", canal):
            return {"ok": False, "erro": "Visualização inválida."}
        with self._trava:
            if self._encerrando or canal in self._auxiliares or any(entrada[0] == tipo for entrada in self._auxiliares.values()):
                return {"ok": False, "erro": "Essa visualização já está aberta."}
            self._auxiliares[canal] = (tipo, None)
        api = ApiVisualizacao()
        janela = None
        try:
            janela = self._criar_janela(
                title=f"OrçaCore · {TITULOS[tipo]}",
                url=f"{self._origem}/static/janela_auxiliar.html#{tipo}={canal}",
                js_api=api, width=1100, height=760, min_size=(600, 420),
                # O sistema operacional controla todo o arraste, inclusive a soltura fora do WebView.
                frameless=False, easy_drag=False, background_color="#0d1117",
                text_select=True, zoomable=True,
            )
            api.vincular(janela)

            def fechada():
                with self._trava:
                    self._auxiliares.pop(canal, None)
                    encerrando = self._encerrando
                if not encerrando:
                    try:
                        self._janela.evaluate_js(f"window.AbrirNaApp?.nativaFechada('{canal}')")
                    except Exception as erro:
                        logger.debug("Não foi possível avisar a janela principal: %s", erro)

            janela.events.closed += fechada
            with self._trava:
                self._auxiliares[canal] = (tipo, api)
                encerrando = self._encerrando
            if encerrando:
                self.janela_auxiliar_fechar(canal)
                return {"ok": False, "erro": "O OrçaCore está fechando."}
            return {"ok": True}
        except Exception as erro:
            with self._trava:
                self._auxiliares.pop(canal, None)
            if janela is not None:
                try:
                    api._fechar()
                except Exception as erro_fechar:
                    logger.warning("Falha ao descartar visualização auxiliar: %s", erro_fechar)
            logger.warning("Falha ao abrir visualização auxiliar: %s", erro)
            return {"ok": False, "erro": "Não foi possível abrir a janela nativa."}

    def janela_auxiliar_fechar(self, canal):
        if not isinstance(canal, str):
            return
        with self._trava:
            entrada = self._auxiliares.get(canal)
        if entrada and entrada[1] is not None:
            entrada[1]._fechar()
