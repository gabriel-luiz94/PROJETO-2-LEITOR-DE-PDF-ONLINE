/* ═══════════════════════════════════════════════════════════════════════════
   Voz no chat de IA do Resumo (TASK-027).
   • Microfone (#btn-mic-chat): ditado em pt-BR pelo reconhecimento de voz do NAVEGADOR (Web Speech API). O texto vai para
     #chat-input e NUNCA é enviado sozinho — o usuário confere, edita e envia (o chat pode alterar as tabelas).
   • Leitura em voz alta (opcional, DESLIGADA por padrão): speechSynthesis; interruptor #btn-voz-resposta (localStorage),
     botão por resposta e botão de parar.
   Sem backend, sem custo e sem guardar áudio (no Chrome/Edge o áudio é processado pelo serviço de voz do navegador).
   Sem suporte (ex.: janela de desktop): botões desabilitados com explicação, sem erro. Texto sempre via textContent/value.
═══════════════════════════════════════════════════════════════════════════ */
(function () {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const TTS = window.speechSynthesis && window.SpeechSynthesisUtterance ? window.speechSynthesis : null;
    const PREF_LER = 'chat_voz_resposta';
    const SILENCIO_MS = 4000;       // para sozinho depois deste silêncio
    const LIMITE_FALA = 700;        // caracteres lidos de uma resposta (o resto fica só no texto)
    const el = id => document.getElementById(id);
    const entrada = el('chat-input'), btnMic = el('btn-mic-chat'), btnLer = el('btn-voz-resposta'), btnParar = el('btn-voz-parar');
    const estado = el('voz-estado'), mensagens = el('chat-messages');
    if (!entrada || !btnMic) return;

    function avisar(texto, tipo) {
        if (!estado) return;
        estado.textContent = texto || '';
        estado.className = 'voz-estado' + (tipo ? ' ' + tipo : '');
    }

    /* ── ditado ─────────────────────────────────────────────────────────── */
    let rec = null, ouvindo = false, base = '', finais = '', timer = null;
    const MENSAGENS_ERRO = {
        'not-allowed': 'Permissão de microfone negada. Libere o microfone nas configurações do navegador e tente de novo.',
        'service-not-allowed': 'O navegador bloqueou o serviço de voz. Libere o microfone e tente de novo.',
        'audio-capture': 'Nenhum microfone encontrado.',
        'no-speech': 'Não ouvi nada. Clique no microfone e fale de novo.',
        'network': 'Sem conexão com o serviço de voz do navegador (precisa de internet).',
        'language-not-supported': 'O navegador não oferece reconhecimento em português do Brasil.'
    };

    function montarTexto(provisorio) {
        const dito = (finais + (provisorio || '')).trim();
        return base + (base && dito && !/\s$/.test(base) ? ' ' : '') + dito;
    }

    function ajustarBotao() {
        btnMic.classList.toggle('gravando', ouvindo);
        btnMic.setAttribute('aria-pressed', ouvindo ? 'true' : 'false');
        btnMic.title = ouvindo ? 'Parar o ditado' : 'Ditar por voz (o texto vai para o campo; você confere antes de enviar)';
        btnMic.textContent = ouvindo ? '⏹' : '🎤';
    }

    function reiniciarSilencio() {
        clearTimeout(timer);
        timer = setTimeout(() => { if (rec && ouvindo) rec.stop(); }, SILENCIO_MS);
    }

    function iniciar() {
        if (ouvindo || entrada.disabled) return;
        pararLeitura();
        base = entrada.value; finais = '';
        rec = new SR();
        rec.lang = 'pt-BR';
        rec.interimResults = true;
        rec.continuous = true;
        rec.maxAlternatives = 1;
        rec.onresult = e => {
            let provisorio = '';
            for (let i = e.resultIndex; i < e.results.length; i++) {
                const r = e.results[i];
                if (r.isFinal) finais += r[0].transcript; else provisorio += r[0].transcript;
            }
            entrada.value = montarTexto(provisorio);
            reiniciarSilencio();
        };
        rec.onerror = e => {
            if (e.error === 'aborted') return;
            avisar(MENSAGENS_ERRO[e.error] || `Não foi possível usar o microfone (${e.error}).`, 'erro');
        };
        rec.onend = () => {
            clearTimeout(timer);
            const tinhaErro = estado && estado.classList.contains('erro');
            ouvindo = false; rec = null;
            ajustarBotao();
            entrada.value = montarTexto('');
            entrada.focus();
            entrada.setSelectionRange(entrada.value.length, entrada.value.length);
            if (!tinhaErro) avisar(finais.trim() ? 'Confira o texto, corrija se precisar e envie.' : '', '');
        };
        try {
            rec.start();
            ouvindo = true;
            ajustarBotao();
            avisar('Ouvindo… fale agora. Clique de novo (ou Esc) para parar.', 'ouvindo');
            reiniciarSilencio();
        } catch (e) {
            rec = null; ouvindo = false; ajustarBotao();
            avisar('Não foi possível iniciar o microfone.', 'erro');
        }
    }

    if (!SR) {
        btnMic.disabled = true;
        btnMic.title = 'Ditado por voz indisponível neste ambiente — abra o programa no navegador (Chrome ou Edge).';
        avisar('Ditado por voz indisponível neste ambiente (use o Chrome ou o Edge).', 'mudo');
    } else {
        ajustarBotao();
        btnMic.addEventListener('click', () => { if (ouvindo) rec.stop(); else iniciar(); });
        document.addEventListener('keydown', e => { if (e.key === 'Escape' && ouvindo) rec.stop(); });
    }

    /* ── leitura em voz alta (opcional) ─────────────────────────────────── */
    function lerLigado() { try { return localStorage.getItem(PREF_LER) === 'sim'; } catch (e) { return false; } }
    function ajustarLer() {
        if (!btnLer) return;
        const on = lerLigado();
        btnLer.textContent = on ? '🔊' : '🔇';
        btnLer.setAttribute('aria-pressed', on ? 'true' : 'false');
        btnLer.title = on ? 'Leitura das respostas em voz alta: LIGADA (clique para desligar)' : 'Leitura das respostas em voz alta: desligada (clique para ligar)';
    }

    /** Texto falável de uma resposta: sem blocos de código, linhas de tabela, marcações e links; limitado. */
    function textoParaFalar(texto) {
        const limpo = (texto || '')
            .replace(/```[\s\S]*?```/g, ' ')
            .split('\n').map(l => l.trim()).filter(l => l && !l.startsWith('|')).map(l => (/[.!?:;]$/.test(l) ? l : l + '.')).join(' ')
            .replace(/https?:\/\/\S+/g, ' ')
            .replace(/[*_#`>]+/g, '')
            .replace(/\s+/g, ' ').trim();
        return limpo.length > LIMITE_FALA ? limpo.slice(0, LIMITE_FALA).replace(/\s+\S*$/, '') + '…' : limpo;
    }

    /** Texto da bolha com as quebras de linha (resumo.js troca \n por <br>), sem botões nossos. */
    function textoDaBolha(bolha) {
        const c = bolha.cloneNode(true);
        c.querySelectorAll('.voz-ler').forEach(b => b.remove());
        c.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
        return c.textContent;
    }

    let falando = false;
    function ajustarParar() { if (btnParar) btnParar.classList.toggle('hidden', !falando); }
    function pararLeitura() {
        if (TTS && (falando || TTS.speaking)) TTS.cancel();
        falando = false; ajustarParar();
        mensagens.querySelectorAll('.voz-ler[data-lendo="1"]').forEach(b => { b.dataset.lendo = '0'; b.textContent = '🔊 Ouvir'; });
    }
    function falar(texto, botao) {
        if (!TTS) return;
        const fala = textoParaFalar(texto);
        if (!fala) return;
        pararLeitura();
        const u = new SpeechSynthesisUtterance(fala);
        u.lang = 'pt-BR';
        const v = (TTS.getVoices ? TTS.getVoices() : []).find(x => /^pt-BR/i.test(x.lang)) || (TTS.getVoices ? TTS.getVoices() : []).find(x => /^pt/i.test(x.lang));
        if (v) u.voice = v;
        u.onend = u.onerror = () => { falando = false; ajustarParar(); if (botao) { botao.dataset.lendo = '0'; botao.textContent = '🔊 Ouvir'; } };
        falando = true; ajustarParar();
        if (botao) { botao.dataset.lendo = '1'; botao.textContent = '⏹ Parar'; }
        TTS.speak(u);
    }

    if (!TTS) {
        [btnLer, btnParar].forEach(b => { if (b) b.classList.add('hidden'); });
    } else {
        ajustarLer(); ajustarParar();
        if (btnLer) btnLer.addEventListener('click', () => {
            const novo = !lerLigado();
            try { localStorage.setItem(PREF_LER, novo ? 'sim' : 'nao'); } catch (e) { /* vale só nesta sessão */ }
            if (!novo) pararLeitura();
            ajustarLer();
        });
        if (btnParar) btnParar.addEventListener('click', pararLeitura);

        // Cada resposta da IA ganha "🔊 Ouvir" quando termina; com a leitura ligada, a última é lida sozinha.
        // O fim da resposta = o campo do chat volta a ficar habilitado (resumo.js o desabilita enquanto a IA gera).
        new MutationObserver(() => {
            if (entrada.disabled) return;
            const ultima = Array.from(mensagens.querySelectorAll('.chat-message.ai')).pop();
            if (!ultima || ultima.dataset.voz === '1' || !ultima.textContent.trim() || ultima.querySelector('.loading-dots')) return;
            ultima.dataset.voz = '1';
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'voz-ler'; b.textContent = '🔊 Ouvir'; b.dataset.lendo = '0';
            const texto = textoDaBolha(ultima);
            b.addEventListener('click', () => { if (b.dataset.lendo === '1') pararLeitura(); else falar(texto, b); });
            ultima.appendChild(b);
            if (lerLigado()) falar(texto, b);
        }).observe(entrada, { attributes: true, attributeFilter: ['disabled'] });
    }
    window.vozChat = { textoParaFalar, textoDaBolha };
})();
