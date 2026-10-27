/*
 * vinculacao_ui.js (TASK-084) — modal de vinculação cabo <-> estrutura em abas: Vínculo (painel visual), Tabela de Cabos, Tabela de Estruturas,
 * Script e Mapa. A lógica pura está em vinculacao_regras.js (testada em Node e em Python). Quem altera as tabelas do Resumo é o host (resumo.js).
 * Cores: preto = sem vínculo · roxo = parcial · verde = completo · vermelho = inválido · tracejado = cabo fictício.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document, VR = raiz.VinculacaoRegras;
    let host = null, st = null, overlay = null;

    const CSS = `
    .vz-ov{position:fixed;inset:0;background:rgba(1,4,9,.78);z-index:9000;display:flex;align-items:center;justify-content:center}
    .vz{width:min(1500px,96vw);height:min(900px,94vh);background:#0d1117;border:1px solid #30363d;border-radius:12px;display:flex;flex-direction:column;overflow:hidden;color:#c9d1d9;font:13px/1.35 -apple-system,Segoe UI,Roboto,sans-serif}
    .vz *{box-sizing:border-box}
    .vz-top{display:flex;align-items:center;gap:10px;padding:12px 18px;background:#161b22;border-bottom:1px solid #30363d}.vz-top h3{margin:0;font-size:15px;color:#fff}
    .vz-chip{background:#21262d;border:1px solid #30363d;border-radius:20px;padding:2px 10px;font-size:11px;color:#8b949e;cursor:default}.vz-chip.on{color:#fff;border-color:#58a6ff}.vz-chip.btn{cursor:pointer}
    .vz-x{margin-left:auto;color:#8b949e;font-size:18px;cursor:pointer;background:none;border:none}
    .vz-tabs{display:flex;gap:4px;padding:10px 18px 0;background:#161b22;border-bottom:1px solid #30363d}
    .vz-tab{padding:8px 16px;border:1px solid transparent;border-bottom:none;border-radius:8px 8px 0 0;color:#8b949e;font-weight:600;cursor:pointer;background:none}
    .vz-tab.on{background:#0d1117;border-color:#30363d;color:#fff;position:relative;top:1px}.vz-tab b{font-size:10px;background:#30363d;border-radius:10px;padding:1px 6px;margin-left:6px;color:#c9d1d9}
    .vz-corpo{flex:1;min-height:0;display:flex;flex-direction:column;padding:14px 18px}.vz-scroll{flex:1;min-height:0;overflow:auto;position:relative}
    .vz-foot{display:flex;gap:12px;align-items:center;padding:10px 18px;background:#161b22;border-top:1px solid #30363d;font-size:12px}
    .vz-btn{background:#21262d;border:1px solid #30363d;border-radius:6px;padding:5px 12px;color:#c9d1d9;font-size:12px;cursor:pointer}.vz-btn:hover{border-color:#8b949e}.vz-btn.p{background:#238636;border-color:#2ea043;color:#fff}.vz-btn.b{border-color:#58a6ff;color:#58a6ff}.vz-btn[disabled]{opacity:.45;cursor:default}
    .vz-leg{display:flex;gap:12px;align-items:center;color:#8b949e}.vz-dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:5px;vertical-align:-1px;border:2px solid #6e7681;background:#010409}
    .vz-dot.pu{border-color:#a371f7;background:#2a1d4d}.vz-dot.gr{border-color:#3fb950;background:#12361c}.vz-dot.rd{border-color:#f85149;background:#47161a}.vz-dot.fic{border-style:dashed;border-color:#a371f7}
    .vz-cols{display:grid;grid-template-columns:minmax(300px,430px) minmax(120px,1fr) minmax(360px,560px);position:relative;align-items:start}
    .vz-col h4{margin:0 0 8px;font-size:11px;letter-spacing:.08em;color:#8b949e;text-transform:uppercase;position:sticky;top:0;background:#0d1117;z-index:2;padding:2px 0}
    .vz-card{display:flex;align-items:center;gap:10px;height:62px;margin-bottom:10px;padding:0 12px;border-radius:8px;border:2px solid #6e7681;background:#010409;cursor:pointer;position:relative;overflow:hidden}
    .vz-card.pu{border-color:#a371f7;background:#1d1537}.vz-card.gr{border-color:#3fb950;background:#0f2a17}.vz-card.rd{border-color:#f85149;background:#3a1216}.vz-card.fic{border-style:dashed;border-color:#a371f7;background:#161027}
    .vz-card.sel,.vz-ebox.sel{box-shadow:0 0 0 3px rgba(88,166,255,.65)}.vz-card.glow,.vz-ebox.glow{box-shadow:0 0 0 3px rgba(63,185,80,.40)}.vz-dim{opacity:.35}
    .vz-card .t{font-weight:700;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:260px}.vz-card .s{font-size:11px;color:#8b949e;white-space:nowrap}.vz-card .r{margin-left:auto;text-align:right;font-size:11px;white-space:nowrap}
    .vz-op{font-size:11px;font-weight:700;border-radius:4px;padding:1px 6px;background:#1f6feb33;color:#58a6ff}.vz-op.m{background:#d2992233;color:#d29922}.vz-op.r{background:#f8514933;color:#f85149}
    .vz-erow{display:flex;align-items:center;gap:12px;height:62px;margin-bottom:10px}
    .vz-prow{display:flex;align-items:center;gap:10px;margin-bottom:10px;padding:8px 10px;border:1px solid #21262d;border-radius:10px;background:#0d1117}.vz-blocos{display:flex;flex-wrap:wrap;gap:8px 18px;align-items:flex-start;flex:1;min-width:0}.vz-bloco{display:flex;flex-direction:column;gap:4px;width:150px}.vz-bloco .vz-ebox{min-width:0;width:100%;box-sizing:border-box}.vz-bloco .vz-info{padding:0 2px}
    .vz-ebox{display:flex;align-items:center;justify-content:center;height:56px;min-width:92px;padding:0 16px;border-radius:9px;border:2px solid #6e7681;background:#010409;font-weight:800;font-size:15px;color:#fff;cursor:pointer;position:relative}
    .vz-ebox.pu{border-color:#a371f7;background:#1d1537}.vz-ebox.gr{border-color:#3fb950;background:#0f2a17}.vz-ebox.rd{border-color:#f85149;background:#3a1216}
    .vz-info{display:flex;flex-direction:column;gap:2px;min-width:0}.vz-info .a{font-size:12px;color:#c9d1d9;white-space:nowrap}.vz-info .b{font-size:11px;color:#8b949e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .vz-pbtn{margin-left:auto;border:1px dashed #484f58;color:#6e7681;border-radius:6px;padding:4px 9px;font-size:12px;font-weight:700;background:#0d1117;cursor:default}
    .vz-liga{position:absolute;left:0;top:0;pointer-events:none;overflow:visible}.vz-liga path.hit{pointer-events:stroke;cursor:pointer;stroke:transparent;stroke-width:12;fill:none}
    .vz-menu{position:fixed;background:#1c2230;border:1px solid #484f58;border-radius:8px;padding:6px;min-width:260px;max-width:380px;z-index:9100;box-shadow:0 10px 30px rgba(0,0,0,.6);font-size:12px}
    .vz-menu .it{padding:6px 10px;border-radius:5px;cursor:pointer;white-space:nowrap}.vz-menu .it:hover{background:#2d3548}.vz-menu .it.off{color:#6e7681;cursor:default}.vz-menu .hd{padding:4px 10px;color:#8b949e;font-size:10.5px;text-transform:uppercase;letter-spacing:.06em}
    .vz-menu .sep{height:1px;background:#30363d;margin:4px 0}.vz-menu .rec{color:#3fb950;font-size:10.5px;margin-left:6px}
    .vz-nota{background:#1c2230;border:1px solid #30363d;border-radius:8px;padding:10px 12px;color:#8b949e;margin:0 0 12px;font-size:12px}.vz-nota b{color:#fff}
    .vz-tb{width:100%;border-collapse:collapse;font-size:12px}.vz-tb th{background:#21262d;color:#8b949e;text-align:left;padding:6px 8px;border:1px solid #30363d;font-weight:600;position:sticky;top:0;z-index:1}
    .vz-tb td{padding:4px 6px;border:1px solid #30363d;vertical-align:middle}.vz-tb input[type=text],.vz-tb input[type=number],.vz-tb select{width:100%;background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:4px 6px;font:12px Consolas,Menlo,monospace}
    .vz-tb input[type=checkbox]{transform:scale(1.15)}.vz-tb .x{color:#f85149;cursor:pointer;background:none;border:none;font-size:14px}.vz-tb .mv{color:#8b949e;cursor:pointer;background:none;border:none}
    .vz-erro{color:#f85149;font-size:12px;margin:6px 0;white-space:pre-wrap}.vz-ok{color:#3fb950}.vz-av{color:#d29922}
    .vz-pre{background:#010409;border:1px solid #30363d;border-radius:8px;padding:10px;font:12px/1.5 Consolas,Menlo,monospace;color:#c9d1d9;width:100%;min-height:360px;resize:vertical;white-space:pre}
    .vz-grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;min-height:0}.vz-painel{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 14px}.vz-painel h5{margin:0 0 8px;font-size:12px;color:#fff}
    .vz-mapa{width:100%;height:100%;background:#010409;border:1px solid #30363d;border-radius:8px;cursor:grab}
    `;

    const clone = o => JSON.parse(JSON.stringify(o));
    const esc = t => String(t == null ? '' : t).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[m]));
    function el(tag, props, filhos) {
        const e = doc.createElement(tag);
        Object.keys(props || {}).forEach(k => {
            if (k === 'style') e.style.cssText = props[k];
            else if (k === 'text') e.textContent = props[k];
            else if (k === 'html') e.innerHTML = props[k];
            else if (k === 'class') e.className = props[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]);
            else e.setAttribute(k, props[k]);
        });
        (filhos || []).forEach(f => f != null && e.appendChild(typeof f === 'string' ? doc.createTextNode(f) : f));
        return e;
    }
    const toast = m => { if (raiz.showToast) raiz.showToast(m); };
    const cabos = () => host.cabos(), outros = () => host.outros();
    const analise = () => VR.analisar(cabos(), outros(), st.cfg, host.prefixoCabo, host.codigoEstrutura, host.codigosEstrutura);
    // TASK-094: cada estrutura da linha (9400D, N1, N4 → N1 e N4) é uma unidade; o id é a linha (+ posição/100 a partir da 2ª)
    const linhaDe = ie => VR.linhaDe(ie), slotDe = ie => VR.slotDe(ie), linhaOutros = ie => outros()[VR.linhaDe(ie)];
    const codigoDe = ie => { const l = linhaOutros(ie); return VR.codigosDaLinha(l, host.codigoEstrutura, host.codigosEstrutura)[VR.slotDe(ie)] || ''; };
    const ficticiosDe = ie => VR.ficticiosDe(linhaOutros(ie), VR.slotDe(ie));
    function gravarFicticios(ie, lista) {
        const l = linhaOutros(ie), k = slotDe(ie);
        if (!k) { if (lista.length) l.cabosFicticios = lista; else delete l.cabosFicticios; return; }
        l.cabosFicticiosSlots = l.cabosFicticiosSlots || {};
        if (lista.length) l.cabosFicticiosSlots[k] = lista; else delete l.cabosFicticiosSlots[k];
        if (!Object.keys(l.cabosFicticiosSlots).length) delete l.cabosFicticiosSlots;
    }
    const trim = t => String(t == null ? '' : t).trim();

    // ── abrir / fechar ──
    async function abrir(h, opcoes) {
        host = h;
        if (!doc.getElementById('vz-css')) doc.head.appendChild(el('style', { id: 'vz-css', text: CSS }));
        let cfg = raiz.__vinculacaoConfig;
        if (!cfg) {
            try { const r = await fetch('/api/vinculacao-config?projeto_codigo=' + encodeURIComponent(host.projeto())); cfg = r.ok ? (await r.json()).config : null; } catch (e) { cfg = null; }
        }
        st = { cfg: clone(VR.normalizarConfig(cfg)), aba: (opcoes && opcoes.aba) || 'vinculo', sel: null, filtro: 'todos', dirty: false, scriptIdx: 0, previa: null, textoScript: null, erroScript: '', teste: {} };
        if (!overlay) {
            overlay = el('div', { class: 'vz-ov', onkeydown: e => { if (e.key === 'Escape') { if (menuAberto) fecharMenu(); else fechar(); } } });
            overlay.setAttribute('tabindex', '-1');
            doc.body.appendChild(overlay);
        }
        overlay.style.display = 'flex';
        render();
        overlay.focus();
    }

    function fechar() {
        if (st && st.dirty && !raiz.confirm('Há alterações nas tabelas/scripts que não foram salvas. Fechar mesmo assim?')) return;
        fecharMenu();
        if (overlay) overlay.style.display = 'none';
        if (host && host.depois) host.depois();
    }

    function marcarSujo() { st.dirty = true; const b = overlay.querySelector('#vz-salvar'); if (b) b.disabled = false; }

    async function salvar() {
        const erros = [];
        try {
            const r = await fetch('/api/vinculacao-config', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ projeto_codigo: host.projeto(), config: st.cfg }) });
            if (r.ok) { raiz.__vinculacaoConfig = clone(st.cfg); st.dirty = false; toast('✓ Tabelas de vinculação salvas!'); render(); if (host.depois) host.depois(); return true; }
            const d = await r.json().catch(() => ({}));
            const e = d && d.detail && d.detail.erros;
            erros.push(Array.isArray(e) ? e.join('\n') : ((d && d.detail) || 'Erro ao salvar.'));
        } catch (e) { erros.push('Erro de conexão ao salvar.'); }
        raiz.alert(erros.join('\n'));
        return false;
    }

    // ── casca ──
    function render() {
        fecharMenu();
        const a = analise();
        const pend = Object.values(a.cabos).filter(c => c.estado !== 'gr').length + Object.values(a.estruturas).filter(e => e.estado === 'k' || e.estado === 'pu').length;
        overlay.innerHTML = '';
        const abas = [['vinculo', 'Vínculo', pend ? pend + ' pend.' : ''], ['cabos', 'Tabela de Cabos', st.cfg.cabos.length], ['estruturas', 'Tabela de Estruturas', st.cfg.estruturas.length],
            ['script', 'Script', st.cfg.scripts.length], ['mapa', 'Mapa', '']];
        const vz = el('div', { class: 'vz' }, [
            el('div', { class: 'vz-top' }, [el('h3', { text: '🔗 Vinculação cabo ⇄ estrutura' }), el('span', { class: 'vz-chip', text: 'Projeto ' + host.projeto() }),
                st.dirty ? el('span', { class: 'vz-chip', style: 'color:#d29922;border-color:#d29922', text: 'alterações não salvas' }) : null, el('button', { class: 'vz-x', text: '✕', onclick: fechar })]),
            el('div', { class: 'vz-tabs' }, abas.map(([id, nome, n]) => el('button', { class: 'vz-tab' + (st.aba === id ? ' on' : ''), onclick: () => { st.aba = id; st.sel = null; render(); } }, [nome, n !== '' ? el('b', { text: String(n) }) : null]))),
            el('div', { class: 'vz-corpo', id: 'vz-corpo' }),
            rodape()]);
        overlay.appendChild(vz);
        const corpo = vz.querySelector('#vz-corpo');
        ({ vinculo: abaVinculo, cabos: abaCabos, estruturas: abaEstruturas, script: abaScript, mapa: abaMapa })[st.aba](corpo, a);
    }

    function rodape() {
        return el('div', { class: 'vz-foot' }, [
            el('div', { class: 'vz-leg' }, [el('span', {}, [el('i', { class: 'vz-dot' }), 'sem vínculo']), el('span', {}, [el('i', { class: 'vz-dot pu' }), 'parcial']), el('span', {}, [el('i', { class: 'vz-dot gr' }), 'completo']),
                el('span', {}, [el('i', { class: 'vz-dot rd' }), 'inválido']), el('span', {}, [el('i', { class: 'vz-dot fic' }), 'fictício'])]),
            el('span', { style: 'margin-left:auto' }),
            el('button', { class: 'vz-btn b', text: '⚙ Vincular com o Script…', onclick: () => { st.aba = 'script'; render(); } }),
            el('button', { class: 'vz-btn', text: 'Conectores…', title: 'Regras de conectores (cabo + estrutura → ativo)', onclick: () => { if (raiz.abrirModalConectores) raiz.abrirModalConectores(); } }),
            el('button', { class: 'vz-btn', text: '↶ Desfazer', onclick: () => { host.desfazer(); render(); } }),
            el('button', { class: 'vz-btn b', id: 'vz-salvar', text: 'Salvar tabelas', disabled: st.dirty ? null : 'disabled', onclick: salvar }),
            el('button', { class: 'vz-btn p', text: 'Concluir', onclick: fechar })]);
    }

    // ── menu de contexto ──
    let menuAberto = null;
    function fecharMenu() { if (menuAberto) { menuAberto.remove(); menuAberto = null; } }
    function abrirMenu(e, itens) {
        e.preventDefault(); fecharMenu();
        const m = el('div', { class: 'vz-menu' }, itens.map(i => i === '-' ? el('div', { class: 'sep' }) : (i.hd ? el('div', { class: 'hd', text: i.hd }) :
            el('div', { class: 'it' + (i.off ? ' off' : ''), onclick: i.off ? null : () => { fecharMenu(); i.fn(); }, title: i.title || '' }, [i.t, i.rec ? el('span', { class: 'rec', text: i.rec }) : null]))));
        m.style.left = Math.min(e.clientX, raiz.innerWidth - 400) + 'px'; m.style.top = Math.min(e.clientY, raiz.innerHeight - 300) + 'px';
        doc.body.appendChild(m); menuAberto = m;
        setTimeout(() => doc.addEventListener('click', fecharMenu, { once: true }), 0);
    }

    // ── ABA: VÍNCULO ──
    function infoCabo(ic) { const c = cabos()[ic]; return VR.infoCabo(host.prefixoCabo(c.ativo) || '', c.operacao, st.cfg, VR.extraCabo(c)); }
    function infoEst(ie) { return VR.infoEstrutura(codigoDe(ie), st.cfg); }
    function ocupantes(ie, ignorandoCabo) {
        const lista = [];
        cabos().forEach((c, ic) => { if (VR.ehCaboValido(c) && ic !== ignorandoCabo && (c.vinculoEstruturas || []).includes(ie)) lista.push({ info: infoCabo(ic), ficticio: false }); });
        const e = infoEst(ie);
        ficticiosDe(ie).forEach(cod => lista.push({ info: VR.infoCabo(cod, VR.opFicticio(e), st.cfg), ficticio: true }));
        return lista;
    }
    function podeLigar(ic, ie) {
        const c = infoCabo(ic), e = infoEst(ie), oc = ocupantes(ie, ic);
        return VR.podeComOcupantes(c, e, oc);
    }

    function ligar(ic, ie) {
        const c = cabos()[ic], ja = (c.vinculoEstruturas || []).includes(ie);
        host.antes();
        if (ja) { c.vinculoEstruturas = c.vinculoEstruturas.filter(x => x !== ie); toast('Vínculo removido (Ctrl+Z desfaz).'); render(); return; }
        const info = infoCabo(ic);
        if (!info.vincula) { toast(`O cabo ${info.codigo} não vincula (acompanha o cabo principal).`); return; }
        if ((c.vinculoEstruturas || []).length + VR.estruturasFicticiasDe(c).length >= info.pontas) { toast(`O cabo ${info.codigo} já tem as ${info.pontas} ponta(s); remova uma antes.`); return; }
        const [ok, motivo] = podeLigar(ic, ie);
        if (!ok) { toast('✖ ' + motivo); return; }
        c.vinculoEstruturas = (c.vinculoEstruturas || []).concat([ie]);
        c.vinculoOrigem = 'manual';
        st.sel = null;
        render();
    }

    function clicar(tipo, i) {
        if (st.sel && st.sel.t !== tipo) { const ic = tipo === 'c' ? i : st.sel.i, ie = tipo === 'e' ? i : st.sel.i; ligar(ic, ie); return; }
        st.sel = (st.sel && st.sel.t === tipo && st.sel.i === i) ? null : { t: tipo, i };
        render();
    }

    const faixaCabos = a => (a.maximo != null && a.maximo !== a.precisa ? `${a.precisa}–${a.maximo}` : String(a.precisa));

    function candidatosFicticios(ie) {
        const e = infoEst(ie), oc = ocupantes(ie), codigos = [], vistos = new Set();
        const add = c => { c = trim(c).toUpperCase(); if (c && !vistos.has(c)) { vistos.add(c); codigos.push(c); } };
        cabos().forEach(c => { if (VR.ehCaboValido(c)) add(host.prefixoCabo(c.ativo)); });
        st.cfg.cabos.forEach(r => { if (typeof r.padrao === 'string' && !/[\\^$()\[\]|+*?{}]/.test(r.padrao)) r.padrao.split(',').forEach(add); });
        return codigos.map(cod => { const ic = VR.infoCabo(cod, VR.opFicticio(e), st.cfg), [ok, motivo] = VR.podeComOcupantes(ic, e, oc, true); return { cod, ok, motivo, combina: ok && oc.length > 0 && oc[0].info.codigo === cod }; })
            .sort((a, b) => (b.combina - a.combina) || (b.ok - a.ok));
    }

    function addFicticio(ie, cod) {
        host.antes();
        gravarFicticios(ie, ficticiosDe(ie).concat([cod.toUpperCase()]));
        render();
    }

    function menuEstrutura(e, ie) {
        const info = infoEst(ie), a = analise().estruturas[ie], cands = candidatosFicticios(ie);
        const itens = [{ hd: `Estrutura ${info.codigo} · ${a ? a.tem : 0}${a && a.precisa ? '/' + faixaCabos(a) : ''} cabo(s)` }];
        const aceitos = cands.filter(c => c.ok);
        itens.push({ hd: 'Associar cabo fictício' });
        if (!aceitos.length) itens.push({ t: 'Nenhum cabo aceito por esta estrutura agora', off: true });
        aceitos.slice(0, 8).forEach(c => itens.push({ t: '➕ ' + c.cod, rec: c.combina ? '✔ combina com o outro lado' : '', fn: () => addFicticio(ie, c.cod) }));
        itens.push({ t: '✎ Outro código…', fn: () => { const cod = raiz.prompt('Código do cabo fictício (ex.: C16):'); if (cod) { const c = VR.infoCabo(cod, VR.opFicticio(info), st.cfg), [ok, motivo] = VR.podeComOcupantes(c, info, ocupantes(ie), true); if (ok) addFicticio(ie, cod); else toast('✖ ' + motivo); } } });
        const recusados = cands.filter(c => !c.ok && c.motivo);
        if (recusados.length) itens.push({ t: `(${recusados.length} código(s) não aceito(s): ${recusados.slice(0, 3).map(c => c.cod).join(', ')}…)`, off: true, title: recusados.slice(0, 5).map(c => c.cod + ': ' + c.motivo).join('\n') });
        itens.push('-');
        ficticiosDe(ie).forEach((cod, k) => itens.push({ t: `✖ Remover fictício ${cod}`, fn: () => { host.antes(); const l = ficticiosDe(ie).slice(); l.splice(k, 1); gravarFicticios(ie, l); render(); } }));
        itens.push({ t: 'Desvincular todos os cabos desta estrutura', fn: () => { host.antes(); cabos().forEach(c => { if (c && c.vinculoEstruturas) c.vinculoEstruturas = c.vinculoEstruturas.filter(x => x !== ie); }); render(); } });
        abrirMenu(e, itens);
    }

    // TASK-095: estrutura fictícia do CABO (só o código): ocupa uma ponta do cabo, sem linha em Outros, e segue para a Totalizadora como CABO_ESTRUTURA
    function candidatosEstruturasFicticias(ic) {
        const info = infoCabo(ic), codigos = [], vistos = new Set();
        const add = x => { x = trim(x).toUpperCase(); if (x && !vistos.has(x)) { vistos.add(x); codigos.push(x); } };
        VR.estruturasDisponiveis(outros(), host.codigoEstrutura, host.codigosEstrutura).forEach(ie => add(codigoDe(ie)));
        st.cfg.estruturas.forEach(r => { if (typeof r.padrao === 'string' && !/[\\^$()\[\]|+*?{}]/.test(r.padrao)) r.padrao.split(',').forEach(add); });
        return codigos.map(cod => { const [ok, motivo] = VR.podeVincular(info, VR.infoEstrutura(cod, st.cfg)); return { cod, ok, motivo }; }).sort((a, b) => (b.ok - a.ok));
    }
    function menuCabo(e, ic) {
        const c = cabos()[ic], info = infoCabo(ic), fict = VR.estruturasFicticiasDe(c), cheio = (c.vinculoEstruturas || []).length + fict.length >= info.pontas;
        const itens = [{ hd: 'Cabo ' + (c.ativo || '') }, { t: 'Desvincular este cabo', fn: () => { host.antes(); c.vinculoEstruturas = []; render(); } }, '-', { hd: 'Estrutura fictícia' }];
        const add = cod => { host.antes(); c.estruturasFicticias = fict.concat([cod.toUpperCase()]); render(); };
        if (cheio) itens.push({ t: `O cabo já tem as ${info.pontas} ponta(s)`, off: true });
        else {
            const cands = candidatosEstruturasFicticias(ic), aceitas = cands.filter(x => x.ok);
            if (!aceitas.length) itens.push({ t: 'Nenhuma estrutura aceita por este cabo', off: true });
            aceitas.slice(0, 8).forEach(x => itens.push({ t: '➕ ' + x.cod, fn: () => add(x.cod) }));
            itens.push({ t: '✎ Outro código…', fn: () => { const cod = raiz.prompt('Código da estrutura fictícia (ex.: N1):'); if (!cod) return; const [ok, motivo] = VR.podeVincular(info, VR.infoEstrutura(cod, st.cfg)); if (ok) add(cod); else toast('✖ ' + motivo); } });
            const rec = cands.filter(x => !x.ok && x.motivo);
            if (rec.length) itens.push({ t: `(${rec.length} código(s) não aceito(s): ${rec.slice(0, 3).map(x => x.cod).join(', ')}…)`, off: true, title: rec.slice(0, 5).map(x => x.cod + ': ' + x.motivo).join('\n') });
        }
        fict.forEach((cod, k) => itens.push({ t: `✖ Remover estrutura fictícia ${cod}`, fn: () => { host.antes(); const l = fict.slice(); l.splice(k, 1); if (l.length) c.estruturasFicticias = l; else delete c.estruturasFicticias; render(); } }));
        abrirMenu(e, itens);
    }

    const CORES = { k: '', pu: ' pu', gr: ' gr', rd: ' rd' };
    const TXT_C = { k: 'sem vínculo', pu: '◐ parcial', gr: '● completo', rd: '✖ inválido' };

    function abaVinculo(corpo, a) {
        const tem = cabos().some(VR.ehCaboValido), ests = VR.estruturasDisponiveis(outros(), host.codigoEstrutura, host.codigosEstrutura);
        const filtros = [['todos', 'Todos'], ['pendentes', 'Só pendentes'], ['conflitos', 'Só conflitos']];
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-bottom:10px;align-items:center;flex-wrap:wrap' },
            filtros.map(([id, nome]) => el('span', { class: 'vz-chip btn' + (st.filtro === id ? ' on' : ''), text: nome, onclick: () => { st.filtro = id; render(); } }))
                .concat([el('span', { style: 'margin-left:auto;color:#8b949e', text: 'Clique no cabo e na estrutura (ou o contrário) · botão direito na estrutura: cabo fictício · no cabo: estrutura fictícia · clique na linha: desfazer' })])));
        if (!tem || !ests.length) {
            corpo.appendChild(el('div', { style: 'color:#8b949e;padding:20px', text: !tem ? 'Nenhum cabo na tabela.' : 'Nenhuma estrutura/poste na tabela Outros ainda.' }));
            return;
        }
        const sc = el('div', { class: 'vz-scroll' }), cols = el('div', { class: 'vz-cols' });
        const colC = el('div', { class: 'vz-col' }, [el('h4', { text: 'Cabos' })]), colE = el('div', { class: 'vz-col' }, [el('h4', { text: 'Estruturas (a linha liga na estrutura; o poste é só contexto)' })]);
        const selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selE = st.sel && st.sel.t === 'e' ? st.sel.i : null;
        const ficticiosLista = [];
        cabos().forEach((c, ic) => {
            if (!VR.ehCaboValido(c)) return;
            const r = a.cabos[ic];
            if (st.filtro === 'pendentes' && r.estado === 'gr') return;
            if (st.filtro === 'conflitos' && r.estado !== 'rd') return;
            let cls = 'vz-card' + CORES[r.estado] + (selC === ic ? ' sel' : ''), dim = false, tip = '';
            if (selE !== null) { const [ok, motivo] = podeLigar(ic, selE), ja = (c.vinculoEstruturas || []).includes(selE); if (ok || ja) cls += ' glow'; else { dim = true; tip = motivo; } }
            const op = trim(c.operacao).toUpperCase(), opc = op.indexOf('M') >= 0 ? ' m' : (op.indexOf('R') >= 0 ? ' r' : '');
            const card = el('div', { class: cls + (dim ? ' vz-dim' : ''), 'data-c': ic, title: tip || (c.ativo || ''), onclick: () => clicar('c', ic), oncontextmenu: e => menuCabo(e, ic) }, [
                el('span', { class: 'vz-op' + opc, text: op || 'I' }),
                el('div', {}, [el('div', { class: 't', text: c.ativo || '' }), el('div', { class: 's', text: `${r.codigo}${r.valor_cabo ? ' · tipocabo=' + r.valor_cabo : ''} · ${r.tem}/${r.precisa} ponta(s)${r.ficticias ? ' · ' + r.ficticias + ' estrutura fictícia' : ''}` })]),
                el('div', { class: 'r', style: { gr: 'color:#3fb950', pu: 'color:#a371f7', rd: 'color:#f85149', k: 'color:#8b949e' }[r.estado], text: r.vincula ? TXT_C[r.estado] : 'não vincula' })]);
            colC.appendChild(card);
        });
        // TASK-094: uma linha por POSTE; as estruturas dela (N1, N4…) ficam lado a lado, em blocos separados. Poste sem estrutura não aparece; estrutura sem poste não tem o card do poste.
        const bloco = (ie, o) => {
            const r = a.estruturas[ie];
            if (st.filtro === 'pendentes' && (r.estado === 'gr')) return null;
            if (st.filtro === 'conflitos' && r.estado !== 'rd') return null;
            let cls = 'vz-ebox' + CORES[r.estado] + (selE === ie ? ' sel' : ''), dim = false, tip = '';
            if (selC !== null) { const [ok, motivo] = podeLigar(selC, ie), ja = (cabos()[selC].vinculoEstruturas || []).includes(ie); if (ok || ja) cls += ' glow'; else { dim = true; tip = motivo; } }
            const nome = r.codigo || '?', info = infoEst(ie);
            const ebox = el('div', { class: cls, 'data-e': ie, onclick: () => clicar('e', ie), oncontextmenu: e => menuEstrutura(e, ie), title: tip || (o.ativo || '') }, [nome]);
            const sub = r.erros.length ? '✖ ' + r.erros[0] : (info.regra === null ? 'sem regra na Tabela de Estruturas' : (info.lados === 'IGUAIS' ? `exige ${info.qtd || ''} cabos iguais` : (info.amarracao ? 'amarração' : `${info.qtd || ''} cabo(s)`)));
            return el('div', { class: 'vz-bloco' + (dim ? ' vz-dim' : '') }, [ebox,
                el('div', { class: 'vz-info' }, [el('div', { class: 'a', style: r.estado === 'rd' ? 'color:#f85149' : '', text: `${r.tem}${r.precisa ? '/' + faixaCabos(r) : ''}${r.valor_cabo ? ' · tipocabo=' + r.valor_cabo : ''}${r.ficticios ? ' · ' + r.ficticios + ' fictício' : ''}` }), el('div', { class: 'b', text: sub, title: sub })])]);
        };
        const grupos = [], porLinha = {};
        ests.forEach(ie => { const l = linhaDe(ie); if (!(l in porLinha)) { porLinha[l] = { l, ids: [] }; grupos.push(porLinha[l]); } porLinha[l].ids.push(ie); });
        grupos.forEach(g => {
            const o = outros()[g.l], soPoste = host.linhaSoPoste(o.ativo), postes = host.postesDaLinha(o.ativo);
            if (soPoste && !g.ids.some(ie => cabos().some(c => (c.vinculoEstruturas || []).includes(ie)))) return;          // poste sem estrutura e sem cabo ligado: não aparece
            const blocos = g.ids.map(ie => bloco(ie, o)).filter(Boolean);
            if (!blocos.length) return;
            colE.appendChild(el('div', { class: 'vz-prow' }, [el('div', { class: 'vz-blocos' }, blocos),
                postes.length && !soPoste ? el('span', { class: 'vz-pbtn', title: 'poste da linha (não vincula): ' + (o.ativo || ''), text: postes.join(' ') }) : null]));
            g.ids.forEach(ie => ficticiosDe(ie).forEach((cod, k) => ficticiosLista.push({ ie, cod, k })));
        });
        ficticiosLista.forEach(f => colC.appendChild(el('div', { class: 'vz-card fic', 'data-f': f.ie + '|' + f.k }, [el('span', { class: 'vz-op', text: 'I' }),
            el('div', {}, [el('div', { class: 't' }, [f.cod + ' ', el('span', { style: 'font-size:10px;font-weight:800;color:#a371f7;background:#a371f733;border-radius:4px;padding:1px 6px', text: 'FICTÍCIO' })]),
                el('div', { class: 's', text: 'sem comprimento · só preenche a lacuna' })]),
            el('div', { class: 'r', style: 'color:#a371f7;cursor:pointer', text: '✎ remover', onclick: ev => { ev.stopPropagation(); host.antes(); const l = ficticiosDe(f.ie).slice(); l.splice(f.k, 1); gravarFicticios(f.ie, l); render(); } })])));
        const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('class', 'vz-liga');
        cols.appendChild(colC); cols.appendChild(el('div')); cols.appendChild(colE); cols.appendChild(svg);
        sc.appendChild(cols); corpo.appendChild(sc);
        const desenhar = () => desenharLinhas(cols, svg, a);
        raiz.requestAnimationFrame(desenhar);
        sc.addEventListener('scroll', () => {}, { passive: true });
    }

    function desenharLinhas(cols, svg, a) {
        const ar = cols.getBoundingClientRect();
        svg.setAttribute('width', cols.scrollWidth); svg.setAttribute('height', cols.scrollHeight);
        let h = '';
        const pt = (elm, lado) => { const r = elm.getBoundingClientRect(); if (lado === 't') return [(r.left + r.right) / 2 - ar.left, r.top - ar.top]; return [(lado === 'r' ? r.right : r.left) - ar.left, r.top + r.height / 2 - ar.top]; };      // 't' (TASK-094): a linha chega por cima nas estruturas que não são a 1ª do poste, sem cruzar a vizinha
        const curva = (p, q) => { const mx = (p[0] + q[0]) / 2; return `M${p[0]},${p[1]} C${mx},${p[1]} ${mx},${q[1]} ${q[0]},${q[1]}`; };
        a.links.forEach(l => {
            const ce = cols.querySelector(`[data-c="${l.cabo}"]`), ee = cols.querySelector(`[data-e="${l.estrutura}"]`);
            if (!ce || !ee) return;
            const p = pt(ce, 'r'), q = pt(ee, slotDe(l.estrutura) > 0 ? 't' : 'l'), est = a.cabos[l.cabo].estado;
            const cor = !l.ok || est === 'rd' ? '#f85149' : (est === 'gr' ? '#3fb950' : '#a371f7');
            const d = curva(p, q);
            h += `<path d="${d}" fill="none" stroke="${cor}" stroke-width="2.5" ${l.ok ? '' : 'stroke-dasharray="6 4"'}/><circle cx="${p[0]}" cy="${p[1]}" r="4" fill="${cor}"/><circle cx="${q[0]}" cy="${q[1]}" r="4" fill="${cor}"/>`;
            h += `<path class="hit" d="${d}" data-lc="${l.cabo}" data-le="${l.estrutura}"><title>${esc(l.motivo || 'clique para desfazer o vínculo')}</title></path>`;
        });
        cols.querySelectorAll('[data-f]').forEach(fe => {
            const [ie] = fe.getAttribute('data-f').split('|'), ee = cols.querySelector(`[data-e="${ie}"]`);
            if (!ee) return;
            const p = pt(fe, 'r'), q = pt(ee, slotDe(Number(ie)) > 0 ? 't' : 'l');
            h += `<path d="${curva(p, q)}" fill="none" stroke="#a371f7" stroke-width="2.5" stroke-dasharray="3 5"/><circle cx="${p[0]}" cy="${p[1]}" r="4" fill="#a371f7"/><circle cx="${q[0]}" cy="${q[1]}" r="4" fill="#a371f7"/>`;
        });
        svg.innerHTML = h;
        svg.querySelectorAll('path.hit').forEach(pth => pth.addEventListener('click', () => ligar(Number(pth.getAttribute('data-lc')), Number(pth.getAttribute('data-le')))));
    }

    // ── tabelas editáveis (Cabos / Estruturas) ──
    function celula(r, campo, tipo, extra) {
        const v = r[campo];
        if (tipo === 'bool') return el('input', { type: 'checkbox', ...(v ? { checked: 'checked' } : {}), onchange: e => { r[campo] = e.target.checked; marcarSujo(); } });
        if (tipo === 'select') return el('select', { onchange: e => { r[campo] = e.target.value; marcarSujo(); } }, extra.map(o => el('option', { value: o, text: o, ...(String(v == null ? extra[0] : v) === o ? { selected: 'selected' } : {}) })));
        if (tipo === 'number') return el('input', { type: 'number', min: '1', value: v == null ? '' : v, onchange: e => { r[campo] = e.target.value === '' ? undefined : parseInt(e.target.value, 10); marcarSujo(); } });
        if (tipo === 'ops') return el('input', { type: 'text', value: Array.isArray(v) ? v.join(', ') : '', placeholder: 'I, M', onchange: e => { const l = e.target.value.split(',').map(trim).filter(Boolean); if (l.length) r[campo] = l; else delete r[campo]; marcarSujo(); } });
        if (tipo === 'pontas') return el('div', { style: 'display:flex;gap:4px' }, ['I', 'M'].map(k => el('input', { type: 'number', min: '0', title: 'pontas do cabo ' + k, placeholder: k, value: r.pontas && r.pontas[k] != null ? r.pontas[k] : '', onchange: e => { r.pontas = r.pontas || {}; if (e.target.value === '') delete r.pontas[k]; else r.pontas[k] = parseInt(e.target.value, 10); marcarSujo(); } })));
        return el('input', { type: 'text', value: Array.isArray(v) ? v.join(', ') : (v == null ? '' : v), placeholder: extra || '', onchange: e => { r[campo] = e.target.value; marcarSujo(); } });
    }

    function tabelaEditavel(corpo, lista, colunas, novaLinha, testeFn, nota, nome) {
        corpo.appendChild(el('div', { class: 'vz-nota', html: nota }));
        const sc = el('div', { class: 'vz-scroll' });
        const tb = el('table', { class: 'vz-tb' });
        tb.appendChild(el('thead', {}, [el('tr', {}, [el('th', { text: '#' })].concat(colunas.map(c => el('th', { text: c.rot, title: c.dica || '' }))).concat([el('th', { text: '' })]))]));
        const tbody = el('tbody');
        lista.forEach((r, i) => {
            tbody.appendChild(el('tr', {}, [el('td', { text: String(i + 1), style: 'color:#8b949e;text-align:center' })].concat(colunas.map(c => el('td', { style: c.w ? 'width:' + c.w : '' }, [celula(r, c.campo, c.tipo, c.extra)])))
                .concat([el('td', { style: 'white-space:nowrap' }, [el('button', { class: 'mv', text: '▲', title: 'a primeira linha que casa vale', onclick: () => { if (i > 0) { lista.splice(i - 1, 0, lista.splice(i, 1)[0]); marcarSujo(); render(); } } }),
                    el('button', { class: 'mv', text: '▼', onclick: () => { if (i < lista.length - 1) { lista.splice(i + 1, 0, lista.splice(i, 1)[0]); marcarSujo(); render(); } } }), el('button', { class: 'x', text: '✖', title: 'Excluir', onclick: () => { lista.splice(i, 1); marcarSujo(); render(); } })])])));
        });
        tb.appendChild(tbody); sc.appendChild(tb); corpo.appendChild(sc);
        const saidaTeste = el('span', { style: 'color:#8b949e;flex:1' });
        const campo = el('input', { type: 'text', style: 'width:280px;background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace', placeholder: nome === 'cabos' ? 'Testar um cabo: 3xA35P(7,94)-13,8kV' : 'Testar uma estrutura: 9200D, N1', value: st.teste[nome] || '',
            oninput: e => { st.teste[nome] = e.target.value; saidaTeste.innerHTML = e.target.value ? testeFn(e.target.value) : ''; } });
        if (st.teste[nome]) saidaTeste.innerHTML = testeFn(st.teste[nome]);
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:10px;align-items:center' }, [el('button', { class: 'vz-btn', text: '＋ Linha', onclick: () => { lista.push(novaLinha()); marcarSujo(); render(); } }), campo, saidaTeste,
            el('button', { class: 'vz-btn', text: 'Importar JSON', onclick: () => importarJson(nome) }), el('button', { class: 'vz-btn', text: 'Exportar JSON', onclick: () => exportarJson(nome) })]));
    }

    function baixar(nome, obj) {
        const a = el('a', { href: URL.createObjectURL(new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' })), download: nome }); doc.body.appendChild(a); a.click(); a.remove();
    }
    function exportarJson(nome) { baixar(`vinculacao-${nome}-${host.projeto()}.json`, { formato: 'vinculacao-' + nome, versao: 1, projeto: host.projeto(), [nome]: st.cfg[nome] }); }
    function importarJson(nome) {
        const inp = el('input', { type: 'file', accept: '.json,application/json', onchange: async e => {
            try {
                const d = JSON.parse(await e.target.files[0].text());
                const lista = Array.isArray(d) ? d : d[nome];
                if (!Array.isArray(lista)) throw new Error('O arquivo não traz a lista "' + nome + '".');
                const teste = clone(st.cfg); teste[nome] = lista;
                const erros = (VR.validarConfig ? VR.validarConfig(teste) : []);
                if (erros.length) throw new Error(erros.join('\n'));
                if (!raiz.confirm(`Substituir as ${st.cfg[nome].length} linha(s) atuais por ${lista.length} do arquivo?`)) return;
                st.cfg[nome] = lista; marcarSujo(); render();
            } catch (err) { raiz.alert('Arquivo inválido: ' + err.message); }
        } });
        inp.click();
    }

    function abaCabos(corpo) {
        tabelaEditavel(corpo, st.cfg.cabos, [
            { campo: 'padrao', rot: 'Padrão do cabo', tipo: 'text', extra: '^C\\d+$ ou C16, C25', dica: 'regex ou lista; o código vem sem o 3x e sem a tensão' }, { campo: 'grupo', rot: 'Grupo', tipo: 'text' },
            { campo: 'valor_cabo', rot: 'Valor de cabo', tipo: 'text', w: '90px', dica: 'valor do parâmetro `cabo` (A, P, C…); vazio = sem parâmetro' }, { campo: 'operacoes', rot: 'Operações', tipo: 'ops', w: '100px' },
            { campo: 'pontas', rot: 'Pontas I · M', tipo: 'pontas', w: '120px' }, { campo: 'estruturas', rot: 'Estruturas que aceita', tipo: 'text', extra: '^[NUB][1-4] (vazio = qualquer)' },
            { campo: 'fases', rot: 'Fases', tipo: 'text', w: '55px', extra: '3 ou 1,2', dica: 'só vale para cabos com essa quantidade de fases (do 3x do texto; sem Nx = 1); vazio = qualquer' },
            { campo: 'nivel', rot: 'Nível', tipo: 'select', extra: ['', 'MT', 'BT'], w: '65px', dica: 'MT = tensão acima de 1 kV no texto (ex.: 13,8kV); BT = sem tensão; vazio = qualquer' },
            { campo: 'padrao_texto', rot: 'Padrão no texto', tipo: 'text', extra: '13,8kV (vazio = qualquer)', dica: 'regex contra o texto ORIGINAL do cabo (com 3x e kV)' },
            { campo: 'vincula', rot: 'Vincula', tipo: 'bool', w: '60px', dica: 'desmarcado = acompanha o cabo principal (ex.: neutro CAZ)' }, { campo: 'ativa', rot: 'Ativa', tipo: 'bool', w: '50px' }, { campo: 'obs', rot: 'Observação', tipo: 'text' }],
            () => ({ padrao: '', grupo: '', valor_cabo: '', operacoes: ['I', 'M'], pontas: { I: 2, M: 1 }, estruturas: '', vincula: true, ativa: true, obs: '' }),
            txt => {
                const r = raiz.AnotacaoCabos && raiz.AnotacaoCabos.ehNotacao(txt) ? raiz.AnotacaoCabos.interpretar(txt) : null;
                const cod = r && r.itens.length ? (r.itens.find(i => !i.neutro) || r.itens[0]).ativo : (host.prefixoCabo(txt) || txt.split(' ')[0]);
                const i = VR.infoCabo(cod, 'I', st.cfg);
                return `código <b style="color:#79c0ff">${esc(i.codigo)}</b> → ${i.regra === null ? 'nenhuma linha casa (aceita tudo)' : 'linha #' + (i.regra + 1) + (i.valor_cabo ? ' · tipocabo=<b>' + esc(i.valor_cabo) + '</b>' : '') + ' · pontas I=' + VR.infoCabo(cod, 'I', st.cfg).pontas + ' M=' + VR.infoCabo(cod, 'M', st.cfg).pontas}`;
            },
            '<b>O que esta tabela diz:</b> para cada tipo de cabo, a que estruturas ele pode se ligar, quantas pontas precisa e que <b>valor do parâmetro <code>cabo</code></b> ele carrega. O código é lido sem o <code>3x</code> e sem a tensão. A <b>primeira linha ativa que casa</b> vale; padrão = regex ou lista (<code>C16, C25</code>).', 'cabos');
    }

    function abaEstruturas(corpo) {
        tabelaEditavel(corpo, st.cfg.estruturas, [
            { campo: 'padrao', rot: 'Padrão da estrutura', tipo: 'text', extra: '^N1$ ou N1, N2' }, { campo: 'qtd_cabos', rot: 'Cabos', tipo: 'number', w: '60px', dica: 'quantidade fixa (mín = máx)' }, { campo: 'qtd_min', rot: 'Mín', tipo: 'number', w: '50px', dica: 'faixa: mínimo para ficar completa' }, { campo: 'qtd_max', rot: 'Máx', tipo: 'number', w: '50px', dica: 'faixa: máximo aceito' }, { campo: 'max_reais', rot: 'Reais', tipo: 'number', w: '55px', dica: 'quantos podem ser cabos REAIS; o excedente só entra como cabo fictício (ex.: SI3 = 1)' },
            { campo: 'lados', rot: 'Os lados…', tipo: 'select', extra: ['LIVRE', 'IGUAIS'], w: '90px' }, { campo: 'amarracao', rot: 'Amarração', tipo: 'bool', w: '80px', dica: 'aceita cabo M' },
            { campo: 'aceita_i', rot: 'Aceita I', tipo: 'bool', w: '70px', dica: 'desmarcado = só cabo M' }, { campo: 'cabos', rot: 'Cabos que aceita', tipo: 'text', extra: '(vazio = qualquer)' },
            { campo: 'parametro', rot: 'Parâmetro', tipo: 'text', w: '80px', dica: 'parâmetro ÚNICO que recebe o VALOR DE CABO da Tabela de Cabos (use tipocabo); vazio = nenhum. Para a descrição completa do cabo use a coluna Parâmetros (vários) com cabo.' }, { campo: 'parametros', rot: 'Parâmetros (vários)', tipo: 'text', w: '150px', extra: 'cabo, cabo, tipocabo, tipocabo', dica: 'um valor por cabo vinculado. cabo = descrição completa do cabo, sem o comprimento (3xS04-7,97kV → 3XS04-7.97KV); tipocabo = a coluna Valor de cabo da Tabela de Cabos; o mesmo nome pode se repetir (N4 com dois cabos). Se preenchido, substitui o Parâmetro único.' }, { campo: 'se_diferem', rot: 'Se diferem', tipo: 'select', extra: ['MAIOR', 'BLOQUEAR', 'AVISAR'], w: '95px' },
            { campo: 'ativa', rot: 'Ativa', tipo: 'bool', w: '50px' }, { campo: 'obs', rot: 'Observação', tipo: 'text' }],
            () => ({ padrao: '', qtd_cabos: 1, lados: 'LIVRE', amarracao: false, aceita_i: true, cabos: '', parametro: 'cabo', se_diferem: 'MAIOR', ativa: true, obs: '' }),
            txt => {
                const cod = host.codigoEstrutura(txt) || txt, i = VR.infoEstrutura(cod, st.cfg);
                return `estrutura <b style="color:#79c0ff">${esc(i.codigo)}</b> → ${i.regra === null ? 'nenhuma linha casa' : 'linha #' + (i.regra + 1) + ' · ' + (i.qtd || '?') + ' cabo(s) · lados ' + i.lados + (i.amarracao ? ' · amarração' : '') + (i.aceita_i ? '' : ' · só cabo M')}`;
            },
            '<b>O que esta tabela diz:</b> para cada tipo de estrutura, quantos cabos ela recebe, se os lados precisam ser <b>iguais</b>, se é de <b>amarração</b> (a única que aceita cabo M), quais cabos aceita e <b>qual parâmetro recebe do cabo</b> (<code>cabo</code>). A estrutura é lida sem o poste (<code>9200D, N1</code> → <code>N1</code>). O vínculo só vale se o cabo aceita a estrutura <u>e</u> a estrutura aceita o cabo.', 'estruturas');
    }

    // ── ABA: SCRIPT ──
    const SCRIPT_EXEMPLO = { versao: 1, conflitos: 'avisar', passos: [
        { nome: 'Cabo M na amarração mais próxima', modo: 'proximidade', raio_m: 8, cabos: { operacao: ['M'] }, estruturas: { amarracao: true }, pontas: 1 },
        { nome: 'Cabo I entre as 2 estruturas mais próximas', modo: 'proximidade', raio_m: 15, cabos: { operacao: ['I', '*I'] }, estruturas: 'aceitas', pontas: 2, aceitar_parcial: true },
        { nome: 'Multiplex BT na sequência', modo: 'sequencia', cabos: { ativo: '^(RAMA|PB)\\d+$' }, estruturas: { ativo: '^(B2|R2)$' } }] };

    // TASK-099: o mesmo vínculo por coordenadas, com RESERVA por sorteio (reproduzível, respeitando as Tabelas de Cabos/Estruturas) para o que as coordenadas não resolvem
    const SCRIPT_COM_SORTEIO = { versao: 1, conflitos: 'avisar', passos: SCRIPT_EXEMPLO.passos.concat([
        { nome: 'Reserva: cabo M sem coordenada ou fora do raio → sorteia uma amarração', modo: 'aleatorio', semente: 1, cabos: { operacao: ['M'] }, estruturas: { amarracao: true }, pontas: 1 },
        { nome: 'Reserva: cabo I sem coordenada, fora do raio ou parcial → completa as pontas por sorteio', modo: 'aleatorio', semente: 1, completar: true, cabos: { operacao: ['I', '*I'] }, estruturas: 'aceitas', pontas: 2, aceitar_parcial: true }]) };

    function scriptAtual() {
        const l = st.cfg.scripts;
        if (st.scriptIdx >= l.length) st.scriptIdx = Math.max(0, l.length - 1);
        return l[st.scriptIdx] || null;
    }

    function lerScript() {
        try { const s = JSON.parse(st.textoScript); const e = VR.validarScript(s); st.erroScript = e.join('\n'); return e.length ? null : s; }
        catch (err) { st.erroScript = 'JSON inválido: ' + err.message; return null; }
    }

    function abaScript(corpo) {
        const l = st.cfg.scripts, atual = scriptAtual();
        if (st.textoScript === null) st.textoScript = JSON.stringify(atual ? atual.script : SCRIPT_EXEMPLO, null, 2);
        const sel = el('select', { style: 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;padding:5px 8px', onchange: e => { st.scriptIdx = Number(e.target.value); st.textoScript = null; st.previa = null; render(); } },
            l.length ? l.map((s, i) => el('option', { value: i, text: s.nome || 'Script ' + (i + 1), ...(i === st.scriptIdx ? { selected: 'selected' } : {}) })) : [el('option', { text: '(nenhum script salvo — o exemplo abaixo é só um modelo)' })]);
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap' }, [sel,
            el('button', { class: 'vz-btn', text: 'Salvar neste script', onclick: () => { const s = lerScript(); if (!s) return render(); if (atual) atual.script = s; else { l.push({ nome: raiz.prompt('Nome do script:', 'Rede aérea padrão') || 'Script ' + (l.length + 1), script: s }); st.scriptIdx = l.length - 1; } marcarSujo(); toast('Script guardado (use "Salvar tabelas" para gravar no projeto).'); render(); } }),
            el('button', { class: 'vz-btn', text: 'Novo', onclick: () => { const nome = raiz.prompt('Nome do novo script:', 'Novo script'); if (!nome) return; l.push({ nome, script: clone(SCRIPT_EXEMPLO) }); st.scriptIdx = l.length - 1; st.textoScript = null; marcarSujo(); render(); } }),
            el('button', { class: 'vz-btn', text: 'Novo (coordenadas + sorteio)', title: 'Vínculo por coordenadas e, para o que sobrar, sorteio reproduzível respeitando as tabelas', onclick: () => { const nome = raiz.prompt('Nome do novo script:', 'Coordenadas com sorteio de reserva'); if (!nome) return; l.push({ nome, script: clone(SCRIPT_COM_SORTEIO) }); st.scriptIdx = l.length - 1; st.textoScript = null; marcarSujo(); render(); } }),
            el('button', { class: 'vz-btn', text: 'Excluir', disabled: atual ? null : 'disabled', onclick: () => { if (raiz.confirm('Excluir o script "' + atual.nome + '"?')) { l.splice(st.scriptIdx, 1); st.textoScript = null; marcarSujo(); render(); } } }),
            el('span', { style: 'margin-left:auto' }),
            el('button', { class: 'vz-btn', text: 'Importar', onclick: () => { const inp = el('input', { type: 'file', accept: '.json', onchange: async e => { try { const d = JSON.parse(await e.target.files[0].text()); st.textoScript = JSON.stringify(d.script || d, null, 2); st.previa = null; render(); } catch (err) { raiz.alert('Arquivo inválido: ' + err.message); } } }); inp.click(); } }),
            el('button', { class: 'vz-btn', text: 'Exportar', onclick: () => { const s = lerScript(); if (s) baixar(`vinculacao-script-${host.projeto()}.json`, { formato: 'vinculacao-script', versao: 1, nome: atual ? atual.nome : 'script', script: s }); else render(); } })]));
        const ta = el('textarea', { class: 'vz-pre', spellcheck: 'false', style: 'height:calc(100% - 60px);min-height:300px', oninput: e => { st.textoScript = e.target.value; } }); ta.value = st.textoScript;
        const erro = el('div', { class: 'vz-erro', text: st.erroScript });
        const painel = el('div', { class: 'vz-painel' }, [el('h5', { text: 'Pré-visualização (nada foi aplicado)' }), previaHtml()]);
        const acoes = el('div', { style: 'display:flex;gap:8px;margin-top:10px;flex-wrap:wrap' }, [
            el('button', { class: 'vz-btn b', text: '▶ Pré-visualizar', onclick: () => { st.previa = rodar(false); render(); } }),
            el('button', { class: 'vz-btn', text: 'Aplicar só os completos', onclick: () => aplicarScript(true) }),
            el('button', { class: 'vz-btn p', text: 'Aplicar tudo', onclick: () => aplicarScript(false) })]);
        painel.appendChild(acoes);
        painel.appendChild(el('div', { class: 'vz-nota', style: 'margin-top:12px', html: 'O script usa as duas tabelas (o que cada lado aceita) e as <b>coordenadas</b> do DXF (<code>_x/_y</code>). Vínculo manual existente nunca é sobrescrito, a menos que o passo diga <code>"substituir": true</code>. Aplicado, cada vínculo fica marcado como <b>script</b> e pode ser desfeito um a um (Ctrl+Z desfaz tudo).' }));
        corpo.appendChild(el('div', { class: 'vz-grid2', style: 'flex:1;min-height:0;overflow:auto' }, [el('div', { style: 'min-height:0' }, [ta, erro]), painel]));
    }

    function rodar(aplicar, soCompletos) {
        const s = lerScript();
        if (!s) return null;
        if (!aplicar) return VR.executarScript(s, clone(cabos()), clone(outros()), st.cfg, host.prefixoCabo, host.codigoEstrutura, false, host.codigosEstrutura);
        const copia = clone(cabos());
        const r = VR.executarScript(s, copia, outros(), st.cfg, host.prefixoCabo, host.codigoEstrutura, false, host.codigosEstrutura);
        host.antes();
        r.vinculos.forEach(v => { if (soCompletos && !v.completo) return; cabos()[v.cabo].vinculoEstruturas = v.estruturas.slice(); cabos()[v.cabo].vinculoOrigem = 'script'; });
        return r;
    }

    function aplicarScript(soCompletos) {
        const r = rodar(true, soCompletos);
        if (!r) { render(); return; }
        const n = r.vinculos.filter(v => !soCompletos || v.completo).length;
        toast(`${n} cabo(s) vinculado(s) pelo script (Ctrl+Z desfaz).`);
        st.previa = null; st.aba = 'vinculo'; render();
    }

    function previaHtml() {
        const r = st.previa;
        if (!r) return el('div', { style: 'color:#8b949e', text: 'Clique em "Pré-visualizar" para ver o que o script faria.' });
        const tb = el('table', { class: 'vz-tb' }, [el('thead', {}, [el('tr', {}, ['Passo', 'Cabos alvo', 'Completos', 'Parciais', 'Sem candidato', 'Conflitos'].map(t => el('th', { text: t })))]),
            el('tbody', {}, r.passos.map(p => el('tr', {}, [el('td', { text: p.nome }), el('td', { text: String(p.alvo) + (p.ja_tinham ? ` (+${p.ja_tinham} já vinculados)` : '') }), el('td', { class: 'vz-ok', text: String(p.completos) }),
                el('td', { class: 'vz-av', text: String(p.parciais) }), el('td', { text: String(p.sem_candidato) }), el('td', { style: p.conflitos ? 'color:#f85149' : '', text: String(p.conflitos) })])))]);
        const conf = r.conflitos.length ? el('div', { class: 'vz-erro', text: r.conflitos.slice(0, 8).map(c => `• cabo ${c.cabo + 1}: ${c.motivo}`).join('\n') + (r.conflitos.length > 8 ? `\n… e mais ${r.conflitos.length - 8}` : '') }) : null;
        return el('div', {}, [tb, el('div', { style: 'margin:8px 0;color:#8b949e', html: `Total: <b style="color:#fff">${r.vinculos.length}</b> cabo(s) seriam vinculados · ${r.sem_candidato.length} sem candidato · ${r.conflitos.length} conflito(s)${r.parou ? ' · <span class="vz-erro">parou no 1º conflito</span>' : ''}` }), conf]);
    }

    // ── ABA: MAPA ──
    function abaMapa(corpo, a) {
        const pts = [];
        // TASK-091: cabo = linha da rede (REDE_MT/BT) associada à anotação; estrutura = poste associado. Sem associação, cai no ponto da anotação.
        const meio = g => { let tot = 0; const seg = []; for (let i = 0; i < g.length - 1; i++) { const d = Math.hypot(g[i + 1][0] - g[i][0], g[i + 1][1] - g[i][1]); seg.push(d); tot += d; } let alvo = tot / 2; for (let i = 0; i < seg.length; i++) { if (alvo <= seg[i] || i === seg.length - 1) { const f = seg[i] ? Math.min(1, alvo / seg[i]) : 0; return [g[i][0] + (g[i + 1][0] - g[i][0]) * f, g[i][1] + (g[i + 1][1] - g[i][1]) * f]; } alvo -= seg[i]; } return g[0]; };
        cabos().forEach((c, ic) => { if (!VR.ehCaboValido(c)) return; if (Array.isArray(c._geo) && c._geo.length > 1) { const m = meio(c._geo); pts.push({ t: 'c', i: ic, x: m[0], y: m[1], geo: c._geo, tipo: c._geoTipo }); } else if (c._x != null && c._y != null) pts.push({ t: 'c', i: ic, x: c._x, y: c._y }); });
        const estsMapa = VR.estruturasDisponiveis(outros(), host.codigoEstrutura, host.codigosEstrutura);
        estsMapa.forEach(ie => { const o = linhaOutros(ie), off = slotDe(ie); if (Array.isArray(o._poste)) pts.push({ t: 'e', i: ie, x: o._poste[0], y: o._poste[1], poste: true, off }); else if (o._x != null && o._y != null) pts.push({ t: 'e', i: ie, x: o._x, y: o._y, off }); });      // estruturas do mesmo poste: círculos lado a lado
        let fundo = null; try { fundo = JSON.parse(raiz.localStorage.getItem('mapa_geometria_fundo') || 'null'); } catch (e) { fundo = null; }
        if (st.fundo === undefined) st.fundo = true;
        const semXY = cabos().filter(c => VR.ehCaboValido(c) && (c._x == null || c._y == null)).length + estsMapa.filter(ie => linhaOutros(ie)._x == null || linhaOutros(ie)._y == null).length;
        corpo.appendChild(el('div', { class: 'vz-nota', html: `<b>Mapa pelas coordenadas do desenho (DXF).</b> Cabo = linha da rede (vermelha MT / verde BT) ou quadrado quando não achei a linha; estrutura = círculo no poste. Cores de estado como no painel. Clique para selecionar (vale também na aba Vínculo). Roda do mouse = zoom · arrastar = mover.${semXY ? ` <span class="vz-av">${semXY} item(ns) sem coordenada (PDF ou digitados) não aparecem.</span>` : ''}` }));
        if (fundo && ((fundo.redes || []).length || (fundo.postes || []).length)) corpo.appendChild(el('label', { class: 'vz-nota', style: 'display:block;cursor:pointer' }, [el('input', { type: 'checkbox', ...(st.fundo ? { checked: 'checked' } : {}), onchange: e => { st.fundo = e.target.checked; render(); } }), ' Mostrar a rede e os postes ao fundo (esmaecidos)']));
        if (!pts.length) { corpo.appendChild(el('div', { style: 'color:#8b949e;padding:20px', text: 'Nenhum item com coordenada.' })); return; }
        const xs = [], ys = [];
        pts.forEach(p => { xs.push(p.x); ys.push(p.y); if (p.geo) p.geo.forEach(q => { xs.push(q[0]); ys.push(q[1]); }); });
        const minX = Math.min(...xs), maxX = Math.max(...xs), minY = Math.min(...ys), maxY = Math.max(...ys);
        const W = 1200, H = 600, pad = 40, sx = (W - 2 * pad) / Math.max(maxX - minX, 1e-9), sy = (H - 2 * pad) / Math.max(maxY - minY, 1e-9), k = Math.min(sx, sy);
        const px = p => pad + (p.x - minX) * k + (p.off || 0) * 18, py = p => H - pad - (p.y - minY) * k;     // Y do desenho cresce para cima
        const cor = e => ({ k: '#6e7681', pu: '#a371f7', gr: '#3fb950', rd: '#f85149' }[e]);
        let s = '';
        const pt = (x, y) => `${(pad + (x - minX) * k).toFixed(1)},${(H - pad - (y - minY) * k).toFixed(1)}`;
        if (fundo && st.fundo) {
            (fundo.redes || []).forEach(r => { s += `<polyline points="${r.p.map(q => pt(q[0], q[1])).join(' ')}" fill="none" stroke="${r.t === 'MT' ? '#f85149' : '#3fb950'}" stroke-opacity="0.22" stroke-width="1"/>`; });
            (fundo.postes || []).forEach(q => { s += `<circle cx="${pad + (q[0] - minX) * k}" cy="${H - pad - (q[1] - minY) * k}" r="1.6" fill="#8b949e" fill-opacity="0.45"/>`; });
        }
        const ponto = {}; pts.forEach(p => { ponto[p.t + p.i] = p; });
        a.links.forEach(l => { const c = ponto['c' + l.cabo], e = ponto['e' + l.estrutura]; if (c && e) s += `<line x1="${px(c)}" y1="${py(c)}" x2="${px(e)}" y2="${py(e)}" stroke="${l.ok ? '#3fb950' : '#f85149'}" stroke-width="1.8" ${l.ok ? '' : 'stroke-dasharray="5 4"'}/>`; });
        pts.forEach(p => {
            const sel = st.sel && st.sel.t === p.t && st.sel.i === p.i, est = (p.t === 'c' ? a.cabos : a.estruturas)[p.i].estado, c = cor(est);
            const rot = p.t === 'c' ? (cabos()[p.i].ativo || '') : (a.estruturas[p.i].codigo || '');
            const corRede = p.tipo === 'MT' ? '#f85149' : '#3fb950';
            const linha = p.geo ? `<polyline points="${p.geo.map(q => pt(q[0], q[1])).join(' ')}" fill="none" stroke="transparent" stroke-width="14"/><polyline points="${p.geo.map(q => pt(q[0], q[1])).join(' ')}" fill="none" stroke="${corRede}" stroke-width="${sel ? 4.5 : 3}" stroke-linejoin="round"/><polyline points="${p.geo.map(q => pt(q[0], q[1])).join(' ')}" fill="none" stroke="${c}" stroke-width="1.2" stroke-dasharray="1 5" stroke-linecap="round"/>` : '';
            s += `<g data-t="${p.t}" data-i="${p.i}" style="cursor:pointer">${linha}${p.t === 'c' && p.geo ? '' : p.t === 'c' ? `<rect x="${px(p) - 6}" y="${py(p) - 6}" width="12" height="12" fill="${c}33" stroke="${c}" stroke-width="2"/>` : `<circle cx="${px(p)}" cy="${py(p)}" r="7" fill="${c}33" stroke="${c}" stroke-width="2"/>`}${sel ? `<circle cx="${px(p)}" cy="${py(p)}" r="12" fill="none" stroke="#58a6ff" stroke-width="2"/>` : ''}<text x="${px(p) + 10}" y="${py(p) - 8}" fill="#8b949e" font-size="10">${esc(String(rot).slice(0, 22))}</text><title>${esc(rot)}</title></g>`;
        });
        const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('class', 'vz-mapa'); svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.innerHTML = s;
        let vb = [0, 0, W, H], arrasto = null;
        const aplicarVb = () => svg.setAttribute('viewBox', vb.join(' '));
        svg.addEventListener('wheel', e => { e.preventDefault(); const f = e.deltaY > 0 ? 1.15 : 1 / 1.15, r = svg.getBoundingClientRect(), mx = vb[0] + (e.clientX - r.left) / r.width * vb[2], my = vb[1] + (e.clientY - r.top) / r.height * vb[3]; vb = [mx - (mx - vb[0]) * f, my - (my - vb[1]) * f, vb[2] * f, vb[3] * f]; aplicarVb(); }, { passive: false });
        svg.addEventListener('mousedown', e => { arrasto = { x: e.clientX, y: e.clientY, vb: vb.slice(), mov: false }; });
        raiz.addEventListener('mousemove', e => { if (!arrasto) return; const r = svg.getBoundingClientRect(), dx = (e.clientX - arrasto.x) / r.width * arrasto.vb[2], dy = (e.clientY - arrasto.y) / r.height * arrasto.vb[3]; if (Math.abs(e.clientX - arrasto.x) + Math.abs(e.clientY - arrasto.y) > 3) arrasto.mov = true; vb = [arrasto.vb[0] - dx, arrasto.vb[1] - dy, arrasto.vb[2], arrasto.vb[3]]; aplicarVb(); });
        raiz.addEventListener('mouseup', () => { setTimeout(() => { arrasto = null; }, 0); });
        svg.addEventListener('click', e => { const g = e.target.closest('g[data-t]'); if (!g || (arrasto && arrasto.mov)) return; const t = g.getAttribute('data-t'), i = Number(g.getAttribute('data-i')); if (st.sel && st.sel.t !== t) { const ic = t === 'c' ? i : st.sel.i, ie = t === 'e' ? i : st.sel.i; ligar(ic, ie); } else { st.sel = (st.sel && st.sel.t === t && st.sel.i === i) ? null : { t, i }; render(); } });
        corpo.appendChild(el('div', { class: 'vz-scroll', style: 'overflow:hidden' }, [svg]));
        // coordenadas do item selecionado (só aqui no mapa, como pedido)
        const f2 = n => Number(n).toFixed(2), sp = st.sel && pts.find(p => p.t === st.sel.t && p.i === st.sel.i);
        corpo.appendChild(el('div', { class: 'vz-nota', style: 'margin-top:6px;word-break:break-all', html: !sp ? 'Clique num cabo ou estrutura para ver as coordenadas.'
            : sp.t === 'c' && sp.geo ? `<b>${esc(cabos()[sp.i].ativo || '')}</b> — rede ${sp.tipo || ''} · ${sp.geo.length} ponto(s): ${esc(sp.geo.map(q => `(${f2(q[0])}; ${f2(q[1])})`).join(' → '))}`
            : sp.t === 'e' && sp.poste ? `<b>${esc(a.estruturas[sp.i].codigo || '')}</b> — poste em (${f2(sp.x)}; ${f2(sp.y)})`
            : `<b>${esc(sp.t === 'c' ? (cabos()[sp.i].ativo || '') : (a.estruturas[sp.i].codigo || ''))}</b> — ${sp.t === 'c' ? 'linha da rede' : 'poste'} não encontrado; usando o ponto da anotação (${f2(sp.x)}; ${f2(sp.y)})` }));
    }

    raiz.VinculacaoUI = { abrir, fechar };
})(typeof window !== 'undefined' ? window : globalThis);
