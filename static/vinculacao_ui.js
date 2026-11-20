/*
 * vinculacao_ui.js (TASK-084) — modal de vinculação cabo <-> estrutura em abas: Vínculo (painel visual), Tabela de Cabos, Tabela de Estruturas,
 * Script e Mapa. A lógica pura está em vinculacao_regras.js (testada em Node e em Python). Quem altera as tabelas do Resumo é o host (resumo.js).
 * Cores: preto = sem vínculo · roxo = parcial · verde = completo · vermelho = inválido · tracejado = cabo fictício.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document, VR = raiz.VinculacaoRegras;
    let host = null, st = null, overlay = null, mapaAtualizar = null;

    const CSS = `
    .vz-ov{position:fixed;inset:0;background:#05080d;z-index:9000;display:flex;align-items:center;justify-content:center}
    .vz{width:min(1500px,96vw);height:min(900px,94vh);background:#0d1117;border:1px solid #30363d;border-radius:12px;display:flex;flex-direction:column;overflow:hidden;color:#c9d1d9;font:13px/1.35 -apple-system,Segoe UI,Roboto,sans-serif}
    .vz *{box-sizing:border-box}
    .vz-top{display:flex;align-items:center;gap:10px;padding:12px 18px;background:#161b22;border-bottom:1px solid #30363d}.vz-top h3{margin:0;font-size:15px;color:#fff}
    .vz-chip{background:#21262d;border:1px solid #30363d;border-radius:20px;padding:2px 10px;font-size:11px;color:#8b949e;cursor:default}.vz-chip.on{color:#fff;border-color:#58a6ff}.vz-chip.btn{cursor:pointer}
    .vz-x{margin-left:auto;color:#8b949e;font-size:18px;cursor:pointer;background:none;border:none}
    .vz-tabs{display:flex;gap:4px;overflow-x:auto;padding:10px 18px 0;background:#161b22;border-bottom:1px solid #30363d}
    .vz-tab{flex-shrink:0;white-space:nowrap;padding:8px 16px;border:1px solid transparent;border-bottom:none;border-radius:8px 8px 0 0;color:#8b949e;font-weight:600;cursor:pointer;background:none}
    .vz-tab.on{background:#0d1117;border-color:#30363d;color:#fff;position:relative;top:1px}.vz-tab b{font-size:10px;background:#30363d;border-radius:10px;padding:1px 6px;margin-left:6px;color:#c9d1d9}
    .vz-corpo{flex:1;min-height:0;display:flex;flex-direction:column;padding:14px 18px}.vz-scroll{flex:1;min-height:0;overflow:auto;position:relative}
    .vz-foot{display:flex;gap:12px;align-items:center;padding:10px 18px;background:#161b22;border-top:1px solid #30363d;font-size:12px}
    .vz-btn{background:#21262d;border:1px solid #30363d;border-radius:6px;padding:5px 12px;color:#c9d1d9;font-size:12px;cursor:pointer}.vz-btn:hover{border-color:#8b949e}.vz-btn.p{background:#238636;border-color:#2ea043;color:#fff}.vz-btn.b{border-color:#58a6ff;color:#58a6ff}.vz-btn[disabled]{opacity:.45;cursor:default}
    .vz-leg{display:flex;gap:12px;align-items:center;color:#8b949e}.vz-dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:5px;vertical-align:-1px;border:2px solid #6e7681;background:#010409}
    .vz-dot.pu{border-color:#a371f7;background:#2a1d4d}.vz-dot.gr{border-color:#3fb950;background:#12361c}.vz-dot.rd{border-color:#f85149;background:#47161a}.vz-dot.fic{border-style:dashed;border-color:#a371f7}
    /* TASK-185: painel virtualizado — só as linhas visíveis existem no DOM; cada linha tem altura fixa (compacto: 1 cabo = 1 linha, 1 poste = 1 linha) */
    .vl-cab{display:flex;align-items:center;gap:8px;padding:0 10px;border-radius:6px;border:1.5px solid #6e7681;background:#010409;cursor:pointer;position:absolute;left:0;width:var(--lw,430px);overflow:hidden;white-space:nowrap}
    .vl-cab.pu{border-color:#a371f7;background:#1d1537}.vl-cab.gr{border-color:#3fb950;background:#0f2a17}.vl-cab.rd{border-color:#f85149;background:#3a1216}.vl-cab.fic{border-style:dashed;border-color:#a371f7;background:#161027}
    .vl-cab.sel,.vl-chip.sel{box-shadow:0 0 0 2px rgba(88,166,255,.8)}.vl-cab.glow,.vl-chip.glow{box-shadow:0 0 0 2px rgba(63,185,80,.5)}.vl-dim{opacity:.3}
    .vl-cab .t{font-weight:700;font-size:12px;overflow:hidden;text-overflow:ellipsis;flex:1;min-width:0}.vl-cab .s{font-size:10.5px;color:#8b949e}.vl-cab .r{font-size:10.5px;font-weight:700}.vl-cab .pp{font-size:10.5px;color:#58a6ff}
    .vl-av{display:flex;align-items:center;gap:6px;padding:0 8px 0 18px;position:absolute;left:0;width:var(--lw,430px);font-size:10.5px;color:#8b949e;background:#161b22;border:1px solid #30363d;border-radius:5px;overflow:hidden;white-space:nowrap}
    .vl-av .n{color:#c9d1d9}.vl-av .t{flex:1;overflow:hidden;text-overflow:ellipsis}.vl-av .v{color:#d29922}.vl-av .ok{color:#3fb950}.vl-av .vz-btn{padding:0 8px;font-size:10.5px}
    .vl-pst{position:absolute;display:flex;align-items:center;border:1px solid #21262d;border-radius:6px;background:#0d1117;left:var(--rl,570px);right:0}
    .vl-pst .lab{width:112px;padding:0 8px;font-size:11px;color:#8b949e;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:none}
    .vl-chip{position:absolute;top:2px;bottom:2px;width:84px;display:flex;align-items:center;justify-content:center;gap:5px;border-radius:5px;border:1.5px solid #6e7681;background:#010409;font-weight:800;font-size:12px;color:#fff;cursor:pointer;overflow:hidden;white-space:nowrap}
    .vl-chip i{font-style:normal;font-weight:600;font-size:10px;color:#8b949e}.vl-chip.eqp{border-style:dashed;background:#161b22;color:#8b949e;cursor:default}.vl-chip.pu{border-color:#a371f7;background:#1d1537}.vl-chip.gr{border-color:#3fb950;background:#0f2a17}.vl-chip.rd{border-color:#f85149;background:#3a1216}
    .vl-chip.poste{position:static;width:auto;min-width:84px;padding:0 8px;margin:2px 3px}
    .vl-box{position:relative;flex:1;min-height:0;overflow:auto;will-change:transform;contain:layout paint}.vl-in{position:relative;min-width:var(--mw,1300px)}.vl-hd{display:flex;gap:0;font-size:11px;letter-spacing:.08em;color:#8b949e;text-transform:uppercase;padding:0 0 6px}
    .vl-ach{max-height:92px;overflow:auto;margin-top:8px;padding:6px 10px;border:1px solid #30363d;border-radius:8px;background:#161b22;font-size:11px;color:#8b949e}.vl-ach b{color:#d29922}.vl-ach .a{padding:1px 0}.vl-ach .a.incompleta,.vl-ach .a.sobra{color:#c9d1d9}
    .vl-liga{position:absolute;left:0;top:0;pointer-events:none;overflow:visible}.vl-liga path.hit{pointer-events:stroke;cursor:pointer;stroke:transparent;stroke-width:10;fill:none}
    .vz-op{font-size:11px;font-weight:700;border-radius:4px;padding:1px 6px;background:#1f6feb33;color:#58a6ff}.vz-op.m{background:#d2992233;color:#d29922}.vz-op.r{background:#f8514933;color:#f85149}
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
    .vz--grafo,.vz-no-dialog{
        --g-bg:#1a191f;--g-panel:#222128;--g-raised:#2a2832;--g-border:#393641;
        --g-text:#eeedf3;--g-muted:#aaa6b5;--g-accent:#bda6ff;--g-accent-bg:#3a2a56;
        --g-green:#78b89a;--g-red:#ed909a;--g-gray:#787481;
        font-family:Inter,system-ui,sans-serif;color:var(--g-text);color-scheme:dark
    }
    .vz--grafo{background:var(--g-bg);border-color:var(--g-border);box-shadow:0 24px 80px #09080db3}
    .vz--grafo .vz-top,.vz--grafo .vz-tabs,.vz--grafo .vz-foot{
        background:var(--g-panel);border-color:var(--g-border)
    }
    .vz--grafo .vz-chip,.vz--grafo .vz-tab b{background:var(--g-raised);border-color:var(--g-border);color:var(--g-muted)}
    .vz--grafo .vz-tab{color:var(--g-muted);border-radius:7px 7px 0 0}
    .vz--grafo .vz-tab.on{background:var(--g-bg);color:var(--g-accent);border-color:var(--g-border);box-shadow:inset 0 2px var(--g-accent)}
    .vz--grafo .vz-leg{color:var(--g-muted);flex-wrap:wrap}
    .vz--grafo .vz-dot{background:transparent;border-color:var(--g-gray)}
    .vz--grafo .vz-dot.pu,.vz--grafo .vz-dot.fic{border-color:var(--g-accent)}
    .vz--grafo .vz-dot.gr{border-color:var(--g-green)}.vz--grafo .vz-dot.rd{border-color:var(--g-red)}
    .vz--grafo .vz-btn,.vz-no-dialog .vz-btn{
        background:var(--g-raised);color:var(--g-text);border:1px solid var(--g-border);border-radius:7px;padding:6px 11px
    }
    .vz--grafo .vz-btn:hover,.vz-no-dialog .vz-btn:hover{border-color:var(--g-accent)}
    .vz--grafo .vz-btn.b{color:var(--g-accent)}
    .vz--grafo .vz-btn.p,.vz-no-dialog .vz-btn.p{background:var(--g-accent-bg);color:#f4ebff;border-color:#71559b}
    .vz--grafo .vz-btn:focus-visible,.vz-no-dialog .vz-btn:focus-visible,.vz-no-dialog select:focus-visible{outline:2px solid var(--g-accent);outline-offset:2px}
    .vz-grafo-barra{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px 16px;margin-bottom:14px}
    .vz-grafo-barra h4{margin:0;font-size:20px;letter-spacing:-.024em}
    .vz-grafo-barra p{margin:2px 0 0;color:var(--g-muted);font-size:12px}
    .vz-grafo-legenda{display:flex;gap:10px}
    .vz-grafo-total{background:var(--g-panel);border:1px solid var(--g-border);border-radius:8px;padding:6px 12px;color:var(--g-muted);white-space:nowrap}
    .vz-grafo-total strong{color:var(--g-text);margin-right:5px;font-variant-numeric:tabular-nums}
    .vz-grafo-vista{border-radius:10px;border:1px solid var(--g-border);overflow:hidden}
    .vz-grafo-acoes{position:absolute;bottom:12px;right:12px;display:flex;gap:3px;padding:4px;background:var(--g-panel);border:1px solid var(--g-border);border-radius:9px}
    .vz-grafo-acoes .vz-btn{min-width:32px;padding:5px 8px;background:transparent;border-color:transparent}
    .vz-grafo-aviso{color:var(--g-muted);margin:7px 0 0;font-size:11px}.vz-grafo-aviso:empty{display:none}
    .vz-mapa{display:block;width:100%;height:100%;background:radial-gradient(circle,#3b3645 1px,transparent 1.2px) 0 0 / 22px 22px,#17161c;cursor:grab;touch-action:none}
    .vz-mapa.arrastando{cursor:grabbing}.vz-mapa [data-no]{cursor:pointer;outline:none}
    .vz-mapa [data-no] .vz-alvo{fill:transparent}.vz-mapa [data-no] .vz-halo{fill:#9871e94a;opacity:.5}
    .vz-mapa [data-no] .vz-marcador{fill:#21192e;stroke:var(--g-accent);stroke-width:1.7;vector-effect:non-scaling-stroke}
    .vz-mapa [data-no]:hover .vz-marcador,.vz-mapa [data-no]:focus-visible .vz-marcador{fill:var(--g-accent);stroke:#fff}
    .vz-mapa [data-no] text{fill:#e6dff3;font-size:11px;font-weight:600;paint-order:stroke;stroke:#17161c;stroke-width:3px;pointer-events:none}
    .vz-mapa .vz-rotulo-denso text{visibility:hidden}.vz-mapa .vz-rotulo-denso:hover text,.vz-mapa .vz-rotulo-denso:focus-visible text{visibility:visible}
    .vz-mapa .vz-aresta{stroke:var(--g-gray);stroke-width:1.65;stroke-linecap:round;stroke-linejoin:round}
    .vz-mapa .vz-aresta[data-vao]{cursor:pointer}.vz-mapa .vz-aresta[data-vao]:focus-visible{outline:none;stroke-width:6}
    .vz-mapa .vz-aresta.mt{stroke:#b95f6a}.vz-mapa .vz-aresta.bt{stroke:#4e9875}
    .vz-mapa .vz-aresta.completo{stroke-width:3.6}.vz-mapa .vz-aresta.mt.completo{stroke:#ff6478}
    .vz-mapa .vz-aresta.bt.completo{stroke:#54e1a3}
    .vz-mapa [data-no].completo .vz-marcador{fill:#25513e;stroke:#54e1a3}
    .vz-mapa [data-no].completo .vz-halo{fill:#54e1a34d}
    .vz-mapa [data-no].completo:hover .vz-marcador,.vz-mapa [data-no].completo:focus-visible .vz-marcador{fill:#54e1a3;stroke:#fff}
    .vz-linha-legenda{display:inline-block;width:17px;height:0;border-top:2px solid;vertical-align:middle;margin-right:5px}
    .vz-linha-legenda.mt{color:#ff6478}.vz-linha-legenda.bt{color:#54e1a3}
    .vz--grafo .vz-dot.poste-completo{border-color:#54e1a3;background:#25513e}
    .vz-no-topo .vz-no-completo{color:#54e1a3}
    .vz-no-dialog{width:min(880px,94vw);max-height:84vh;margin:auto;padding:0;background:var(--g-bg);border:1px solid #514663;border-radius:14px;font-size:13px;line-height:1.45;overflow:hidden;box-shadow:0 24px 85px #08060cbd}
    .vz-no-dialog *, .vz-no-dialog{box-sizing:border-box}.vz-no-dialog[open]{display:flex;flex-direction:column}
    .vz-no-dialog::backdrop{background:#111017b5}
    .vz-no-dialog .vz-no-topo{animation:none;text-align:left;position:static;flex:none;display:flex;gap:12px;align-items:center;padding:17px 20px;background:var(--g-panel);border-bottom:1px solid var(--g-border)}
    .vz-no-topo h4,.vz-no-dialog h5{margin:0;color:var(--g-text)}.vz-no-topo h4{font-size:19px}
    .vz-no-topo p{margin:3px 0 0;color:var(--g-muted);font-size:12px;font-variant-numeric:tabular-nums}
    .vz-no-topo-acoes{display:flex;gap:6px;margin-left:auto;align-items:center}
    .vz-no-conteudo{min-height:0;overflow:auto;padding:16px 20px;display:grid;grid-template-columns:minmax(150px,1fr) minmax(0,2.7fr);align-items:start;gap:18px}
    .vz-no-conteudo h5{font-size:13px;margin-bottom:10px}
    .vz-no-lateral{border:1px solid var(--g-border);border-radius:9px;padding:12px;background:var(--g-panel)}
    .vz-no-estruturas{display:grid;gap:5px}.vz-no-estrutura{padding:5px 0;border-top:1px solid var(--g-border);overflow-wrap:anywhere}
    .vz-no-estrutura span,.vz-no-ajuda{color:var(--g-muted)}
    .vz-no-vinculos{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
    .vz-no-cabo{padding:11px 12px;border:1px solid var(--g-border);border-radius:8px;margin-bottom:8px;background:var(--g-panel)}
    .vz-no-cabo h6{margin:0 0 3px;font-size:13px;overflow-wrap:anywhere}.vz-no-cabo p{margin:0 0 7px}
    .vz-no-cabo .vz-op{margin-right:4px}
    .vz-no-acoes{display:flex;gap:7px;align-items:center;flex-wrap:wrap;margin-top:8px}
    .vz-no-acoes select{min-width:140px;max-width:100%;background:var(--g-bg);color:var(--g-text);border:1px solid var(--g-border);border-radius:7px;padding:5px 7px}
    .vz-no-acoes .vz-no-ajuda{flex-basis:100%;font-size:11px;color:#e1bc8d}.vz-no-acoes .vz-no-ajuda:empty{display:none}
    .vz-no-status{margin:0 0 8px;color:var(--g-accent)}.vz-no-status:empty{display:none}
    @media(max-width:630px){
        .vz-no-dialog .vz-no-topo{padding:12px;align-items:flex-start}
        .vz-no-conteudo{grid-template-columns:1fr;padding:12px;gap:12px}
        .vz-no-estruturas{display:flex;gap:8px;flex-wrap:wrap}.vz-no-estrutura{border:0;padding:0}
        .vz--grafo .vz-foot{flex-wrap:wrap;gap:6px;padding:7px 12px}
    }
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
    // ── contexto memoizado (TASK-185): análise + índices calculados UMA vez por estado dos dados, nunca a cada clique/rolagem ──
    let cx = null;
    const metaAtivo = new Map();          // texto da linha de Outros → { soPoste, postes } (a interpretação é cara e só depende do texto)
    function assinatura() {
        let s = st.cfgVer + '#';
        cabos().forEach(c => { s += c ? `${c.ativo || ''}|${c.operacao || ''}|${c.vinculoEstruturas || ''}|${c.vinculoPostes || ''}|${c.estruturasFicticias || ''}|${c.textoOriginal || ''}|${c.fases || ''}|${c.tensao || ''}|${c.nivel || ''}\n` : '\n'; });
        outros().forEach(o => { s += o ? `${o.ativo || ''}|${o.cabosFicticios || ''}|${o.cabosFicticiosSlots ? JSON.stringify(o.cabosFicticiosSlots) : ''}|${o.origemAvulso || ''}\n` : '\n'; });
        return s;
    }
    function ctx() {
        const sig = assinatura();
        if (cx && cx.sig === sig && cx.cfg === st.cfg) return cx;
        const C = cabos(), a = VR.analisar(C, outros(), st.cfg, host.prefixoCabo, host.codigoEstrutura, host.codigosEstrutura);
        cx = { sig, cfg: st.cfg, a, infoC: {}, infoE: {}, ocup: {}, avul: {}, it: {} };
        C.forEach((c, ic) => { if (VR.ehCaboValido(c)) (c.vinculoEstruturas || []).forEach(ie => (cx.ocup[ie] = cx.ocup[ie] || []).push(ic)); });
        return cx;
    }
    const analise = () => ctx().a;
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
        st = { cfg: clone(VR.normalizarConfig(cfg)), aba: (opcoes && opcoes.aba) || 'vinculo', sel: null, filtro: 'todos', filtroCP: 'todos', kbModo: 'c', kbP: null, elAch: null, auto: true, cfgVer: 0, resultadoPoste: null, elRes: null, dirty: false, scriptIdx: 0, previa: null, textoScript: null, erroScript: '', teste: {} };
        if (!overlay) {
            overlay = el('div', { class: 'vz-ov', onkeydown: e => { if (tecladoCabosPostes(e)) return; if (e.defaultPrevented || e.target.closest('dialog')) return; if (e.key === 'Escape') { if (menuAberto) fecharMenu(); else fechar(); } } });
            overlay.setAttribute('tabindex', '-1');
            doc.body.appendChild(overlay);
        }
        overlay.style.display = 'flex';
        cx = null; painel = null; metaAtivo.clear();
        render();
        overlay.focus();
    }

    function fechar() {
        if (st && st.dirty && !raiz.confirm('Há alterações nas tabelas/scripts que não foram salvas. Fechar mesmo assim?')) return false;
        fecharMenu();
        const dialogoMapa = overlay?.querySelector('.vz-no-dialog[open]');
        if (dialogoMapa) dialogoMapa.close();
        if (overlay) overlay.style.display = 'none';
        if (host && host.depois) host.depois();
        return true;
    }

    function marcarSujo() { st.dirty = true; st.cfgVer++; const b = overlay.querySelector('#vz-salvar'); if (b) b.disabled = false; }

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
        if (painel) { painel.destruir(); painel = null; }
        mapaAtualizar = null;
        st.elRes = null; st.elAch = null;
        const a = analise();
        const pend = Object.values(a.cabos).filter(c => c.estado !== 'gr').length + Object.values(a.estruturas).filter(e => e.estado === 'k' || e.estado === 'pu').length;
        overlay.innerHTML = '';
        const abas = [['vinculo', 'Vínculo', pend ? pend + ' pend.' : ''], ['cabopostes', 'Cabos ⇄ Postes', ''], ['mapa', 'Mapa', ''], ['cabos', 'Tabela de Cabos', st.cfg.cabos.length], ['estruturas', 'Tabela de Estruturas', st.cfg.estruturas.length], ['dicionario', 'Dicionário de cabos', (st.cfg.dicionario_cabos || []).length], ['postes', 'Postes', st.cfg.postes ? '✓' : ''], ['desenho', 'Desenho', ''], ['exportacao', 'Exportação', st.cfg.exportacao ? '✓' : ''], ['casos', 'Casos', ''], ['pacote', 'Pacote', ''],
            ['script', 'Script', st.cfg.scripts.length]];
        const vz = el('div', { class: st.aba === 'mapa' ? 'vz vz--grafo' : 'vz' }, [
            el('div', { class: 'vz-top' }, [el('h3', { text: '🔗 Vinculação cabo ⇄ estrutura' }), el('span', { class: 'vz-chip', text: 'Projeto ' + host.projeto() }),
                st.dirty ? el('span', { class: 'vz-chip', style: 'color:#d29922;border-color:#d29922', text: 'alterações não salvas' }) : null, el('button', { class: 'vz-x', text: '✕', onclick: fechar })]),
            el('div', { class: 'vz-tabs' }, abas.map(([id, nome, n]) => el('button', { class: 'vz-tab' + (st.aba === id ? ' on' : ''), onclick: () => { st.aba = id; st.sel = null; render(); } }, [nome, id === 'vinculo' ? el('b', { id: 'vz-pend', text: String(n), style: n === '' ? 'display:none' : '' }) : (n !== '' ? el('b', { text: String(n) }) : null)]))),
            el('div', { class: 'vz-corpo', id: 'vz-corpo' }),
            rodape()]);
        overlay.appendChild(vz);
        const corpo = vz.querySelector('#vz-corpo');
        ({ vinculo: abaVinculo, cabopostes: abaCabosPostes, cabos: abaCabos, estruturas: abaEstruturas, dicionario: abaDicionario, postes: abaPostes, desenho: abaDesenho, exportacao: abaExportacao, casos: abaCasos, pacote: abaPacote, script: abaScript, mapa: abaMapa })[st.aba](corpo, a);
    }

    function rodape() {
        return el('div', { class: 'vz-foot' }, [
            st.aba === 'mapa' ? el('div', { class: 'vz-leg' }, [
                el('span', {}, [el('i', { class: 'vz-linha-legenda mt' }), 'MT']),
                el('span', {}, [el('i', { class: 'vz-linha-legenda bt' }), 'BT']),
                el('span', {}, [el('i', { class: 'vz-dot poste-completo' }), 'Poste completo']),
                el('span', {}, [el('i', { class: 'vz-dot pu' }), 'Poste pendente']),
                el('span', { text: 'Traço forte indica cabos vinculados' })
            ]) : el('div', { class: 'vz-leg' }, [el('span', {}, [el('i', { class: 'vz-dot' }), 'sem vínculo']), el('span', {}, [el('i', { class: 'vz-dot pu' }), 'parcial']), el('span', {}, [el('i', { class: 'vz-dot gr' }), 'completo']),
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
    function infoCabo(ic) { const k = cx || ctx(), c = cabos()[ic]; return k.infoC[ic] || (k.infoC[ic] = VR.infoCabo(host.prefixoCabo(c.ativo) || '', c.operacao, st.cfg, VR.extraCabo(c))); }
    function infoEst(ie) { const k = cx || ctx(); return k.infoE[ie] || (k.infoE[ie] = VR.infoEstrutura(codigoDe(ie), st.cfg)); }
    function ocupantes(ie, ignorandoCabo) {
        const k = cx || ctx(), lista = [];
        (k.ocup[ie] || []).forEach(ic => { if (ic !== ignorandoCabo) lista.push({ info: infoCabo(ic), ficticio: false }); });
        const e = infoEst(ie);
        ficticiosDe(ie).forEach(cod => lista.push({ info: VR.infoCabo(cod, VR.opFicticio(e), st.cfg), ficticio: true }));
        return lista;
    }
    function podeLigar(ic, ie) {
        const c = infoCabo(ic), e = infoEst(ie), oc = ocupantes(ie, ic);
        return VR.podeComOcupantes(c, e, oc);
    }

    /** Atualização leve: refaz só os dados do painel (a análise é memoizada) em vez de reconstruir o modal inteiro. */
    function atualizar() {
        if (!painel && !mapaAtualizar) { render(); return; }
        const k = ctx();
        if (painel) painel.refazer();
        const b = overlay.querySelector('#vz-pend');
        if (b) { const pend = Object.values(k.a.cabos).filter(c => c.estado !== 'gr').length + Object.values(k.a.estruturas).filter(e => e.estado === 'k' || e.estado === 'pu').length; b.textContent = pend ? pend + ' pend.' : ''; b.style.display = pend ? '' : 'none'; }
        if (st.elRes) st.elRes.textContent = textoResumoPoste();
        if (mapaAtualizar) mapaAtualizar();
        atualizarAchados();
    }

    function motivoParaLigar(ic, ie) {
        const c = cabos()[ic], info = infoCabo(ic);
        if (!info.vincula) return `O cabo ${info.codigo} não vincula (acompanha o cabo principal).`;
        if ((c.vinculoEstruturas || []).length + VR.estruturasFicticiasDe(c).length >= info.pontas)
            return `O cabo ${info.codigo} já tem as ${info.pontas} ponta(s); remova uma antes.`;
        const [ok, motivo] = podeLigar(ic, ie);
        return ok ? '' : motivo || 'Estrutura incompatível.';
    }

    function ligar(ic, ie) {
        const c = cabos()[ic], ja = (c.vinculoEstruturas || []).includes(ie);
        const motivo = ja ? '' : motivoParaLigar(ic, ie);
        if (motivo) { toast('✖ ' + motivo); return; }
        host.antes();
        if (ja) {
            c.vinculoEstruturas = c.vinculoEstruturas.filter(x => x !== ie);
            toast('Vínculo removido (Ctrl+Z desfaz).');
        } else {
            c.vinculoEstruturas = (c.vinculoEstruturas || []).concat([ie]);
            c.vinculoOrigem = 'manual';
            st.sel = null;
        }
        atualizar();
    }

    function clicar(tipo, i) {
        if (tipo === 'e' && infoEst(i).ocupa === false) { toast('Equipamento (' + infoEst(i).codigo + ') lê os cabos do poste: não recebe vínculo de cabo.'); return; }          // TASK-169
        if (st.sel && st.sel.t !== tipo) { const ic = tipo === 'c' ? i : st.sel.i, ie = tipo === 'e' ? i : st.sel.i; ligar(ic, ie); return; }
        st.sel = (st.sel && st.sel.t === tipo && st.sel.i === i) ? null : { t: tipo, i };
        if (painel) painel.pintar(true); else render();
    }

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
        atualizar();
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
        ficticiosDe(ie).forEach((cod, k) => itens.push({ t: `✖ Remover fictício ${cod}`, fn: () => { host.antes(); const l = ficticiosDe(ie).slice(); l.splice(k, 1); gravarFicticios(ie, l); atualizar(); } }));
        itens.push({ t: 'Desvincular todos os cabos desta estrutura', fn: () => { host.antes(); cabos().forEach(c => { if (c && c.vinculoEstruturas) c.vinculoEstruturas = c.vinculoEstruturas.filter(x => x !== ie); }); atualizar(); } });
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
        const itens = [{ hd: 'Cabo ' + (c.ativo || '') }, { t: 'Desvincular este cabo', fn: () => { host.antes(); c.vinculoEstruturas = []; atualizar(); } }, '-', { hd: 'Estrutura fictícia' }];
        const add = cod => { host.antes(); c.estruturasFicticias = fict.concat([cod.toUpperCase()]); atualizar(); };
        if (cheio) itens.push({ t: `O cabo já tem as ${info.pontas} ponta(s)`, off: true });
        else {
            const cands = candidatosEstruturasFicticias(ic), aceitas = cands.filter(x => x.ok);
            if (!aceitas.length) itens.push({ t: 'Nenhuma estrutura aceita por este cabo', off: true });
            aceitas.slice(0, 8).forEach(x => itens.push({ t: '➕ ' + x.cod, fn: () => add(x.cod) }));
            itens.push({ t: '✎ Outro código…', fn: () => { const cod = raiz.prompt('Código da estrutura fictícia (ex.: N1):'); if (!cod) return; const [ok, motivo] = VR.podeVincular(info, VR.infoEstrutura(cod, st.cfg)); if (ok) add(cod); else toast('✖ ' + motivo); } });
            const rec = cands.filter(x => !x.ok && x.motivo);
            if (rec.length) itens.push({ t: `(${rec.length} código(s) não aceito(s): ${rec.slice(0, 3).map(x => x.cod).join(', ')}…)`, off: true, title: rec.slice(0, 5).map(x => x.cod + ': ' + x.motivo).join('\n') });
        }
        fict.forEach((cod, k) => itens.push({ t: `✖ Remover estrutura fictícia ${cod}`, fn: () => { host.antes(); const l = fict.slice(); l.splice(k, 1); if (l.length) c.estruturasFicticias = l; else delete c.estruturasFicticias; atualizar(); } }));
        abrirMenu(e, itens);
    }

    const CORES = { k: '', pu: ' pu', gr: ' gr', rd: ' rd' };
    const TXT_C = { k: 'sem vínculo', pu: '◐ parcial', gr: '● completo', rd: '✖ inválido' };
    const COR_T = { gr: '#3fb950', pu: '#a371f7', rd: '#f85149', k: '#8b949e' };
    const faixaCabos = a => (a.maximo != null && a.maximo !== a.precisa ? `${a.precisa}–${a.maximo}` : String(a.precisa));

    // ── painel virtualizado (TASK-185): compartilhado pelas abas Vínculo e Cabos ⇄ Postes ──
    // Só as linhas visíveis (+ margem) existem no DOM; cada linha tem altura fixa, então posição e linhas de ligação saem por aritmética (sem medir o layout).
    const RH = 30, RA = 24, LW = 430, GUT = 140, LAB = 112, CW = 84, CG = 6;
    let painel = null;

    function criarPainel(corpo, def) {
        const box = el('div', { class: 'vl-box' }), inn = el('div', { class: 'vl-in' }), L = el('div'), R = el('div');
        const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.setAttribute('class', 'vl-liga');
        inn.appendChild(L); inn.appendChild(R); inn.appendChild(svg); box.appendChild(inn);
        corpo.appendChild(el('div', { class: 'vl-hd' }, [el('div', { style: `width:${LW + GUT}px`, text: def.tituloE }), el('div', { text: def.tituloD })]));
        corpo.appendChild(box);
        inn.style.setProperty('--lw', LW + 'px'); inn.style.setProperty('--rl', (LW + GUT) + 'px');
        const larg = LW + GUT + LAB + 8 * (CW + CG) + 40;
        inn.style.setProperty('--mw', larg + 'px'); svg.setAttribute('width', larg);
        let dados = null, jan = null, raf = 0, ro = null;
        function faixa(lista, y0, y1) {
            let lo = 0, hi = lista.length;
            while (lo < hi) { const m = (lo + hi) >> 1; if (lista[m].y + lista[m].h < y0) lo = m + 1; else hi = m; }
            let i1 = lo; while (i1 < lista.length && lista[i1].y <= y1) i1++;
            return [lo, i1];
        }
        function refazer() { dados = def.itens(); inn.style.height = (Math.max(dados.hE, dados.hD) + 8) + 'px'; jan = null; pintar(true); }
        function pintar(forcar) {
            if (!dados) return;
            const top = box.scrollTop, alt = box.clientHeight || 700, OV = 400;
            if (!forcar && jan && top >= jan[0] && top + alt <= jan[1]) return;
            const y0 = top - OV, y1 = top + alt + OV; jan = [y0, y1];
            const [a0, a1] = faixa(dados.esq, y0, y1), [b0, b1] = faixa(dados.dir, y0, y1);
            let hE = '', hD = '';
            for (let i = a0; i < a1; i++) hE += def.htmlE(dados.esq[i]);
            for (let i = b0; i < b1; i++) hD += def.htmlD(dados.dir[i]);
            L.innerHTML = hE; R.innerHTML = hD;
            const sy = Math.max(0, y0), alturaSvg = Math.max(1, Math.min(Math.max(dados.hE, dados.hD) + 8, y1) - sy);
            svg.style.top = sy + 'px'; svg.setAttribute('height', alturaSvg);          // o SVG cobre só a janela visível (não a lista inteira)
            let g = '';
            const curva = (x1, y1b, x2, y2) => { const mx = (x1 + x2) / 2; return `M${x1},${y1b} C${mx},${y1b} ${mx},${y2} ${x2},${y2}`; };
            dados.links.forEach(l => {
                const ie = dados.esq[l.e], id = dados.dir[l.d];
                if (!ie || !id) return;
                const yE = ie.y + (ie.h - 4) / 2, topo = l.slot > 0;
                const yD = topo ? id.y + 1 : id.y + (id.h - 4) / 2, xD = LW + GUT + (topo ? LAB + l.slot * (CW + CG) + CW / 2 : 0);
                if ((yE < y0 || yE > y1) && (yD < y0 || yD > y1)) return;
                const d = curva(LW, yE - sy, xD, yD - sy), cor = l.cor, cyE = yE - sy, cyD = yD - sy;
                g += `<path d="${d}" fill="none" stroke="${cor}" stroke-width="2" ${l.dash ? 'stroke-dasharray="5 4"' : ''}/><circle cx="${LW}" cy="${cyE}" r="3.5" fill="${cor}"/><circle cx="${xD}" cy="${cyD}" r="3.5" fill="${cor}"/>`;
                if (l.hit) g += `<path class="hit" d="${d}" ${l.hit}><title>${esc(l.motivo || 'clique para desfazer o vínculo')}</title></path>`;
            });
            svg.innerHTML = g;
        }
        box.addEventListener('scroll', () => { if (!raf) raf = raiz.requestAnimationFrame(() => { raf = 0; pintar(false); }); }, { passive: true });
        inn.addEventListener('click', ev => def.clique(ev));
        inn.addEventListener('contextmenu', ev => def.menu(ev));
        if (raiz.ResizeObserver) { ro = new raiz.ResizeObserver(() => pintar(true)); ro.observe(box); }
        return { refazer, pintar: f => pintar(f !== false), box, dados: () => dados, rolarPara(y, h) { if (y < box.scrollTop) box.scrollTop = Math.max(0, y - 8); else if (y + h > box.scrollTop + box.clientHeight) box.scrollTop = y + h - box.clientHeight + 8; }, destruir() { if (ro) ro.disconnect(); if (raf) raiz.cancelAnimationFrame(raf); } };
    }

    const metaDe = l => { const at = outros()[l].ativo || '', m = metaAtivo.get(at); if (m) return m; const n = { soPoste: host.linhaSoPoste(at), postes: host.postesDaLinha(at) }; metaAtivo.set(at, n); return n; };
    const postesDe = l => metaDe(l).postes;
    const avulsos = ic => { const k = cx || ctx(); return k.avul[ic] || (k.avul[ic] = host.avulsosDoCabo ? host.avulsosDoCabo(ic) : []); };
    const attrNum = (alvo, nome) => { const e = alvo.closest('[' + nome + ']'); return e ? Number(e.getAttribute(nome)) : null; };
    const opClasse = op => (op.indexOf('M') >= 0 ? ' m' : (op.indexOf('R') >= 0 ? ' r' : ''));

    function textoLinhaAvulso(av) {
        const completos = av.itens.filter(i => !i.completar), faltam = av.itens.filter(i => i.completar);
        return completos.map(i => `${i.qtd} × ${i.ativo}`).join(' · ') + (faltam.length ? (completos.length ? ' · ' : '') + faltam.map(i => `${i.qtd} × ${i.ativo} — complete o cabo existente`).join(' · ') : '');
    }
    function htmlAvulso(it) {
        const av = it.av, texto = textoLinhaAvulso(av), completos = av.itens.some(i => !i.completar);
        const estado = av.estado === 'adicionado' ? '<span class="ok">adicionado ✓</span>' : `<button class="vz-btn" data-av="${it.ic}|${esc(av.ficticia)}" ${completos ? '' : 'disabled'} title="adiciona estes itens à tabela Outros (você pode editar ou apagar)">+ Outros</button>`;
        return `<div class="vl-av" style="top:${it.y}px;height:${it.h - 3}px" title="${esc(texto)}"><span class="n">${esc(av.ficticia)} (fictícia) →</span><span class="t">${esc(texto)}</span>${av.estado === 'desatualizado' ? '<span class="v">avulso desatualizado</span>' : ''}${estado}</div>`;
    }

    // ── ABA: VÍNCULO ──
    function abaVinculo(corpo) {
        const filtros = [['todos', 'Todos'], ['pendentes', 'Só pendentes'], ['conflitos', 'Só conflitos']];
        const chips = filtros.map(([id, nome]) => el('span', { class: 'vz-chip btn' + (st.filtro === id ? ' on' : ''), text: nome, onclick: () => { st.filtro = id; chips.forEach((c, i) => c.classList.toggle('on', filtros[i][0] === id)); painel.refazer(); } }));
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-bottom:8px;align-items:center;flex-wrap:wrap' },
            chips.concat([el('span', { style: 'margin-left:auto;color:#8b949e;font-size:11px', text: 'Clique no cabo e na estrutura (ou o contrário) · botão direito: fictício · clique na linha: desfazer' })])));
        const ests = VR.estruturasDisponiveis(outros(), host.codigoEstrutura, host.codigosEstrutura);
        if (!ests.length) {                          // TASK-126
            corpo.appendChild(el('div', { style: 'color:#8b949e;padding:20px', text: 'Nenhuma estrutura/poste na tabela Outros ainda.' }));
            return;
        }
        if (!cabos().some(VR.ehCaboValido)) corpo.appendChild(el('div', { style: 'color:#8b949e;font-size:12px;padding:2px 2px 8px', text: 'Nenhum cabo na tabela — use o botão direito na estrutura para associar um cabo fictício.' }));
        painel = criarPainel(corpo, {
            tituloE: 'Cabos', tituloD: 'Estruturas — um poste por linha',
            itens() {
                const k = cx || ctx(), a = k.a, chave = 'v|' + st.filtro;
                if (k.it[chave]) return k.it[chave];
                const esq = [], dir = [], links = [], eDe = {}, dDe = {}, fict = [], C = cabos(), O = outros();
                let yE = 0, yD = 0;
                C.forEach((cb, ic) => {
                    if (!VR.ehCaboValido(cb)) return;
                    const r = a.cabos[ic];
                    if (st.filtro === 'pendentes' && r.estado === 'gr') return;
                    if (st.filtro === 'conflitos' && r.estado !== 'rd') return;
                    eDe[ic] = esq.length; esq.push({ t: 'c', ic, y: yE, h: RH }); yE += RH;
                    avulsos(ic).forEach(av => { esq.push({ t: 'a', ic, av, y: yE, h: RA }); yE += RA; });
                });
                const ests = VR.estruturasDisponiveis(O, host.codigoEstrutura, host.codigosEstrutura), grupos = [], porLinha = {};
                ests.forEach(ie => { const l = linhaDe(ie); if (!(l in porLinha)) { porLinha[l] = { l, ids: [] }; grupos.push(porLinha[l]); } porLinha[l].ids.push(ie); });
                grupos.forEach(g => {
                    const o = O[g.l], m = metaDe(g.l);
                    if (m.soPoste && !g.ids.some(ie => (k.ocup[ie] || []).length)) return;                                      // poste sem estrutura e sem cabo ligado: não aparece
                    const ids = g.ids.filter(ie => { const r = a.estruturas[ie]; return !(st.filtro === 'pendentes' && r.estado === 'gr') && !(st.filtro === 'conflitos' && r.estado !== 'rd'); });
                    if (!ids.length) return;
                    dir.push({ t: 'p', l: g.l, ids, label: m.postes.length && !m.soPoste ? m.postes.join(' ') : (o.ativo || ''), y: yD, h: RH });
                    ids.forEach((ie, kk) => { dDe[ie] = [dir.length - 1, kk]; });
                    yD += RH;
                    g.ids.forEach(ie => ficticiosDe(ie).forEach((cod, kk) => fict.push({ ie, cod, k: kk })));
                });
                fict.forEach(f => { esq.push({ t: 'f', ie: f.ie, cod: f.cod, k: f.k, y: yE, h: RH }); yE += RH; if (dDe[f.ie]) links.push({ e: esq.length - 1, d: dDe[f.ie][0], slot: dDe[f.ie][1], cor: '#a371f7', dash: true }); });
                a.links.forEach(l => {
                    const e = eDe[l.cabo], d = dDe[l.estrutura];
                    if (e === undefined || !d) return;
                    const est = a.cabos[l.cabo].estado;
                    links.push({ e, d: d[0], slot: d[1], cor: !l.ok || est === 'rd' ? '#f85149' : (est === 'gr' ? '#3fb950' : '#a371f7'), dash: !l.ok, motivo: l.motivo, hit: `data-lc="${l.cabo}" data-le="${l.estrutura}"` });
                });
                return (k.it[chave] = { esq, dir, links, hE: yE, hD: yD });
            },
            htmlE(it) {
                const k = cx, a = k.a;
                if (it.t === 'a') return htmlAvulso(it);
                if (it.t === 'f') return `<div class="vl-cab fic" data-f="${it.ie}|${it.k}" style="top:${it.y}px;height:${it.h - 4}px"><span class="vz-op">I</span><span class="t">${esc(it.cod)} <span style="font-size:10px;font-weight:800;color:#a371f7">FICTÍCIO</span></span><span class="s">sem comprimento</span><span class="r" style="color:#a371f7" data-rmf="${it.ie}|${it.k}">✎ remover</span></div>`;
                const c = cabos()[it.ic], r = a.cabos[it.ic], selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selE = st.sel && st.sel.t === 'e' ? st.sel.i : null;
                let cls = 'vl-cab' + CORES[r.estado] + (selC === it.ic ? ' sel' : ''), tip = '';
                if (selE !== null) { const [ok, motivo] = podeLigar(it.ic, selE), ja = (c.vinculoEstruturas || []).includes(selE); if (ok || ja) cls += ' glow'; else { cls += ' vl-dim'; tip = motivo; } }
                const op = trim(c.operacao).toUpperCase();
                return `<div class="${cls}" data-c="${it.ic}" style="top:${it.y}px;height:${it.h - 4}px" title="${esc(tip || c.ativo || '')}"><span class="vz-op${opClasse(op)}">${esc(op || 'I')}</span><span class="t">${esc(c.ativo || '')}</span>`
                    + `<span class="s">${esc(r.codigo)}${r.valor_cabo ? ' · tc=' + esc(r.valor_cabo) : ''} · ${r.tem}/${r.precisa}${r.ficticias ? ' · +' + r.ficticias + ' fict.' : ''}</span><span class="r" style="color:${COR_T[r.estado]}">${r.vincula ? TXT_C[r.estado] : 'não vincula'}</span></div>`;
            },
            htmlD(it) {
                const k = cx, a = k.a, o = outros()[it.l], selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selE = st.sel && st.sel.t === 'e' ? st.sel.i : null;
                let h = `<div class="vl-pst" data-p="${it.l}" style="top:${it.y}px;height:${it.h - 4}px"><span class="lab" title="${esc(o.ativo || '')}">${esc(it.label)}</span>`;
                it.ids.forEach((ie, kk) => {
                    const r = a.estruturas[ie], eqp = r.ocupa === false;                                                         // TASK-169: equipamento (lê do poste): cinza, sem alvo de cabo
                    let cls = 'vl-chip' + (eqp ? ' eqp' : CORES[r.estado]) + (selE === ie ? ' sel' : ''), tip = '';
                    if (selC !== null && !eqp) { const [ok, motivo] = podeLigar(selC, ie), ja = (cabos()[selC].vinculoEstruturas || []).includes(ie); if (ok || ja) cls += ' glow'; else { cls += ' vl-dim'; tip = motivo; } }
                    const inf = infoEst(ie);
                    const sub = eqp ? (r.erros.length ? '✖ ' + r.erros[0] : `lê do poste ${(r.le_do_poste && r.le_do_poste.poste) || ''}: ${r.le_do_poste && r.le_do_poste.cabos.length ? r.le_do_poste.cabos.join(', ') : 'nenhum cabo' + (r.avisos.length ? ' (' + r.avisos[0] + ')' : '')}`)
                        : r.erros.length ? '✖ ' + r.erros[0] : (inf.regra === null ? 'sem regra na Tabela de Estruturas' : (inf.lados === 'IGUAIS' ? `exige ${inf.qtd || ''} cabos iguais` : (inf.amarracao ? 'amarração' : `${inf.qtd || ''} cabo(s)`)));
                    const tit = (tip || (o.ativo || '')) + ' — ' + sub + (r.valor_cabo ? ' · tipocabo=' + r.valor_cabo : '') + (r.ficticios ? ' · ' + r.ficticios + ' fictício' : '');
                    h += `<div class="${cls}" data-e="${ie}" style="left:${LAB + kk * (CW + CG)}px" title="${esc(tit)}"><b>${esc(r.codigo || '?')}</b>${eqp ? '' : `<i>${r.tem}${r.precisa ? '/' + faixaCabos(r) : ''}</i>`}</div>`;
                });
                return h + '</div>';
            },
            clique(ev) {
                const t = ev.target, hit = t.closest && t.closest('path.hit');
                if (hit) { ligar(Number(hit.getAttribute('data-lc')), Number(hit.getAttribute('data-le'))); return; }
                const bt = t.closest('[data-av]'); if (bt) { ev.stopPropagation(); const [ic, fic] = bt.getAttribute('data-av').split('|'); host.adicionarAvulsos(Number(ic), fic); atualizar(); return; }
                const rm = t.closest('[data-rmf]'); if (rm) { const [ie, kk] = rm.getAttribute('data-rmf').split('|'); host.antes(); const l = ficticiosDe(Number(ie)).slice(); l.splice(Number(kk), 1); gravarFicticios(Number(ie), l); atualizar(); return; }
                const c = attrNum(t, 'data-c'); if (c !== null) { clicar('c', c); return; }
                const e = attrNum(t, 'data-e'); if (e !== null) clicar('e', e);
            },
            menu(ev) {
                const c = attrNum(ev.target, 'data-c'); if (c !== null) { ctx(); menuCabo(ev, c); return; }
                const e = attrNum(ev.target, 'data-e'); if (e !== null) { ctx(); if (infoEst(e).ocupa !== false) menuEstrutura(ev, e); else ev.preventDefault(); }
            }
        });
        painel.refazer();
    }

    // ── ABA: CABOS ⇄ POSTES (TASK-185) ──
    // O usuário só diz a que poste(s) cada cabo chega; o algoritmo por poste (o mesmo da obra do plugin) escolhe a estrutura de cada poste que recebe o cabo.
    function atribuirPoste(ic, l) {
        ctx();
        const cb = cabos()[ic], info = infoCabo(ic), at = VR.postesAtribuidos(cb), ja = at.includes(l);
        if (!info.vincula) { toast(`O cabo ${info.codigo} não vincula (acompanha o cabo principal).`); return; }
        if (!ja && at.length >= info.pontas) { toast(`O cabo ${info.codigo} já chega a ${info.pontas} poste(s); remova um antes.`); return; }
        host.antes();
        const nova = ja ? at.filter(x => x !== l) : at.concat([l]);
        if (nova.length) cb.vinculoPostes = nova;
        else { delete cb.vinculoPostes; if (cb.vinculoOrigem === 'poste') { cb.vinculoEstruturas = []; delete cb.vinculoOrigem; } }
        st.sel = null;
        if (st.auto) aplicarAlgoritmoPoste();
        atualizar();
    }
    function aplicarAlgoritmoPoste() {
        const r = VR.vincularPorPoste(cabos(), outros(), st.cfg, host.prefixoCabo, host.codigoEstrutura, host.codigosEstrutura, undefined, true);
        st.resultadoPoste = r;
        return r;
    }
    function textoResumoPoste() {
        const s = st.resultadoPoste && st.resultadoPoste.resumo;
        return s ? `${s.completos} cabo(s) completo(s) · ${s.parciais} parcial(is) · ${s.estruturas_incompletas} estrutura(s) incompleta(s)` : '';
    }


    const ROTULO_ACHADO = { sobra: 'sobra', incompleta: 'incompleta', ponta_solta: 'ponta solta', sem_geometria: 'sem geometria', escolha: 'escolha' };
    /** Achados do último «Vincular estruturas»: o que ficou incompleto ou sobrando, para o usuário saber onde mexer. */
    function atualizarAchados() {
        const box = st.elAch;
        if (!box) return;
        const ach = (st.resultadoPoste && st.resultadoPoste.achados) || [];
        box.style.display = ach.length ? '' : 'none';
        box.innerHTML = ach.length ? `<b>${ach.length} aviso(s) do algoritmo por poste</b>` + ach.slice(0, 200).map(a => `<div class="a ${esc(a.tipo)}">${esc(ROTULO_ACHADO[a.tipo] || a.tipo)} — ${esc(a.mensagem)}</div>`).join('') : '';
    }
    /** Teclado da aba Cabos ⇄ Postes: ↑↓ escolhem o cabo; → vai para os postes (↑↓ escolhem, Enter liga/desliga); ← volta aos cabos. */
    function tecladoCabosPostes(e) {
        if (st.aba !== 'cabopostes' || !painel || menuAberto || e.ctrlKey || e.metaKey || e.altKey) return false;
        const t = e.target && e.target.tagName;
        if (t === 'INPUT' || t === 'SELECT' || t === 'TEXTAREA') return false;
        const d = painel.dados();
        if (!d || !['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Enter'].includes(e.key)) return false;
        e.preventDefault();
        const passo = e.key === 'ArrowUp' ? -1 : 1;
        if (e.key === 'ArrowRight') { st.kbModo = 'p'; if (st.kbP === null || !d.dir.some(x => x.l === st.kbP)) st.kbP = d.dir.length ? d.dir[0].l : null; painel.pintar(true); return true; }
        if (e.key === 'ArrowLeft') { st.kbModo = 'c'; painel.pintar(true); return true; }
        if (st.kbModo === 'c') {
            if (e.key === 'Enter' || !d.esq.length) return true;
            const cur = st.sel && st.sel.t === 'c' ? d.esq.findIndex(x => x.ic === st.sel.i) : -1;
            const nx = d.esq[Math.max(0, Math.min(d.esq.length - 1, cur < 0 ? 0 : cur + passo))];
            st.sel = { t: 'c', i: nx.ic }; painel.rolarPara(nx.y, nx.h); painel.pintar(true);
            return true;
        }
        if (!d.dir.length) return true;
        const i = d.dir.findIndex(x => x.l === st.kbP);
        if (e.key === 'Enter') { if (st.sel && st.sel.t === 'c' && st.kbP !== null) { const ic = st.sel.i; atribuirPoste(ic, st.kbP); st.sel = { t: 'c', i: ic }; painel.pintar(true); } else toast('Escolha um cabo com ↑↓ antes (← volta aos cabos).'); return true; }
        const nx = d.dir[Math.max(0, Math.min(d.dir.length - 1, i < 0 ? 0 : i + passo))];
        st.kbP = nx.l; painel.rolarPara(nx.y, nx.h); painel.pintar(true);
        return true;
    }

    function abaCabosPostes(corpo) {
        const filtros = [['todos', 'Todos'], ['sem', 'Sem poste'], ['pend', 'Pendentes']];
        if (!filtros.some(f => f[0] === st.filtroCP)) st.filtroCP = 'todos';
        const res = el('span', { style: 'color:#8b949e;font-size:11px', text: textoResumoPoste() });
        const chips = filtros.map(([id, nome]) => el('span', { class: 'vz-chip btn' + (st.filtroCP === id ? ' on' : ''), text: nome, onclick: () => { st.filtroCP = id; chips.forEach((c, i) => c.classList.toggle('on', filtros[i][0] === id)); painel.refazer(); } }));
        const auto = el('input', { type: 'checkbox', onchange: e => { st.auto = e.target.checked; if (st.auto) { host.antes(); aplicarAlgoritmoPoste(); res.textContent = textoResumoPoste(); atualizar(); } } });
        auto.checked = st.auto;
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-bottom:8px;align-items:center;flex-wrap:wrap' }, chips.concat([
            el('label', { style: 'display:flex;gap:5px;align-items:center;color:#c9d1d9;font-size:12px;cursor:pointer', title: 'Cada vez que você liga um cabo a um poste, o algoritmo do poste já escolhe as estruturas.' }, [auto, 'Vincular as estruturas automaticamente']),
            el('button', { class: 'vz-btn b', text: '🏗️ Vincular estruturas agora', title: 'Roda o algoritmo por poste (Ctrl+Z desfaz)', onclick: () => { host.antes(); const r = aplicarAlgoritmoPoste(), s = r.resumo; res.textContent = textoResumoPoste(); toast(`Por poste: ${s.completos} completo(s), ${s.parciais} parcial(is), ${s.sem_vinculo} sem vínculo (Ctrl+Z desfaz).`); atualizar(); } }),
            res, el('span', { style: 'margin-left:auto;color:#8b949e;font-size:11px', text: 'Clique no cabo e no poste (ou o contrário) · ↑↓ cabo, → postes, Enter liga · clique na linha: desfazer · botão direito: limpar' })])));
        st.elRes = res;
        st.elAch = el('div', { class: 'vl-ach' });
        painel = criarPainel(corpo, {
            tituloE: 'Cabos', tituloD: 'Postes — a estrutura de cada um é escolhida pelo algoritmo',
            itens() {
                const k = cx || ctx(), a = k.a, chave = 'p|' + st.filtroCP;
                if (k.it[chave]) return k.it[chave];
                const esq = [], dir = [], links = [], eDe = {}, dDe = {}, C = cabos(), O = outros(), cont = {};
                C.forEach(cb => { if (VR.ehCaboValido(cb)) VR.postesAtribuidos(cb).forEach(l => { cont[l] = (cont[l] || 0) + 1; }); });
                let yE = 0, yD = 0;
                C.forEach((cb, ic) => {
                    if (!VR.ehCaboValido(cb)) return;
                    const r = a.cabos[ic], at = VR.postesAtribuidos(cb);
                    if (st.filtroCP === 'sem' && at.length) return;
                    if (st.filtroCP === 'pend' && r.estado === 'gr') return;
                    eDe[ic] = esq.length; esq.push({ t: 'c', ic, y: yE, h: RH }); yE += RH;
                });
                const ests = VR.estruturasDisponiveis(O, host.codigoEstrutura, host.codigosEstrutura), grupos = [], porLinha = {};
                ests.forEach(ie => { const l = linhaDe(ie); if (!(l in porLinha)) { porLinha[l] = { l, ids: [] }; grupos.push(porLinha[l]); } porLinha[l].ids.push(ie); });
                grupos.forEach(g => {
                    const mm = metaDe(g.l), m = mm.postes;
                    if (mm.soPoste) return;                                                              // linha só com poste (sem estrutura): não há o que o algoritmo ligar
                    dir.push({ t: 'p', l: g.l, ids: g.ids, n: cont[g.l] || 0, label: m.length ? m.join(' ') : (O[g.l].ativo || ''), y: yD, h: RH }); dDe[g.l] = dir.length - 1; yD += RH;
                });
                C.forEach((cb, ic) => {
                    if (eDe[ic] === undefined) return;
                    const est = a.cabos[ic].estado;
                    VR.postesAtribuidos(cb).forEach(l => { if (dDe[l] !== undefined) links.push({ e: eDe[ic], d: dDe[l], slot: 0, cor: COR_T[est] === COR_T.k ? '#58a6ff' : COR_T[est], motivo: 'clique para tirar o cabo deste poste', hit: `data-lc="${ic}" data-lp="${l}"` }); });
                });
                return (k.it[chave] = { esq, dir, links, hE: yE, hD: yD });
            },
            htmlE(it) {
                const k = cx, c = cabos()[it.ic], r = k.a.cabos[it.ic], at = VR.postesAtribuidos(c), info = infoCabo(it.ic);
                const selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selP = st.sel && st.sel.t === 'p' ? st.sel.i : null;
                let cls = 'vl-cab' + CORES[r.estado] + (selC === it.ic ? ' sel' : ''), tip = '';
                if (selP !== null) { if (at.includes(selP) || (info.vincula && at.length < info.pontas)) cls += ' glow'; else { cls += ' vl-dim'; tip = info.vincula ? 'o cabo já chega ao máximo de postes' : 'cabo que não vincula'; } }
                const op = trim(c.operacao).toUpperCase(), nomes = at.map(l => (postesDe(l).length ? postesDe(l).join(' ') : '#' + (l + 1))).join(' · ');
                const es = (c.vinculoEstruturas || []).map(ie => codigoDe(ie)).filter(Boolean).join(' · ');
                return `<div class="${cls}" data-c="${it.ic}" style="top:${it.y}px;height:${it.h - 4}px" title="${esc(tip || c.ativo || '')}"><span class="vz-op${opClasse(op)}">${esc(op || 'I')}</span><span class="t">${esc(c.ativo || '')}</span>`
                    + `<span class="pp">${nomes ? '⇒ ' + esc(nomes) : (info.vincula ? '—' : 'não vincula')}</span><span class="r" style="color:${COR_T[r.estado]}">${es ? esc(es) : ''}</span></div>`;
            },
            htmlD(it) {
                const k = cx, a = k.a, o = outros()[it.l], selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selP = st.sel && st.sel.t === 'p' ? st.sel.i : null;
                let cls = 'vl-pst';
                if (selC !== null) { const cb = cabos()[selC], at = VR.postesAtribuidos(cb), info = infoCabo(selC); if (at.includes(it.l) || (info.vincula && at.length < info.pontas)) cls += ' glow'; else cls += ' vl-dim'; }
                let h = `<div class="${cls}${selP === it.l || (st.kbModo === 'p' && st.kbP === it.l) ? ' sel' : ''}" data-p="${it.l}" style="top:${it.y}px;height:${it.h - 4}px;cursor:pointer" title="${esc(o.ativo || '')}"><span class="lab">${esc(it.label)}</span>`;
                it.ids.forEach((ie, kk) => { const r = a.estruturas[ie], eqp = r.ocupa === false; h += `<div class="vl-chip${eqp ? ' eqp' : CORES[r.estado]}" style="left:${LAB + kk * (CW + CG)}px;cursor:inherit" title="${esc(r.codigo || '')}"><b>${esc(r.codigo || '?')}</b>${eqp ? '' : `<i>${r.tem}${r.precisa ? '/' + faixaCabos(r) : ''}</i>`}</div>`; });
                return h + `<span class="pp" style="margin-left:auto;padding-right:8px;font-size:10.5px;color:${it.n ? '#58a6ff' : '#6e7681'}">${it.n} cabo(s)</span></div>`;
            },
            clique(ev) {
                const t = ev.target, hit = t.closest && t.closest('path.hit');
                if (hit) { atribuirPoste(Number(hit.getAttribute('data-lc')), Number(hit.getAttribute('data-lp'))); return; }
                const c = attrNum(t, 'data-c'), p = c === null ? attrNum(t, 'data-p') : null, tipo = c !== null ? 'c' : (p !== null ? 'p' : null);
                if (!tipo) return;
                const i = c !== null ? c : p;
                if (st.sel && st.sel.t !== tipo) { const ic = tipo === 'c' ? i : st.sel.i, l = tipo === 'p' ? i : st.sel.i; atribuirPoste(ic, l); return; }
                st.sel = (st.sel && st.sel.t === tipo && st.sel.i === i) ? null : { t: tipo, i };
                painel.pintar(true);
            },
            menu(ev) {
                const c = attrNum(ev.target, 'data-c'); if (c === null) return;
                ctx(); const cb = cabos()[c];
                abrirMenu(ev, [{ hd: 'Cabo ' + (cb.ativo || '') }, VR.postesAtribuidos(cb).length ? { t: 'Tirar este cabo dos postes', fn: () => { host.antes(); delete cb.vinculoPostes; if (cb.vinculoOrigem === 'poste') { cb.vinculoEstruturas = []; delete cb.vinculoOrigem; } if (st.auto) aplicarAlgoritmoPoste(); atualizar(); } } : { t: 'Cabo sem poste atribuído', off: true }]);
            }
        });
        painel.refazer();
        corpo.appendChild(st.elAch); atualizarAchados();
    }

    // ── tabelas editáveis (Cabos / Estruturas) ──
    function celula(r, campo, tipo, extra) {
        const v = r[campo];
        if (tipo === 'bool') return el('input', { type: 'checkbox', ...((v === undefined ? !!extra : v) ? { checked: 'checked' } : {}), onchange: e => { r[campo] = e.target.checked; marcarSujo(); } });
        if (tipo === 'select') return el('select', { onchange: e => { r[campo] = e.target.value; marcarSujo(); } }, extra.map(o => el('option', { value: o, text: o, ...(String(v == null ? extra[0] : v) === o ? { selected: 'selected' } : {}) })));
        if (tipo === 'json') return el('input', { type: 'text', value: v === undefined || v === null || (Array.isArray(v) && !v.length) ? '' : JSON.stringify(v), placeholder: extra || '', style: 'min-width:150px',
            onchange: e => { const t = e.target.value.trim(); if (!t) { delete r[campo]; marcarSujo(); return; } try { r[campo] = JSON.parse(t); marcarSujo(); } catch (err) { raiz.alert('JSON inválido em ' + campo + ': ' + err.message); e.target.value = v === undefined ? '' : JSON.stringify(v); } } });          // TASK-170
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
            tbody.appendChild(el('tr', {}, [el('td', { text: String(i + 1), style: 'color:#8b949e;text-align:center' })].concat(colunas.map(c => el('td', { style: (c.w ? 'width:' + c.w : '') + (c.cinza && c.cinza(r) ? ';opacity:.4' : '') }, [celula(r, c.campo, c.tipo, c.extra)])))
                .concat([el('td', { style: 'white-space:nowrap' }, [el('button', { class: 'mv', text: '▲', title: 'a primeira linha que casa vale', onclick: () => { if (i > 0) { lista.splice(i - 1, 0, lista.splice(i, 1)[0]); marcarSujo(); render(); } } }),
                    el('button', { class: 'mv', text: '▼', onclick: () => { if (i < lista.length - 1) { lista.splice(i + 1, 0, lista.splice(i, 1)[0]); marcarSujo(); render(); } } }), el('button', { class: 'x', text: '✖', title: 'Excluir', onclick: () => { lista.splice(i, 1); marcarSujo(); render(); } })])])));
        });
        tb.appendChild(tbody); sc.appendChild(tb); corpo.appendChild(sc);
        const saidaTeste = el('span', { style: 'color:#8b949e;flex:1' });
        const campo = el('input', { type: 'text', style: 'width:280px;background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace', placeholder: nome === 'dicionario_cabos' ? 'Testar um texto: 3xA35P(7,94)-13,8kV 45 m' : nome === 'cabos' ? 'Testar um cabo: 3xA35P(7,94)-13,8kV' : 'Testar uma estrutura: 9200D, N1', value: st.teste[nome] || '',
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

    /** TASK-173: dicionário de cabos do projeto (o que o cabo É: família, valor, seção, nível, fases, apelidos e neutro). A Tabela de Cabos continua dizendo em quais estruturas ele entra. */
    function abaDicionario(corpo) {
        if (!Array.isArray(st.cfg.dicionario_cabos)) st.cfg.dicionario_cabos = [];
        const DC = raiz.DicionarioCabos, AC = raiz.AnotacaoCabos;
        tabelaEditavel(corpo, st.cfg.dicionario_cabos, [
            { campo: 'padrao', rot: 'Padrão do código', tipo: 'text', extra: '^A(\\d+)P$', dica: 'regex contra o código do cabo principal (sem Nx, neutro e tensão); a primeira linha que casa vale' }, { campo: 'familia', rot: 'Família', tipo: 'text', w: '130px' },
            { campo: 'valor_cabo', rot: 'Valor de cabo', tipo: 'text', w: '80px', dica: 'vale quando a Tabela de Cabos não tem valor para o cabo' }, { campo: 'secao', rot: 'Seção (mm²)', tipo: 'text', w: '80px', extra: '$1 ou 50', dica: '$1 = grupo da regex; vazio = regra fixa do código' },
            { campo: 'awg', rot: 'Tabela AWG', tipo: 'json', w: '130px', extra: '{"4":21.2}', dica: 'seção por texto do grupo (AWG)' },
            { campo: 'nivel', rot: 'Nível', tipo: 'select', extra: ['', 'MT', 'BT'], w: '65px', dica: 'vale quando o texto não traz tensão (sem isso: BT)' }, { campo: 'fases_padrao', rot: 'Fases', tipo: 'number', w: '55px', dica: 'fases quando o texto não traz Nx' },
            { campo: 'apelidos', rot: 'Apelidos', tipo: 'json', extra: '["^PB\\\\d+$", {"padrao":"^(\\\\d+)$","codigo":"C$1"}]', dica: 'textos curtos que viram o código' }, { campo: 'neutro', rot: 'Neutro', tipo: 'text', w: '100px', extra: 'CAZ{valor}', dica: 'como montar o neutro do parêntese com vírgula' },
            { campo: 'notacao_solta', rot: 'Código solto', tipo: 'bool', w: '70px', dica: 'o código sozinho (ALU-50 30 m) já conta como notação do LABEL' }, { campo: 'ativa', rot: 'Ativa', tipo: 'bool', w: '50px' }, { campo: 'obs', rot: 'Observação', tipo: 'text' }],
            () => ({ padrao: '', familia: '', valor_cabo: '', secao: '', nivel: '', ativa: true, obs: '' }),
            txt => {
                if (!DC || !AC) return 'dicionário indisponível';
                const dic = st.cfg.dicionario_cabos, r = AC.interpretar(txt, dic), p = r.itens.find(i => !i.neutro) || r.itens[0];
                if (!p) return 'nada reconhecido' + (r.erros.length ? ' — ' + esc(r.erros.join('; ')) : '');
                const [li] = DC.linha(dic, p.ativo), neutro = r.itens.find(i => i.neutro), tens = r.tensao ? Number(r.tensao.replace(',', '.')) : null;
                const nivel = tens !== null ? (tens > 1 ? 'MT' : 'BT') : (DC.nivelPadrao(dic, p.ativo) || 'BT'), sec = DC.secao(dic, p.ativo);
                return `código <b style="color:#79c0ff">${esc(p.ativo)}</b> · ${p.fases} fase(s)${neutro ? ' · neutro <b>' + esc(neutro.ativo) + '</b>' : ''} · ${nivel}${sec !== null ? ' · ' + sec + ' mm²' : ''}${DC.valorCabo(dic, p.ativo) ? ' · valor <b>' + esc(DC.valorCabo(dic, p.ativo)) + '</b>' : ''} · ${li === null ? 'nenhuma linha do dicionário casa' : 'linha #' + (li + 1)}`;
            },
            '<b>O que esta tabela diz:</b> o que cada cabo <b>é</b> (família, valor de cabo, seção para o "maior primeiro", nível e fases quando o texto não traz, apelidos e neutro). Vazia = regras de Pernambuco. A <b>Tabela de Cabos</b> segue dizendo em quais estruturas ele entra. A primeira linha que casa vale.',
            'dicionario_cabos');
    }

    /** TASK-174: convenção de postes do projeto (o que é poste na linha de Outros: tipos, formatos de nome, notação compacta e faixas). Sem tabela, vale a de Pernambuco. */
    function abaPostes(corpo) {
        const CP = raiz.ConvencaoPostes;
        if (!CP) { corpo.appendChild(el('div', { class: 'vz-nota', text: 'Convenção de postes indisponível nesta tela.' })); return; }
        const caixa = 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace';
        corpo.appendChild(el('div', { class: 'vz-nota', html: '<b>O que esta tela diz:</b> o que é <b>poste</b> na linha de Outros — os <b>tipos</b> (código e letra da notação compacta), as <b>formas de escrever</b> o poste (a primeira é o nome oficial) e as faixas de altura e de dígitos do esforço usadas para separar <code>11600D</code> em altura e esforço. Use <code>{tipo}</code>, <code>{sufixo}</code>, <code>{altura}</code> e <code>{esforco}</code>. <b>Sem tabela, vale a de Pernambuco.</b>' }));
        if (!st.cfg.postes) {
            corpo.appendChild(el('div', { style: 'margin:10px 0' }, [el('span', { style: 'color:#8b949e', text: 'Este projeto usa a convenção de Pernambuco (DT/RC/CV/FV, 11600D). ' }),
                el('button', { class: 'vz-btn', text: 'Personalizar a partir do padrão', onclick: () => { st.cfg.postes = JSON.parse(JSON.stringify(CP.PADRAO)); marcarSujo(); render(); } })]));
        } else {
            const c = st.cfg.postes;
            if (!Array.isArray(c.tipos)) c.tipos = [];
            const sc = el('div', { class: 'vz-scroll' }), tb = el('table', { class: 'vz-tb' });
            tb.appendChild(el('thead', {}, [el('tr', {}, ['Código', 'Letra (compacto)', 'Nome', ''].map(x => el('th', { text: x })))]));
            const tbody = el('tbody');
            c.tipos.forEach((tp, i) => tbody.appendChild(el('tr', {}, [celula(tp, 'codigo', 'text', 'DT'), celula(tp, 'sufixo', 'text', 'D'), celula(tp, 'nome', 'text', 'duplo T')].map(x => el('td', {}, [x]))
                .concat([el('td', {}, [el('button', { class: 'x', text: '✖', title: 'Excluir o tipo', onclick: () => { c.tipos.splice(i, 1); marcarSujo(); render(); } })])]))));
            tb.appendChild(tbody); sc.appendChild(tb); corpo.appendChild(sc);
            const ln = (rot, ...x) => el('div', { style: 'display:flex;gap:8px;align-items:center;margin-top:8px' }, [el('span', { style: 'min-width:170px;color:#8b949e', text: rot })].concat(x));
            const faixa = (campo, rot) => ln(rot, ...[0, 1].map(k => el('input', { type: 'number', min: '0', style: 'width:70px;' + caixa, value: Array.isArray(c[campo]) ? c[campo][k] : '',
                onchange: e => { const v = Array.isArray(c[campo]) ? c[campo].slice() : [0, 0]; v[k] = parseInt(e.target.value, 10); c[campo] = v; marcarSujo(); } })));
            corpo.appendChild(el('div', { style: 'margin-top:8px' }, [el('button', { class: 'vz-btn', text: '＋ Tipo', onclick: () => { c.tipos.push({ codigo: '', sufixo: '', nome: '' }); marcarSujo(); render(); } })]));
            corpo.appendChild(ln('Formas de escrever (uma por linha)', el('textarea', { rows: '3', style: 'flex:1;' + caixa, text: (c.formatos || []).join('\n'), placeholder: '{tipo}{altura}/{esforco}',
                onchange: e => { c.formatos = e.target.value.split('\n').map(trim).filter(Boolean); marcarSujo(); } })));
            corpo.appendChild(ln('Notação compacta', el('input', { type: 'text', style: 'width:260px;' + caixa, value: c.compacto || '', placeholder: '{altura}{esforco}{sufixo}', onchange: e => { c.compacto = e.target.value; marcarSujo(); } })));
            corpo.appendChild(faixa('altura', 'Altura (m): mínimo · máximo'));
            corpo.appendChild(faixa('esforco_digitos', 'Dígitos do esforço: mín · máx'));
            const erros = CP.validar(c);
            if (erros.length) corpo.appendChild(el('div', { style: 'color:#f85149;margin-top:8px', html: erros.map(esc).join('<br>') }));
            corpo.appendChild(el('div', { style: 'margin-top:10px' }, [el('button', { class: 'vz-btn', text: 'Voltar ao padrão de Pernambuco', onclick: () => { if (raiz.confirm('Remover a convenção deste projeto e voltar ao padrão de Pernambuco?')) { delete st.cfg.postes; marcarSujo(); render(); } } })]));
        }
        const saida = el('span', { style: 'color:#8b949e;flex:1' });
        const mostrar = v => {
            if (!v) { saida.innerHTML = ''; return; }
            const r = CP.testar(v, st.cfg.postes || undefined);
            saida.innerHTML = r.poste ? `<b style="color:#3fb950">poste</b> · tipo ${esc(r.tipo)} · nome <b style="color:#79c0ff">${esc(r.ativo)}</b> · compacto <b>${esc(r.compacto)}</b>` : (r.erro ? `<b style="color:#f85149">${esc(r.erro)}</b>` : 'não é poste (entra como estrutura)');
        };
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:12px;align-items:center' }, [el('input', { type: 'text', style: 'width:280px;' + caixa, placeholder: 'Testar um nome: DT11/600 ou 11600D', value: st.teste.postes || '',
            oninput: e => { st.teste.postes = e.target.value; mostrar(e.target.value); } }), saida]));
        mostrar(st.teste.postes);
    }

    /** TASK-175: leitura do desenho (DXF) do projeto — layers e cores que dizem o que é rede MT/BT. Mesmos campos do plugin do nanoCAD (/api/plugin/leitura); grava à parte da vinculação. */
    function abaDesenho(corpo) {
        const caixa = 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace';
        const url = '/api/plugin/leitura?projeto=' + encodeURIComponent(host.projeto());
        if (!st.desenho) {
            st.desenho = { cfg: null, erro: '', msg: '' };
            fetch(url).then(r => r.ok ? r.json() : Promise.reject(new Error('Projeto sem cadastro (HTTP ' + r.status + ')'))).then(d => { st.desenho.cfg = d.leitura || {}; render(); })
                .catch(e => { st.desenho.erro = String(e.message || e); st.desenho.cfg = {}; render(); });
        }
        const d = st.desenho, c = d.cfg;
        corpo.appendChild(el('div', { class: 'vz-nota', html: '<b>O que esta tela diz:</b> quais <b>layers</b> e <b>cores</b> do desenho (DXF) são rede <b>MT</b> e <b>BT</b>. Os mesmos campos do plugin do nanoCAD. Layer: o nome <b>contém</b> o trecho; <code>=NOME</code> = nome exato; <code>MT_*</code> = curinga. Cor: número ACI (1 = vermelho), <code>#RRGGBB</code> ou faixa. <b>Sem configuração, vale Pernambuco</b> (REDE_MT/REDE_BT, vermelho = MT, verde = BT).' }));
        if (c === null) { corpo.appendChild(el('div', { style: 'color:#8b949e;margin-top:10px', text: 'Carregando…' })); return; }
        if (d.erro) corpo.appendChild(el('div', { style: 'color:#d29922;margin-top:8px', text: 'Aviso: ' + d.erro }));
        const lista = k => el('input', { type: 'text', style: 'flex:1;' + caixa, value: (c[k] || []).join(', '), onchange: e => { const l = e.target.value.split(',').map(trim).filter(Boolean); if (l.length) c[k] = l; else delete c[k]; } });
        const cor = k => el('input', { type: 'text', style: 'width:300px;' + caixa, value: c[k] === undefined ? '' : (typeof c[k] === 'string' ? c[k] : JSON.stringify(c[k])), placeholder: '1 · #FF0000 · {"r":[170,255],"g":[0,100],"b":[0,100]}',
            onchange: e => { const v = e.target.value.trim(); if (!v) { delete c[k]; return; } try { c[k] = /^[#\d]/.test(v) && !/^\d+$/.test(v) ? v : JSON.parse(v); } catch (x) { raiz.alert('Cor inválida: ' + x.message); } } });
        const ln = (rot, x) => el('div', { style: 'display:flex;gap:8px;align-items:center;margin-top:8px' }, [el('span', { style: 'min-width:180px;color:#8b949e', text: rot }), x]);
        corpo.appendChild(ln('Layers da rede MT', lista('layers_rede_mt')));
        corpo.appendChild(ln('Layers da rede BT', lista('layers_rede_bt')));
        corpo.appendChild(ln('Cor da rede MT', cor('cor_mt')));
        corpo.appendChild(ln('Cor da rede BT', cor('cor_bt')));
        corpo.appendChild(ln('Regra', el('select', { style: caixa, onchange: e => { c.regra = e.target.value; } }, [['COR_MANDA', 'COR_MANDA — a layer confirma a cor; divergência descarta (padrão)'], ['LAYER_MANDA', 'LAYER_MANDA — a layer decide; sem layer de rede, vale a cor'], ['SO_LAYER', 'SO_LAYER — só a layer; a cor é ignorada']]
            .map(([v, n]) => el('option', { value: v, text: n, ...((c.regra || 'COR_MANDA') === v ? { selected: 'selected' } : {}) })))));
        corpo.appendChild(ln('Blocos de poste (nome contém)', lista('blocos_poste')));
        corpo.appendChild(ln('Layers de poste', lista('layers_poste')));
        corpo.appendChild(ln('Layers de guia (vazio = qualquer)', lista('layers_guia')));
        const saida = el('span', { style: 'color:#8b949e;flex:1' }), lay = el('input', { type: 'text', style: 'width:200px;' + caixa, placeholder: 'Layer: MT_PRIMARIA', value: st.teste.desenho_layer || '' }), cr = el('input', { type: 'text', style: 'width:110px;' + caixa, placeholder: 'Cor: 1 ou #0000FF', value: st.teste.desenho_cor || '' });
        const testar = async () => {
            st.teste.desenho_layer = lay.value; st.teste.desenho_cor = cr.value;
            try {
                const r = await fetch('/api/plugin/leitura/testar', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ leitura: c, layer: lay.value, cor: /^\d+$/.test(cr.value.trim()) ? Number(cr.value.trim()) : cr.value.trim() }) });
                const j = await r.json();
                if (!r.ok) throw new Error(((j.detail && j.detail.erros) || [j.detail || 'erro']).join('; '));
                const nome = { MT: '<b style="color:#f85149">rede MT</b>', BT: '<b style="color:#3fb950">rede BT</b>', DESCARTADA: '<b style="color:#d29922">linha descartada</b>', NAO_E_REDE: 'não é rede' }[j.resultado] || esc(j.resultado);
                saida.innerHTML = nome + ' · ' + esc(j.motivo);
            } catch (e) { saida.innerHTML = '<span style="color:#f85149">' + esc(String(e.message || e)) + '</span>'; }
        };
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:14px;align-items:center' }, [lay, cr, el('button', { class: 'vz-btn', text: 'Testar', onclick: testar }), saida]));
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:14px;align-items:center' }, [el('button', { class: 'vz-btn b', text: 'Salvar leitura do desenho', onclick: async () => {
            try {
                const r = await fetch(url, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ leitura: c }) }), j = await r.json();
                if (!r.ok) throw new Error(((j.detail && j.detail.erros) || [j.detail || 'erro']).join('; '));
                d.cfg = j.leitura; d.msg = j.aviso || 'Salvo.'; d.erro = ''; render();
            } catch (e) { raiz.alert('Não salvou: ' + (e.message || e)); }
        } }), el('span', { style: 'color:#3fb950', text: d.msg || '' })]));
    }

    /** TASK-176: perfil de exportação do projeto — campos do cabeçalho do relatório, centro por tipo, PDF/XLSX e arquivos oferecidos. Sem perfil, vale o de Neoenergia. */
    function abaExportacao(corpo) {
        const PE = raiz.PerfilExportacao;
        if (!PE) { corpo.appendChild(el('div', { class: 'vz-nota', text: 'Perfil de exportação indisponível nesta tela.' })); return; }
        const caixa = 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace';
        corpo.appendChild(el('div', { class: 'vz-nota', html: '<b>O que esta tela diz:</b> como saem os relatórios do projeto: os <b>campos do cabeçalho</b> (rótulo no PDF, em qual coluna, se viram parâmetro <code>OBRA.&lt;CHAVE&gt;</code> da base e se o programa lembra o valor), o <b>centro</b> de material e de mão de obra, o PDF (cor do título, rodapé) e as <b>colunas do Excel</b>. O <code>.obra</code> grava as chaves do cabeçalho. <b>Sem perfil, vale o de Neoenergia.</b>' }));
        if (!st.cfg.exportacao) {
            corpo.appendChild(el('div', { style: 'margin:10px 0' }, [el('span', { style: 'color:#8b949e', text: 'Este projeto usa o perfil de Neoenergia (centros 1600/1000, cabeçalho da nota SAP). ' }),
                el('button', { class: 'vz-btn', text: 'Personalizar a partir do padrão', onclick: () => { st.cfg.exportacao = JSON.parse(JSON.stringify(PE.PADRAO)); marcarSujo(); render(); } })]));
        } else {
            const x = st.cfg.exportacao;
            if (!Array.isArray(x.cabecalho)) x.cabecalho = [];
            x.centro = x.centro && typeof x.centro === 'object' ? x.centro : {}; x.pdf = x.pdf || {}; x.xlsx = x.xlsx || {};
            const sc = el('div', { class: 'vz-scroll' }), tb = el('table', { class: 'vz-tb' });
            tb.appendChild(el('thead', {}, [el('tr', {}, ['Chave', 'Rótulo no PDF', 'Rótulo na tela', 'Coluna', 'OBRA.*', 'Lembrar', ''].map(h => el('th', { text: h })))]));
            const tbody = el('tbody');
            x.cabecalho.forEach((c, i) => tbody.appendChild(el('tr', {}, [celula(c, 'chave', 'text', 'contrato'), celula(c, 'rotulo', 'text', 'CONTRATO:'), celula(c, 'tela', 'text', 'Contrato'), celula(c, 'coluna', 'select', ['esq', 'meio', 'id']),
                celula(c, 'obra', 'bool'), celula(c, 'lembrar', 'bool')].map(y => el('td', {}, [y])).concat([el('td', { style: 'white-space:nowrap' }, [
                    el('button', { class: 'mv', text: '▲', onclick: () => { if (i > 0) { x.cabecalho.splice(i - 1, 0, x.cabecalho.splice(i, 1)[0]); marcarSujo(); render(); } } }),
                    el('button', { class: 'mv', text: '▼', onclick: () => { if (i < x.cabecalho.length - 1) { x.cabecalho.splice(i + 1, 0, x.cabecalho.splice(i, 1)[0]); marcarSujo(); render(); } } }),
                    el('button', { class: 'x', text: '✖', title: 'Excluir o campo', onclick: () => { x.cabecalho.splice(i, 1); marcarSujo(); render(); } })])]))));
            tb.appendChild(tbody); sc.appendChild(tb); corpo.appendChild(sc);
            corpo.appendChild(el('div', { style: 'margin-top:8px' }, [el('button', { class: 'vz-btn', text: '＋ Campo', onclick: () => { x.cabecalho.push({ chave: '', rotulo: '', tela: '', coluna: 'meio', obra: false, lembrar: false }); marcarSujo(); render(); } })]));
            const ln = (rot, ...y) => el('div', { style: 'display:flex;gap:8px;align-items:center;margin-top:8px' }, [el('span', { style: 'min-width:190px;color:#8b949e', text: rot })].concat(y));
            const txt = (obj, k, larg, ph) => el('input', { type: 'text', style: 'width:' + larg + 'px;' + caixa, value: obj[k] === undefined ? '' : obj[k], placeholder: ph || '', onchange: e => { if (e.target.value.trim()) obj[k] = e.target.value.trim(); else delete obj[k]; marcarSujo(); } });
            corpo.appendChild(ln('Centro · material / mão de obra', txt(x.centro, 'MATERIAL', 90, '1600'), txt(x.centro, 'MAO-DE-OBRA', 90, '1000')));
            corpo.appendChild(ln('PDF · cor do título / rodapé', txt(x.pdf, 'cor_titulo', 100, '#00A443'), txt(x.pdf, 'rodape', 200, 'Internal Use')));
            corpo.appendChild(ln('PDF · rótulos da ordem / do ID', txt(x.pdf, 'rotulo_ordem', 160, 'TIPO DE ORDEM'), txt(x.pdf, 'rotulo_id', 160, 'ID INVESTIMENTO')));
            corpo.appendChild(ln('Excel · colunas (ordem)', el('input', { type: 'text', style: 'width:320px;' + caixa, value: (x.xlsx.colunas || PE.PADRAO.xlsx.colunas).join(', '), placeholder: 'mdo, codigo, qtd, centro, desc',
                onchange: e => { const l = e.target.value.split(',').map(trim).filter(Boolean); if (l.length) x.xlsx.colunas = l; else delete x.xlsx.colunas; marcarSujo(); } })));
            corpo.appendChild(ln('Arquivos oferecidos', ...PE.ARQUIVOS.map(a => el('label', { style: 'color:#c9d1d9' }, [el('input', { type: 'checkbox', ...((x.arquivos || PE.ARQUIVOS).includes(a) ? { checked: 'checked' } : {}),
                onchange: e => { const atual = new Set(x.arquivos || PE.ARQUIVOS); if (e.target.checked) atual.add(a); else atual.delete(a); x.arquivos = PE.ARQUIVOS.filter(k => atual.has(k)); marcarSujo(); } }), ' ' + a]))));
            const erros = PE.validar(x);
            if (erros.length) corpo.appendChild(el('div', { style: 'color:#f85149;margin-top:8px', html: erros.map(esc).join('<br>') }));
            corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:12px' }, [
                el('button', { class: 'vz-btn b', text: 'Pré-visualizar uma página', onclick: () => previaExportacao(corpo, x) }),
                el('button', { class: 'vz-btn', text: 'Voltar ao padrão de Neoenergia', onclick: () => { if (raiz.confirm('Remover o perfil deste projeto e voltar ao padrão de Neoenergia?')) { delete st.cfg.exportacao; marcarSujo(); render(); } } })]));
        }
    }
    async function previaExportacao(corpo, perfil) {
        const PE = raiz.PerfilExportacao, erros = PE.validar(perfil);
        if (erros.length) { raiz.alert(erros.join('\n')); return; }
        const cab = {}; PE.normalizar(perfil).cabecalho.forEach(c => { cab[c.chave] = c.chave === 'titulo' ? 'PROJETO DE EXEMPLO' : c.chave === 'nota' ? '000000' : (c.tela || c.chave) + ' (exemplo)'; });
        const linhas = ['M', 'N'].flatMap(m => [1, 2, 3].map(i => ({ mdo: m, codigo: String(m === 'M' ? 2460000 + i : 500100 + i), desc: (m === 'M' ? 'MATERIAL DE EXEMPLO ' : 'SERVIÇO DE EXEMPLO ') + i, qtd: i * 1.5 })));
        try {
            const r = await fetch('/api/orcamento/exportar-sap-pdf', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cabecalho: cab, odi: linhas, odd: [], perfil }) });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || ('HTTP ' + r.status)); }
            const url = URL.createObjectURL(await r.blob());
            const antigo = corpo.querySelector('.previa-exportacao'); if (antigo) antigo.remove();
            corpo.appendChild(el('iframe', { class: 'previa-exportacao', src: url + '#toolbar=0&view=FitH', style: 'width:100%;height:520px;border:1px solid #30363d;border-radius:6px;margin-top:12px;background:#fff' }));
        } catch (e) { raiz.alert('Não consegui gerar a pré-visualização: ' + (e.message || e)); }
    }

    /** TASK-178: casos de referência do projeto — pedidos com o resultado esperado. "Conferir base" roda todos (em lotes, com progresso) e mostra o que mudou código a código; "Conferir um CSV" faz o mesmo
     *  com uma base candidata, sem gravar; "Mudança esperada" aceita o resultado novo (com a nota de quem e por quê). */
    async function abaCasos(corpo) {
        const caixa = 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace';
        const projeto = host.projeto(), req = (url, corpoJson) => fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corpoJson) });
        if (!st.casos) {
            st.casos = { lista: null, res: null, progresso: '', msg: '' };
            fetch('/api/casos?projeto=' + encodeURIComponent(projeto)).then(r => r.ok ? r.json() : Promise.reject(new Error('HTTP ' + r.status))).then(d => { st.casos.lista = d.casos; render(); })
                .catch(e => { st.casos.lista = []; st.casos.msg = 'Não consegui ler os casos: ' + (e.message || e); render(); });
        }
        const ck = st.casos;
        corpo.appendChild(el('div', { class: 'vz-nota', html: '<b>O que esta aba faz:</b> guarda <b>pedidos com o resultado esperado</b> (no Resultado: <b>Guardar como caso de referência</b>) e <b>confere a base</b>: roda todos os casos e mostra o que mudou, código a código (antes → depois). Use depois de importar ou editar a base ou a vinculação. A importação de pacote também mostra o impacto nos casos antes de gravar.' }));
        if (!ck.lista) { corpo.appendChild(el('div', { style: 'color:#8b949e;margin-top:10px', text: 'Carregando…' })); return; }
        const rodar = async (baseCsv) => {
            ck.res = null; ck.msg = ''; const ids = ck.lista.map(c => c.id), todos = { casos: [], resumo: { igual: 0, mudou: 0, erro: 0 }, total: 0, segundos: 0 };
            for (let i = 0; i < ids.length; i += 40) {
                ck.progresso = `Conferindo ${Math.min(i + 40, ids.length)} de ${ids.length}…`; render();
                try {
                    const r = await req('/api/casos/conferir', { projeto, ids: ids.slice(i, i + 40), base_csv: baseCsv || undefined }), d = await r.json();
                    if (!r.ok) throw new Error(d.detail || ('HTTP ' + r.status));
                    todos.casos = todos.casos.concat(d.casos); ['igual', 'mudou', 'erro'].forEach(k => { todos.resumo[k] += d.resumo[k]; }); todos.total += d.total; todos.segundos += d.segundos;
                } catch (e) { ck.msg = 'Não conferiu: ' + (e.message || e); ck.progresso = ''; render(); return; }
            }
            ck.progresso = ''; ck.res = todos; ck.baseCsv = baseCsv || ''; render();
        };
        corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin:10px 0;align-items:center;flex-wrap:wrap' }, [
            el('button', { class: 'vz-btn b', text: 'Conferir base', ...(ck.lista.length ? {} : { disabled: 'disabled' }), onclick: () => rodar('') }),
            el('label', { style: 'color:#8b949e' }, ['ou conferir um CSV antes de importar: ', el('input', { type: 'file', accept: '.csv,text/csv', style: caixa, onchange: async e => { const f = e.target.files[0]; if (f) await rodar(await f.text()); } })]),
            el('span', { style: 'color:#d29922', text: ck.progresso || '' })]));
        const tb = el('table', { class: 'vz-tb' });
        tb.appendChild(el('thead', {}, [el('tr', {}, ['Caso', 'Origem', 'Itens esperados', 'Situação', ''].map(h => el('th', { text: h })))]));
        const sit = id => ck.res ? ck.res.casos.find(c => c.id === id) : null, cor = { igual: '#3fb950', mudou: '#d29922', erro: '#f85149' };
        tb.appendChild(el('tbody', {}, ck.lista.map(c => { const s = sit(c.id); return el('tr', {}, [el('td', { text: c.nome }), el('td', { text: c.origem }), el('td', { text: String(c.itens) }),
            el('td', { style: s ? 'font-weight:600;color:' + cor[s.situacao] : 'color:#8b949e', text: s ? s.situacao : '—' }),
            el('td', {}, [el('button', { class: 'x', text: '✖', title: 'Excluir o caso', onclick: async () => { if (!raiz.confirm('Excluir o caso “' + c.nome + '”?')) return; await fetch('/api/casos/' + c.id + '?projeto=' + encodeURIComponent(projeto), { method: 'DELETE' }); st.casos = null; render(); } })])]); })));
        corpo.appendChild(el('div', { class: 'vz-scroll' }, [tb]));
        if (!ck.lista.length) corpo.appendChild(el('div', { style: 'color:#8b949e;margin-top:8px', text: 'Nenhum caso ainda. No Resultado do Orçamento, clique em “Guardar como caso de referência”.' }));
        if (ck.res) {
            const r = ck.res.resumo;
            corpo.appendChild(el('div', { style: 'margin-top:10px;font-weight:600', text: `${r.igual} igual(is) · ${r.mudou} mudou · ${r.erro} com erro — ${ck.res.total} caso(s) em ${ck.res.segundos.toFixed(1)} s${ck.baseCsv ? ' (base candidata, nada foi gravado)' : ''}` }));
            ck.res.casos.filter(c => c.situacao !== 'igual').forEach(c => {
                corpo.appendChild(el('div', { style: 'margin-top:8px;border-left:3px solid ' + cor[c.situacao] + ';padding-left:10px' }, [el('div', { style: 'font-weight:600', text: c.nome + ' — ' + c.situacao }),
                    c.erro ? el('div', { style: 'color:#f85149', text: c.erro }) : null,
                    ...c.diferencas.map(d => el('div', { style: 'font-family:Consolas,monospace;font-size:12px', html: `<b>${esc(d.codigo)}</b> (${esc(d.operacao)}${d.mdo ? ' · ' + esc(d.mdo) : ''}): ${d.antes === null ? '—' : d.antes} → ${d.depois === null ? '—' : d.depois} <span style="color:#8b949e">${esc(d.tipo)}</span>` })),
                    c.situacao === 'mudou' ? el('button', { class: 'vz-btn', text: 'Mudança esperada', onclick: async () => {
                        const nota = raiz.prompt('Por que esta mudança é esperada? (fica no histórico do caso)', '');
                        if (!nota) return;
                        try { const rr = await req('/api/casos/aceitar', { projeto, ids: [c.id], nota, base_csv: ck.baseCsv || undefined }), d = await rr.json(); if (!rr.ok) throw new Error(d.detail || ('HTTP ' + rr.status)); st.casos = null; render(); }
                        catch (e) { raiz.alert('Não aceitou: ' + (e.message || e)); }
                    } }) : null].filter(Boolean)));
            });
        }
        if (ck.msg) corpo.appendChild(el('div', { style: 'color:#f85149;margin-top:8px', text: ck.msg }));
    }

    /** TASK-177: pacote do projeto — exportar tudo (ou partes) num .zip e importar num projeto com prévia do que muda. Só o administrador importa; nada é gravado antes de confirmar. */
    async function abaPacote(corpo) {
        const caixa = 'background:#010409;color:#c9d1d9;border:1px solid #30363d;border-radius:4px;padding:5px 8px;font:12px Consolas,monospace';
        if (!st.pacote) {
            st.pacote = { partes: null, marcadas: null, previa: null, arquivo: null, msg: '' };
            fetch('/api/pacote/partes').then(r => r.json()).then(d => { st.pacote.partes = d.partes; st.pacote.marcadas = new Set(d.partes.map(x => x.parte)); render(); }).catch(() => { st.pacote.partes = []; st.pacote.marcadas = new Set(); render(); });
        }
        const pk = st.pacote, projeto = host.projeto();
        corpo.appendChild(el('div', { class: 'vz-nota', html: '<b>O que esta aba faz:</b> o <b>pacote</b> é um arquivo <code>.zip</code> com tudo o que um projeto precisa (base técnica, vinculação, dicionário de cabos, postes, desenho, exportação, regras do leitor e tabela regional). Quem distribui o programa manda o pacote; quem recebe <b>importa</b> num projeto vazio e começa a orçar. A importação mostra antes <b>o que muda</b> em cada parte e só grava depois de você confirmar.' }));
        if (!pk.partes) { corpo.appendChild(el('div', { style: 'color:#8b949e;margin-top:10px', text: 'Carregando…' })); return; }
        corpo.appendChild(el('div', { style: 'margin-top:10px;font-weight:600', text: 'Exportar o projeto ' + projeto }));
        const lista = el('div', { style: 'display:grid;grid-template-columns:1fr 1fr;gap:4px;margin:6px 0' }, pk.partes.map(x => el('label', { style: 'color:#c9d1d9' }, [el('input', { type: 'checkbox', ...(pk.marcadas.has(x.parte) ? { checked: 'checked' } : {}),
            onchange: e => { if (e.target.checked) pk.marcadas.add(x.parte); else pk.marcadas.delete(x.parte); } }), ' ' + x.rotulo])));
        corpo.appendChild(lista);
        corpo.appendChild(el('button', { class: 'vz-btn b', text: 'Exportar pacote (.zip)', onclick: async () => {
            try {
                const r = await fetch('/api/pacote/exportar?projeto=' + encodeURIComponent(projeto) + '&partes=' + encodeURIComponent(Array.from(pk.marcadas).join(',')));
                if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || ('HTTP ' + r.status)); }
                const a = el('a', { href: URL.createObjectURL(await r.blob()), download: 'pacote_' + projeto + '.zip' }); doc.body.appendChild(a); a.click(); a.remove();
                pk.msg = 'Pacote exportado.'; render();
            } catch (e) { raiz.alert('Não exportou: ' + (e.message || e)); }
        } }));
        corpo.appendChild(el('div', { style: 'margin-top:18px;font-weight:600', text: 'Importar um pacote no projeto ' + projeto }));
        if (pk.prontos === undefined) { pk.prontos = null; fetch('/api/pacote/embutidos').then(r => r.json()).then(d => { pk.prontos = d.pacotes || []; render(); }).catch(() => { pk.prontos = []; }); }
        if (pk.prontos && pk.prontos.length) {
            const sel = el('select', { style: caixa }, pk.prontos.map(x => el('option', { value: x.id, text: x.nome })));
            corpo.appendChild(el('div', { style: 'display:flex;gap:8px;align-items:center;margin:6px 0' }, [el('span', { style: 'color:#8b949e', text: 'Começar de um pacote pronto:' }), sel, el('button', { class: 'vz-btn', text: 'Ver a prévia', onclick: async () => {
                try {
                    const z = await fetch('/api/pacote/embutidos/' + encodeURIComponent(sel.value) + '.zip'); if (!z.ok) throw new Error('HTTP ' + z.status);
                    pk.arquivo = new File([await z.blob()], 'pacote_' + sel.value + '.zip', { type: 'application/zip' });
                    const fd = new FormData(); fd.append('file', pk.arquivo); fd.append('projeto', projeto);
                    const r = await fetch('/api/pacote/previa', { method: 'POST', body: fd }), d = await r.json();
                    if (!r.ok) throw new Error(d.detail || ('HTTP ' + r.status));
                    pk.previa = d; pk.importar = new Set(d.partes.filter(x => x.situacao !== 'igual').map(x => x.parte)); pk.msg = '';
                } catch (err) { pk.msg = 'Não li o pacote pronto: ' + (err.message || err); }
                render();
            } })]));
            corpo.appendChild(el('div', { style: 'color:#8b949e;font-size:12px;margin-bottom:4px', text: pk.prontos.map(x => x.nome + ': ' + x.descricao).join(' · ') }));
        }
        corpo.appendChild(el('input', { type: 'file', accept: '.zip,application/zip', style: 'margin:6px 0;' + caixa, onchange: async e => {
            pk.arquivo = e.target.files[0] || null; pk.previa = null; pk.msg = '';
            if (!pk.arquivo) { render(); return; }
            try {
                const fd = new FormData(); fd.append('file', pk.arquivo); fd.append('projeto', projeto);
                const r = await fetch('/api/pacote/previa', { method: 'POST', body: fd }), d = await r.json();
                if (!r.ok) throw new Error(d.detail || ('HTTP ' + r.status));
                pk.previa = d; pk.importar = new Set(d.partes.filter(x => x.situacao !== 'igual').map(x => x.parte));
            } catch (err) { pk.msg = 'Não li o pacote: ' + (err.message || err); }
            render();
        } }));
        if (pk.previa) {
            const rot = { novo: ['novo', '#3fb950'], substitui: ['substitui', '#d29922'], igual: ['igual', '#8b949e'] };
            corpo.appendChild(el('div', { style: 'color:#8b949e;margin:6px 0', text: 'Pacote' + (pk.previa.origem ? ' do projeto ' + pk.previa.origem.nome + ' (' + pk.previa.origem.codigo + ')' : '') + (pk.previa.descricao ? ' — ' + pk.previa.descricao : '') }));
            const tb = el('table', { class: 'vz-tb' });
            tb.appendChild(el('thead', {}, [el('tr', {}, ['Importar', 'Parte', 'Situação', 'Hoje', 'Pacote', 'Avisos'].map(h => el('th', { text: h })))]));
            tb.appendChild(el('tbody', {}, pk.previa.partes.map(x => el('tr', {}, [
                el('td', {}, [el('input', { type: 'checkbox', ...(pk.importar.has(x.parte) ? { checked: 'checked' } : {}), onchange: e => { if (e.target.checked) pk.importar.add(x.parte); else pk.importar.delete(x.parte); } })]),
                el('td', { text: x.rotulo }), el('td', { style: 'color:' + rot[x.situacao][1] + ';font-weight:600', text: rot[x.situacao][0] }), el('td', { text: x.antes }), el('td', { text: x.depois }),
                el('td', { style: 'color:#f85149', html: x.avisos.map(esc).join('<br>') })]))));
            corpo.appendChild(el('div', { class: 'vz-scroll' }, [tb]));
            corpo.appendChild(el('div', { style: 'display:flex;gap:8px;margin-top:10px' }, [
                el('button', { class: 'vz-btn b', text: 'Confirmar e importar', onclick: async () => {
                    if (!pk.importar.size) { raiz.alert('Marque ao menos uma parte.'); return; }
                    if (!raiz.confirm('Importar ' + pk.importar.size + ' parte(s) no projeto ' + projeto + '? As partes marcadas como "substitui" serão trocadas.')) return;
                    try {
                        const fd = new FormData(); fd.append('file', pk.arquivo); fd.append('projeto', projeto); fd.append('partes', Array.from(pk.importar).join(',')); fd.append('confirmar', 'true');
                        const r = await fetch('/api/pacote/importar', { method: 'POST', body: fd }), d = await r.json();
                        if (!r.ok) throw new Error(d.detail || ('HTTP ' + r.status));
                        pk.msg = 'Importado: ' + d.gravadas.join(', ') + '. Reabra a tela para usar a configuração nova.'; pk.previa = null; pk.arquivo = null;
                        if (d.gravadas.some(p => ['vinculacao', 'dicionario_cabos', 'postes', 'exportacao'].includes(p))) { try { const rc = await fetch('/api/vinculacao-config?projeto_codigo=' + encodeURIComponent(projeto)); if (rc.ok) st.cfg = clone(VR.normalizarConfig((await rc.json()).config)); } catch (x) { /* a tela recarrega ao reabrir */ } }
                    } catch (err) { raiz.alert('Não importou: ' + (err.message || err)); }
                    render();
                } }),
                el('button', { class: 'vz-btn', text: 'Cancelar', onclick: () => { pk.previa = null; pk.arquivo = null; render(); } })]));
        }
        if (pk.msg) corpo.appendChild(el('div', { style: 'color:#3fb950;margin-top:10px', text: pk.msg }));
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
            { campo: 'padrao', rot: 'Padrão da estrutura', tipo: 'text', extra: '^N1$ ou N1, N2' }, { campo: 'ocupa', rot: 'Ocupa', tipo: 'bool', extra: true, w: '55px', dica: 'marcado (padrão) = consome pontas dos cabos. Desmarcado = EQUIPAMENTO (RL, CFU, para-raio…): não consome ponta, não conta capacidade e lê os cabos do poste (veja Herda).' }, { campo: 'herda', rot: 'Herda', tipo: 'select', extra: ['', 'POSTE'], w: '75px', dica: 'só com Ocupa desmarcado: POSTE = recebe os cabos que chegam no poste (desenho), os ligados às estruturas do poste e os fictícios' }, { campo: 'nivel', rot: 'Nível', tipo: 'select', extra: ['', 'MT', 'BT'], w: '60px', dica: 'filtra os cabos do poste pelo nível do cabo' }, { campo: 'ordem', rot: 'Ordem', tipo: 'select', extra: ['ORIGEM', 'MAIOR_PRIMEIRO'], w: '120px', dica: 'MAIOR_PRIMEIRO: ordena pela seção do cabo, o maior primeiro (tronco em cabo, ramal em cabo2)' }, { campo: 'distintos', rot: 'Distintos', tipo: 'bool', w: '65px', dica: 'junta os cabos de mesma descrição antes de distribuir nos parâmetros (vale também para estruturas que ocupam)' }, { campo: 'se_faltar', rot: 'Se faltar', tipo: 'select', extra: ['AVISAR', 'BLOQUEAR'], w: '90px', dica: 'poste sem cabo que sirva: AVISAR (parâmetro no padrão da base) ou BLOQUEAR (estado vermelho)' },
            { campo: 'qtd_cabos', rot: 'Cabos', tipo: 'number', w: '60px', dica: 'quantidade fixa (mín = máx)', cinza: r => r.ocupa === false }, { campo: 'qtd_min', rot: 'Mín', tipo: 'number', w: '50px', dica: 'faixa: mínimo para ficar completa', cinza: r => r.ocupa === false }, { campo: 'qtd_max', rot: 'Máx', tipo: 'number', w: '50px', dica: 'faixa: máximo aceito', cinza: r => r.ocupa === false }, { campo: 'max_reais', rot: 'Reais', tipo: 'number', w: '55px', cinza: r => r.ocupa === false, dica: 'quantos podem ser cabos REAIS; o excedente só entra como cabo fictício (ex.: SI3 = 1)' },
            { campo: 'lados', rot: 'Os lados…', tipo: 'select', extra: ['LIVRE', 'IGUAIS', 'DERIVACAO'], w: '100px', cinza: r => r.ocupa === false, dica: 'LIVRE · IGUAIS (cabos iguais) · DERIVACAO (1 cabo de derivação obrigatório + até 2 do tronco, iguais entre si; o 1º parâmetro é sempre a derivação)' }, { campo: 'amarracao', rot: 'Amarração', tipo: 'bool', w: '80px', dica: 'aceita cabo M', cinza: r => r.ocupa === false },
            { campo: 'aceita_i', rot: 'Aceita I', tipo: 'bool', w: '70px', dica: 'desmarcado = só cabo M', cinza: r => r.ocupa === false }, { campo: 'cabos', rot: 'Cabos que aceita', tipo: 'text', extra: '(vazio = qualquer)' },
            { campo: 'parametro', rot: 'Parâmetro', tipo: 'text', w: '80px', dica: 'parâmetro ÚNICO que recebe o VALOR DE CABO da Tabela de Cabos (use tipocabo); vazio = nenhum. Para a descrição completa do cabo use a coluna Parâmetros (vários) com cabo.' }, { campo: 'parametros', rot: 'Parâmetros (vários)', tipo: 'text', w: '150px', extra: 'cabo, cabo, tipocabo, tipocabo', dica: 'um valor por cabo vinculado. cabo = descrição completa do cabo, sem o comprimento (3xS04-7,97kV → 3XS04-7.97KV); tipocabo = a coluna Valor de cabo da Tabela de Cabos; o mesmo nome pode se repetir (N4 com dois cabos). Se preenchido, substitui o Parâmetro único.' }, { campo: 'se_diferem', rot: 'Se diferem', tipo: 'select', extra: ['MAIOR', 'BLOQUEAR', 'AVISAR'], w: '95px' },
            { campo: 'cabos_vaos', rot: 'Cabos por vão', tipo: 'json', extra: '["^PB\\\\d+$", "^C\\\\d+$"]', w: '170px', dica: 'lista de regex, uma por vão: cada uma precisa de um cabo DIFERENTE (IBIT: 1 multiplex + 1 BT nua). Com ela a estrutura aceita tantos cabos quantos vãos e os parâmetros seguem a ordem dos vãos. Vazio = vale a coluna Cabos que aceita.' },
            { campo: 'define', rot: 'Parâmetros automáticos', tipo: 'json', extra: '[{"linha_tem":"RL","param":"tipo=RL"}]', w: '220px', dica: 'lista de {linha_tem?, cabo?, param:"nome=valor"}: linha_tem = outra estrutura da mesma linha do poste; cabo = regex do 1º cabo da estrutura. O 1º item que casa para cada nome vence; não sobrescreve o digitado nem o vínculo de cabo.' },
            { campo: 'ativa', rot: 'Ativa', tipo: 'bool', w: '50px' }, { campo: 'obs', rot: 'Observação', tipo: 'text' }],
            () => ({ padrao: '', qtd_cabos: 1, lados: 'LIVRE', amarracao: false, aceita_i: true, cabos: '', parametro: 'cabo', se_diferem: 'MAIOR', ativa: true, obs: '' }),
            txt => {
                const cod = host.codigoEstrutura(txt) || txt, i = VR.infoEstrutura(cod, st.cfg);
                if (i.regra !== null && !i.ocupa) return `equipamento <b style="color:#79c0ff">${esc(i.codigo)}</b> → linha #${i.regra + 1} · <b>não ocupa pontas</b>${i.herda === 'POSTE' ? ' · lê os cabos ' + (i.nivel || '') + ' do poste' : ''} · ordem ${i.ordem === 'MAIOR_PRIMEIRO' ? 'maior primeiro' : 'de origem'}${i.distintos ? ' · cabos distintos' : ''} · parâmetros ${esc((i.parametros || []).map(p => p.nome.toLowerCase()).join(', ') || i.parametro.toLowerCase() || '—')} · sem cabo: ${i.se_faltar === 'BLOQUEAR' ? 'bloqueia' : 'avisa'}`;
                return `estrutura <b style="color:#79c0ff">${esc(i.codigo)}</b> → ${i.regra === null ? 'nenhuma linha casa' : 'linha #' + (i.regra + 1) + ' · ' + (i.qtd || '?') + ' cabo(s) · lados ' + i.lados + (i.amarracao ? ' · amarração' : '') + (i.aceita_i ? '' : ' · só cabo M')}${i.distintos ? ' · cabos distintos' : ''}${i.ordem === 'MAIOR_PRIMEIRO' ? ' · maior primeiro' : ''}`;
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
            el('button', { class: 'vz-btn p', text: 'Aplicar tudo', onclick: () => aplicarScript(false) }),
            el('button', { class: 'vz-btn b', text: '🏗️ Vincular por poste (desenho)', title: 'Cada estrutura do poste liga nos cabos que CHEGAM nele (pela geometria do desenho do plugin), respeitando as Tabelas de Cabos e de Estruturas. Vínculo manual é mantido; os demais são refeitos.', onclick: () => aplicarPorPoste() })]);
        painel.appendChild(acoes);
        painel.appendChild(previaPosteHtml());
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

    // TASK-166: vinculação por poste — cada estrutura liga nos cabos que chegam no poste (geometria `_geo` do vão × `_poste`)
    function aplicarPorPoste() {
        host.antes();
        const r = VR.vincularPorPoste(cabos(), outros(), st.cfg, host.prefixoCabo, host.codigoEstrutura, host.codigosEstrutura, undefined, true);
        st.resultadoPoste = r;
        const s = r.resumo;
        toast(`Por poste: ${s.completos} cabo(s) completo(s), ${s.parciais} parcial(is), ${s.sem_vinculo} sem vínculo; ${s.estruturas_incompletas} estrutura(s) incompleta(s) (Ctrl+Z desfaz).`);
        render();
    }

    function previaPosteHtml() {
        const r = st.resultadoPoste;
        if (!r) return el('div', { style: 'color:#8b949e;margin-top:10px', text: 'Vincular por poste usa a geometria do desenho (obra vinda do plugin do nanoCAD).' });
        const s = r.resumo, rot = { incompleta: 'Estrutura incompleta', sobra: 'Cabo sobrando no poste', ponta_solta: 'Ponta solta', escolha: 'Escolhido e avisado', sem_geometria: 'Sem geometria' };
        const lista = r.achados.slice(0, 40).map(a => `• ${rot[a.tipo] || a.tipo}: ${a.mensagem}`).join('\n') + (r.achados.length > 40 ? `\n… e mais ${r.achados.length - 40}` : '');
        return el('div', { style: 'margin-top:12px' }, [
            el('div', { style: 'color:#8b949e', html: `Por poste: <b style="color:#3fb950">${s.completos}</b> completo(s) · <b style="color:#a371f7">${s.parciais}</b> parcial(is) · ${s.sem_vinculo} sem vínculo · ${s.estruturas_incompletas} estrutura(s) incompleta(s) · ${s.pontas_soltas} ponta(s) solta(s) · ${s.manuais_mantidos} manual(is) mantido(s)` }),
            r.achados.length ? el('div', { class: 'vz-erro', style: 'white-space:pre-wrap;max-height:220px;overflow:auto', text: lista }) : null]);
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
    function abaMapa(corpo) {
        const W = 1200, H = 600, margem = 54;
        const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('class', 'vz-mapa');
        svg.setAttribute('aria-label', 'Grafo da rede. Selecione um poste para editar os vínculos');
        const posicoes = new Map(), vaosPorCabo = new Map();
        const totalPostes = el('strong'), totalVaos = el('strong');
        const contagem = el('div', { class: 'vz-grafo-legenda' }, [
            el('span', { class: 'vz-grafo-total' }, [totalPostes, 'Postes']),
            el('span', { class: 'vz-grafo-total' }, [totalVaos, 'Vãos'])
        ]);
        let grafo, vb = st.grafoVista || [0, 0, W, H], arrasto, ultimoNo, mover = false;
        let ultimoStatus = '';
        const aplicarVista = () => { st.grafoVista = vb; svg.setAttribute('viewBox', vb.join(' ')); };
        const zoom = (f, x = vb[0] + vb[2] / 2, y = vb[1] + vb[3] / 2) => {
            f = Math.max(120 / vb[2], Math.min(3600 / vb[2], f));
            vb = [x + (vb[0] - x) * f, y + (vb[1] - y) * f, vb[2] * f, vb[3] * f];
            aplicarVista();
        };
        const botao = (text, onclick, title = text) => el('button', { class: 'vz-btn', text, title, 'aria-label': title, onclick });
        const barra = el('div', { class: 'vz-grafo-barra' }, [
            el('div', {}, [el('h4', { text: 'Grafo da rede' }), el('p', { text: 'Selecione um poste para editar os vínculos ou um vão para consultar no assistente.' })]), contagem
        ]);
        const acoesMapa = el('div', { class: 'vz-grafo-acoes' }, [
            botao('−', () => zoom(1.25), 'Afastar'), botao('+', () => zoom(.8), 'Aproximar'),
            botao('Enquadrar', () => { vb = [0, 0, W, H]; aplicarVista(); })
        ]);
        const suporte = el('div', { class: 'vz-scroll vz-grafo-vista' }, [svg, acoesMapa]);
        const vazio = el('p', { class: 'vz-nota', text: 'As coordenadas da leitura ou as atribuições de Cabos ⇄ Postes formam a rede no mapa.' });
        const avisos = el('p', { class: 'vz-grafo-aviso', role: 'status' });
        const dialogo = el('dialog', { class: 'vz-no-dialog', 'aria-labelledby': 'vz-no-titulo' });
        corpo.append(barra, vazio, suporte, avisos, dialogo);

        function pintar() {
            grafo = VR.grafoDaRede(cabos(), outros(), host.codigoEstrutura, host.codigosEstrutura);
            let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
            const limitar = ([x, y]) => { x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y); };
            grafo.nos.forEach(n => { if (n.pos) limitar(n.pos); });
            grafo.vaos.forEach(v => (v.pontos || v.pontas || []).forEach(limitar));
            const sinteticos = [...grafo.nos.values()].filter(n => !n.pos);
            const temGeo = Number.isFinite(x0);
            const largura = temGeo && sinteticos.length ? W * .74 : W;
            const k = temGeo ? Math.min((largura - 2 * margem) / Math.max(x1 - x0, 1), (H - 2 * margem) / Math.max(y1 - y0, 1)) : 1;
            const ponto = ([x, y]) => [largura / 2 + (x - (x0 + (x1 - x0) / 2)) * k, H / 2 - (y - (y0 + (y1 - y0) / 2)) * k];
            posicoes.clear();
            vaosPorCabo.clear();
            grafo.vaos.forEach(v => v.cabos.forEach(ic => vaosPorCabo.set(ic, v)));
            grafo.nos.forEach(n => { if (n.pos) posicoes.set(n.id, ponto(n.pos)); });
            const colunas = Math.ceil(Math.sqrt(sinteticos.length));
            sinteticos.forEach((n, i) => {
                const ini = temGeo ? largura : 0, linhas = Math.ceil(sinteticos.length / colunas);
                posicoes.set(n.id, [ini + margem + (i % colunas) * (W - ini - 2 * margem) / Math.max(1, colunas - 1),
                    margem + Math.floor(i / colunas) * (H - 2 * margem) / Math.max(1, linhas - 1)]);
            });
            const a = analise();
            const rede = grafo.vaos.map((v, indice) => {
                const extremidades = v.ids.map((id, i) => id ? posicoes.get(id) : ponto(v.pontas[i]));
                const pts = v.pontos ? v.pontos.map(ponto) : extremidades.slice();
                pts[0] = extremidades[0]; pts[pts.length - 1] = extremidades[1];
                const grupos = VR.gruposVisuaisDoVao(v.cabos, cabos(), a, st.cfg.dicionario_cabos);
                const [x0, y0] = pts[0], [x1, y1] = pts[pts.length - 1];
                const dx = x1 - x0, dy = y1 - y0, comprimento = Math.hypot(dx, dy) || 1;
                return grupos.map((grupo, i) => {
                    const deslocamento = (i - (grupos.length - 1) / 2) * 4;
                    const linha = pts.map(([x, y]) => `${(x - dy * deslocamento / comprimento).toFixed(2)},${(y + dx * deslocamento / comprimento).toFixed(2)}`).join(' ');
                    const estado = grupo.completo ? 'Vinculado' : 'Pendente';
                    const rotulo = `${grupo.tipo} · ${estado}\n${grupo.cabos.map(ic => cabos()[ic].ativo).join('\n')}`;
                    const classeRede = grupo.tipo === 'MT' ? 'mt' : grupo.tipo === 'BT' ? 'bt' : '';
                    const interacao = host.consultarAssistente ? `data-vao="${indice}" data-grupo="${i}" tabindex="0" role="button" aria-label="Consultar vão ${esc(grupo.tipo)} ${indice + 1} no assistente"` : '';
                    return `<polyline ${interacao} class="vz-aresta ${classeRede}${grupo.completo ? ' completo' : ''}" points="${linha}" fill="none" vector-effect="non-scaling-stroke" ${v.ids.includes(null) ? 'stroke-dasharray="5 4"' : ''}><title>${esc(rotulo)}</title></polyline>`;
                }).join('');
            }).join('');
            const labels = [];
            const nos = [...grafo.nos.values()].map(n => {
                const [x, y] = posicoes.get(n.id), rotulo = nomeDoNo(n), curto = `P${n.linhas[0] + 1}`;
                const direita = x + 15 + curto.length * 7;
                const denso = labels.some(([a, b, c]) => Math.abs(y - c) < 14 && x + 10 < b && direita > a);
                if (!denso) labels.push([x + 10, direita, y]);
                const completo = VR.posteCompletoNoGrafo(estruturasDoNo(n), a);
                const classe = [denso ? 'vz-rotulo-denso' : '', completo ? 'completo' : 'pendente'].join(' ');
                return `<g class="${classe}" data-no="${esc(n.id)}" tabindex="0" role="button" aria-label="${esc(rotulo)}" transform="translate(${x} ${y})"><title>${esc(rotulo)} · ${completo ? 'Vinculação completa' : 'Vinculação pendente'}</title><circle class="vz-alvo" r="14"/><circle class="vz-halo" r="8.5"/><circle class="vz-marcador" r="4.8"/><text x="11" y="-6">${curto}</text></g>`;
            }).join('');
            svg.innerHTML = rede + nos;
            totalPostes.textContent = grafo.nos.size;
            totalVaos.textContent = grafo.vaos.length;
            avisos.textContent = [grafo.pontasSoltas ? `${grafo.pontasSoltas} ponta(s) sem poste identificado` : '',
                grafo.semVao ? `${grafo.semVao} cabo(s) sem dados suficientes para traçar um vão` : ''].filter(Boolean).join(' · ');
            vazio.hidden = !!(grafo.nos.size || grafo.vaos.length);
            suporte.hidden = !vazio.hidden;
        }

        function nomeDoNo(no) {
            const identificacao = no.linhas.flatMap(postesDe)[0];
            return 'Poste ' + (no.linhas[0] + 1) + (identificacao ? ' · ' + identificacao : '');
        }

        function editarNo(id, alvo) {
            ultimoNo = alvo;
            ultimoStatus = '';
            st.grafoNo = id;
            preencherNo();
            if (!dialogo.open) dialogo.showModal();
        }

        const estruturasDoNo = no => no.estruturas.filter(ie => !metaDe(linhaDe(ie)).soPoste && codigoDe(ie));

        function preencherNo() {
            const no = grafo.nos.get(st.grafoNo);
            if (!no) { if (dialogo.open) dialogo.close(); return; }
            const a = analise();
            const estruturas = estruturasDoNo(no);
            const completo = VR.posteCompletoNoGrafo(estruturas, a);
            const fecharNo = botao('Fechar', () => dialogo.close());
            const consultarNo = host.consultarAssistente && botao('Consultar no assistente', () => consultar({
                tipo: 'poste', rotulo: nomeDoNo(no), outros: no.linhas, cabos: [...no.cabos]
            }));
            const desfazer = botao('Desfazer', () => {
                host.desfazer();
                ultimoStatus = 'Histórico atualizado.';
                atualizar();
                dialogo.querySelector('[data-desfazer]')?.focus();
            });
            desfazer.setAttribute('data-desfazer', '');
            fecharNo.setAttribute('autofocus', '');
            const status = el('p', { class: 'vz-no-status', role: 'status', text: ultimoStatus });
            function mudarVinculo(ic, ie, acao, foco) {
                const motivo = acao === 'ligar' ? motivoParaLigar(ic, ie) : '';
                if (motivo) { status.textContent = motivo; return; }
                ultimoStatus = acao === 'ligar' ? 'Vínculo registrado.' : 'Vínculo removido.';
                ligar(ic, ie);
                if (host.depois) host.depois();
                dialogo.querySelector(`[data-foco="${foco}"]`)?.focus();
            }
            const lateral = el('aside', { class: 'vz-no-lateral' }, [
                el('h5', { text: 'Estruturas do poste' }),
                estruturas.length ? el('div', { class: 'vz-no-estruturas' }, estruturas.map(ie => {
                    const r = a.estruturas[ie];
                    return el('div', { class: 'vz-no-estrutura' }, [
                        el('strong', { text: codigoDe(ie) }),
                        el('span', { text: infoEst(ie).ocupa === false ? ' · equipamento' : ` · ${r.tem} ligado(s)` })
                    ]);
                })) : el('p', { class: 'vz-no-ajuda', text: 'Nenhuma estrutura cadastrada neste poste.' })
            ]);
            const painelNo = el('section', {}, [el('h5', { text: 'Cabos do poste' }), status]);
            if (!no.cabos.size) painelNo.appendChild(el('p', { class: 'vz-no-ajuda', text: 'Nenhum cabo chega a este poste pela geometria ou pelas atribuições da obra.' }));
            no.cabos.forEach(ic => {
                const c = cabos()[ic], r = a.cabos[ic], existentes = (r && c.vinculoEstruturas || []).filter(ie => no.linhas.includes(linhaDe(ie)));
                const vao = vaosPorCabo.get(ic);
                const outroNo = vao && vao.ids.includes(no.id) && vao.ids.find(id => id && id !== no.id);
                const origem = outroNo ? nomeDoNo(grafo.nos.get(outroNo)) : (vao && vao.ids.includes(no.id) ? 'Sem outro poste identificado' : 'Sem vão associado neste poste');
                const cartao = el('section', { class: 'vz-no-cabo' }, [
                    el('h6', {}, [el('span', { class: 'vz-op' + opClasse(trim(c.operacao)), text: c.operacao || 'I' }), ' ', c.ativo]),
                    el('p', { class: 'vz-no-ajuda', style: 'margin:0 0 6px', text: `Cabo ${ic + 1} · ${origem}` }),
                    el('div', { class: 'vz-no-vinculos' }, [el('span', { class: 'vz-no-ajuda', text: existentes.length ? 'Vínculos neste poste' : 'Sem vínculo neste poste' }),
                        ...existentes.map(ie => botao(codigoDe(ie) + ' ×', () => mudarVinculo(ic, ie, 'remover', 'c' + ic), 'Remover vínculo com ' + codigoDe(ie)))])
                ]);
                if (!r) {
                    cartao.appendChild(el('div', { class: 'vz-no-ajuda', text: 'Operação de retirada, fora da vinculação.' }));
                    painelNo.appendChild(cartao);
                    return;
                }
                cartao.title = TXT_C[r.estado];
                const candidatas = estruturas.filter(ie => !existentes.includes(ie) && infoEst(ie).ocupa !== false);
                const seletor = el('select', { 'aria-label': 'Estrutura para ' + c.ativo, 'data-foco': 'c' + ic }, candidatas.map(ie =>
                    el('option', { value: ie, text: codigoDe(ie) + ' · linha ' + (linhaDe(ie) + 1) })));
                if (candidatas.length) {
                    const dica = el('span', { class: 'vz-no-ajuda' });
                    const atualizarDica = () => { dica.textContent = motivoParaLigar(ic, Number(seletor.value)); };
                    seletor.addEventListener('change', atualizarDica);
                    const vincular = botao('Vincular', () => mudarVinculo(ic, Number(seletor.value), 'ligar', 'c' + ic));
                    vincular.classList.add('p');
                    atualizarDica();
                    cartao.appendChild(el('div', { class: 'vz-no-acoes' }, [seletor,
                        vincular, dica]));
                }
                painelNo.appendChild(cartao);
            });
            dialogo.replaceChildren(el('header', { class: 'vz-no-topo' }, [
                el('div', {}, [el('h4', { id: 'vz-no-titulo', text: nomeDoNo(no) }),
                    el('p', { class: completo ? 'vz-no-completo' : '', text: completo ? 'Vinculação completa' : 'Vinculação pendente' }),
                    el('p', { text: no.pos ? `X ${no.pos[0].toFixed(2)} · Y ${no.pos[1].toFixed(2)}` : 'Poste atribuído sem coordenadas' })]), el('div', { class: 'vz-no-topo-acoes' }, [consultarNo, desfazer, fecharNo].filter(Boolean))
            ]), el('div', { class: 'vz-no-conteudo' }, [lateral, painelNo]));
        }

        dialogo.addEventListener('close', () => {
            st.grafoNo = null;
            [...svg.querySelectorAll('[data-no]')].find(n => n.dataset.no === ultimoNo?.dataset.no)?.focus();
        });
        function coordenada(e) {
            const p = svg.createSVGPoint();
            p.x = e.clientX; p.y = e.clientY;
            return p.matrixTransform(svg.getScreenCTM().inverse());
        }
        svg.addEventListener('wheel', e => { e.preventDefault(); const p = coordenada(e); zoom(e.deltaY > 0 ? 1.15 : 1 / 1.15, p.x, p.y); }, { passive: false });
        svg.addEventListener('pointerdown', e => {
            if (e.button !== 0) return;
            mover = false;
            arrasto = { id: e.pointerId, x: e.clientX, y: e.clientY, vb, matriz: svg.getScreenCTM().inverse() };
        });
        svg.addEventListener('pointermove', e => {
            if (!arrasto || e.pointerId !== arrasto.id) return;
            const dx = e.clientX - arrasto.x, dy = e.clientY - arrasto.y;
            if (!mover && Math.hypot(dx, dy) < 4) return;
            if (!mover) { svg.setPointerCapture(e.pointerId); svg.classList.add('arrastando'); mover = true; }
            vb = [arrasto.vb[0] - dx * arrasto.matriz.a, arrasto.vb[1] - dy * arrasto.matriz.d, arrasto.vb[2], arrasto.vb[3]];
            aplicarVista();
        });
        for (const evento of ['pointerup', 'pointerleave', 'pointercancel']) svg.addEventListener(evento, () => { arrasto = null; svg.classList.remove('arrastando'); });
        function consultar(selecao) {
            if (host.consultarAssistente && fechar()) host.consultarAssistente(selecao);
        }
        function ativarMapa(alvo) {
            const no = alvo.closest('[data-no]');
            if (no) { editarNo(no.dataset.no, no); return; }
            const aresta = alvo.closest('[data-vao]');
            if (!aresta) return;
            const v = grafo.vaos[Number(aresta.dataset.vao)];
            const grupos = VR.gruposVisuaisDoVao(v.cabos, cabos(), analise(), st.cfg.dicionario_cabos);
            const grupo = grupos[Number(aresta.dataset.grupo)];
            if (!grupo) return;
            const nos = v.ids.map(id => grafo.nos.get(id)).filter(Boolean);
            consultar({ tipo: 'vão', rotulo: ['Vão ' + grupo.tipo + ' ' + (Number(aresta.dataset.vao) + 1), nos.map(nomeDoNo).join(' ↔ ')].filter(Boolean).join(' · '),
                cabos: grupo.cabos, outros: nos.flatMap(n => n.linhas) });
        }
        svg.addEventListener('click', e => { if (!mover) ativarMapa(e.target); });
        svg.addEventListener('keydown', e => {
            if (e.target.closest('[data-no], [data-vao]') && (e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                ativarMapa(e.target);
            }
        });

        mapaAtualizar = () => { pintar(); if (dialogo.open) preencherNo(); };
        pintar();
        aplicarVista();
        if (st.grafoNo && grafo.nos.has(st.grafoNo)) editarNo(st.grafoNo, [...svg.querySelectorAll('[data-no]')].find(n => n.dataset.no === st.grafoNo));
    }

    raiz.VinculacaoUI = { abrir, fechar };
})(typeof window !== 'undefined' ? window : globalThis);
