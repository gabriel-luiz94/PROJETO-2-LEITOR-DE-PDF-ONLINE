/*
 * vinculacao_ui.js (TASK-084) — modal de vinculação cabo <-> estrutura em abas: Vínculo (painel visual), Tabela de Cabos, Tabela de Estruturas,
 * Script e Mapa. A lógica pura está em vinculacao_regras.js (testada em Node e em Python). Quem altera as tabelas do Resumo é o host (resumo.js).
 * Cores: preto = sem vínculo · roxo = parcial · verde = completo · vermelho = inválido · tracejado = cabo fictício.
 */
(function (raiz) {
    'use strict';
    const doc = raiz.document, VR = raiz.VinculacaoRegras;
    let host = null, st = null, overlay = null, mapaAtualizar = null, mapaDesmontar = null;

    const CSS = `
    .vz-ov{position:fixed;inset:0;background:#05080d;z-index:9000;display:flex;align-items:center;justify-content:center}
    .vz{width:min(1500px,96vw);height:min(900px,94vh);background:#0d1117;border:1px solid #30363d;border-radius:12px;display:flex;flex-direction:column;overflow:hidden;color:#c9d1d9;font:13px/1.35 -apple-system,Segoe UI,Roboto,sans-serif}
    .vz-ov.vz-destacada{position:relative;inset:auto;z-index:auto;flex:1;min-height:0;background:transparent}
    .vz-ov.vz-destacada.vz-pagina{position:fixed;inset:0;z-index:9000;background:#0d1117}
    .vz-ov.vz-destacada:not(.vz-pagina) .vz-top .vz-x{display:none}
    .vz-destacada > .vz{width:100%;height:100%;border:0;border-radius:0}
    .vz-destacada .vz-destacar{display:none}
    .vz-destacada .vz-foot{gap:6px;flex-wrap:wrap}
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
    .vz-grafo-total{background:var(--g-panel);border:1px solid var(--g-border);border-radius:8px;padding:6px 12px;color:var(--g-muted);white-space:nowrap}
    .vz-grafo-total strong{color:var(--g-text);margin-right:5px;font-variant-numeric:tabular-nums}
    .vz-grafo-area{flex:1;min-height:0;display:grid;grid-template-columns:minmax(0,1fr) minmax(205px,24%);gap:12px}
    .vz-grafo-area.sem-lateral{grid-template-columns:minmax(0,1fr)}
    .vz-grafo-vista{border-radius:10px;border:1px solid var(--g-border);overflow:hidden}
    .vz-grafo-cabos{min-height:0;display:flex;flex-direction:column;gap:8px;padding:10px;background:var(--g-panel);border:1px solid var(--g-border);border-radius:10px}
    .vz-grafo-cabos[hidden],.vz-grafo-area[hidden],.vz-grafo-revisao[hidden]{display:none}
    .vz-grafo-cabos h5{margin:0}.vz-grafo-lista{overflow:auto;display:grid;align-content:start;gap:6px}
    .vz-grafo-cabecalho,.vz-grafo-legenda,.vz-grafo-ferramentas{display:flex;align-items:center;flex-wrap:wrap;gap:6px}
    .vz-grafo-item{display:grid;gap:6px;padding:7px;border-radius:7px;border:1px solid var(--g-border);border-left-width:3px;background:var(--g-raised);overflow-wrap:anywhere}
    .vz-grafo-item.mt{border-left-color:#ff6478}.vz-grafo-item.bt{border-left-color:#54e1a3}
    .vz-grafo-item small{display:block;color:var(--g-muted)}
    .vz-grafo-pontas{display:flex;flex-wrap:wrap;gap:4px}
    .vz-grafo-pontas .vz-btn{flex:1 1 auto;padding:4px;white-space:normal}
    .vz-grafo-item.selecionado{border-color:var(--g-accent);background:var(--g-accent-bg)}
    .vz-grafo-item:not(.selecionado) .vz-grafo-consulta{display:none}
    .vz-grafo-cabo-titulo{width:100%;text-align:left;overflow-wrap:anywhere}
    .vz-mapa [data-cabo-mapa].selecionado .vz-aresta{stroke-width:5;filter:drop-shadow(0 0 3px #d9cbff)}
    .vz-mapa [data-no].ficticio .vz-marcador{stroke:#e3c9ff;stroke-dasharray:3 2}
    .vz-edicao-cartao .vz-composer-form{padding:6px 0;gap:6px;overflow:visible}
    .vz-edicao-cartao .vz-btn{padding:4px 7px}
    .vz-no-card-foco{border-color:var(--g-accent)!important}
    .vz-no-controles{display:flex;align-items:center;gap:4px;flex-wrap:wrap;margin:3px 0}
    .vz-no-dialog .vz-no-conteudo .vz-btn{font-size:11px;padding:3px 6px;line-height:1.25}
    .vz-grafo-ferramentas select{color:var(--g-text);background:var(--g-bg);border:1px solid var(--g-border);border-radius:6px;padding:5px}

    .vz-mapa [data-no].avulsa .vz-marcador{stroke-dasharray:2 1}
    .vz-mapa [data-no] .vz-ficticio-rotulo{fill:#e3c9ff}
    .vz-no-dialog.vz-composer-dialog{width:min(460px,94vw)}
    .vz-composer-form{display:grid;gap:10px;padding:14px 20px 18px;min-height:0;overflow:auto}
    .vz-composer-form label{display:grid;gap:3px}
    .vz-composer-form input,.vz-composer-form select,.vz-no-editor select{width:100%;min-width:0;color:var(--g-text);background:var(--g-bg);border:1px solid var(--g-border);border-radius:6px;padding:7px}
    .vz-no-editor{display:grid;gap:5px;margin-top:5px}
    .vz-no-editor small{color:var(--g-muted)}
    .vz-composer-erro{margin:0;color:#ffb7ae}
    .vz-composer-erro:empty{display:none}
    .vz-ficticio-opcoes{display:flex;flex-wrap:wrap;gap:4px}
    .vz-grafo-revisao{display:flex;align-items:center;flex-wrap:wrap;gap:8px;margin-top:8px;padding:8px 10px;background:var(--g-panel);border:1px solid var(--g-border);border-radius:8px}
    .vz-grafo-revisao details{flex-basis:100%}.vz-grafo-revisao ul{max-height:130px;overflow:auto;margin:6px 0}
    .vz-mapa [data-arrastar]{cursor:grab}.vz-mapa.arrastando [data-arrastar]{cursor:grabbing}
    .vz-mapa [data-no].alvo .vz-halo{fill:#bb99ff;opacity:.9}
    .vz-mapa .vz-alvo-aresta{fill:none;stroke:transparent;stroke-width:14;pointer-events:stroke}
    .vz-mapa [data-cabo-mapa]:focus-visible .vz-aresta{stroke-width:6}
    .vz-mapa .vz-aresta.sugerida{stroke:#e9ae68;stroke-width:3;pointer-events:none}
    .vz-mapa [data-no].aproximada .vz-marcador{stroke:#e9ae68;stroke-dasharray:2 2}.vz-mapa [data-no].aproximada .vz-halo{fill:#e9ae6833}
    .vz-mapa [data-no].escolhido .vz-halo{fill:#5fb3ff;opacity:.8}.vz-mapa [data-no].origem-vao .vz-halo{fill:#ffd166;opacity:.95}
    .vz-mapa [data-no].iman .vz-halo{fill:#ffd166;opacity:.95}.vz-mapa [data-no].iman .vz-marcador{stroke:#ffd166}
    .vz-mapa .vz-camada-fundo{pointer-events:none}.vz-mapa .vz-fundo-rede{fill:none;stroke:#9a93ad;stroke-opacity:.3;stroke-width:1;vector-effect:non-scaling-stroke}
    .vz-mapa .vz-fundo-rede.mt{stroke:#ff6478;stroke-opacity:.26}.vz-mapa .vz-fundo-rede.bt{stroke:#54e1a3;stroke-opacity:.26}
    .vz-mapa .vz-fundo-bloco{fill:#17161c;stroke:#c9c0dc;stroke-opacity:.6;stroke-width:1.2;vector-effect:non-scaling-stroke;pointer-events:all;cursor:copy}
    .vz-mapa .vz-fundo-bloco:hover{stroke:#ffd166;stroke-opacity:1}
    .vz-mapa .vz-alca{fill:#ffd166;stroke:#17161c;stroke-width:1.5;vector-effect:non-scaling-stroke;cursor:crosshair}
    .vz-mapa .vz-elastico{stroke:#ffd166;stroke-width:2;stroke-dasharray:5 4;fill:none;vector-effect:non-scaling-stroke;pointer-events:none}
    .vz-mapa .vz-retangulo{fill:#5fb3ff1f;stroke:#5fb3ff;stroke-dasharray:4 3;vector-effect:non-scaling-stroke;pointer-events:none}
    .vz-mapa.tracando,.vz-mapa.tracando [data-no]{cursor:crosshair}
    .vz-btn.ligado{background:var(--g-accent-bg)!important;border-color:var(--g-accent)!important}
    .vz-pend{border:1px solid var(--g-border);border-radius:8px;padding:6px 8px;background:var(--g-raised);display:grid;gap:6px;max-height:46%;overflow:auto;flex:none}
    .vz-pend[hidden]{display:none}.vz-pend summary{cursor:pointer;font-weight:600}
    .vz-pend-barra{display:grid;gap:2px;font-size:11px;color:var(--g-muted)}.vz-pend-barra i{display:block;height:5px;border-radius:3px;background:var(--g-border);overflow:hidden}
    .vz-pend-barra i b{display:block;height:100%;background:#54e1a3}
    .vz-pend ul{list-style:none;margin:0;padding:0;display:grid;gap:4px}.vz-pend li{display:grid;gap:3px;font-size:11px;padding:4px 0;border-top:1px solid var(--g-border)}
    .vz-pend li > button.vz-pend-alvo{all:unset;cursor:pointer;color:var(--g-text);overflow-wrap:anywhere}.vz-pend li > button.vz-pend-alvo:focus-visible{outline:1px solid var(--g-accent)}
    .vz-pend-acoes{display:flex;gap:4px;flex-wrap:wrap}.vz-pend-acoes .vz-btn{font-size:11px;padding:2px 6px}
    .vz-grafo-mais{width:100%}
    .vz-ops{display:inline-flex;gap:2px;background:var(--g-bg);border:1px solid var(--g-border);border-radius:7px;padding:2px;flex:none}
    .vz-op-chip{min-width:26px;padding:3px 5px;border:0;border-radius:5px;background:transparent;color:var(--g-muted);font:700 11px/1.2 Inter,system-ui,sans-serif;cursor:pointer}
    .vz-op-chip.sel{background:var(--g-accent-bg);color:#fff;box-shadow:inset 0 0 0 1px var(--g-accent)}.vz-op-chip.f.sel{background:#4a3a1c;color:#f6d39a;box-shadow:inset 0 0 0 1px #e9ae68}
    .vz-op-chip.r.sel{background:#4a1f27;color:#ffb3bd;box-shadow:inset 0 0 0 1px #ff6478}.vz-op-chip.m.sel{background:#1f3a4a;color:#9fd4ff;box-shadow:inset 0 0 0 1px #5fb3ff}
    .vz-op-chip:focus-visible{outline:2px solid var(--g-accent)}.vz-ops-sep{width:1px;background:var(--g-border);margin:2px}
    .vz-ponto-dialog.vz-no-dialog.vz-composer-dialog{width:min(720px,96vw)}
    .vz-nome-linha{display:grid;grid-template-columns:minmax(120px,190px) 14px 1fr;gap:6px;align-items:end}.vz-sep{color:var(--g-muted);font-size:18px;text-align:center;padding-bottom:6px}
    .vz-previa-titulo{font:12px ui-monospace,Consolas,monospace;color:var(--g-accent);background:var(--g-panel);border:1px dashed var(--g-border);border-radius:6px;padding:5px 8px;overflow-wrap:anywhere}
    .vz-tipo-ponto{font-size:11px;border:1px solid var(--g-border);border-radius:20px;padding:1px 9px;width:max-content;margin-left:6px;vertical-align:middle;font-weight:600}
    .vz-tipo-ponto.poste{border-color:#54e1a3;color:#54e1a3}.vz-tipo-ponto.avulso{border-color:#e9ae68;color:#e9ae68}.vz-tipo-ponto.fic{border-color:#e3c9ff;color:#e3c9ff}.vz-tipo-ponto:empty{display:none}
    .vz-rotulo-linhas{font-size:12px;color:var(--g-muted)}
    .vz-linhas-ponto{display:grid;gap:5px}.vz-lp{display:grid;grid-template-columns:18px 1fr auto 22px;gap:6px;align-items:center}
    .vz-lp-n{color:var(--g-muted);font-size:11px;text-align:right}
    .vz-lp-x{background:transparent;border:0;color:var(--g-muted);cursor:pointer;font-size:14px}.vz-lp-x:hover{color:#ff8293}
    .vz-secao-linhas h5{display:flex;align-items:center;gap:8px}.vz-secao-linhas h5 .vz-btn{margin-left:auto}
    .vz-lps{display:grid;gap:4px;margin-bottom:10px}
    .vz-lp-painel{display:grid;grid-template-columns:34px 1fr auto 20px;gap:8px;align-items:center;padding:4px 8px;border:1px solid var(--g-border);border-radius:8px;background:var(--g-panel);cursor:pointer}
    .vz-lp-painel:hover,.vz-lp-painel:focus-visible{border-color:var(--g-accent);outline:none}.vz-lp-painel.fic{border-style:dashed}.vz-lp-painel.bloqueada{cursor:default;opacity:.75}
    .vz-lp-painel.editando{grid-template-columns:1fr auto auto auto;border-color:var(--g-accent);cursor:default}
    .vz-lp-painel input{background:var(--g-bg);color:var(--g-text);border:1px solid var(--g-border);border-radius:6px;padding:5px 7px;min-width:0}
    .vz-lp-painel small{color:var(--g-muted);font-size:11px}.vz-lp-texto{overflow-wrap:anywhere}
    .vz-op-badge{display:inline-block;min-width:28px;text-align:center;font-weight:700;font-size:11px;border-radius:5px;padding:2px 5px;background:var(--g-accent-bg);color:#fff}
    .vz-op-badge.r{background:#4a1f27;color:#ffb3bd}.vz-op-badge.m{background:#1f3a4a;color:#9fd4ff}.vz-op-badge.f{background:#4a3a1c;color:#f6d39a}
    .vz-aviso-antigo{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:8px 12px;margin-bottom:10px;border:1px solid #946d34;background:#2b2418;border-radius:9px;color:#f4dcb2}
    .vz-aviso-antigo[hidden]{display:none}.vz-aviso-antigo span{flex:1 1 260px}
    .vz-no-cabo-sel{cursor:pointer}.vz-no-cabo.escolhido{border-color:#ffd166;box-shadow:0 0 0 1px #ffd166}
    .vz-dica-vinculo{font-size:11px;margin:0 0 8px}
    .vz-no-topo h4 small{font-weight:400;color:var(--g-muted);font-size:14px}
    .vz-esc-sel{background:var(--g-bg);color:var(--g-text);border:1px solid var(--g-border);border-radius:6px;padding:4px 6px;font-size:12px;max-width:128px}
    .vz-grafo-escala{position:absolute;bottom:12px;left:12px;display:flex;align-items:center;gap:8px;padding:5px 9px;background:var(--g-panel);border:1px solid var(--g-border);border-radius:9px;font-size:11px;color:var(--g-text);pointer-events:none}
    .vz-grafo-escala[hidden]{display:none}.vz-esc-barra{display:block;height:6px;border:1.5px solid var(--g-text);border-top:0;min-width:2px}.vz-esc-barra-rot{color:var(--g-muted,#8b949e)}
    .vz-mapa.vz-movendo text{display:none}
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
    .vz-mapa .vz-aresta.diverge{stroke-dasharray:10 3 2 3}.vz-grafo-alerta{display:block;margin:2px 0 0;font-size:11px;color:#e9ae68}
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
    .vz-no-topo-acoes{display:flex;gap:6px;margin-left:auto;align-items:center;justify-content:flex-end;flex-wrap:wrap}
    .vz-no-conteudo{min-height:0;overflow:auto;padding:16px 20px}
    .vz-no-rede{display:grid;grid-template-columns:minmax(0,1.45fr) minmax(130px,1fr);align-items:start;gap:36px;position:relative}
    .vz-no-linhas{position:absolute;inset:0;width:100%;height:100%;overflow:visible;pointer-events:none}
    .vz-no-rede .vz-no-acoes select{min-width:0}
    .vz-no-seletor{width:100%;margin:0 0 8px;padding:6px;background:var(--g-panel);color:var(--g-text);border:1px solid var(--g-border);border-radius:7px}
    .vz-no-conteudo h5{font-size:13px;margin-bottom:10px}
    .vz-no-filtros{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px}
    .vz-no-filtros label{display:flex;align-items:center;gap:5px;margin-right:auto;color:var(--g-text)}
    .vz-no-filtros input{accent-color:var(--g-accent)}
    .vz-no-filtros .vz-btn{font-size:11px;padding:3px 7px}
    .vz-no-detalhes{margin-top:14px;padding:8px 10px;border-top:1px solid var(--g-border);color:var(--g-muted)}
    .vz-no-detalhes summary,.vz-no-mais summary{cursor:pointer;width:max-content;max-width:100%}
    .vz-no-detalhes p{margin:6px 0;font-size:12px}
    .vz-no-linha{display:flex;align-items:center;gap:5px;flex-wrap:wrap;min-height:21px}
    .vz-no-linha > button{font-weight:600}
    .vz-no-linha .vz-edicao-cartao,.vz-no-estado{margin-left:auto}
    .vz-no-ocupacao,.vz-no-estado{font-size:11px;font-variant-numeric:tabular-nums;color:var(--g-muted)}
    .vz-no-mais{font-size:11px;color:var(--g-muted);margin-top:3px}
    .vz-no-mais[open] summary{margin-bottom:5px}
    .vz-no-conectar{margin-bottom:7px}
    .vz-no-rotulo{font-size:11px;color:var(--g-muted);overflow-wrap:anywhere}
    .vz-no-cabo.pendente{border-color:#946d34}
    .vz-no-cabo.pendente .vz-no-estado,.vz-no-cabo.pendente summary{color:#f4c888}
    .vz-no-cabo.vinculado .vz-no-estado{color:#73d6ab}
    .vz-no-estrutura .vz-no-ocupacao.pendente{color:#f4c888}
    .vz-no-cabo.destaque,.vz-no-estrutura.destaque{border-color:var(--g-accent);box-shadow:0 0 0 1px var(--g-accent)}
    .vz-no-linha .vz-nivel{font-size:10px;font-weight:700;color:var(--g-muted)}
    .vz-no-linha .vz-nivel.mt{color:#ff8293}.vz-no-linha .vz-nivel.bt{color:#73d6ab}
    .vz-no-linha .vz-op{font-size:10px}
    .vz-ponto{display:grid;gap:8px;margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid var(--g-border)}
    .vz-ponto label{display:grid;gap:3px;font-size:12px;color:var(--g-text)}.vz-ponto input{background:var(--g-bg);color:var(--g-text);border:1px solid var(--g-border);border-radius:6px;padding:5px 7px}
    .vz-mapa.posicionando{cursor:crosshair}.vz-mapa.posicionando [data-arrastar]{pointer-events:none}
    .vz-ponto-botoes{display:flex;flex-wrap:wrap;gap:6px}.vz-ponto h6{margin:4px 0 0;font-size:11px;color:var(--g-muted,#8b949e);text-transform:uppercase;letter-spacing:.04em}
    .vz-ponto-linhas{list-style:none;margin:0;padding:0;display:grid;gap:3px;font-size:12px}.vz-ponto-linhas li{display:flex;gap:6px;align-items:baseline;overflow:hidden}
    .vz-ponto-linhas small{flex:0 0 auto;color:var(--g-muted,#8b949e)}.vz-ponto-linhas code{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .vz-no-estruturas{display:grid;gap:5px}.vz-no-estrutura{padding:5px 7px;border:1px solid var(--g-border);border-radius:8px;background:var(--g-panel);overflow-wrap:anywhere}
    .vz-no-estrutura span,.vz-no-ajuda{color:var(--g-muted)}
    .vz-no-estrutura.instalando{border-color:#2ea043;border-left-width:3px;background:#0f2a1a}.vz-no-estrutura.instalando .vz-no-linha>button:first-of-type{color:#7ee787;font-weight:700;font-size:14px}
    .vz-no-estrutura.removendo{opacity:.55}.vz-no-estrutura.removendo strong{text-decoration:line-through}
    .vz-no-estrutura .vz-op-badge{flex:none}
    .vz-no-cabo{padding:6px 8px;border:1px solid var(--g-border);border-radius:8px;margin-bottom:4px;background:var(--g-panel)}
    .vz-no-cabo h6{margin:0;font-size:12px;overflow-wrap:anywhere}.vz-no-cabo p{margin:0 0 3px}
    .vz-no-cabo .vz-op{margin-right:4px}
    .vz-no-acoes{display:flex;gap:7px;align-items:center;flex-wrap:wrap;margin-top:4px}
    .vz-no-acoes select{min-width:140px;max-width:100%;background:var(--g-bg);color:var(--g-text);border:1px solid var(--g-border);border-radius:7px;padding:5px 7px}
    .vz-no-acoes .vz-no-ajuda{flex-basis:100%;font-size:11px;color:#e1bc8d}.vz-no-acoes .vz-no-ajuda:empty{display:none}
    .vz-no-status{margin:0 0 8px;color:var(--g-accent)}.vz-no-status:empty{display:none}
    @media(max-width:630px){
        .vz-no-dialog .vz-no-topo{padding:12px;align-items:flex-start}
        .vz-no-conteudo{padding:12px}.vz-no-rede{grid-template-columns:minmax(0,1fr) minmax(94px,30%);gap:14px}
        .vz-no-dialog .vz-no-topo{flex-wrap:wrap}.vz-no-topo-acoes{margin-left:0;justify-content:flex-start}
        .vz-grafo-area{grid-template-columns:minmax(0,1fr) minmax(107px,34%);gap:5px}
        .vz-grafo-cabos{padding:5px}.vz-grafo-lista .vz-btn{padding:5px}
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
        cabos().forEach(c => { s += c ? `${c.ativo || ''}|${c.operacao || ''}|${!!c.comentario}|${c.vinculoEstruturas || ''}|${c.vinculoPostes || ''}|${c.estruturasFicticias || ''}|${c.textoOriginal || ''}|${c.fases || ''}|${c.tensao || ''}|${c.nivel || ''}\n` : '\n'; });
        outros().forEach(o => { s += o ? `${o.ativo || ''}|${o.entidade || ''}|${o.operacao || ''}|${!!o.congelada}|${o.cabosFicticios || ''}|${o.cabosFicticiosSlots ? JSON.stringify(o.cabosFicticiosSlots) : ''}|${o.origemAvulso || ''}\n` : '\n'; });
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
    let abrindo = false;
    async function abrir(h, opcoes) {
        if (abrindo) return;
        if (overlay && overlay.style.display !== 'none') {
            if (host === h && st.projeto === h.projeto() && !opcoes?.aba) {
                if (raiz.AbrirNaApp?.nativaAtiva?.('mapa')) { toast('O Mapa está aberto em outra janela.'); return; }
                raiz.AbrirNaApp?.trazerConteudo?.('mapa');
                overlay.focus();
                return;
            }
            if (!fechar()) return;
        }
        abrindo = true;
        host = h;
        const projeto = host.projeto();
        if (!doc.getElementById('vz-css')) doc.head.appendChild(el('style', { id: 'vz-css', text: CSS }));
        let cfg = raiz.__vinculacaoConfig;
        if (!cfg) {
            try { const r = await fetch('/api/vinculacao-config?projeto_codigo=' + encodeURIComponent(projeto)); cfg = r.ok ? (await r.json()).config : null; } catch (e) { cfg = null; }
        }
        abrindo = false;
        if (projeto !== host.projeto()) return;
        st = { projeto, cfg: clone(VR.normalizarConfig(cfg)), aba: (opcoes && opcoes.aba) || 'mapa', sel: null, filtro: 'todos', filtroCP: 'todos', kbModo: 'c', kbP: null, elAch: null, auto: true, cfgVer: 0, resultadoPoste: null, elRes: null, dirty: false, scriptIdx: 0, previa: null, textoScript: null, erroScript: '', teste: {} };
        if (!overlay) {
            // TASK-241: Esc no Mapa só deseleciona; na janela (visualização) nunca fecha — só o modal antigo (sem AbrirNaApp) fecha fora do mapa
            overlay = el('div', { class: 'vz-ov', onkeydown: e => { if (tecladoCabosPostes(e)) return; if (e.defaultPrevented || e.target.closest('dialog')) return; if (e.key === 'Escape') { if (menuAberto) fecharMenu(); else if (st.aba !== 'mapa' && !overlay.classList.contains('vz-destacada')) fechar(); } } });
            overlay.setAttribute('tabindex', '-1');
            doc.body.appendChild(overlay);
        }
        if (!overlay.isConnected) doc.body.appendChild(overlay);
        overlay.style.display = 'flex';
        montarJanela();
        cx = null; painel = null; metaAtivo.clear();
        render();
        overlay.focus();
    }

    // TASK-241: o Mapa abre como as janelas de Resultado e Base de Orçamento (área principal, minimiza para o rodapé, maximizar/restaurar
    // e «Abrir em nova janela» do Windows). Na janela nativa, o Resumo roda em modo janela (__modoJanelaMapa) e mostra só o Mapa, a página toda.
    let assinaturaNativa = '', envioNativa = 0;
    function montarJanela() {
        if (raiz.__modoJanelaMapa) { overlay.classList.add('vz-destacada', 'vz-pagina'); return; }
        if (!raiz.AbrirNaApp?.abrirConteudo) return;                        // sem a moldura do programa: o modal de sempre
        overlay.classList.add('vz-destacada');
        assinaturaNativa = '';
        raiz.AbrirNaApp.abrirConteudo('mapa', 'Mapa', overlay, {
            aoFechar: fechar,
            capturar: () => (host?.instantaneo ? { dados: host.instantaneo(), fundo: host.fundoCad?.() || null } : null),
            restaurar: estado => {
                if (estado?.dados) aplicarDaNativa(estado.dados);
                raiz.requestAnimationFrame(() => atualizarObra(false, true));      // voltou da janela nativa: redesenha com o que mudou lá
            },
            aoDados: dados => aplicarDaNativa(dados)
        });
    }
    /** A janela nativa mudou a obra: aplica aqui (um passo de Ctrl+Z) sem devolver o mesmo estado para ela. */
    function aplicarDaNativa(dados) {
        if (!dados || !host?.aplicarInstantaneo) return;
        assinaturaNativa = JSON.stringify(dados);
        host.aplicarInstantaneo(dados);
    }
    /** A obra mudou nesta janela com o Mapa aberto na janela nativa: manda o estado para lá (uma vez por quadro, só se mudou). */
    function sincronizarNativa() {
        if (!raiz.AbrirNaApp?.nativaAtiva?.('mapa') || !host?.instantaneo || envioNativa) return;
        envioNativa = raiz.requestAnimationFrame(() => {
            envioNativa = 0;
            const dados = host.instantaneo(), assinatura = JSON.stringify(dados);
            if (assinatura === assinaturaNativa) return;
            assinaturaNativa = assinatura;
            raiz.AbrirNaApp.enviarNativa('mapa', 'dados', { dados });
        });
    }

    let atualizacao = 0;
    function atualizarObra(substituida = false, forcar = false) {
        if (!overlay?.classList.contains('vz-destacada') || overlay.style.display === 'none') return;
        sincronizarNativa();
        if (raiz.AbrirNaApp?.nativaAtiva?.('mapa')) return;                 // o Mapa está na janela nativa: aqui fica escondido
        if (substituida) {
            overlay.querySelectorAll('.vz-no-dialog[open]').forEach(d => d.close());
            st.grafoNo = null;
            st.grafoVista = null;
            st.sel = null;
        } else if (!forcar && overlay.contains(doc.activeElement)) return;
        if (atualizacao) return;
        atualizacao = raiz.requestAnimationFrame(() => {
            atualizacao = 0;
            if (!overlay?.classList.contains('vz-destacada')) return;
            cx = null;
            metaAtivo.clear();
            render();
        });
    }

    function fechar() {
        if (st && st.dirty && !raiz.confirm('Há alterações nas tabelas/scripts que não foram salvas. Fechar mesmo assim?')) return false;
        fecharMenu();
        overlay?.querySelectorAll('.vz-no-dialog[open]').forEach(d => d.close());
        if (painel) { painel.destruir(); painel = null; }
        mapaDesmontar?.();
        mapaDesmontar = mapaAtualizar = null;
        if (overlay) {
            overlay.style.display = 'none';
            if (raiz.__modoJanelaMapa) { raiz.AbrirNaApp?.voltarOuFechar?.(); }      // janela nativa: fechar o Mapa fecha a janela (a obra volta antes)
            else if (overlay.classList.contains('vz-destacada')) {
                overlay.classList.remove('vz-destacada');
                raiz.AbrirNaApp?.removerConteudo?.('mapa');
            }
        }
        if (host && host.depois) host.depois();
        return true;
    }

    function marcarSujo() { st.dirty = true; st.cfgVer++; const b = overlay.querySelector('#vz-salvar'); if (b) b.disabled = false; }

    async function salvar() {
        if (st.projeto !== host.projeto()) { render(); return false; }
        const sessao = st;
        const erros = [];
        try {
            const r = await fetch('/api/vinculacao-config', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ projeto_codigo: sessao.projeto, config: sessao.cfg }) });
            if (st !== sessao || !overlay?.isConnected || overlay.style.display === 'none') return r.ok;
            if (r.ok && sessao.projeto !== host.projeto()) { sessao.dirty = false; render(); return true; }
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
        mapaDesmontar?.();
        mapaDesmontar = mapaAtualizar = null;
        st.elRes = null; st.elAch = null;
        overlay.innerHTML = '';
        if (st.projeto !== host.projeto()) {
            overlay.appendChild(el('div', { class: 'vz' }, [
                el('div', { class: 'vz-top' }, [el('h3', { text: 'Mapa' }), el('button', { class: 'vz-x', text: '✕', 'aria-label': 'Fechar', onclick: fechar })]),
                el('div', { class: 'vz-corpo' }, [el('p', { role: 'status', text: 'O projeto do Resumo mudou. Feche esta janela e abra o Mapa do projeto atual.' })])
            ]));
            return;
        }
        const abas = [['mapa', 'Mapa', ''], ['vinculo', 'Vínculo', ''], ['cabopostes', 'Cabos ⇄ Postes', ''], ['cabos', 'Tabela de Cabos', st.cfg.cabos.length], ['estruturas', 'Tabela de Estruturas', st.cfg.estruturas.length], ['dicionario', 'Dicionário de cabos', (st.cfg.dicionario_cabos || []).length], ['postes', 'Postes', st.cfg.postes ? '✓' : ''], ['desenho', 'Desenho', ''], ['exportacao', 'Exportação', st.cfg.exportacao ? '✓' : ''], ['casos', 'Casos', ''], ['pacote', 'Pacote', ''],
            ['script', 'Script', st.cfg.scripts.length]];
        const vz = el('div', { class: st.aba === 'mapa' ? 'vz vz--grafo' : 'vz' }, [
            el('div', { class: 'vz-top' }, [el('h3', { text: '🗺 Mapa' }), el('span', { class: 'vz-chip', text: 'Projeto ' + host.projeto() }),
                st.dirty ? el('span', { class: 'vz-chip', style: 'color:#d29922;border-color:#d29922', text: 'alterações não salvas' }) : null,
                el('button', { class: 'vz-x', text: '✕', 'aria-label': 'Fechar mapa', onclick: fechar })]),
            el('div', { class: 'vz-tabs' }, abas.map(([id, nome, n]) => el('button', { class: 'vz-tab' + (st.aba === id ? ' on' : ''), onclick: () => { st.aba = id; st.sel = null; render(); } }, [nome, id === 'vinculo' ? el('b', { id: 'vz-pend', text: String(n), style: n === '' ? 'display:none' : '' }) : (n !== '' ? el('b', { text: String(n) }) : null)]))),
            el('div', { class: 'vz-corpo', id: 'vz-corpo' }),
            rodape()]);
        overlay.appendChild(vz);
        const corpo = vz.querySelector('#vz-corpo');
        try {
            const a = analise();
            atualizarPendencias(a);
            ({ vinculo: abaVinculo, cabopostes: abaCabosPostes, cabos: abaCabos, estruturas: abaEstruturas, dicionario: abaDicionario, postes: abaPostes, desenho: abaDesenho, exportacao: abaExportacao, casos: abaCasos, pacote: abaPacote, script: abaScript, mapa: abaMapa })[st.aba](corpo, a);
        } catch (erro) {
            console.error('Falha ao exibir a vinculação de cabos e estruturas.', erro);
            if (painel) { painel.destruir(); painel = null; }
            mapaDesmontar?.();
            mapaDesmontar = mapaAtualizar = null;
            corpo.replaceChildren(el('p', { class: 'vz-erro', role: 'alert', text: 'Não foi possível exibir esta aba da vinculação. Use Concluir para voltar ao Resumo.' }));
        }
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
    function infoCabo(ic) { const k = cx || ctx(), c = cabos()[ic]; return k.infoC[ic] || (k.infoC[ic] = VR.infoCabo(host.prefixoCabo(c.ativo) || '', c.operacao, st.cfg, VR.extraCabo(c, st.cfg.dicionario_cabos))); }
    function infoEst(ie) { const k = cx || ctx(); return k.infoE[ie] || (k.infoE[ie] = VR.infoEstrutura(codigoDe(ie), st.cfg)); }
    function ocupantes(ie, ignorandoCabo) {
        const k = cx || ctx(), lista = [];
        (k.ocup[ie] || []).forEach(ic => { if (ic !== ignorandoCabo) lista.push({ info: infoCabo(ic), ficticio: false }); });
        const e = infoEst(ie);
        ficticiosDe(ie).forEach(cod => lista.push({ info: VR.infoCaboFicticio(cod, e.codigo, VR.opFicticio(e), st.cfg), ficticio: true }));
        return lista;
    }
    function podeLigar(ic, ie) {
        if (!estruturaDisponivel(ie)) return [false, 'A estrutura está fora da vinculação.'];
        if (!infoEst(ie).ocupa) return [false, 'O equipamento lê os cabos do ponto e não recebe pontas.'];
        const c = infoCabo(ic), e = infoEst(ie), oc = ocupantes(ie, ic);
        return VR.podeComOcupantes(c, e, oc);
    }

    function atualizarPendencias(a) {
        const pend = Object.values(a.cabos).filter(c => c.estado !== 'gr').length + Object.values(a.estruturas).filter(e => e.estado === 'k' || e.estado === 'pu').length;
        const badge = overlay.querySelector('#vz-pend');
        badge.textContent = pend ? pend + ' pend.' : '';
        badge.style.display = pend ? '' : 'none';
    }

    /** Atualização leve: refaz só os dados do painel (a análise é memoizada) em vez de reconstruir o modal inteiro. */
    function atualizar() {
        if (!painel && !mapaAtualizar) { render(); return; }
        const k = ctx();
        if (painel) painel.refazer();
        atualizarPendencias(k.a);
        if (st.elRes) st.elRes.textContent = textoResumoPoste();
        if (mapaAtualizar) mapaAtualizar();
        atualizarAchados();
    }

    function estruturaDisponivel(ie) { return Boolean((cx || ctx()).a.estruturas[ie]); }

    function motivoParaLigar(ic, ie) {
        const c = cabos()[ic];
        if (!VR.ehCaboValido(c) || c.comentario) return 'O cabo está fora da vinculação.';
        const info = infoCabo(ic);
        if (!info.vincula) return `O cabo ${info.codigo} não vincula (acompanha o cabo principal).`;
        if ((c.vinculoEstruturas || []).length + VR.estruturasFicticiasDe(c).length >= info.pontasMax)                  // TASK-234: o manual vai até pontasMax (cabo M: 2ª ponta)
            return `O cabo ${info.codigo} já tem as ${info.pontasMax} ponta(s); remova uma antes.`;
        const [ok, motivo] = podeLigar(ic, ie);
        return ok ? '' : motivo || 'Estrutura incompatível.';
    }

    function ligar(ic, ie, antesDeAlterar) {
        const c = cabos()[ic], ja = (c.vinculoEstruturas || []).includes(ie);
        const motivo = ja ? '' : motivoParaLigar(ic, ie);
        if (motivo) { toast('✖ ' + motivo); return; }
        host.antes();
        antesDeAlterar?.();
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
        return codigos.map(cod => { const ic = VR.infoCaboFicticio(cod, e.codigo, VR.opFicticio(e), st.cfg), [ok, motivo] = VR.podeComOcupantes(ic, e, oc, true); return { cod, ok, motivo, combina: ok && oc.length > 0 && oc[0].info.codigo === cod }; })
            .sort((a, b) => (b.combina - a.combina) || (b.ok - a.ok));
    }

    function addFicticio(ie, entrada) {
        const cod = trim(entrada).toUpperCase();
        if (!cod || cod.length > 120 || /[\r\n]/.test(cod)) return 'Informe o código do cabo fictício.';
        if (!estruturaDisponivel(ie) || !infoEst(ie).ocupa) return 'Esta estrutura não recebe cabos.';
        const e = infoEst(ie);
        const info = VR.infoCaboFicticio(cod, e.codigo, VR.opFicticio(e), st.cfg);
        const [ok, motivo] = VR.podeComOcupantes(info, e, ocupantes(ie), true);
        if (!ok) return motivo || 'Cabo fictício incompatível.';
        host.antes();
        gravarFicticios(ie, ficticiosDe(ie).concat(cod));
        atualizar();
        host.depois?.();
        return '';
    }

    function removerFicticio(ie, k) {
        host.antes();
        gravarFicticios(ie, ficticiosDe(ie).filter((_, i) => i !== k));
        atualizar();
        host.depois?.();
    }

    function menuEstrutura(e, ie) {
        const info = infoEst(ie), a = analise().estruturas[ie], cands = candidatosFicticios(ie);
        const itens = [{ hd: `Estrutura ${info.codigo} · ${a ? a.tem : 0}${a && a.precisa ? '/' + faixaCabos(a) : ''} cabo(s)` }];
        const aceitos = cands.filter(c => c.ok);
        itens.push({ hd: 'Associar cabo fictício' });
        if (!aceitos.length) itens.push({ t: 'Nenhum cabo aceito por esta estrutura agora', off: true });
        aceitos.slice(0, 8).forEach(c => itens.push({ t: '➕ ' + c.cod, rec: c.combina ? '✔ combina com o outro lado' : '', fn: () => { const erro = addFicticio(ie, c.cod); if (erro) toast('✖ ' + erro); } }));
        itens.push({ t: '✎ Outro código…', fn: () => { const cod = raiz.prompt('Código do cabo fictício (ex.: C16):'); if (cod) { const erro = addFicticio(ie, cod); if (erro) toast('✖ ' + erro); } } });
        const recusados = cands.filter(c => !c.ok && c.motivo);
        if (recusados.length) itens.push({ t: `(${recusados.length} código(s) não aceito(s): ${recusados.slice(0, 3).map(c => c.cod).join(', ')}…)`, off: true, title: recusados.slice(0, 5).map(c => c.cod + ': ' + c.motivo).join('\n') });
        itens.push('-');
        ficticiosDe(ie).forEach((cod, k) => itens.push({ t: `✖ Remover fictício ${cod}`, fn: () => removerFicticio(ie, k) }));
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
        const c = cabos()[ic], info = infoCabo(ic), fict = VR.estruturasFicticiasDe(c), cheio = (c.vinculoEstruturas || []).length + fict.length >= info.pontasMax;
        const itens = [{ hd: 'Cabo ' + (c.ativo || '') }, { t: 'Desvincular este cabo', fn: () => { host.antes(); c.vinculoEstruturas = []; atualizar(); } }, '-', { hd: 'Estrutura fictícia' }];
        const add = cod => { host.antes(); c.estruturasFicticias = fict.concat([cod.toUpperCase()]); atualizar(); };
        if (cheio) itens.push({ t: `O cabo já tem as ${info.pontasMax} ponta(s)`, off: true });
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
        if (!ja && at.length >= info.pontasMax) { toast(`O cabo ${info.codigo} já chega a ${info.pontasMax} poste(s); remova um antes.`); return; }
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
                if (selP !== null) { if (at.includes(selP) || (info.vincula && at.length < info.pontasMax)) cls += ' glow'; else { cls += ' vl-dim'; tip = info.vincula ? 'o cabo já chega ao máximo de postes' : 'cabo que não vincula'; } }
                const op = trim(c.operacao).toUpperCase(), nomes = at.map(l => (postesDe(l).length ? postesDe(l).join(' ') : '#' + (l + 1))).join(' · ');
                const es = (c.vinculoEstruturas || []).map(ie => codigoDe(ie)).filter(Boolean).join(' · ');
                return `<div class="${cls}" data-c="${it.ic}" style="top:${it.y}px;height:${it.h - 4}px" title="${esc(tip || c.ativo || '')}"><span class="vz-op${opClasse(op)}">${esc(op || 'I')}</span><span class="t">${esc(c.ativo || '')}</span>`
                    + `<span class="pp">${nomes ? '⇒ ' + esc(nomes) : (info.vincula ? '—' : 'não vincula')}</span><span class="r" style="color:${COR_T[r.estado]}">${es ? esc(es) : ''}</span></div>`;
            },
            htmlD(it) {
                const k = cx, a = k.a, o = outros()[it.l], selC = st.sel && st.sel.t === 'c' ? st.sel.i : null, selP = st.sel && st.sel.t === 'p' ? st.sel.i : null;
                let cls = 'vl-pst';
                if (selC !== null) { const cb = cabos()[selC], at = VR.postesAtribuidos(cb), info = infoCabo(selC); if (at.includes(it.l) || (info.vincula && at.length < info.pontasMax)) cls += ' glow'; else cls += ' vl-dim'; }
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
        if (tipo === 'pontas') return el('div', { style: 'display:flex;gap:4px' }, ['I', 'M'].map(k => el('input', { type: 'number', min: '0', title: campo + ' do cabo ' + k, placeholder: k, value: r[campo] && r[campo][k] != null ? r[campo][k] : '', onchange: e => { r[campo] = r[campo] || {}; if (e.target.value === '') delete r[campo][k]; else r[campo][k] = parseInt(e.target.value, 10); if (!Object.keys(r[campo]).length) delete r[campo]; marcarSujo(); } })));
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
            { campo: 'pontas', rot: 'Pontas I · M', tipo: 'pontas', w: '120px', dica: 'pontas que o vínculo automático liga (e que deixam o cabo completo); vazio = I 2 · M 1' },
            { campo: 'pontas_max', rot: 'Máx. à mão I · M', tipo: 'pontas', w: '120px', dica: 'até quantas pontas o usuário pode ligar à mão; vazio = I igual às pontas · M 2 (a 2ª ponta do cabo M é sempre manual)' }, { campo: 'estruturas', rot: 'Estruturas que aceita', tipo: 'text', extra: '^[NUB][1-4] (vazio = qualquer)' },
            { campo: 'fases', rot: 'Fases', tipo: 'text', w: '55px', extra: '3 ou 1,2', dica: 'só vale para cabos com essa quantidade de fases (do 3x do texto; sem Nx = 1); vazio = qualquer' },
            { campo: 'nivel', rot: 'Nível', tipo: 'select', extra: ['', 'MT', 'BT'], w: '65px', dica: 'MT = tensão acima de 1 kV no texto (ex.: 13,8kV); BT = sem tensão; vazio = qualquer' },
            { campo: 'padrao_texto', rot: 'Padrão no texto', tipo: 'text', extra: '13,8kV (vazio = qualquer)', dica: 'regex contra o texto ORIGINAL do cabo (com 3x e kV)' },
            { campo: 'vincula', rot: 'Vincula', tipo: 'bool', w: '60px', dica: 'desmarcado = acompanha o cabo principal (ex.: neutro CAZ)' }, { campo: 'ativa', rot: 'Ativa', tipo: 'bool', w: '50px' }, { campo: 'obs', rot: 'Observação', tipo: 'text' }],
            () => ({ padrao: '', grupo: '', valor_cabo: '', operacoes: ['I', 'M'], pontas: { I: 2, M: 1 }, estruturas: '', vincula: true, ativa: true, obs: '' }),
            txt => {
                const r = raiz.AnotacaoCabos && raiz.AnotacaoCabos.ehNotacao(txt) ? raiz.AnotacaoCabos.interpretar(txt) : null;
                const cod = r && r.itens.length ? (r.itens.find(i => !i.neutro) || r.itens[0]).ativo : (host.prefixoCabo(txt) || txt.split(' ')[0]);
                const i = VR.infoCabo(cod, 'I', st.cfg);
                return `código <b style="color:#79c0ff">${esc(i.codigo)}</b> → ${i.regra === null ? 'nenhuma linha casa (aceita tudo)' : 'linha #' + (i.regra + 1) + (i.valor_cabo ? ' · tipocabo=<b>' + esc(i.valor_cabo) + '</b>' : '') + ' · pontas I=' + VR.infoCabo(cod, 'I', st.cfg).pontas + ' M=' + VR.infoCabo(cod, 'M', st.cfg).pontas + ' (à mão até I=' + VR.infoCabo(cod, 'I', st.cfg).pontasMax + ' M=' + VR.infoCabo(cod, 'M', st.cfg).pontasMax + ')'}`;
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
        const ns = 'http://www.w3.org/2000/svg';
        const svg = doc.createElementNS(ns, 'svg');
        svg.setAttribute('class', 'vz-mapa');
        svg.setAttribute('aria-label', 'Grafo da rede');
        const posicoes = new Map(), vaosPorCabo = new Map(), nosSvg = new Map();
        const totalPostes = el('strong'), totalAvulsas = el('strong'), totalVaos = el('strong');
        const contagem = el('div', { class: 'vz-grafo-legenda' }, [
            el('span', { class: 'vz-grafo-total' }, [totalPostes, 'Postes']),
            el('span', { class: 'vz-grafo-total' }, [totalAvulsas, 'Pontos avulsos']),
            el('span', { class: 'vz-grafo-total' }, [totalVaos, 'Vãos'])
        ]);
        const ME = raiz.MapaEscala;                                    // TASK-224: escala 1:500, recorte e barra de escala (static/mapa_escala.js)
        const lerEscala = () => { try { return raiz.localStorage.getItem('vz_mapa_escala'); } catch (e) { return null; } };
        let escala = ME ? ME.escalaValida(st.grafoEscala || lerEscala() || 'ajustar') : 'ajustar';
        const escalar = () => escala !== 'ajustar';
        let kLayout = null, bases = new Map(), limites = [18, 18, W - 18, H - 18];       // px do mapa por metro do desenho · posição «de fábrica» de cada ponto · onde o ponto pode ser arrastado
        let itensRecorte = [], itensPorNo = new Map(), areaRecorte = null, recortar = false;
        let grafo, vb, arrasto, mover = false;
        const vistaPadrao = () => {
            if (!escalar()) return [0, 0, W, H];
            const larg = svg.clientWidth || 1000, alt = svg.clientHeight || 560;                  // 1 px da tela = 1 unidade do mapa: o desenho sai em 1:500
            return ME.vistaCentrada(W / 2, H / 2, larg, alt);
        };
        vb = st.grafoVista && st.grafoVistaEscala === escala ? st.grafoVista : (!escalar() ? [0, 0, W, H] : null);
        let ultimoNo, ultimoStatus = '', caboEscolhido = null, pontaEscolhida = null;
        let soPendenciasNo = false, estruturaEmFoco = null, caboSelNo = null;          // TASK-235: cabo escolhido no painel (clique no cabo e depois na estrutura)
        let desenhosPorNo = new Map(), desenhosPorCabo = new Map(), camadaVaos = null, camadaNos = null, cabosSvg = [];
        let desenhos = [], ponto, sugestoes = [], observadorNo, quadro, caboEmFoco = null, filtroMapa = 'todos';
        let cabosFiltrados = new Set(), fecharEdicaoAtiva;
        // TASK-233: montagem rápida — seleção por caixa, ferramenta «Traçar vão», alças das pontas, fundo do CAD e pendências
        const MM = raiz.MapaMontagem;
        const selecao = new Set();
        let ferramenta = null, vaoOrigem = null, camadaAlcas = null, fundoLivres = [], pendencias = [], cartoes = null;
        const lerFundo = () => { try { return raiz.localStorage.getItem('vz_mapa_fundo') !== '0'; } catch (e) { return true; } };
        let fundoLigado = lerFundo();
        const posCadNo = n => n && (n.pos || n.posAprox);
        const RAIO_BLOCO_LIVRE = 30;                                         // metros: anotação sem bloco ↔ bloco livre do CAD (sugestão nas pendências)
        const escalaInfo = el('div', { class: 'vz-grafo-escala', hidden: '' }, [el('span', { class: 'vz-esc-rot' }), el('i', { class: 'vz-esc-barra' }), el('small', { class: 'vz-esc-barra-rot' })]);
        function atualizarEscala() {
            const pxTela = Math.min((svg.clientWidth || W) / vb[2], (svg.clientHeight || H) / vb[3]);
            const d = ME && kLayout ? ME.denominadorDe(kLayout, pxTela) : null, barra = d ? ME.barraDeEscala(kLayout * pxTela) : null;
            escalaInfo.hidden = !d;
            if (!d) return;
            escalaInfo.querySelector('.vz-esc-rot').textContent = 'Escala ≈ ' + ME.rotuloDaEscala(d);
            escalaInfo.querySelector('.vz-esc-barra').style.width = Math.round(barra.px) + 'px';
            escalaInfo.querySelector('.vz-esc-barra-rot').textContent = barra.rotulo;
        }
        // Recorte: só pontos/vãos perto da janela ficam visíveis (redes grandes); refaz quando a janela sai da área já preparada
        function recortarVista(forcar) {
            if (!recortar || !ME) return;
            const agora = ME.retanguloDaVista(vb, 0);
            if (!forcar && areaRecorte && ME.contem(areaRecorte, agora)) return;
            areaRecorte = ME.retanguloDaVista(vb, .5);
            // fora da janela o elemento SAI do DOM (display:none ainda custava layout/estilo a cada quadro em redes grandes)
            ME.recortar(itensRecorte, areaRecorte, (it, v) => { if (v) it.camada.appendChild(it.el); else it.el.remove(); });
        }
        let quadroVista = 0, fimMovimento = 0;
        const aplicarVistaJa = () => {
            st.grafoVista = vb; st.grafoVistaEscala = escala;
            svg.setAttribute('viewBox', vb.join(' '));
            atualizarEscala(); recortarVista();
            const sep = separacaoCabos();                                                         // TASK-240: zoom mudou bastante → refaz o espaço entre cabos lado a lado
            if (unidadesDaSeparacao && Math.abs(sep - unidadesDaSeparacao) > .25 * unidadesDaSeparacao) desenharPosicoes();
            if (caboEmFoco !== null) desenharAlcas();
        };
        // pan/zoom chegam várias vezes por quadro (mouse/roda): aplica uma vez por quadro; em redes grandes some o texto enquanto mexe (o texto é o que mais pesa)
        const aplicarVista = () => {
            st.grafoVista = vb; st.grafoVistaEscala = escala;
            if (!quadroVista) quadroVista = raiz.requestAnimationFrame(() => { quadroVista = 0; aplicarVistaJa(); });
            if (recortar) { svg.classList.add('vz-movendo'); raiz.clearTimeout(fimMovimento); fimMovimento = raiz.setTimeout(() => svg.classList.remove('vz-movendo'), 160); }
        };
        const zoom = (f, x = vb[0] + vb[2] / 2, y = vb[1] + vb[3] / 2) => {
            f = Math.max(120 / vb[2], Math.min((escalar() ? 9000 : 3600) / vb[2], f));
            vb = [x + (vb[0] - x) * f, y + (vb[1] - y) * f, vb[2] * f, vb[3] * f];
            aplicarVista();
        };
        const botao = (text, onclick, title = text) => el('button', { class: 'vz-btn', text, title, 'aria-label': title, onclick });
        const instrucao = el('p');
        const sugerir = botao('Revisar conexões detectadas', prepararSugestoes);
        const barra = el('div', { class: 'vz-grafo-barra' }, [
            el('div', {}, [el('h4', { text: 'Grafo da rede' }), instrucao]),
            el('div', { class: 'vz-grafo-cabecalho' }, [contagem, sugerir])
        ]);
        // TASK-235: só três ações — + Ponto (linhas com operação, F = fictícia), + Cabo e Traçar cabo. Poste/estrutura/fictícios viram linhas do ponto.
        const ficticio = botao('+ Cabo fictício', () => abrirFicticio(), 'Novo cabo fictício');               // (fora da barra; mantido para quem ainda chama)
        const btnPonto = botao('+ Ponto', () => abrirPonto(), 'Novo ponto: nome - comentário e as linhas (poste, estruturas), cada uma com a operação. Atalho: P');
        btnPonto.classList.add('p');
        const ferramentas = el('div', { class: 'vz-grafo-ferramentas' }, [
            btnPonto,
            botao('+ Cabo', () => abrirCriacao('cabo'), 'Novo cabo (operação F = cabo fictício). Atalho: C')
        ]);
        const btnVao = botao('Traçar cabo', () => alternarVao(), 'Clique no ponto de origem e depois no de destino: cria um cabo do último tipo usado, com o comprimento medido, ligado nas duas pontas. Shift+clique no destino cria sem perguntar. Atalho: V');
        if (host.inserirNoMapa) ferramentas.append(btnVao);
        const btnAgrupar = botao('Agrupar pontos', () => agruparPontos(), 'Cria um grupo «@nome» com as linhas de cada ponto em Demais Ativos');    // (fora da barra: a conversão do formato antigo faz isso)
        const filtro = el('select', { 'aria-label': 'Filtrar mapa' }, opcoes([
            ['todos', 'Tudo'], ['I', 'A instalar'], ['R', 'A remover'], ['M', 'Mantendo'], ['ficticios', 'Fictícios'], ['conferir', 'A conferir (nível/geometria)']
        ]));
        filtro.addEventListener('change', () => { filtroMapa = filtro.value; sugestoes = []; pintar(); });
        ferramentas.append(filtro);
        if (host.inserirNoMapa) barra.append(ferramentas);
        else barra.append(ficticio);
        const avisoAntigo = el('div', { class: 'vz-aviso-antigo', hidden: '' });                             // TASK-235: pontos no formato antigo
        const seletorEscala = el('select', { class: 'vz-esc-sel', 'aria-label': 'Escala do mapa',
            title: '1:500 mostra o desenho em tamanho real de prancha (1 m = 2 mm; aproximado, depende do monitor). Ajustar encaixa toda a rede na tela.' }, opcoes(ME ? ME.ESCALAS : [['ajustar', 'Ajustar à tela']]));
        seletorEscala.value = escala;
        seletorEscala.addEventListener('change', () => {
            escala = ME.escalaValida(seletorEscala.value);
            st.grafoEscala = escala;
            try { raiz.localStorage.setItem('vz_mapa_escala', escala); } catch (e) { /* sem localStorage: vale só nesta abertura */ }
            st.grafoVista = null; vb = null; areaRecorte = null;
            pintar();
            vb = vistaPadrao();
            aplicarVistaJa();
        });
        const acoesMapa = el('div', { class: 'vz-grafo-acoes' }, [
            ...(ME ? [seletorEscala] : []),
            botao('−', () => zoom(1.25), 'Afastar'), botao('+', () => zoom(.8), 'Aproximar'),
            botao('Enquadrar', () => { vb = vistaPadrao(); aplicarVista(); }, escalar() ? 'Centralizar a rede na escala escolhida' : 'Enquadrar'),
            botao('Fundo', () => alternarFundo(), 'Mostra/oculta, esmaecidos, a rede e os blocos de poste do CAD que ainda não viraram linha da obra. Clique num bloco para transformá-lo em ponto.'),
            botao('Posições do CAD', () => restaurarPosicoesDoCad(), 'Descarta as posições ajustadas à mão (arrastar, criar ponto, Rearranjar) e volta cada ponto para a coordenada do CAD. Ctrl+Z desfaz.'),
            botao('Rearranjar', () => rearranjarMapa(), 'Reorganiza os pontos de forma uniforme e limpa (vãos de tamanho parecido, sem sobreposição). Em 1:500 volta os pontos para a posição do CAD. Ctrl+Z desfaz.')
        ]);
        const suporte = el('div', { class: 'vz-scroll vz-grafo-vista' }, [svg, escalaInfo, acoesMapa]);
        const lista = el('div', { class: 'vz-grafo-lista' });
        const tituloPontas = el('h5');
        const painelPend = el('details', { class: 'vz-pend', open: '' });
        const lateralMapa = el('aside', { class: 'vz-grafo-cabos', 'aria-label': 'Cabos do mapa' }, [painelPend, tituloPontas, lista]);
        const area = el('div', { class: 'vz-grafo-area' }, [suporte, lateralMapa]);
        const revisao = el('div', { class: 'vz-grafo-revisao' });
        const vazio = el('p', { class: 'vz-nota', text: 'Adicione postes ou cabos à obra para montar a rede.' });
        const avisos = el('p', { class: 'vz-grafo-aviso', role: 'status' });
        const dialogo = el('dialog', { class: 'vz-no-dialog', 'aria-labelledby': 'vz-no-titulo' });
        const editor = el('dialog', { class: 'vz-no-dialog vz-composer-dialog', 'aria-labelledby': 'vz-composer-titulo' });
        corpo.append(barra, avisoAntigo, vazio, area, revisao, avisos, dialogo, editor);

        const linhaNo = no => no.poste ?? no.linhas[0];
        const linhaFicticia = no => outros()[linhaNo(no)]._mapaPosteFicticio === true;
        const pontoSoFicticio = no => !linhaFicticia(no) && no.linhas.length > 0 && no.linhas.every(l => VR.ehFicticia(outros()[l]));      // TASK-235: ponto só com linhas F
        const estruturasDoNo = no => no.estruturas.filter(ie => !metaDe(linhaDe(ie)).soPoste && codigoDe(ie));
        const ligacoes = ic => cabos()[ic].vinculoEstruturas || [];
        const jaLigadoNo = (ic, no) => ligacoes(ic).some(ie => no.linhas.includes(linhaDe(ie)));
        const pontosDoCabo = ic => [...new Set(grafo.cabosDoMapa.get(ic)?.ids.filter(Boolean))];
        const livres = ic => grafo?.cabosDoMapa.has(ic) ? Math.max(0, 2 - pontosDoCabo(ic).length) : 0;
        function gravarPontos(ic, ids) {
            cabos()[ic].vinculoPostes = ids.map(id => linhaNo(grafo.nos.get(id)));
            cabos()[ic]._mapaTopologia = true;
        }
        function alterarPontoDoCabo(ic, no, conectar) {
            const pontos = pontosDoCabo(ic), ja = pontos.includes(no.id);
            if (conectar && ja) return 'Este cabo já chega a este ponto.';
            if (conectar && !livres(ic)) return 'O cabo já tem dois pontos. Solte uma ponta para mudar o traçado.';
            if (!conectar && (!ja || jaLigadoNo(ic, no))) return 'Remova primeiro o vínculo com a estrutura deste ponto.';
            const linhasAfetadas = [...new Set([...pontos, no.id])].map(id => linhaNo(grafo.nos.get(id)));
            host.antes();
            gravarPontos(ic, conectar ? pontos.concat(no.id) : pontos.filter(id => id !== no.id));
            if (!conectar && linhaFicticia(no)) {
                const c = cabos()[ic], id = outros()[linhaNo(no)]._mapaId;
                c._mapaAncoras = (c._mapaAncoras || []).filter(([anterior]) => anterior !== id);
            }
            caboEscolhido = pontaEscolhida = null;
            ultimoStatus = conectar
                ? (linhaFicticia(no) ? 'Ponta conectada. Escolha a estrutura fictícia no painel do poste.'
                    : 'Ponta conectada ao ponto.')
                : 'Ponta desconectada do ponto.';
            ultimoStatus = comVinculosAutomaticos(ultimoStatus, automatizarPontos(linhasAfetadas));
            atualizar();
            host.depois?.();
            avisos.textContent = ultimoStatus;
            return '';
        }

        function estruturasCompativeis(ic, no) {
            if (jaLigadoNo(ic, no)) return [];
            return estruturasDoNo(no).filter(ie => infoEst(ie).ocupa && !motivoParaLigar(ic, ie));
        }

        function automatizarPontos(linhas) {
            if (!linhas.length) return 0;
            cx = null;
            grafo = montarGrafo();
            const pontos = new Set(linhas.map(l => grafo.nosPorLinha.get(l)).filter(Boolean));
            const chegadas = [];
            for (const no of pontos) {
                if (linhaFicticia(no)) continue;
                for (const ic of no.cabos) {
                    const c = cabos()[ic];
                    if (c.bloqueada || c.congelada || !VR.ehCaboValido(c) || !infoCabo(ic).vincula ||
                        !VR.postesAtribuidos(c).some(l => no.linhas.includes(l))) continue;
                    const candidatas = estruturasCompativeis(ic, no).filter(ie =>
                        !linhaOutros(ie).bloqueada && infoEst(ie).regra !== null && infoEst(ie).qtd !== null);
                    chegadas.push({ ic, candidatas });
                }
            }
            const propostas = VR.planejarVinculosDoMapa(chegadas, infoEst, ocupantes, infoCabo, ic =>
                infoCabo(ic).pontas - ligacoes(ic).length - VR.estruturasFicticiasDe(cabos()[ic]).length);
            for (const { ic, ie } of propostas) {
                const c = cabos()[ic], anteriores = ligacoes(ic);
                if (!anteriores.length) c.vinculoOrigem = 'mapa';
                c.vinculoEstruturas = [...anteriores, ie];
            }
            if (propostas.length) cx = null;
            return propostas.length;
        }
        const comVinculosAutomaticos = (texto, total) => total ? texto + ' · ' + total + ' vínculo(s) automático(s).' : texto;

        function motivoNo(ic, no) {
            if (jaLigadoNo(ic, no)) return 'Este cabo já tem uma ponta vinculada neste ponto.';
            const estruturas = estruturasDoNo(no);
            if (!estruturas.length) return 'O ponto não tem estrutura disponível. Adicione uma estrutura ao ponto.';
            return [...new Set(estruturas.map(ie => motivoParaLigar(ic, ie)).filter(Boolean))].join(' · ');
        }

        function atualizarSelecao() {
            instrucao.textContent = caboEscolhido !== null
                ? 'Cabo ' + (caboEscolhido + 1) + (pontaEscolhida ? ' · ponta livre ' + pontaEscolhida : '') + ' · clique num ponto destacado.'
                : 'Selecione um cabo para ver suas pontas. Arraste os pontos para organizar o mapa.';
            lista.querySelectorAll('[data-cabo-livre]').forEach(b => {
                const ativo = Number(b.dataset.caboLivre) === caboEscolhido && Number(b.dataset.ponta) === pontaEscolhida;
                b.classList.toggle('p', ativo);
                b.setAttribute('aria-pressed', String(ativo));
            });
            lista.querySelectorAll('[data-cartao-cabo]').forEach(card => {
                const selecionado = Number(card.dataset.cartaoCabo) === caboEmFoco;
                card.classList.toggle('selecionado', selecionado);
                card.querySelector('.vz-grafo-cabo-titulo').setAttribute('aria-pressed', String(selecionado));
            });
            cabosSvg.forEach(g => g.classList.toggle('selecionado', Number(g.dataset.caboMapa) === caboEmFoco));
            const ligados = caboEscolhido === null ? null : new Set(pontosDoCabo(caboEscolhido));
            nosSvg.forEach((g, id) => g.classList.toggle('alvo', ligados !== null && !ligados.has(id)));
        }
        function destacarCabo(ic, rolar = false) {
            caboEmFoco = ic;
            caboEscolhido = pontaEscolhida = null;
            if (rolar) garantirCartao(ic);
            atualizarSelecao();
            desenharAlcas();
            if (rolar) lista.querySelector(`[data-cartao-cabo="${ic}"]`)?.scrollIntoView({ block: 'nearest' });
        }

        function selecionarCabo(ic, ponta = 1) {
            if (!livres(ic)) return;
            caboEmFoco = ic;
            const desmarcar = ic === caboEscolhido && ponta === pontaEscolhida;
            caboEscolhido = desmarcar ? null : ic;
            pontaEscolhida = desmarcar ? null : ponta;
            avisos.textContent = '';
            garantirCartao(ic);
            atualizarSelecao();
            desenharAlcas();
        }

        function calcularSugestoes() {
            return [...grafo.cabosDoMapa].flatMap(([ic, dados]) => {
                if (!cabosFiltrados.has(ic) || cabos()[ic]._mapaTopologia || cabos()[ic].congelada) return [];
                const ids = pontosDoCabo(ic);
                if (!ids.length || ids.some(id => !dados.referencias.includes(id))) return [];
                return [{ ic, ids }];
            });
        }
        function prepararSugestoes() {
            sugestoes = calcularSugestoes();
            pintar();
            if (!sugestoes.length) avisos.textContent = 'Nenhuma conexão detectada para confirmar. Você pode conectar as pontas manualmente.';
            else revisao.querySelector('summary')?.focus();
        }
        function aplicarSugestoes() {
            if (!sugestoes.length) return;
            // Refaz a proposta a partir dos dados atuais antes de confirmar o lote.
            cx = null;
            grafo = montarGrafo();
            const revistas = calcularSugestoes();
            if (JSON.stringify(revistas) !== JSON.stringify(sugestoes)) {
                sugestoes = revistas;
                pintar();
                avisos.textContent = 'A obra mudou. Confira as sugestões atualizadas.';
                return;
            }
            host.antes();
            for (const { ic, ids } of sugestoes) gravarPontos(ic, ids);
            const total = sugestoes.length;
            const auto = automatizarPontos(sugestoes.flatMap(({ ic }) => VR.postesAtribuidos(cabos()[ic])));
            atualizar();
            host.depois?.();
            avisos.textContent = comVinculosAutomaticos(total + ' cabo(s) com conexões confirmadas.', auto);
            sugerir.focus();
        }
        function renderRevisao() {
            revisao.hidden = !sugestoes.length;
            revisao.replaceChildren();
            if (!sugestoes.length) return;
            revisao.append(el('strong', { text: `${sugestoes.length} cabo(s) com conexões detectadas` }),
                botao('Aplicar sugestões', aplicarSugestoes),
                botao('Descartar', () => { sugestoes = []; pintar(); sugerir.focus(); }),
                el('details', {}, [
                    el('summary', { text: 'Conferir conexões detectadas' }),
                    el('ul', {}, sugestoes.map(({ic, ids}) => el('li', { text: `Cabo ${ic + 1} · ${cabos()[ic].ativo} → ${ids.map(id => nomeDoNo(grafo.nos.get(id))).join(' ↔ ')}` })))
                ]));
        }

        function montarGrafo() {
            return VR.grafoDaRede(cabos(), outros(), host.codigoEstrutura, host.codigosEstrutura, (_, i) => postesDe(i).length > 0);
        }

        // TASK-225: cada ponto é um grupo «@nome» em Demais Ativos; o nome do grupo é o nome do ponto (sem grupo: número automático pela linha)
        let titulosGrupo = new Map();
        const GLm = raiz.GruposLinhas;
        const nomeGrupoDaLinha = l => { const r = outros()[l], t = r && r.grupoDe ? titulosGrupo.get(r.grupoDe) : null; return t && GLm ? GLm.nomeDoPonto(t) : ''; };      // TASK-235: só o nome (o comentário fica no painel)
        const nomeGrupoDoNo = no => nomeGrupoDaLinha(linhaNo(no)) || no.linhas.map(nomeGrupoDaLinha).find(Boolean) || '';
        const numeroDoNo = no => nomeGrupoDoNo(no) || `P${linhaNo(no) + 1}`;                  // TASK-239: o mesmo nome no mapa e no cartão do cabo
        const limitar01 = v => Math.max(0, Math.min(1, v));
        // grava a posição que o ponto tem na tela (Ajustar: fração da tela · 1:500: deslocamento em metros)
        function guardarPosicao(id) {
            const no = grafo.nos.get(id), p = posicoes.get(id);
            if (!no || !p) return;
            const linha = outros()[linhaNo(no)];
            if (emEscalaAgora()) { const b = bases.get(id) || [0, 0]; linha._mapaPoste500 = [(p[0] - b[0]) / kLayout, (p[1] - b[1]) / kLayout]; }
            else linha._mapaPoste = [limitar01(p[0] / W), limitar01(p[1] / H)];
        }
        // Antes de mudar a rede (novo ponto…), prende todos os pontos onde estão: o encaixe automático dependia do conjunto de pontos e reorganizava tudo ao adicionar um
        function fixarPosicoes() {
            if (!grafo) return;
            grafo.nos.forEach((no, id) => { const linha = outros()[linhaNo(no)]; if (linha && !linha[campoPonto()]) guardarPosicao(id); });
        }
        // lugar livre perto do centro da janela para o ponto recém-criado
        function posicaoLivre(excetoId) {
            const cx = vb ? vb[0] + vb[2] / 2 : W / 2, cy = vb ? vb[1] + vb[3] / 2 : H / 2;
            const passo = emEscalaAgora() ? Math.max(48, 12 * kLayout) : 64;
            const ocupados = [...posicoes].filter(([id]) => id !== excetoId).map(([, p]) => p);
            const dentro = (x, y) => x >= limites[0] && x <= limites[2] && y >= limites[1] && y <= limites[3];
            for (let r = 0; r < 90; r++) {
                const n = r === 0 ? 1 : 6 * r;
                for (let i = 0; i < n; i++) {
                    const a = 2 * Math.PI * i / n + r * .5, x = cx + Math.cos(a) * r * passo * .6, y = cy + Math.sin(a) * r * passo * .6;
                    if (dentro(x, y) && ocupados.every(([ox, oy]) => Math.hypot(ox - x, oy - y) >= passo * .8)) return [x, y];
                }
            }
            return [Math.max(limites[0], Math.min(limites[2], cx)), Math.max(limites[1], Math.min(limites[3], cy))];
        }
        function colocarNovoPonto(linha) {
            const no = grafo.nosPorLinha.get(linha);
            if (!no) return;
            const [x, y] = posicaoLivre(no.id);
            posicoes.set(no.id, [x, y]);
            guardarPosicao(no.id);
            desenharPosicoes({ idNo: no.id });
            recortarVista(true);
        }

        // Posicionar: depois de criar (ou copiar) um ponto, ele acompanha o mouse e o clique fixa o lugar; Esc mantém o lugar sugerido
        let colocando = null, quadroPosicionar = 0, alvoPosicionar = null;
        const teclaPosicionar = e => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); encerrarPosicionamento(false); } };
        function iniciarPosicionamento(linha) {
            const no = grafo.nosPorLinha.get(linha);
            if (!no) return;
            encerrarPosicionamento(true);
            colocando = { id: no.id, original: posicoes.get(no.id).slice(), nome: nomeGrupoDoNo(no) || nomeDoNo(no) };
            svg.classList.add('posicionando');
            doc.addEventListener('keydown', teclaPosicionar, true);
            avisos.textContent = 'Posicione «' + colocando.nome + '»: mova o mouse e clique no mapa para fixar o ponto (Esc mantém o lugar sugerido).';
        }
        function encerrarPosicionamento(confirmou, silencioso) {
            if (!colocando) return;
            const c = colocando;
            colocando = null; alvoPosicionar = null;
            svg.classList.remove('posicionando');
            doc.removeEventListener('keydown', teclaPosicionar, true);
            if (quadroPosicionar) { raiz.cancelAnimationFrame(quadroPosicionar); quadroPosicionar = 0; }
            if (silencioso) return;
            if (!confirmou && grafo.nos.has(c.id)) { posicoes.set(c.id, c.original); guardarPosicao(c.id); desenharPosicoes({ idNo: c.id }); }
            recortarVista(true);
            avisos.textContent = confirmou ? 'Ponto posicionado. Arraste para ajustar.' : 'Ponto mantido no lugar sugerido.';
        }
        function copiarPonto(no, status) {
            fixarPosicoes();
            const r = host.copiarPontoNoMapa(no.linhas);
            if (r.erro) { status.textContent = r.erro; return; }
            dialogo.close();
            cx = null; metaAtivo.clear();
            atualizar();
            host.depois?.();
            colocarNovoPonto(r.indice);
            iniciarPosicionamento(r.indice);
        }
        function copiarCabo(ic) {
            const r = host.copiarCaboNoMapa(ic);
            if (r.erro) { avisos.textContent = r.erro; return; }
            cx = null; metaAtivo.clear();
            atualizar();
            caboEmFoco = r.indice;
            caboEscolhido = livres(r.indice) ? r.indice : null;
            pontaEscolhida = caboEscolhido === null ? null : 1;
            garantirCartao(r.indice);
            atualizarSelecao();
            lista.querySelector(`[data-cartao-cabo="${r.indice}"]`)?.scrollIntoView({ block: 'nearest' });
            avisos.textContent = 'Cabo copiado (linha ' + (r.indice + 1) + ' de Cabos). Clique num ponto destacado para ligar a ponta.';
        }

        function rearranjarMapa() {
            if (!grafo || !grafo.nos.size) return;
            host.antes();
            if (emEscalaAgora()) {                                           // 1:500: a escala manda — descarta os ajustes manuais (volta ao CAD; pontos sem CAD ficam numa grade ao lado)
                grafo.nos.forEach(no => { delete outros()[linhaNo(no)]._mapaPoste500; });
                cabos().forEach(c => { if (c) delete c._mapaCabo500; });
            } else if (raiz.MapaLayout) {
                const ids = [...grafo.nos.keys()], vistos = new Set(), arestas = [];
                desenhos.forEach(d => { const a = d.v.ids[0] || (d.ancoras && d.ancoras[0]), b = d.v.ids[1] || (d.ancoras && d.ancoras[1]); if (a && b && a !== b) { const k = a < b ? a + '|' + b : b + '|' + a; if (!vistos.has(k)) { vistos.add(k); arestas.push([a, b]); } } });
                const nos = ids.map(id => ({ id, x: posicoes.get(id)[0], y: posicoes.get(id)[1] }));
                const L = Math.max(36, Math.min(110, Math.sqrt((W - 2 * margem) * (H - 2 * margem) / ids.length) * .9));
                const fixos = ids.filter(id => { const no = grafo.nos.get(id); return !linhaFicticia(no) && no.linhas.some(l => outros()[l].bloqueada); });     // grupo bloqueado não sai do lugar
                const res = raiz.MapaLayout.rearranjar(nos, arestas, { comprimento: L, fixos });
                let c = raiz.MapaLayout.caixa([...res.values()]), w = Math.max(c[2] - c[0], 1), h = Math.max(c[3] - c[1], 1);
                const util = [W - 2 * margem, H - 2 * margem];
                // rede «de pé» numa tela deitada: gira 90° quando assim cabe maior (sem pontos bloqueados, que não podem sair do lugar)
                const girar = !fixos.length && Math.min(util[0] / h, util[1] / w) > Math.min(util[0] / w, util[1] / h) * 1.15;
                if (girar) { res.forEach((p, id) => res.set(id, [p[1], -p[0]])); c = raiz.MapaLayout.caixa([...res.values()]); w = Math.max(c[2] - c[0], 1); h = Math.max(c[3] - c[1], 1); }
                const f = fixos.length ? 1 : Math.min(util[0] / w, util[1] / h, 2);                            // cabe na área do mapa, todos na mesma proporção
                const ox = fixos.length ? 0 : margem + (util[0] - w * f) / 2 - c[0] * f, oy = fixos.length ? 0 : margem + (util[1] - h * f) / 2 - c[1] * f;
                ids.forEach(id => { const [x, y] = res.get(id); posicoes.set(id, [Math.max(18, Math.min(W - 18, x * f + ox)), Math.max(18, Math.min(H - 18, y * f + oy))]); guardarPosicao(id); });
                cabos().forEach(cb => { if (cb) delete cb._mapaCabo; });                                    // vãos retos
            }
            st.grafoVista = null;
            cx = null;
            atualizar();
            vb = vistaPadrao();
            aplicarVistaJa();
            avisos.textContent = emEscalaAgora() ? 'Pontos na posição do CAD. Ctrl+Z desfaz.' : 'Pontos rearranjados. Ctrl+Z desfaz.';
        }
        // posições ajustadas à mão (arrastar, criar/copiar ponto, Rearranjar, Agrupar) valem no lugar das coordenadas do CAD; aqui dá para voltar às do CAD
        function pontosAjustados() {
            const campo = campoPonto();
            return [...grafo.nos.values()].filter(no => outros()[linhaNo(no)] && outros()[linhaNo(no)][campo]).length;
        }
        function restaurarPosicoesDoCad() {
            const n = pontosAjustados();
            if (!n) { avisos.textContent = 'Nenhum ponto está com posição ajustada à mão.'; return; }
            host.antes();
            const campo = campoPonto(), campoC = campoCabo();
            grafo.nos.forEach(no => { const linha = outros()[linhaNo(no)]; if (linha) delete linha[campo]; });
            cabos().forEach(c => { if (c) delete c[campoC]; });
            st.grafoVista = null;
            cx = null;
            atualizar();
            vb = vistaPadrao();
            aplicarVistaJa();
            avisos.textContent = n + ' ponto(s) voltaram à posição do CAD (os sem coordenada ficam na coluna ao lado). Ctrl+Z desfaz.';
        }
        function pontosSemGrupo() {
            return [...grafo.nos.values()].filter(no => !nomeGrupoDoNo(no) && no.linhas.some(l => outros()[l] && !outros()[l].grupoDe))
                .sort((a, b) => linhaNo(a) - linhaNo(b));
        }
        function agruparPontos() {
            const lista = pontosSemGrupo();
            if (!lista.length) { avisos.textContent = 'Todos os pontos já são grupos.'; return; }
            fixarPosicoes();
            const n = host.agruparPontosNoMapa(lista.map(no => ({ linhas: no.linhas, nome: 'P' + (linhaNo(no) + 1) })));          // mantém o nome que o ponto já tinha no mapa (P<linha>)
            cx = null;
            atualizar();
            host.depois?.();
            avisos.textContent = n + ' ponto(s) viraram grupo em Demais Ativos. Ctrl+Z desfaz.';
        }

        // TASK-227: diagnóstico por cabo (nível do rótulo × linha do desenho, e por que não aparece) — uma vez por repintada
        let diagnosticos = new Map();
        const diagDe = ic => { if (!diagnosticos.has(ic)) diagnosticos.set(ic, VR.diagnosticoDoCabo(cabos()[ic], grafo.cabosDoMapa.get(ic), st.cfg.dicionario_cabos)); return diagnosticos.get(ic); };
        const precisaConferir = ic => { const d = diagDe(ic); return d.diverge || !!d.motivo; };

        function pintar() {
            diagnosticos = new Map();
            titulosGrupo = new Map();
            outros().forEach(r => { if (r && r.grupoDecl) titulosGrupo.set(r.gid, r); });
            grafo = montarGrafo();
            for (const id of [...selecao]) if (!grafo.nos.has(id)) selecao.delete(id);
            if (vaoOrigem && !grafo.nos.has(vaoOrigem)) vaoOrigem = null;
            let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
            const limitar = ([x, y]) => { x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y); };
            const posCad = n => n.pos || n.posAprox;                                                       // TASK-233: sem bloco no CAD, vale a posição aproximada da anotação
            grafo.nos.forEach(n => { if (posCad(n)) limitar(posCad(n)); });
            grafo.vaos.forEach(v => (v.pontos || v.pontas || []).forEach(limitar));
            const sinteticos = [...grafo.nos.values()].filter(n => !posCad(n));
            const temGeo = Number.isFinite(x0);
            const emEscala = escalar() && temGeo && !!ME;                                                       // TASK-224: 1:500 = fator fixo (px por metro), o desenho pode passar da área do «Ajustar»
            const largura = temGeo && sinteticos.length && !emEscala ? W * .74 : W;
            const k = emEscala ? ME.pxPorMetro(500) : temGeo ? Math.min((largura - 2 * margem) / Math.max(x1 - x0, 1), (H - 2 * margem) / Math.max(y1 - y0, 1)) : 1;
            kLayout = temGeo ? k : null;
            ponto = ([x, y]) => [largura / 2 + (x - (x0 + (x1 - x0) / 2)) * k, H / 2 - (y - (y0 + (y1 - y0) / 2)) * k];
            const gx1 = temGeo ? ponto([x1, y0])[0] : 0, gy0 = temGeo ? ponto([x0, y1])[1] : 0;
            posicoes.clear(); bases = new Map();
            const posicaoSalva = (n, base) => {
                const linha = outros()[linhaNo(n)];
                if (emEscala) return linha._mapaPoste500 ? [base[0] + linha._mapaPoste500[0] * k, base[1] + linha._mapaPoste500[1] * k] : base;      // em escala o ajuste manual é um deslocamento em metros
                return linha._mapaPoste ? [linha._mapaPoste[0] * W, linha._mapaPoste[1] * H] : base;
            };
            vaosPorCabo.clear();
            nosSvg.clear();
            desenhos = [];
            grafo.vaos.forEach(v => v.cabos.forEach(ic => vaosPorCabo.set(ic, v)));
            grafo.nos.forEach(n => {
                if (!posCad(n)) return;
                const base = ponto(posCad(n));
                bases.set(n.id, base);
                posicoes.set(n.id, posicaoSalva(n, base));
            });
            const colunas = Math.ceil(Math.sqrt(sinteticos.length));
            const linhas = Math.ceil(sinteticos.length / Math.max(colunas, 1));
            sinteticos.forEach((n, i) => {
                const ini = temGeo ? (emEscala ? gx1 + 140 : largura) : 0;
                const coluna = (i % colunas + 1) / (colunas + 1);
                const linha = (Math.floor(i / colunas) + 1) / (linhas + 1);
                const base = emEscala ? [ini + (i % colunas) * 96, gy0 + Math.floor(i / colunas) * 56] : [ini + coluna * (W - ini), linha * H];
                bases.set(n.id, base);
                posicoes.set(n.id, posicaoSalva(n, base));
            });
            limites = emEscala ? (() => { const c = ME.caixaDe([...posicoes.values()], 700); return [c[0], c[1], c[2], c[3]]; })() : [18, 18, W - 18, H - 18];
            // TASK-233: fundo do CAD (rede e blocos de poste sem anotação) — só o que encosta na rede; o bloco que já é ponto não aparece
            const fundoBruto = temGeo && MM && host.fundoCad ? host.fundoCad() : null;
            const fundo = fundoBruto ? MM.recortarFundo(fundoBruto, [x0, y0, x1, y1]) : null;
            fundoLivres = fundo ? MM.blocosLivres(fundo.postes, [...grafo.nos.values()].map(n => n.pos).filter(Boolean), 1) : [];       // as pendências usam mesmo com o fundo oculto
            const svgFundo = !fundo || !fundoLigado ? '' : fundo.redes.map(r => `<polyline class="vz-fundo-rede ${r.t === 'MT' ? 'mt' : r.t === 'BT' ? 'bt' : ''}" points="${r.p.map(q => ponto(q).map(v => v.toFixed(1)).join(',')).join(' ')}"/>`).join('')
                + fundoLivres.map(b => { const [bx, by] = ponto(b.p); return `<rect class="vz-fundo-bloco" data-bloco="${b.id}" x="${(bx - 4).toFixed(1)}" y="${(by - 4).toFixed(1)}" width="8" height="8"><title>Bloco de poste do CAD sem anotação · clique para transformar em ponto</title></rect>`; }).join('');

            const a = analise();
            const rede = [];
            const nosFicticios = new Set([...grafo.nos.values()].filter(no => linhaFicticia(no) ||
                estruturasDoNo(no).some(ie => ficticiosDe(ie).length)).map(no => no.id));
            const operacaoDoMapa = row => trim(row.operacao).toUpperCase().replace(/^\*/, '') || 'I';
            const cabosVisiveis = new Set([...grafo.cabosDoMapa].filter(([ic, d]) => filtroMapa === 'todos' ||
                (filtroMapa === 'conferir' ? precisaConferir(ic) : filtroMapa === 'ficticios' ? d.ids.some(id => nosFicticios.has(id)) || VR.estruturasFicticiasDe(cabos()[ic]).length
                    : operacaoDoMapa(cabos()[ic]) === filtroMapa)).map(([ic]) => ic));
            const nosVisiveis = new Set([...grafo.nos.values()].filter(no => filtroMapa === 'todos' ||
                (filtroMapa === 'conferir' ? false : filtroMapa === 'ficticios' ? nosFicticios.has(no.id)
                    : !linhaFicticia(no) && no.linhas.some(l => operacaoDoMapa(outros()[l]) === filtroMapa))).map(no => no.id));
            cabosFiltrados = cabosVisiveis;
            for (const ic of cabosVisiveis) pontosDoCabo(ic).forEach(id => nosVisiveis.add(id));
            const exibir = (v, candidato = false) => {
                const indices = v.cabos.filter(ic => cabosVisiveis.has(ic));
                indices.forEach((ic, i) => {
                    const grupo = VR.gruposVisuaisDoVao([ic], cabos(), a, st.cfg.dicionario_cabos)[0];
                    const nome = `Cabo ${ic + 1} · ${cabos()[ic].ativo}`;
                    const rotulo = `${nome} · ${grupo.tipo || 'Cabo'} · ${grupo.completo ? 'Vinculado' : 'Pendente'}`;
                    const classe = (grupo.tipo === 'MT' ? 'mt' : grupo.tipo === 'BT' ? 'bt' : '') + (diagDe(ic).diverge ? ' diverge' : '');
                    desenhos.push({v, grupo, ic: candidato ? null : ic, faixa: 0, inverte: false});
                    const tracado = `<polyline class="vz-aresta ${classe}${grupo.completo ? ' completo' : ''}${candidato ? ' sugerida' : ''}" fill="none" vector-effect="non-scaling-stroke" ${candidato || (v.ids.includes(null) && !grupo.completo) ? 'stroke-dasharray="5 4"' : ''}><title>${esc(rotulo)}</title></polyline>`;
                    rede.push(candidato ? tracado : `<g data-cabo-mapa="${ic}" data-arrastar="cabo" tabindex="0" role="button" aria-label="${esc(nome)}"><title>${esc(rotulo)}</title>${tracado}<polyline class="vz-alvo-aresta" vector-effect="non-scaling-stroke"/></g>`);
                });
            };
            grafo.vaos.forEach(v => exibir(v));
            grafo.cabosDoMapa.forEach(({ids}, ic) => {
                if (!vaosPorCabo.has(ic) && ids.some(Boolean)) exibir({ids, pontos:null, pontas:null, cabos:[ic]});
            });
            for (const { ic, ids } of sugestoes) {
                exibir({ ids: [ids[0], ids[1] || null], pontos: null, pontas: null, cabos: [ic] }, true);
            }
            faixasDosCabos(desenhos);

            const labels = [];
            const nos = [...grafo.nos.values()].filter(n => nosVisiveis.has(n.id)).map(n => {
                const [x, y] = posicoes.get(n.id), rotulo = nomeDoNo(n);
                const avulsa = n.poste === null;
                const estruturas = estruturasDoNo(n);
                const tituloAvulsa = avulsa ? estruturas.slice(0, 2).map(codigoDe).join(' / ') || trim(outros()[linhaNo(n)].ativo) : '';
                const nomeGrupo = nomeGrupoDoNo(n);
                const numero = nomeGrupo ? (nomeGrupo.length > 16 ? nomeGrupo.slice(0, 15) + '…' : nomeGrupo) : numeroDoNo(n);
                const curto = avulsa ? numero + ' · ' + tituloAvulsa.slice(0, 18) + (tituloAvulsa.length > 18 ? '…' : '')
                    : linhaFicticia(n) ? numero + ' · Fictício' : numero;
                const ficticios = estruturas.reduce((q, ie) => q + ficticiosDe(ie).length, 0);
                const direita = x + 15 + curto.length * 7;
                const denso = labels.some(([a, b, c]) => Math.abs(y - c) < 14 && x + 10 < b && direita > a);
                if (!denso) labels.push([x + 10, direita, y]);
                const completo = VR.posteCompletoNoGrafo(estruturas, a);
                const classe = [denso ? 'vz-rotulo-denso' : '', completo ? 'completo' : 'pendente', avulsa ? 'avulsa' : '', linhaFicticia(n) || pontoSoFicticio(n) ? 'ficticio' : '',
                    !n.pos && n.posAprox ? 'aproximada' : '', selecao.has(n.id) ? 'escolhido' : '', n.id === vaoOrigem ? 'origem-vao' : ''].join(' ');
                const estado = (linhaFicticia(n) ? 'Âncora visual' : completo ? 'Vinculação completa' : 'Vinculação pendente')
                    + (!n.pos && n.posAprox ? ' · posição aproximada (anotação sem bloco de poste no CAD)' : '');
                return `<g class="${classe}" data-no="${esc(n.id)}" data-arrastar="poste" tabindex="0" role="button" aria-label="${esc(rotulo)}"><title>${esc(rotulo)} · ${estado}</title><circle class="vz-alvo" r="19"/><circle class="vz-halo" r="10"/><circle class="vz-marcador" r="6"/><text x="14" y="-7">${esc(curto)}</text>${ficticios ? `<text class="vz-ficticio-rotulo" x="13" y="10">F × ${ficticios}</text>` : ''}</g>`;
            }).join('');
            svg.innerHTML = '<g class="vz-camada-fundo">' + svgFundo + '</g><g class="vz-camada-vaos">' + rede.join('') + '</g><g class="vz-camada-nos">' + nos + '</g><g class="vz-camada-alcas"></g>';
            camadaVaos = svg.querySelector('.vz-camada-vaos'); camadaNos = svg.querySelector('.vz-camada-nos'); camadaAlcas = svg.querySelector('.vz-camada-alcas');
            cabosSvg = [...svg.querySelectorAll('[data-cabo-mapa]')];
            svg.querySelectorAll('[data-no]').forEach(g => nosSvg.set(g.dataset.no, g));
            svg.querySelectorAll('.vz-aresta').forEach((linha, i) => { desenhos[i].linha = linha; });
            // TASK-228: ponta de cabo sem poste (a mais de 5 m de qualquer um) segue visualmente o ponto mais perto (até 25 m) — assim o cabo anda junto quando o ponto é movido/rearranjado.
            // A ligação de verdade (d.v.ids) continua exigindo 5 m; isto só ancora o desenho.
            const comPos = [...grafo.nos.values()].filter(posCad);                                       // TASK-233: o ponto de posição aproximada também ancora (costuma ser o poste certo da ponta solta)
            for (const d of desenhos) {
                d.ancoras = [null, null];
                if (!d.v.pontas) continue;
                d.v.ids.forEach((id, i) => {
                    if (id || !d.v.pontas[i]) return;
                    let melhor = null, dmin = RAIO_ANCORA_VISUAL;
                    for (const n of comPos) { const q = posCad(n), dd = Math.hypot(q[0] - d.v.pontas[i][0], q[1] - d.v.pontas[i][1]); if (dd <= dmin) { dmin = dd; melhor = n.id; } }
                    d.ancoras[i] = melhor;
                });
            }
            desenhosPorNo = new Map(); desenhosPorCabo = new Map();
            for (const d of desenhos) {
                new Set([...d.v.ids, ...(d.ancoras || [])]).forEach(id => { if (id) { if (!desenhosPorNo.has(id)) desenhosPorNo.set(id, []); desenhosPorNo.get(id).push(d); } });     // a ponta ancorada (TASK-228) anda junto ao arrastar
                if (d.ic !== null) { if (!desenhosPorCabo.has(d.ic)) desenhosPorCabo.set(d.ic, []); desenhosPorCabo.get(d.ic).push(d); }
            }
            desenharPosicoes();
            // TASK-224: rede grande → recorte por janela (itens: cada ponto e cada vão com a caixa onde ele aparece)
            recortar = !!ME && nosSvg.size + desenhos.length > 400;
            itensPorNo = new Map(); areaRecorte = null;
            itensRecorte = recortar ? [...nosSvg].map(([id, g]) => { const it = { el: g, camada: camadaNos, b: caixaDoNo(id), v: true, id }; itensPorNo.set(id, it); return it; })
                .concat(desenhos.map(d => { const el = d.linha.parentNode.hasAttribute('data-cabo-mapa') ? d.linha.parentNode : d.linha; d.item = { el, camada: camadaVaos, b: d.b, v: true }; return d.item; })) : [];
            svg.classList.toggle('vz-movendo', false);
            if (vb) recortarVista(true);
            const pontosVisiveis = [...nosVisiveis].map(id => grafo.nos.get(id));
            totalPostes.textContent = pontosVisiveis.filter(no => no.poste !== null && !linhaFicticia(no)).length;
            totalAvulsas.textContent = pontosVisiveis.filter(no => no.poste === null).length;
            totalVaos.textContent = grafo.vaos.filter(v => v.cabos.some(ic => cabosVisiveis.has(ic))).length;
            if (!cabosVisiveis.has(caboEmFoco)) caboEmFoco = null;
            if (caboEscolhido !== null && (!cabosVisiveis.has(caboEscolhido) || !livres(caboEscolhido)))
                caboEscolhido = pontaEscolhida = null;
            // TASK-233: lista lateral em lotes (rede grande abre rápido); o resto entra ao rolar até o fim
            const fabricar = ic => {
                const c = cabos()[ic], grupo = VR.gruposVisuaisDoVao([ic], cabos(), a, st.cfg.dicionario_cabos)[0];
                const nome = `Cabo ${ic + 1} · ${c.ativo}`, qtd = livres(ic);
                const botoes = pontosDoCabo(ic).map(id => {
                    const no = grafo.nos.get(id);
                    return botao(numeroDoNo(no), () => {
                        destacarCabo(ic);
                        editarNo(id, nosSvg.get(id));
                    }, 'Abrir ' + nomeDoNo(no) + ' · ' + nome);
                });
                for (let ponta = 1; ponta <= qtd; ponta++) {
                    const b = botao(qtd === 1 ? 'Ponta livre' : 'Ponta livre ' + ponta, () => selecionarCabo(ic, ponta), nome + ' · ponta livre ' + ponta);
                    b.dataset.caboLivre = ic;
                    b.dataset.ponta = ponta;
                    botoes.push(b);
                }
                const titulo = botao(nome, () => destacarCabo(ic));
                const consulta = host.consultarAssistente ? botao('Consultar cabo', () => consultarCabo(ic)) : null;
                consulta?.classList.add('vz-grafo-consulta');
                titulo.classList.add('vz-grafo-cabo-titulo');
                return el('article', { class: 'vz-grafo-item ' + (grupo.tipo || '').toLowerCase(), 'data-cartao-cabo': ic }, [
                    titulo, el('small', { text: [grupo.tipo, c.operacao].filter(Boolean).join(' · ') }),
                    diagDe(ic).diverge ? el('small', { class: 'vz-grafo-alerta', text: '⚠ Nível ' + diagDe(ic).nivel + ' pelo ' + (diagDe(ic).fonte === 'linha' ? 'nível da linha' : diagDe(ic).fonte) + ', mas a linha do desenho é ' + diagDe(ic).geo + ' (confira a associação no CAD)' }) : null,
                    diagDe(ic).motivo ? el('small', { class: 'vz-grafo-alerta', text: '⚠ ' + diagDe(ic).motivo.charAt(0).toUpperCase() + diagDe(ic).motivo.slice(1) }) : null,
                    el('div', { class: 'vz-grafo-pontas' }, botoes),
                    controleEdicao('cabo', ic), host.copiarCaboNoMapa ? botao('Copiar', () => copiarCabo(ic), 'Copiar ' + nome + ' (cabo novo, sem pontas ligadas)') : null, consulta
                ]);
            };
            cartoes?.observador?.disconnect();
            cartoes = { ordem: [...cabosVisiveis], fabricar, feitos: 0, mais: null, observador: null };
            lista.replaceChildren();
            mostrarMaisCartoes(LOTE_CARTOES);
            if (caboEmFoco !== null) garantirCartao(caboEmFoco);
            if (vb) atualizarEscala();
            const semGrupo = pontosSemGrupo().length;
            // TASK-235: pontos no formato antigo (sem grupo ou com comentário em linha «#») — pergunta antes de converter
            if (host.converterPontosNoMapa && !st.adiarConversao && GLm) {
                const antigos = outros().reduce((n, r, i) => n + (r && r.grupoDecl && GLm.pontoDaLinha(outros(), i)?.formatoAntigo ? 1 : 0), 0);
                const total = antigos + semGrupo;
                avisoAntigo.hidden = !total;
                if (total) {
                    const converter = botao('Converter todos', () => {
                        fixarPosicoes();
                        const n = host.converterPontosNoMapa(pontosSemGrupo().map(no => ({ linhas: no.linhas, nome: 'P' + (linhaNo(no) + 1) })));
                        cx = null; metaAtivo.clear();
                        atualizar();
                        host.depois?.();
                        avisos.textContent = n + ' ponto(s) convertido(s) para «@nome - comentário». Ctrl+Z desfaz.';
                    }, 'Agrupa os pontos sem grupo e leva as linhas «#» para o título (juntas por «; »). Ctrl+Z desfaz.');
                    converter.classList.add('p');
                    avisoAntigo.replaceChildren(el('span', { text: '⚠ ' + total + ' ponto(s) no formato antigo (sem grupo ou com comentário em linha separada).' }), converter,
                        botao('Agora não', () => { st.adiarConversao = true; avisoAntigo.hidden = true; }));
                }
            } else avisoAntigo.hidden = true;
            btnAgrupar.disabled = !semGrupo;
            btnAgrupar.textContent = semGrupo ? 'Agrupar pontos (' + semGrupo + ')' : 'Agrupar pontos';
            tituloPontas.textContent = 'Cabos · ' + cabosVisiveis.size;
            pendencias = calcularPendencias(a, cabosVisiveis, nosVisiveis);
            renderPendencias(a, cabosVisiveis, nosVisiveis);
            lateralMapa.hidden = !cabosVisiveis.size && !pendencias.length;
            area.classList.toggle('sem-lateral', lateralMapa.hidden);
            vazio.hidden = !!(grafo.nos.size || grafo.cabosDoMapa.size);
            area.hidden = !vazio.hidden;
            sugerir.disabled = !grafo.nos.size || !grafo.cabosDoMapa.size;
            ficticio.disabled = ![...grafo.nos.values()].some(no => estruturasDoNo(no).some(ie => infoEst(ie).ocupa));
            avisos.textContent = !area.hidden && !nosVisiveis.size && !cabosVisiveis.size ? 'Nenhum ponto ou cabo para este filtro.'
                : !area.hidden ? 'Arrastar altera somente a disposição visual.' + (() => {
                    const aj = pontosAjustados(), c = [...posicoes.values()].length ? raiz.MapaLayout && raiz.MapaLayout.caixa([...posicoes.values()]) : null;
                    const grande = emEscalaAgora() && c && ((c[2] - c[0]) > (svg.clientWidth || W) * 1.05 || (c[3] - c[1]) > (svg.clientHeight || H) * 1.05);
                    return (aj ? ' ' + aj + ' ponto(s) com posição ajustada à mão (botão «Posições do CAD» volta às coordenadas).' : '')
                        + (grande ? ' Em 1:500 só parte da rede cabe na janela: arraste para andar ou escolha «Ajustar à tela».' : '');
                })() + (() => {
                    const ids = [...cabosVisiveis], div = ids.filter(ic => diagDe(ic).diverge).length, semGeo = ids.filter(ic => diagDe(ic).naoDesenhado).length, curtas = ids.filter(ic => diagDe(ic).motivo && !diagDe(ic).naoDesenhado).length;
                    return (div ? ' ⚠ ' + div + ' cabo(s) com nível do rótulo diferente do desenho.' : '') + (semGeo ? ' ⚠ ' + semGeo + ' cabo(s) sem linha do desenho e sem ponta ligada (não aparecem).' : '')
                        + (curtas ? ' ' + curtas + ' com ponta fora do raio de um poste.' : '') + (div || semGeo || curtas ? ' Use o filtro «A conferir».' : '');
                })() : '';
            atualizarSelecao();
            renderRevisao();
            desenharAlcas();
        }

        // ── TASK-233: montagem rápida ──────────────────────────────────────────────────────────────────────────
        const LOTE_CARTOES = 60;
        function mostrarMaisCartoes(n) {
            if (!cartoes) return;
            const fim = Math.min(cartoes.ordem.length, cartoes.feitos + n);
            const novos = [];
            for (let k = cartoes.feitos; k < fim; k++) novos.push(cartoes.fabricar(cartoes.ordem[k]));
            cartoes.feitos = fim;
            cartoes.mais?.remove();
            cartoes.observador?.disconnect();
            lista.append(...novos);
            const resta = cartoes.ordem.length - cartoes.feitos;
            if (resta > 0) {
                cartoes.mais = botao('Mostrar mais ' + resta + ' cabo(s)', () => mostrarMaisCartoes(LOTE_CARTOES));
                cartoes.mais.classList.add('vz-grafo-mais');
                lista.append(cartoes.mais);
                if (raiz.IntersectionObserver) {
                    cartoes.observador = new raiz.IntersectionObserver(ents => { if (ents.some(x => x.isIntersecting)) mostrarMaisCartoes(LOTE_CARTOES); }, { root: lista });
                    cartoes.observador.observe(cartoes.mais);
                }
            } else cartoes.mais = null;
        }
        function garantirCartao(ic) {
            if (!cartoes) return;
            const k = cartoes.ordem.indexOf(ic);
            if (k >= cartoes.feitos) mostrarMaisCartoes(k - cartoes.feitos + 1 + 10);
        }

        // alças nas pontas do cabo em foco: arrastar até um ponto liga a ponta (ímã)
        function desenharAlcas() {
            if (!camadaAlcas) return;
            const d = caboEmFoco !== null && host.inserirNoMapa && ferramenta !== 'vao' ? (desenhosPorCabo.get(caboEmFoco) || []).find(x => x.caminho) : null;
            if (!d) { if (!arrasto?.ponta) camadaAlcas.replaceChildren(); return; }
            // a alça fica um pouco para dentro do cabo: o ponto da ponta continua clicável/arrastável
            const porPx = vb ? vb[2] / (svg.clientWidth || W) : 1, recuo = 16 * porPx;
            const dentro = (a, b) => { const L = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1, t = Math.min(recuo, L * .4) / L; return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]; };
            const c = d.caminho, fins = [dentro(c[0], c[1]), dentro(c[c.length - 1], c[c.length - 2])];
            let alcas = [...camadaAlcas.querySelectorAll('.vz-alca')];
            if (alcas.length !== 2 || Number(alcas[0].dataset.ic) !== d.ic) {
                camadaAlcas.replaceChildren();
                alcas = [0, 1].map(i => {
                    const c = doc.createElementNS(ns, 'circle');
                    c.setAttribute('class', 'vz-alca'); c.setAttribute('r', '6');
                    c.dataset.arrastar = 'ponta'; c.dataset.ic = d.ic; c.dataset.i = i;
                    const t = doc.createElementNS(ns, 'title');
                    t.textContent = 'Ponta ' + (i + 1) + ' do cabo ' + (d.ic + 1) + ': arraste até um ponto para ligar';
                    c.appendChild(t);
                    camadaAlcas.appendChild(c);
                    return c;
                });
            }
            alcas.forEach((c, i) => { c.setAttribute('cx', fins[i][0].toFixed(1)); c.setAttribute('cy', fins[i][1].toFixed(1)); });
        }
        function idsDoCabo(ic) {
            const d = (desenhosPorCabo.get(ic) || [])[0];
            return (d ? d.v.ids : grafo.cabosDoMapa.get(ic)?.ids || [null, null]).slice(0, 2).concat([null, null]).slice(0, 2);
        }
        /** Liga a ponta `i` do cabo ao ponto `no` (troca o ponto, se ela já chegava em outro). Devolve a mensagem de erro ou ''. */
        function religarPonta(ic, i, no, silencioso) {
            const ids = idsDoCabo(ic), atual = ids[i] ? grafo.nos.get(ids[i]) : null;
            if (!no || (atual && atual.id === no.id)) return '';
            if (ids[1 - i] === no.id) return 'As duas pontas do cabo não podem ficar no mesmo ponto.';
            if (cabos()[ic].bloqueada) return 'O cabo está bloqueado.';
            if (atual && jaLigadoNo(ic, atual)) return 'Remova primeiro o vínculo com a estrutura de ' + nomeDoNo(atual) + '.';
            ids[i] = no.id;
            const linhasAfetadas = [...new Set([...ids, atual?.id].filter(Boolean))].map(id => linhaNo(grafo.nos.get(id)));
            if (!silencioso) host.antes();
            gravarPontos(ic, ids.filter(Boolean));
            if (atual && linhaFicticia(atual)) {
                const c = cabos()[ic], id = outros()[linhaNo(atual)]._mapaId;
                c._mapaAncoras = (c._mapaAncoras || []).filter(([anterior]) => anterior !== id);
            }
            if (silencioso) return linhasAfetadas;
            caboEscolhido = pontaEscolhida = null;
            ultimoStatus = comVinculosAutomaticos('Ponta do cabo ' + (ic + 1) + ' ligada a ' + nomeDoNo(no) + '.', automatizarPontos(linhasAfetadas));
            atualizar();
            host.depois?.();
            avisos.textContent = ultimoStatus;
            return '';
        }
        function ligarSugestoes() {
            const itens = pendencias.filter(p => p.tipo === 'ponta' && p.sugestao);
            if (!itens.length) return;
            host.antes();
            let n = 0; const afetadas = new Set();
            for (const p of itens) {
                const r = religarPonta(p.ic, p.i, grafo.nos.get(p.sugestao), true);
                if (Array.isArray(r)) { n++; r.forEach(l => afetadas.add(l)); }
            }
            const auto = automatizarPontos([...afetadas]);
            atualizar();
            host.depois?.();
            avisos.textContent = comVinculosAutomaticos(n + ' ponta(s) ligada(s) ao ponto mais perto. Ctrl+Z desfaz.', auto);
        }
        // ponto sem bloco no CAD: usa o bloco livre (ou confirma a posição da anotação) — vira coordenada do CAD
        function fixarCoordenada(no, p, texto) {
            host.antes();
            const linhas = no.linhas.filter(l => !outros()[l]._mapaPosteRef);
            linhas.forEach(l => { const r = outros()[l]; r._poste = [Math.round(p[0] * 100) / 100, Math.round(p[1] * 100) / 100]; delete r._posteAprox; delete r._mapaPoste; delete r._mapaPoste500; });
            cx = null;
            atualizar();
            host.depois?.();
            avisos.textContent = texto + ' Ctrl+Z desfaz.';
        }
        function transformarBloco(idBloco) {
            const b = fundoLivres.find(x => x.id === idBloco);
            if (!b || !host.inserirNoMapa) return;
            abrirCriacao('poste', null, null, { cad: b.p });
        }
        function alternarFundo() {
            fundoLigado = !fundoLigado;
            try { raiz.localStorage.setItem('vz_mapa_fundo', fundoLigado ? '1' : '0'); } catch (e) { /* vale só nesta abertura */ }
            pintar();
            avisos.textContent = fundoLigado ? (fundoLivres.length || svg.querySelector('.vz-fundo-rede') ? 'Fundo do CAD visível: ' + fundoLivres.length + ' bloco(s) de poste livre(s). Clique num bloco para transformá-lo em ponto.' : 'Sem fundo do CAD para esta obra (abra o projeto do plugin ou leia o DXF).') : 'Fundo do CAD oculto.';
        }

        // «Traçar vão»: origem → destino cria o cabo ligado nas duas pontas, com o comprimento medido
        function alternarVao(ligar = ferramenta !== 'vao') {
            ferramenta = ligar ? 'vao' : null;
            if (!ligar) { vaoOrigem = null; nosSvg.forEach(g => g.classList.remove('origem-vao')); }
            btnVao.classList.toggle('ligado', ligar);
            btnVao.setAttribute('aria-pressed', String(ligar));
            svg.classList.toggle('tracando', ligar);
            desenharAlcas();
            avisos.textContent = ligar ? 'Traçar vão: clique no ponto de origem (Esc encerra).' : 'Traçar vão encerrado.';
        }
        function metrosEntre(a, b) {
            if (emEscalaAgora()) { const p = posicoes.get(a), q = posicoes.get(b); return p && q ? MM.metros([p[0] / kLayout, p[1] / kLayout], [q[0] / kLayout, q[1] / kLayout]) : null; }
            return MM.metros(posCadNo(grafo.nos.get(a)), posCadNo(grafo.nos.get(b)));
        }
        function marcarOrigemVao(id) {
            vaoOrigem = id;
            nosSvg.forEach((g, k) => g.classList.toggle('origem-vao', k === id));
        }
        function cliqueVao(id, rapido) {
            const no = grafo.nos.get(id);
            if (!no) return;
            if (!vaoOrigem || !grafo.nos.has(vaoOrigem)) { marcarOrigemVao(id); avisos.textContent = 'Origem: ' + nomeDoNo(no) + '. Clique no ponto de destino.'; return; }
            if (vaoOrigem === id) { marcarOrigemVao(null); avisos.textContent = 'Origem desmarcada. Clique no ponto de origem.'; return; }
            const a = vaoOrigem, m = metrosEntre(a, id), modelo = st.modeloCabo;
            if (rapido && modelo && m) {
                fixarPosicoes();
                const r = host.inserirNoMapa('cabo', Object.assign({}, modelo, { metros: String(m), poste: linhaNo(grafo.nos.get(a)) }));
                if (r.erro) { avisos.textContent = r.erro; return; }
                concluirVao(r.indice, a, id);
                return;
            }
            abrirCriacao('cabo', linhaNo(grafo.nos.get(a)), null, { vao: { a, b: id }, metros: m });
        }
        function concluirVao(ic, a, b) {
            const la = linhaNo(grafo.nos.get(a)), lb = linhaNo(grafo.nos.get(b));
            const c = cabos()[ic];
            c.vinculoPostes = [la, lb]; c._mapaTopologia = true;
            cx = null;
            const auto = automatizarPontos([la, lb]);
            atualizar();
            host.depois?.();
            marcarOrigemVao(grafo.nosPorLinha.get(lb)?.id || null);
            avisos.textContent = comVinculosAutomaticos('Vão criado: cabo ' + (ic + 1) + ' (' + c.ativo + '). Clique no próximo ponto para continuar.', auto);
        }

        // seleção (Shift+arrastar no fundo ou Shift+clique num ponto): mover juntos, excluir com Del
        function pintarSelecao() { nosSvg.forEach((g, id) => g.classList.toggle('escolhido', selecao.has(id))); }
        function limparSelecao() { if (!selecao.size) return; selecao.clear(); pintarSelecao(); }
        function avisoSelecao() {
            avisos.textContent = selecao.size ? selecao.size + ' ponto(s) selecionado(s): arraste um deles para mover todos · Del exclui · Esc limpa.' : 'Seleção vazia.';
        }
        function excluirSelecao() {
            if (!host.removerDoMapa) return;
            const nos = [...selecao].map(id => grafo.nos.get(id)).filter(Boolean);
            const alvo = nos.length ? { outros: [...new Set(nos.flatMap(no => no.linhas))], cabos: [] } : caboEmFoco !== null ? { cabos: [caboEmFoco], outros: [] } : null;
            if (!alvo) { avisos.textContent = 'Selecione pontos (Shift+arrastar) ou um cabo para excluir.'; return; }
            const r = host.removerDoMapa(alvo);
            if (r.erro) { avisos.textContent = r.erro; return; }
            selecao.clear(); caboEmFoco = caboEscolhido = pontaEscolhida = null;
            cx = null; metaAtivo.clear();
            atualizar();
            host.depois?.();
            avisos.textContent = (r.outros ? r.outros + ' linha(s) de Demais Ativos' : '') + (r.outros && r.cabos ? ' e ' : '') + (r.cabos ? r.cabos + ' cabo(s)' : '') + ' excluído(s). Ctrl+Z desfaz.';
        }

        // Pendências: o que falta para a rede ficar montada e vinculada, com a correção a um clique
        function centralizarEm(p) {
            if (!p || !vb) return;
            vb = [p[0] - vb[2] / 2, p[1] - vb[3] / 2, vb[2], vb[3]];
            aplicarVista();
        }
        function focarNo(id) { const g = nosSvg.get(id); centralizarEm(posicoes.get(id)); g?.focus(); }
        function focarCabo(ic) {
            const d = (desenhosPorCabo.get(ic) || []).find(x => x.caminho);
            if (d) centralizarEm([(d.caminho[0][0] + d.caminho[d.caminho.length - 1][0]) / 2, (d.caminho[0][1] + d.caminho[d.caminho.length - 1][1]) / 2]);
            destacarCabo(ic, true);
        }
        function calcularPendencias(a, cabosVisiveis, nosVisiveis) {
            const itens = [];
            const visiveis = [...nosVisiveis].map(id => grafo.nos.get(id)).filter(Boolean).filter(no => !linhaFicticia(no));
            const aprox = visiveis.filter(no => !no.pos && no.posAprox);
            const pares = new Map((MM ? MM.casar(aprox.map(no => ({ id: no.id, p: no.posAprox })), fundoLivres, RAIO_BLOCO_LIVRE) : []).map(p => [p.a, p]));
            for (const no of visiveis) {
                if (!posCadNo(no)) itens.push({ tipo: 'sem-posicao', id: no.id, texto: nomeDoNo(no) + ': sem coordenada no CAD (está na coluna ao lado)',
                    acoes: [['Posicionar', () => iniciarPosicionamento(linhaNo(no))]] });
                else if (!no.pos) {
                    const par = pares.get(no.id), bloco = par && fundoLivres.find(b => b.id === par.b);
                    itens.push({ tipo: 'aproximada', id: no.id, texto: nomeDoNo(no) + ': sem bloco de poste no CAD, posição aproximada (a da anotação)',
                        acoes: [...(bloco ? [['Usar bloco livre (' + Math.round(par.d) + ' m)', () => fixarCoordenada(no, bloco.p, nomeDoNo(no) + ' no bloco livre do CAD.')]] : []),
                            ['Confirmar posição', () => fixarCoordenada(no, no.posAprox, 'Posição de ' + nomeDoNo(no) + ' confirmada.')]] });
                }
            }
            for (const ic of cabosVisiveis) {
                const c = cabos()[ic], nome = 'Cabo ' + (ic + 1) + ' · ' + c.ativo;
                const d = (desenhosPorCabo.get(ic) || [])[0], ids = idsDoCabo(ic);
                ids.forEach((id, i) => {
                    if (id) return;
                    const sug = d && d.ancoras && d.ancoras[i] && d.ancoras[i] !== ids[1 - i] ? d.ancoras[i] : null;
                    const noSug = sug && grafo.nos.get(sug);
                    itens.push({ tipo: 'ponta', ic, i, sugestao: sug, texto: nome + ': ponta ' + (i + 1) + ' solta' + (d ? '' : ' (sem desenho)'),
                        acoes: [...(noSug ? [['Ligar a ' + nomeDoNo(noSug), () => { const m = religarPonta(ic, i, noSug); if (m) avisos.textContent = m; }]] : []),
                            ...(livres(ic) ? [['Escolher ponto', () => { focarCabo(ic); selecionarCabo(ic, 1); }]] : [])] });
                });
                if (ids.every(Boolean) && infoCabo(ic).vincula && !VR.gruposVisuaisDoVao([ic], cabos(), a, st.cfg.dicionario_cabos)[0].completo) {
                    const pendente = ids.map(id => grafo.nos.get(id)).find(no => no && !jaLigadoNo(ic, no)) || grafo.nos.get(ids[0]);
                    itens.push({ tipo: 'vinculo', ic, texto: nome + ': vínculo com a estrutura pendente em ' + nomeDoNo(pendente),
                        acoes: [['Abrir ponto', () => { destacarCabo(ic); editarNo(pendente.id, nosSvg.get(pendente.id)); }]] });
                }
                if (diagDe(ic).diverge) itens.push({ tipo: 'nivel', ic, texto: nome + ': nível ' + diagDe(ic).nivel + ' no rótulo, linha ' + diagDe(ic).geo + ' no desenho', acoes: [] });
            }
            return itens;
        }
        function renderPendencias(a, cabosVisiveis, nosVisiveis) {
            const vinculaveis = [...cabosVisiveis].filter(ic => infoCabo(ic).vincula);
            const completos = vinculaveis.filter(ic => VR.gruposVisuaisDoVao([ic], cabos(), a, st.cfg.dicionario_cabos)[0].completo).length;
            const pontas = [...cabosVisiveis].reduce((s, ic) => s + idsDoCabo(ic).filter(Boolean).length, 0);
            const nos = [...nosVisiveis].map(id => grafo.nos.get(id)).filter(no => no && !linhaFicticia(no));
            const barra = (rotulo, feitos, total) => {
                const p = MM ? MM.progresso(feitos, total) : { pct: 0, texto: feitos + '/' + total };
                return el('div', { class: 'vz-pend-barra' }, [el('span', { text: rotulo + ' · ' + p.texto }), el('i', {}, [el('b', { style: 'width:' + p.pct + '%' })])]);
            };
            const sugestoesPontas = pendencias.filter(p => p.tipo === 'ponta' && p.sugestao).length;
            const LIMITE = 150;
            const lista = el('ul', {}, pendencias.slice(0, LIMITE).map(p => el('li', {}, [
                el('button', { class: 'vz-pend-alvo', type: 'button', text: p.texto, title: 'Mostrar no mapa', onclick: () => (p.id ? focarNo(p.id) : focarCabo(p.ic)) }),
                p.acoes.length ? el('div', { class: 'vz-pend-acoes' }, p.acoes.map(([t, f]) => botao(t, f))) : null
            ])));
            painelPend.replaceChildren(...[                                                           // (replaceChildren escreveria «null» como texto)
                el('summary', { text: pendencias.length ? 'Pendências · ' + pendencias.length : 'Rede montada ✓' }),
                barra('Vãos vinculados', completos, vinculaveis.length),
                barra('Pontas ligadas', pontas, cabosVisiveis.size * 2),
                barra('Pontos com coordenada do CAD', nos.filter(no => no.pos).length, nos.length),
                sugestoesPontas && host.inserirNoMapa ? botao('Ligar ' + sugestoesPontas + ' ponta(s) ao ponto mais perto', ligarSugestoes, 'Cada ponta solta vai para o ponto mais perto da ponta do desenho (até 25 m). Ctrl+Z desfaz tudo.') : null,
                pendencias.length ? lista : null,
                pendencias.length > LIMITE ? el('small', { text: '… e mais ' + (pendencias.length - LIMITE) + '. Resolva as primeiras ou use os filtros.' }) : null
            ].filter(x => x));
            painelPend.hidden = !(nos.length || cabosVisiveis.size);
        }

        // atalhos do mapa (fora de campos de texto e diálogos)
        const teclaMapa = e => {
            if (colocando || e.ctrlKey || e.metaKey || e.altKey || dialogo.open || editor.open) return;
            const t = e.target;
            if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName) || t.closest?.('dialog'))) return;
            if (!svg.isConnected || !svg.getClientRects().length) return;                     // TASK-241: Mapa minimizado/escondido não pega teclas
            if (t && t !== doc.body && t !== doc.documentElement && !overlay?.contains(t)) return;       // foco em outra tela do Resumo
            const k = e.key;
            if (k === 'Escape') {
                // TASK-241: Esc no Mapa só deseleciona (ferramenta, seleção, ponta escolhida, cabo em foco) — nunca fecha o Mapa
                e.preventDefault(); e.stopPropagation();
                if (ferramenta) alternarVao(false);
                limparSelecao();
                if (caboEscolhido !== null || caboEmFoco !== null) { caboEscolhido = pontaEscolhida = null; caboEmFoco = null; atualizarSelecao(); desenharAlcas(); }
                return;
            }
            if (!host.inserirNoMapa) return;
            const acao = { v: () => alternarVao(), V: () => alternarVao(), p: () => abrirPonto(), c: () => abrirCriacao('cabo'),
                l: () => { if (caboEmFoco !== null && livres(caboEmFoco)) selecionarCabo(caboEmFoco, 1); else avisos.textContent = 'Selecione um cabo com ponta livre para ligar (L).'; },
                Delete: () => excluirSelecao() }[k];
            if (!acao) return;
            e.preventDefault();
            acao();
        };
        doc.addEventListener('keydown', teclaMapa, true);

        const RAIO_ANCORA_VISUAL = 25;                                       // metros do CAD
        // posição de uma ponta de cabo sem poste: a do desenho + o quanto o ponto mais perto saiu do lugar original
        function pontaAncorada(d, i) {
            const p = ponto(d.v.pontas[i]), a = d.ancoras && d.ancoras[i];
            if (!a) return p;
            const pos = posicoes.get(a), base = bases.get(a);
            return pos && base ? [p[0] + pos[0] - base[0], p[1] + pos[1] - base[1]] : p;
        }
        // caixa de um ponto no mapa (marcador + rótulo à direita)
        const caixaDoNo = id => { const [x, y] = posicoes.get(id); return [x - 26, y - 26, x + 200, y + 26]; };
        const emEscalaAgora = () => escalar() && kLayout !== null;                  // sem coordenadas do CAD não existe escala: vale o modo «Ajustar»
        const campoPonto = () => (emEscalaAgora() ? '_mapaPoste500' : '_mapaPoste'), campoCabo = () => (emEscalaAgora() ? '_mapaCabo500' : '_mapaCabo');
        const escalaCurva = () => (emEscalaAgora() ? [kLayout, kLayout] : [W, H]);
        // TASK-240: cabos entre o MESMO par de pontos ficam lado a lado — juntos por par (e não pela geometria de cada um), cada um com a sua faixa,
        // do mesmo lado qualquer que seja o sentido em que foi traçado — e o espaço entre eles nunca fica abaixo de SEP_CABO_PX na tela.
        const SEP_CABO_PX = 9, SEP_CABO_MIN = 10;
        let unidadesDaSeparacao = 0;
        const separacaoCabos = () => Math.max(SEP_CABO_MIN, SEP_CABO_PX * (vb ? vb[2] / (svg.clientWidth || W) : 1));
        function faixasDosCabos(lista) {
            const grupos = new Map();
            for (const d of lista) {
                const [a, b] = d.v.ids, par = a && b && a !== b ? [a, b].sort() : null;
                const chave = par ? par.join('|') : d.v;
                if (!grupos.has(chave)) grupos.set(chave, []);
                grupos.get(chave).push(d);
                d.inverte = !!par && a !== par[0];
            }
            const ordem = d => (d.grupo?.tipo === 'MT' ? 0 : d.grupo?.tipo === 'BT' ? 1 : 2) * 1e6 + (d.ic ?? 5e5);
            for (const g of grupos.values()) g.sort((x, y) => ordem(x) - ordem(y)).forEach((d, i) => { d.faixa = i - (g.length - 1) / 2; });
        }       // o «puxar» do cabo é guardado em fração da tela (Ajustar) ou em metros (1:500)
        function desenharPosicoes(alvo) {
            const desenharNo = (g, id) => g.setAttribute('transform', 'translate(' + posicoes.get(id).join(' ') + ')');
            const idsAlvo = alvo?.ids || (alvo?.idNo ? [alvo.idNo] : null);                       // TASK-233: vários pontos (seleção) redesenham só o que muda
            if (idsAlvo) {
                for (const id of idsAlvo) {
                    if (nosSvg.has(id)) desenharNo(nosSvg.get(id), id);
                    const it = itensPorNo.get(id);
                    if (it) { it.b = caixaDoNo(id); if (!it.v) { it.v = true; it.camada.appendChild(it.el); } }
                }
            } else if (!alvo) nosSvg.forEach(desenharNo);
            const lista = idsAlvo ? [...new Set(idsAlvo.flatMap(id => desenhosPorNo.get(id) || []))] : alvo?.ic !== undefined ? desenhosPorCabo.get(alvo.ic) || [] : desenhos;
            if ((!idsAlvo && alvo?.ic === undefined) || !unidadesDaSeparacao) unidadesDaSeparacao = separacaoCabos();
            const [cw, ch] = escalaCurva();
            for (const d of lista) {
                const {v, ic, linha} = d;
                const deslocamento = d.faixa * unidadesDaSeparacao * (d.inverte ? -1 : 1);
                const pts = v.ids.map((id, i) => id ? posicoes.get(id) : v.pontas && pontaAncorada(d, i));
                if (!pts.every(Boolean)) {
                    const i = pts[0] ? 0 : 1;
                    const [x, y] = pts[i], sinal = x > W / 2 ? -1 : 1;
                    const lado = d.grupo.tipo === 'MT' ? -sinal : sinal;
                    pts[1 - i] = [x + lado * (95 + ((ic ?? 0) % 5) * 9), y + (y > H / 2 ? -76 : 76)];
                }
                const [x0, y0] = pts[0], [x1, y1] = pts[1];
                const dx = x1 - x0, dy = y1 - y0, comprimento = Math.hypot(dx, dy) || 1;
                const curva = ic !== null && (arrasto?.ic === ic && arrasto.valor ? arrasto.valor : cabos()[ic][campoCabo()]);
                const caminho = v.pontos && !curva ? v.pontos.map((pt, i) => {
                    if (i === 0) return pts[0];
                    if (i === v.pontos.length - 1) return pts[1];
                    const t = i / (v.pontos.length - 1), [px, py] = ponto(pt);
                    const originalA = ponto(v.pontos[0]), originalB = ponto(v.pontos[v.pontos.length - 1]);
                    return [px + (x0 - originalA[0]) * (1 - t) + (x1 - originalB[0]) * t,
                        py + (y0 - originalA[1]) * (1 - t) + (y1 - originalB[1]) * t];
                })
                    : curva ? [pts[0], [x0 + dx / 2 + curva[0] * cw, y0 + dy / 2 + curva[1] * ch], pts[1]] : pts;
                const pontos = caminho.map(([x, y]) =>
                    `${(x - dy * deslocamento / comprimento).toFixed(2)},${(y + dx * deslocamento / comprimento).toFixed(2)}`).join(' ');
                linha.setAttribute('points', pontos);
                d.caminho = caminho;
                if (linha.parentNode.hasAttribute('data-cabo-mapa')) linha.nextElementSibling.setAttribute('points', pontos);
                d.b = ME ? ME.caixaDe(caminho, 14 + Math.abs(deslocamento)) : null;
                if (d.item) { d.item.b = d.b; if (alvo && !d.item.v) { d.item.v = true; d.item.camada.appendChild(d.item.el); } }
            }
            if (caboEmFoco !== null && !arrasto?.ponta) desenharAlcas();
        }

        function campo(rotulo, controle) {
            controle.setAttribute('aria-label', rotulo);
            return el('label', {}, [el('span', { text: rotulo }), controle]);
        }
        function opcoes(valores) {
            return valores.map(([value, text]) => el('option', { value, text }));
        }
        const operacoesEdicao = [
            ['F', 'F · Fictício (fora do orçamento, só para o vínculo)'],
            ['I', 'I · Instalar'], ['*I', '*I · Instalar em linha viva'],
            ['M', 'M · Manter ou reaplicar'], ['*M', '*M · Manter ou reaplicar em linha viva'],
            ['R', 'R · Remover'], ['*R', '*R · Remover em linha viva']
        ];
        function selectComValor(opcoesCampo, valor) {
            const atual = String(valor ?? '');
            const todas = opcoesCampo.some(([v]) => v === atual) ? opcoesCampo : [[atual, atual || 'Sem valor'], ...opcoesCampo];
            const campo = el('select', {}, opcoes(todas));
            campo.value = atual;
            return campo;
        }

        function controleEdicao(tipo, indice, rotulo) {
            if (!host.editarNoMapa) return null;
            const referencia = (tipo === 'cabo' ? cabos() : outros())[indice];
            if (!referencia) return null;
            const nome = rotulo || (tipo === 'cabo' ? 'cabo ' + (indice + 1) : 'linha ' + (indice + 1));
            const caixa = el('div', { class: 'vz-edicao-cartao' });
            const abrir = botao('Editar', iniciar, 'Editar ' + nome);
            abrir.setAttribute('aria-expanded', 'false');
            abrir.disabled = tipo !== 'poste-ficticio' && Boolean(referencia.bloqueada || referencia.comentario);
            caixa.append(abrir);
            function iniciar() {
                fecharEdicaoAtiva?.(false);
                const cabo = tipo === 'cabo', ficticio = tipo === 'poste-ficticio';
                const ativo = el('input', { type: 'text', maxlength: ficticio ? '80' : '500',
                    ...(ficticio ? {} : { required: '', 'data-sugestao-origem': cabo ? 'CABOS' : 'OUTROS' }) });
                ativo.value = ficticio ? (referencia.ativo === 'Poste fictício' ? '' : referencia.ativo.replace(/^Poste fictício · /, '')) : referencia.ativo;
                const campos = [campo(ficticio ? 'Nome no mapa' : cabo ? 'Anotação do cabo (inclui comprimento)' : 'Anotação da linha', ativo)];
                let coletar;
                if (ficticio) {
                    const estruturas = el('input', { type: 'text', maxlength: '440', required: '', 'data-sugestao-origem': 'OUTROS' });
                    estruturas.value = (referencia._mapaEstruturasFicticias || []).join(', ');
                    campos.push(campo('Estruturas fictícias, separadas por vírgula', estruturas));
                    coletar = () => ({ ativo: trim(ativo.value), estruturas: [...new Set(estruturas.value.split(',').map(s => trim(s).toUpperCase()).filter(Boolean))] });
                } else {
                    const operacao = selectComValor(operacoesEdicao, referencia.operacao);
                    campos.push(campo('Operação', operacao));
                    if (cabo) {
                        const nivel = selectComValor([['', 'Pelo ativo'], ['MT', 'MT'], ['BT', 'BT']], referencia.nivel);
                        const fases = el('input', { type: 'text', inputmode: 'numeric', maxlength: '4', placeholder: 'Pelo ativo' });
                        const tensao = el('input', { type: 'text', maxlength: '40', placeholder: 'Pelo ativo' });
                        fases.value = referencia.fases ?? '';
                        tensao.value = referencia.tensao ?? '';
                        campos.push(campo('Fases', fases), campo('Tensão', tensao), campo('Nível', nivel));
                        coletar = () => ({ ativo: ativo.value, operacao: operacao.value, fases: trim(fases.value), tensao: trim(tensao.value), nivel: nivel.value });
                    } else {
                        campos.push(el('small', { class: 'vz-no-ajuda', text: 'A edição usa a linha da tabela, inclusive os itens que compartilham a mesma anotação.' }));
                        coletar = () => ({ ativo: ativo.value, operacao: operacao.value });
                    }
                }
                const cancelar = botao('Cancelar', () => fecharForm(true));
                cancelar.type = 'button';
                const mensagem = el('p', { class: 'vz-composer-erro', role: 'alert' });
                const salvar = el('button', { type: 'submit', class: 'vz-btn p', text: 'Salvar edição' });
                const form = el('form', { class: 'vz-composer-form', onkeydown: e => {
                    if (e.key !== 'Escape') return;
                    e.preventDefault();
                    e.stopPropagation();
                    fecharForm(true);
                }, onsubmit: e => {
                    e.preventDefault();
                    if (st.projeto !== host.projeto()) {
                        mensagem.textContent = 'O projeto mudou. Reabra o mapa para editar.';
                        return;
                    }
                    const dados = coletar(), anterior = JSON.stringify(referencia);
                    const erro = ficticio && (!dados.estruturas.length || dados.estruturas.length > 20 || dados.estruturas.some(s => s.length > 80))
                        ? 'Informe entre 1 e 20 estruturas com nomes de até 80 caracteres.'
                        : host.editarNoMapa(tipo, indice, referencia, dados, st.cfg);
                    if (erro) { mensagem.textContent = erro; return; }
                    fecharForm(false);
                    if (cabo) caboEmFoco = indice;
                    const linhas = cabo ? VR.postesAtribuidos(cabos()[indice]) : [indice];
                    const total = !ficticio && anterior !== JSON.stringify(referencia) ? automatizarPontos(linhas) : 0;
                    ultimoStatus = comVinculosAutomaticos('Linha atualizada na tabela.', total);
                    cx = null;
                    metaAtivo.clear();
                    atualizar();
                    if (total) host.depois?.();
                    avisos.textContent = ultimoStatus;
                    if (dialogo.open) dialogo.querySelector('[autofocus]')?.focus();
                    else lista.querySelector(`[data-cartao-cabo="${indice}"] .vz-grafo-cabo-titulo`)?.focus();
                } }, [...campos, mensagem, el('div', { class: 'vz-no-controles' }, [salvar, cancelar])]);
                function fecharForm(retornarFoco) {
                    form.remove();
                    abrir.hidden = false;
                    abrir.setAttribute('aria-expanded', 'false');
                    if (fecharEdicaoAtiva === fecharForm) fecharEdicaoAtiva = null;
                    if (retornarFoco) abrir.focus();
                }
                fecharEdicaoAtiva = fecharForm;
                abrir.hidden = true;
                abrir.setAttribute('aria-expanded', 'true');
                caixa.append(form);
                ativo.focus();
            }
            return caixa;
        }

        function abrirEditor(titulo, textoBotao, montagem, executar) {
            const erro = el('p', { class: 'vz-composer-erro', role: 'alert' });
            const fecharEditor = botao('Cancelar', () => editor.close());
            fecharEditor.type = 'button';
            const salvar = el('button', { class: 'vz-btn p', type: 'submit', text: textoBotao });
            const form = el('form', { class: 'vz-composer-form', onsubmit: e => {
                e.preventDefault();
                const mensagem = executar();
                if (mensagem) erro.textContent = mensagem;
                else editor.close();
            } }, [...montagem, erro, el('div', { class: 'vz-ficticio-opcoes' }, [salvar, fecharEditor])]);
            editor.replaceChildren(el('header', { class: 'vz-no-topo' }, [
                el('h4', { id: 'vz-composer-titulo', text: titulo })
            ]), form);
            if (!editor.open) editor.showModal();
            form.querySelector('input, select')?.focus();
        }

        function pedirEstruturasFicticias(titulo, salvar, comNome = false) {
            const nome = el('input', { type: 'text', maxlength: '80', placeholder: 'Poste fictício' });
            const codigos = el('input', { type: 'text', maxlength: '440', placeholder: 'N1, N3', 'data-sugestao-origem': 'OUTROS' });
            const campos = comNome ? [campo('Nome no mapa', nome)] : [];
            campos.push(campo('Estruturas fictícias, separadas por vírgula', codigos));
            abrirEditor(titulo, 'Salvar', campos, () => {
                const lista = [...new Set(codigos.value.split(',').map(c => c.trim().toUpperCase()).filter(Boolean))];
                if (lista.length > 20 || lista.some(c => c.length > 80)) return 'Use até 20 estruturas com nomes de até 80 caracteres.';
                if (!comNome && !lista.length) return 'Informe pelo menos uma estrutura fictícia.';
                return salvar(nome.value.trim(), lista.length ? lista : ['Ancoragem']);
            });
        }

        function abrirPosteFicticio() {
            pedirEstruturasFicticias('Novo poste fictício', (nome, estruturas) => {
                fixarPosicoes();
                const indice = host.inserirPosteFicticioNoMapa(nome, estruturas);
                if (filtroMapa !== 'todos' && filtroMapa !== 'ficticios') filtro.value = filtroMapa = 'todos';
                atualizar();
                colocarNovoPonto(indice);
                host.depois?.();
                avisos.textContent = 'Poste fictício criado. Conecte uma ponta para ancorar o cabo.';
                iniciarPosicionamento(indice);
                return '';
            }, true);
        }

        function estruturaDaAncora(ic, no) {
            const id = outros()[linhaNo(no)]._mapaId;
            return cabos()[ic]._mapaAncoras?.find(([anterior]) => anterior === id)?.[1];
        }

        function gravarAncora(ic, no, estrutura) {
            const row = outros()[linhaNo(no)], c = cabos()[ic];
            if (!row._mapaId || !pontosDoCabo(ic).includes(no.id)) return;
            host.antes();
            const ancoras = (c._mapaAncoras || []).filter(([id]) => id !== row._mapaId);
            if (estrutura) ancoras.push([row._mapaId, estrutura]);
            c._mapaAncoras = ancoras;
            ultimoStatus = estrutura ? 'Ponta ancorada a ' + estrutura + '.' : 'Âncora interna removida.';
            atualizar();
        }

        function abrirCriacao(tipo, poste = null, edicao = null, extra = {}) {
            const cabo = tipo === 'cabo';
            const vao = cabo && extra.vao ? extra.vao : null;                                     // TASK-233: «Traçar vão» (origem → destino)
            const titulo = edicao ? 'Editar poste' : vao ? 'Novo vão · ' + nomeDoNo(grafo.nos.get(vao.a)) + ' → ' + nomeDoNo(grafo.nos.get(vao.b))
                : cabo ? 'Novo cabo' : tipo === 'poste' ? (extra.cad ? 'Novo poste no bloco do CAD' : 'Novo poste') : 'Nova estrutura';
            const exemplo = cabo ? '3xA35P-13,8kV ou CAA2' : tipo === 'poste' ? '12600D, N3' : 'N3 ou 2-N3';
            const ativo = el('input', { type: 'text', required: '', maxlength: '440', placeholder: exemplo, 'data-sugestao-origem': cabo ? 'CABOS' : 'OUTROS' });
            const referencia = edicao && outros()[edicao.poste];
            if (referencia) ativo.value = referencia.ativo;
            const operacao = cabo ? seletorOperacao('I') : selectComValor(operacoesEdicao, referencia?.operacao || 'I');        // TASK-235: no cabo, F = cabo fictício
            const metros = el('input', { type: 'text', inputmode: 'decimal', maxlength: '24', required: '', placeholder: '25' });
            const fases = (raiz.GradeTabela?.FASES || ['A', 'B', 'C', 'AB', 'AC', 'BC', 'ABC']);
            const fase = el('select', {}, opcoes([['', 'Do rótulo'], ...fases.map(v => [v, v])]));
            const nivel = el('select', {}, opcoes([['', 'Pelo ativo'], ['MT', 'MT'], ['BT', 'BT']]));
            if (vao && st.modeloCabo) {                                                                // o último cabo criado é o modelo do próximo vão
                ativo.value = st.modeloCabo.ativo || '';
                operacao.value = st.modeloCabo.operacao || operacao.value;
                fase.value = st.modeloCabo.fase || '';
                nivel.value = st.modeloCabo.nivel || '';
            }
            if (cabo && extra.metros) metros.value = String(extra.metros);
            const campos = [campo(cabo ? 'Ativo ou rótulo, sem a metragem' : 'Anotação', ativo), campo('Operação', operacao)];
            if (cabo) campos.push(campo('Comprimento em metros', metros), campo('Fases', fase), campo('Nível', nivel));
            const novoPonto = !edicao && !cabo && poste === null && !!host.definirPontoNoMapa;
            const infoPonto = edicao ? host.pontoNoMapa?.(edicao.linhas) : null;
            const nomePonto = el('input', { type: 'text', maxlength: '60', placeholder: 'Automático' });
            const comentarioPonto = el('input', { type: 'text', maxlength: '300', placeholder: 'Opcional' });
            nomePonto.value = infoPonto?.nome || '';
            comentarioPonto.value = infoPonto?.comentario || '';
            if (novoPonto || edicao) campos.push(campo('Nome do ponto (vazio = número automático)', nomePonto), campo('Comentário do ponto', comentarioPonto));
            abrirEditor(titulo, edicao ? 'Salvar edição' : cabo ? 'Adicionar a Cabos' : 'Adicionar a Demais Ativos', campos, () => {
                if (st.projeto !== host.projeto()) return 'O projeto mudou. Reabra o mapa para editar.';
                if (edicao) {
                    const antes = JSON.stringify(referencia);
                    const posicaoAnterior = posicoes.get(edicao.id);
                    const erro = host.editarNoMapa('poste', edicao.poste, referencia, {
                        ativo: ativo.value, operacao: operacao.value,
                        ponto: { linhas: edicao.linhas, nome: nomePonto.value, comentario: comentarioPonto.value, manterSemGrupo: true }
                    }, st.cfg);
                    if (erro) return erro;
                    const indice = outros().indexOf(referencia);
                    cx = null; metaAtivo.clear();
                    const auto = antes !== JSON.stringify(referencia) ? automatizarPontos([indice]) : 0;
                    grafo = montarGrafo();
                    st.grafoNo = grafo.nosPorLinha.get(indice)?.id || null;
                    ultimoStatus = comVinculosAutomaticos('Poste atualizado em Demais Ativos.', auto);
                    atualizar();
                    const posicaoAtual = posicoes.get(st.grafoNo);
                    if (posicaoAnterior && posicaoAtual && Math.hypot(posicaoAtual[0] - posicaoAnterior[0], posicaoAtual[1] - posicaoAnterior[1]) > 0.01) {
                        // Ao criar um grupo, a reordenação das linhas pode mudar a célula da grade esquemática.
                        posicoes.set(st.grafoNo, posicaoAnterior);
                        guardarPosicao(st.grafoNo);
                        atualizar();
                    }
                    ultimoNo = nosSvg.get(st.grafoNo);
                    host.depois?.();
                    return '';
                }
                fixarPosicoes();                                                                          // os pontos que já existem não saem do lugar
                const resultado = host.inserirNoMapa(tipo, { ativo: ativo.value, operacao: operacao.value,
                    metros: metros.value, fase: fase.value, nivel: nivel.value, poste, nome: nomePonto.value, comentario: comentarioPonto.value });
                if (resultado.erro) return resultado.erro;
                if (cabo) st.modeloCabo = { ativo: ativo.value, operacao: operacao.value, fase: fase.value, nivel: nivel.value };
                if (vao) { concluirVao(resultado.indice, vao.a, vao.b); return ''; }
                if (novoPonto && extra.cad) outros()[resultado.indice]._poste = [Math.round(extra.cad[0] * 100) / 100, Math.round(extra.cad[1] * 100) / 100];
                cx = null;
                const linhas = cabo ? VR.postesAtribuidos(cabos()[resultado.indice]) : poste === null ? [] : [resultado.indice];
                const auto = automatizarPontos(linhas);
                const texto = comVinculosAutomaticos(cabo ? 'Cabo criado na tabela Cabos.' : 'Linha criada em Demais Ativos.', auto);
                if (cabo) {
                    grafo = montarGrafo();
                    caboEscolhido = livres(resultado.indice) ? resultado.indice : null;
                    pontaEscolhida = caboEscolhido === null ? null : 1;
                }
                if (dialogo.open) ultimoStatus = texto;
                atualizar();
                if (auto) host.depois?.();
                avisos.textContent = texto;
                if (novoPonto && extra.cad) { host.depois?.(); avisos.textContent = 'Ponto criado no bloco do CAD. ' + texto; }
                else if (novoPonto) { colocarNovoPonto(resultado.indice); iniciarPosicionamento(resultado.indice); }
                return '';
            });
        }

        // ── TASK-235: o ponto = nome - comentário + linhas, cada uma com a operação (F = fictícia: fora do orçamento, só para o vínculo) ──
        const OPS_PONTO = ['F', '|', 'I', 'M', 'R', '|', '*I', '*M', '*R'];
        function seletorOperacao(valor, aoMudar) {
            const caixa = el('div', { class: 'vz-ops', role: 'radiogroup', 'aria-label': 'Operação' });
            let atual = OPS_PONTO.includes(valor) && valor !== '|' ? valor : 'I';
            const botoes = OPS_PONTO.map(op => op === '|' ? el('span', { class: 'vz-ops-sep' }) : el('button', {
                type: 'button', class: 'vz-op-chip' + (op === 'F' ? ' f' : op.includes('R') ? ' r' : op.includes('M') ? ' m' : ''), text: op, role: 'radio',
                title: op === 'F' ? 'Fictício: fica fora do orçamento e vale só para o vínculo' : op, onclick: () => definir(op, true)
            }));
            function definir(op, doUsuario) {
                atual = op;
                botoes.forEach(b => { if (b.tagName === 'BUTTON') { const sel = b.textContent === op; b.classList.toggle('sel', sel); b.setAttribute('aria-checked', String(sel)); } });
                if (doUsuario) aoMudar?.(op);
            }
            caixa.append(...botoes);
            Object.defineProperty(caixa, 'value', { get: () => atual, set: v => definir(v, false) });
            definir(atual, false);
            return caixa;
        }
        const ehPosteTexto = texto => { try { return (host.postesDaLinha?.(texto) || []).length > 0; } catch (e) { return false; } };
        function tipoDoPonto(itens) {
            const cheias = itens.filter(x => x.ativo);
            if (!cheias.length) return '';
            if (cheias.every(x => x.op === 'F')) return 'Fictício';
            return cheias.some(x => ehPosteTexto(x.ativo)) ? 'Poste' : 'Estrutura avulsa';
        }
        /** Linhas do ponto (as da tabela, sem título nem comentário «#»). */
        function linhasDoPonto(no) {
            const info = host.pontoNoMapa?.(no.linhas);
            const idx = info ? info.membros : no.linhas.filter(l => outros()[l] && !outros()[l].comentario && !outros()[l].grupoDecl);
            return { info, idx: idx.filter(l => outros()[l] && !outros()[l].comentario) };
        }
        const opDaLinha = r => { const o = trim(r.operacao).toUpperCase(); return o === '' ? 'I' : o; };
        const OPS_INSTALANDO = ['I', '*I', 'M', '*M'];                                       // TASK-238: em destaque na lista de estruturas do ponto
        const seloOperacao = op => el('span', { class: 'vz-op-badge' + (/R$/.test(op) ? ' r' : /M$/.test(op) ? ' m' : op === 'F' ? ' f' : ''), text: op, title: 'Operação ' + op });
        /** Lista editável de linhas: [nº, ativo, operação, ×]. Ao preencher a última, entra outra vazia com a operação de cima. */
        function editorDeLinhas(iniciais, aoMudar) {
            const lista = el('div', { class: 'vz-linhas-ponto' });
            const linhas = [];
            function nova(x) {
                const ativo = el('input', { type: 'text', maxlength: '440', placeholder: 'ex.: 11300D, N3 · N3 · CE2', 'data-sugestao-origem': 'OUTROS', 'aria-label': 'Linha do ponto' });
                ativo.value = x.ativo || '';
                const ops = seletorOperacao(x.op || 'I', () => aoMudar?.());
                const n = el('span', { class: 'vz-lp-n' });
                const remover = el('button', { type: 'button', class: 'vz-lp-x', text: '×', title: 'Tirar esta linha do ponto', onclick: () => { lista.removeChild(item.el); linhas.splice(linhas.indexOf(item), 1); renumerar(); garantirVazia(); aoMudar?.(); } });
                const item = { indice: Number.isInteger(x.indice) ? x.indice : null, ativo, ops, el: el('div', { class: 'vz-lp' }, [n, ativo, ops, remover]) };
                ativo.addEventListener('input', () => { garantirVazia(); aoMudar?.(); });
                linhas.push(item); lista.append(item.el);
                return item;
            }
            function renumerar() { linhas.forEach((l, i) => { l.el.firstChild.textContent = i + 1; }); }
            function garantirVazia() {
                const ult = linhas[linhas.length - 1];
                if (!ult || trim(ult.ativo.value)) nova({ op: ult ? ult.ops.value : 'I' });
                renumerar();
            }
            iniciais.forEach(nova);
            garantirVazia();
            return { el: lista, coletar: () => linhas.map(l => ({ indice: l.indice, ativo: trim(l.ativo.value), op: l.ops.value })), focar: () => linhas[0]?.ativo.focus() };
        }
        /** Depois de gravar o ponto: o grafo é refeito e o painel/seleção seguem a linha `indice`. */
        function depoisDoPonto(indice, mensagem, novo) {
            cx = null; metaAtivo.clear();
            grafo = montarGrafo();
            const id = grafo.nosPorLinha.get(indice)?.id || null;
            const auto = id ? automatizarPontos(grafo.nos.get(id).linhas) : 0;
            if (st.grafoNo !== null && !novo) st.grafoNo = id;
            ultimoStatus = comVinculosAutomaticos(mensagem, auto);
            atualizar();
            host.depois?.();
            avisos.textContent = ultimoStatus;
            if (novo) { colocarNovoPonto(indice); iniciarPosicionamento(indice); }
        }
        /** Formulário do ponto (novo ou existente): nome - comentário na mesma linha, prévia do título e as linhas com a operação. */
        function abrirPonto(no = null) {
            if (!host.salvarPontoNoMapa) { abrirCriacao('poste'); return; }
            const { info, idx } = no ? linhasDoPonto(no) : { info: null, idx: [] };
            const nome = el('input', { type: 'text', maxlength: '60', placeholder: 'Automático' });
            const comentario = el('input', { type: 'text', maxlength: '300', placeholder: 'Opcional' });
            nome.value = info?.nome || ''; comentario.value = info?.comentario || '';
            const previa = el('div', { class: 'vz-previa-titulo' }), tipo = el('span', { class: 'vz-tipo-ponto' });
            let forcar = false;
            const linhas = editorDeLinhas(idx.map(l => ({ indice: l, ativo: outros()[l].ativo, op: opDaLinha(outros()[l]) })), () => { forcar = false; atualizarPrevia(); });
            function atualizarPrevia() {
                const n = trim(nome.value).toUpperCase() || (info?.nome || 'P…'), c = trim(comentario.value);
                previa.textContent = '@' + n + (c ? ' - ' + c : '');
                const t = tipoDoPonto(linhas.coletar());
                tipo.textContent = t ? 'Tipo: ' + t : '';
                tipo.className = 'vz-tipo-ponto ' + (t === 'Poste' ? 'poste' : t === 'Fictício' ? 'fic' : 'avulso');
            }
            [nome, comentario].forEach(i => i.addEventListener('input', () => { forcar = false; atualizarPrevia(); }));
            const nomeLinha = el('div', { class: 'vz-nome-linha' }, [campo('Nome (vazio = automático)', nome), el('span', { class: 'vz-sep', text: '-' }), campo('Comentário', comentario)]);
            atualizarPrevia();
            abrirEditor(no ? 'Editar ponto' : 'Novo ponto', no ? 'Salvar ponto' : 'Criar e posicionar', [nomeLinha, previa, tipo, el('span', { class: 'vz-rotulo-linhas', text: 'Linhas do ponto' }), linhas.el], () => {
                if (st.projeto !== host.projeto()) return 'O projeto mudou. Reabra o mapa para editar.';
                if (!no) fixarPosicoes();
                const r = host.salvarPontoNoMapa({ linhas: no ? no.linhas : null, nome: nome.value, comentario: comentario.value, itens: linhas.coletar(), forcar }, st.cfg);
                if (r.erro) return r.erro;
                if (r.conflitos) { forcar = true; return 'Vínculos que cairiam: ' + r.conflitos.join(' · ') + '. Clique em «' + (no ? 'Salvar ponto' : 'Criar') + '» de novo para confirmar (Ctrl+Z desfaz).'; }
                depoisDoPonto(r.indice, no ? 'Ponto atualizado em Demais Ativos' + (r.quebrados ? ' · ' + r.quebrados + ' vínculo(s) removido(s).' : '.') : 'Ponto ' + r.nome + ' criado.', !no);
                return '';
            });
            editor.classList.add('vz-ponto-dialog');
            editor.addEventListener('close', () => editor.classList.remove('vz-ponto-dialog'), { once: true });
            linhas.focar();
        }
        /** Painel do ponto: «Linhas do ponto» com edição na própria linha (ativo + operação). */
        function secaoLinhasDoPonto(no, status) {
            const { idx } = linhasDoPonto(no);
            const caixa = el('section', { class: 'vz-secao-linhas' });
            const salvarItens = (itens, mensagem) => {
                let forcar = !!salvarItens.forcar;
                const r = host.salvarPontoNoMapa({ linhas: no.linhas, itens, forcar }, st.cfg);
                if (r.erro) { status.textContent = r.erro; return false; }
                if (r.conflitos) {
                    status.textContent = 'Vínculos que cairiam: ' + r.conflitos.join(' · ') + '.';
                    const confirmar = botao('Trocar mesmo assim', () => { salvarItens.forcar = true; salvarItens(itens, mensagem); salvarItens.forcar = false; });
                    confirmar.classList.add('p');
                    status.append(' ', confirmar);
                    return false;
                }
                depoisDoPonto(r.indice, mensagem + (r.quebrados ? ' · ' + r.quebrados + ' vínculo(s) removido(s).' : '.'));
                return true;
            };
            const atuais = () => idx.map(l => ({ indice: l, ativo: outros()[l].ativo, op: opDaLinha(outros()[l]) }));
            function linhaEmEdicao(item, k) {
                const ativo = el('input', { type: 'text', maxlength: '440', 'data-sugestao-origem': 'OUTROS', 'aria-label': 'Linha do ponto' });
                ativo.value = item.ativo;
                const ops = seletorOperacao(item.op);
                const ok = botao('OK', () => {
                    const itens = atuais();
                    if (k === null) itens.push({ indice: null, ativo: trim(ativo.value), op: ops.value });
                    else itens[k] = { indice: item.indice, ativo: trim(ativo.value), op: ops.value };
                    salvarItens(itens, k === null ? 'Linha acrescentada ao ponto' : 'Linha atualizada');
                }, 'Gravar a linha');
                ok.classList.add('p');
                const cancelar = botao('×', () => preencherNo(), 'Cancelar');
                ativo.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); ok.click(); } if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); preencherNo(); } });
                const linha = el('div', { class: 'vz-lp-painel editando' }, [ativo, ops, ok, cancelar]);
                setTimeout(() => ativo.focus(), 0);
                return linha;
            }
            const lista = el('div', { class: 'vz-lps' });
            atuais().forEach((item, k) => {
                const r = outros()[item.indice], op = item.op, bloqueada = !!r.bloqueada;
                const linha = el('div', { class: 'vz-lp-painel' + (op === 'F' ? ' fic' : '') + (bloqueada ? ' bloqueada' : ''), ...(bloqueada ? {} : { tabindex: '0', role: 'button' }),
                    title: bloqueada ? 'Linha bloqueada' : 'Clique para editar a linha e a operação' }, [
                    el('span', { class: 'vz-op-badge' + (op === 'F' ? ' f' : op.includes('R') ? ' r' : op.includes('M') ? ' m' : ''), text: op }),
                    el('span', { class: 'vz-lp-texto', text: item.ativo }),
                    el('small', { text: op === 'F' ? 'fictícia · fora do orçamento' : ehPosteTexto(item.ativo) ? 'poste' : '' }),
                    bloqueada ? null : el('button', { type: 'button', class: 'vz-lp-x', text: '×', title: 'Tirar esta linha do ponto', onclick: e => {
                        e.stopPropagation();
                        const itens = atuais(); itens[k] = { indice: item.indice, ativo: '', op };
                        salvarItens(itens, 'Linha tirada do ponto');
                    } })
                ]);
                if (!bloqueada) {
                    const editar = () => linha.replaceWith(linhaEmEdicao(item, k));
                    linha.addEventListener('click', editar);
                    linha.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target === linha) { e.preventDefault(); editar(); } });
                }
                lista.append(linha);
            });
            const ultimaOp = idx.length ? opDaLinha(outros()[idx[idx.length - 1]]) : 'I';
            const mais = botao('+ Linha', () => { mais.disabled = true; lista.append(linhaEmEdicao({ indice: null, ativo: '', op: ultimaOp }, null)); }, 'Acrescentar uma linha ao ponto (operação da linha de cima)');
            caixa.append(el('h5', {}, ['Linhas do ponto', mais]), lista);
            return caixa;
        }
        function abrirFicticio(unidade = null) {
            const camposEstrutura = [];
            grafo.nos.forEach(no => estruturasDoNo(no).forEach(ie => {
                if (infoEst(ie).ocupa) camposEstrutura.push([ie, nomeDoNo(no) + ' · ' + codigoDe(ie) + ' · linha ' + (linhaDe(ie) + 1)]);
            }));
            if (!camposEstrutura.length) return;
            const alvo = el('select', {}, opcoes(camposEstrutura));
            if (unidade !== null) alvo.value = String(unidade);
            const codigo = el('input', { type: 'text', required: '', maxlength: '120', placeholder: 'A35P', 'data-sugestao-origem': 'CABOS' });
            abrirEditor('Novo cabo fictício', 'Adicionar à estrutura', [
                campo('Estrutura', alvo), campo('Código do cabo', codigo)
            ], () => {
                const ie = Number(alvo.value);
                ultimoStatus = 'Cabo fictício registrado.';
                const mensagem = addFicticio(ie, codigo.value);
                if (mensagem) { ultimoStatus = ''; return mensagem; }
                avisos.textContent = 'Cabo fictício vinculado a ' + codigoDe(ie) + '.';
                return '';
            });
        }

        function trocarPosteDaEstrutura(linha, poste) {
            const destino = poste === null ? null : grafo.nosPorLinha.get(poste);
            if (destino && !destino.linhas.includes(linha)) {
                const origem = grafo.nosPorLinha.get(linha), pontos = new Set(destino.linhas);
                const repetido = cabos().findIndex((c, ic) => {
                    if (!c) return false;
                    const internas = c.vinculoEstruturas || [];
                    if (internas.some(ie => linhaDe(ie) === linha) && internas.some(ie => pontos.has(linhaDe(ie)))) return true;
                    const ids = pontosDoCabo(ic);
                    return ids.includes(origem.id) && ids.includes(destino.id) &&
                        (origem.linhas.length === 1 || VR.postesAtribuidos(c).includes(linha));
                });
                if (repetido >= 0) return `O cabo ${repetido + 1} ficaria com duas pontas no mesmo ponto. Solte uma ponta antes.`;
            }
            const origem = grafo.nosPorLinha.get(linha);
            const linhas = origem ? [...origem.linhas] : [];
            if (!host.associarNoMapa(linha, poste)) return 'A estrutura ou o poste escolhido mudou.';
            const auto = automatizarPontos(linhas.concat(linha, ...(poste === null ? [] : [poste])));
            grafo = montarGrafo();
            const no = grafo.nosPorLinha.get(linha);
            st.grafoNo = no?.id || null;
            ultimoStatus = comVinculosAutomaticos(poste === null ? 'Estrutura devolvida ao seu ponto.' : 'Estrutura associada ao poste.', auto);
            atualizar();
            ultimoNo = nosSvg.get(st.grafoNo);
            host.depois?.();
            avisos.textContent = ultimoStatus;
            return '';
        }

        function escolherPosteDaEstrutura(linha, no) {
            const candidatos = [...grafo.nos.values()].filter(n => n.poste !== null && !linhaFicticia(n) && n.id !== no.id);
            const destinos = el('select', { required: '' }, opcoes([
                ['', 'Selecione um poste'], ...candidatos.map(n => [n.poste, nomeDoNo(n)])
            ]));
            abrirEditor('Associar estrutura ao poste', 'Associar', [campo('Poste', destinos)],
                () => trocarPosteDaEstrutura(linha, Number(destinos.value)));
        }

        function nomeDoNo(no) {
            const grupo = nomeGrupoDoNo(no);
            if (grupo) {                                                          // TASK-225: tudo é «ponto»; o nome do grupo identifica
                const detalhe = linhaFicticia(no) ? 'Fictício' : no.poste === null ? estruturasDoNo(no).map(codigoDe).join(', ') || outros()[linhaNo(no)].ativo : postesDe(no.poste)[0] || '';
                return 'Ponto ' + grupo + (detalhe ? ' · ' + detalhe : '');
            }
            if (no.poste === null) {
                const codigos = estruturasDoNo(no).map(codigoDe).join(', ');
                return 'Ponto avulso ' + (linhaNo(no) + 1) + ' · ' + (codigos || outros()[linhaNo(no)].ativo);
            }
            if (linhaFicticia(no)) {
                const nome = outros()[linhaNo(no)].ativo;
                const detalhe = nome === 'Poste fictício' ? '' : ' · ' + nome.replace(/^Poste fictício · /, '');
                return 'Poste fictício ' + (linhaNo(no) + 1) + detalhe;
            }
            const identificacao = postesDe(no.poste)[0];
            return 'Poste ' + (no.poste + 1) + (identificacao ? ' · ' + identificacao : '');
        }

        function editarNo(id, alvo) {
            if (st.grafoNo !== id) { soPendenciasNo = false; estruturaEmFoco = null; caboSelNo = null; }
            ultimoNo = alvo;
            ultimoStatus = '';
            st.grafoNo = id;
            if (!dialogo.open) dialogo.showModal();
            preencherNo();
            const card = caboEmFoco === null ? null : dialogo.querySelector(`[data-cabo-no="${caboEmFoco}"]`);
            card?.scrollIntoView({ block: 'nearest' });
            (card?.querySelector('select, button') || dialogo.querySelector('[autofocus]'))?.focus();
        }

        // TASK-225: caixa «Ponto» do painel — nome (ou número automático), comentário, congelar/bloquear o grupo e as linhas que o ponto tem nas tabelas
        function caixaDoPonto(no, status, ehFicticio) {
            const info = host.pontoNoMapa(no.linhas);
            const nome = el('input', { type: 'text', maxlength: '60', placeholder: 'Automático', 'aria-label': 'Nome do ponto' });
            const comentario = el('input', { type: 'text', maxlength: '300', placeholder: 'Opcional', 'aria-label': 'Comentário do ponto' });
            nome.value = info ? info.nome : ''; comentario.value = info ? info.comentario : '';
            const objeto = outros()[linhaNo(no)];
            const depoisDeMudar = mensagem => {                                  // as linhas podem ter mudado de índice: acha o ponto de novo pela linha
                host.depois?.();
                cx = null; metaAtivo.clear();
                grafo = montarGrafo();
                st.grafoNo = grafo.nosPorLinha.get(outros().indexOf(objeto))?.id || null;
                ultimoStatus = mensagem;
                atualizar();
                if (st.grafoNo) ultimoNo = nosSvg.get(st.grafoNo);
            };
            const salvar = host.salvarPontoNoMapa ? null : botao(info ? 'Salvar nome e comentário' : 'Criar o grupo do ponto', () => {
                const erro = host.definirPontoNoMapa(no.linhas, { nome: nome.value, comentario: comentario.value });
                if (erro) { status.textContent = erro; return; }
                depoisDeMudar('Ponto atualizado em Demais Ativos.');
            }, 'Grava o nome e o comentário do ponto (o grupo @ em Demais Ativos)');
            salvar?.classList.add('p');
            const botoes = [salvar].filter(Boolean);
            if (host.copiarPontoNoMapa) botoes.push(botao('Copiar ponto', () => copiarPonto(no, status), 'Copia o ponto (novo grupo com as mesmas linhas, sem vínculos) e pede onde posicioná-lo'));
            if (info && !ehFicticio && host.marcarPontoNoMapa) {
                botoes.push(botao(info.congelada ? 'Descongelar grupo' : 'Congelar grupo', () => { host.marcarPontoNoMapa(no.linhas, 'congelada'); depoisDeMudar(info.congelada ? 'Grupo descongelado.' : 'Grupo congelado: fora dos cálculos.'); },
                    'Congela todas as linhas do ponto (saem da Totalizadora, do orçamento e da validação)'));
                botoes.push(botao(info.bloqueada ? 'Desbloquear grupo' : 'Bloquear grupo', () => { host.marcarPontoNoMapa(no.linhas, 'bloqueada'); depoisDeMudar(info.bloqueada ? 'Grupo desbloqueado.' : 'Grupo bloqueado: não edita nem exclui; o «Rearranjar» não o move.'); },
                    'Bloqueia todas as linhas do ponto contra edição e exclusão'));
            }
            const linhas = [], vistas = new Set();
            const somar = (tabela, i, texto) => { const k = tabela + i; if (!vistas.has(k) && String(texto || '').trim()) { vistas.add(k); linhas.push([tabela, i, texto]); } };
            if (info) { somar('Demais Ativos', info.tituloIdx, outros()[info.tituloIdx].ativo); if (info.comentarioIdx !== null) somar('Demais Ativos', info.comentarioIdx, outros()[info.comentarioIdx].ativo); }
            no.linhas.forEach(l => { if (!outros()[l].comentario) somar('Demais Ativos', l, outros()[l].ativo); });
            [...no.cabos].sort((a, b) => a - b).forEach(ic => somar('Cabos', ic, cabos()[ic] && cabos()[ic].ativo));
            return el('div', { class: 'vz-ponto' }, [
                el('h5', { text: 'Ponto' }),
                ...(salvar ? [campo('Nome (vazio = número automático)', nome), campo('Comentário', comentario)] : []),
                el('div', { class: 'vz-ponto-botoes' }, botoes),
                info ? null : el('small', { class: 'vz-no-ajuda', text: 'Este ponto ainda não é grupo em Demais Ativos.' }),
                el('h6', { text: 'Linhas nas tabelas' }),
                el('ul', { class: 'vz-ponto-linhas' }, linhas.map(([tabela, i, texto]) => el('li', {}, [el('small', { text: tabela + ' · ' + (i + 1) }), el('code', { text: texto, title: texto })])))
            ]);
        }
        const ehLinhaPosteNoPainel = linha => outros()[linha].entidade === 'POSTE' || postesDe(linha).length;
        function preencherNo() {
            observadorNo?.disconnect();
            const no = grafo.nos.get(st.grafoNo);
            if (!no) { if (dialogo.open) dialogo.close(); return; }
            const a = analise();
            const estruturas = estruturasDoNo(no), ehFicticio = linhaFicticia(no);
            // TASK-238: as estruturas em retirada (R, *R) também aparecem na lista, apagadas e fora do vínculo
            const removendo = no.linhas.filter(l => { const r = outros()[l]; return r && trim(r.ativo) && !r.comentario && !r.grupoDecl && !VR.participaVinculo(r); })
                .flatMap(l => { const cods = VR.codigosDaLinha(outros()[l], host.codigoEstrutura, host.codigosEstrutura).filter(Boolean);
                    return (cods.length ? cods : outros()[l].entidade === 'ESTRUTURA' ? [trim(outros()[l].ativo)] : []).map(c => [l, c]); });
            const row = outros()[linhaNo(no)];
            const estruturasFicticias = ehFicticio ? row._mapaEstruturasFicticias || [] : [];
            const completo = VR.posteCompletoNoGrafo(estruturas, a);
            const estruturaPendente = r => r?.ocupa && r.maximo !== 0 && r.estado !== 'gr' && (r.precisa !== null || r.erros.length);
            const estruturasPendentes = estruturas.filter(ie => estruturaPendente(a.estruturas[ie])).length;
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
                const motivo = acao === 'ligar' ? (jaLigadoNo(ic, no)
                    ? 'Este cabo já tem uma ponta vinculada neste poste.' : motivoParaLigar(ic, ie)) : '';
                if (motivo) { status.textContent = motivo; return; }
                ultimoStatus = acao === 'ligar' ? 'Vínculo registrado.' : 'Vínculo removido.';
                caboEscolhido = pontaEscolhida = null;
                ligar(ic, ie, () => {
                    if (!cabos()[ic]._mapaTopologia) gravarPontos(ic, pontosDoCabo(ic));
                });
                if (host.depois) host.depois();
                dialogo.querySelector(`[data-foco="${foco}"]`)?.focus();
            }
            const lateral = el('aside', { class: 'vz-no-lateral' }, [
                el('h5', { text: 'Estruturas' }),
                ehFicticio ? el('div', { class: 'vz-no-estruturas' }, estruturasFicticias.map((codigo, i) =>
                    el('div', { class: 'vz-no-estrutura', 'data-estrutura-ancora': i }, [
                        el('div', { class: 'vz-no-linha' }, [el('strong', { text: codigo }), el('span', { text: 'Fictícia' })])
                    ]))) : estruturas.length || removendo.length ? el('div', { class: 'vz-no-estruturas' }, estruturas.map(ie => {
                    const r = a.estruturas[ie], e = infoEst(ie), codigo = codigoDe(ie);
                    const ficticios = ficticiosDe(ie);
                    const controlesFicticios = e.ocupa ? el('details', { class: 'vz-no-mais' }, [
                        el('summary', { text: ficticios.length ? 'Cabos fictícios (' + ficticios.length + ')' : 'Cabo fictício' }),
                        el('div', { class: 'vz-no-controles' }, [
                            ...ficticios.map((cod, k) => botao(cod + ' ×',
                                () => { ultimoStatus = 'Cabo fictício removido.'; removerFicticio(ie, k); }, 'Remover cabo fictício ' + cod)),
                            botao('+ Fictício', () => abrirFicticio(ie), 'Adicionar cabo fictício a ' + codigo)
                        ])
                    ]) : null;
                    const editar = controleEdicao(ehLinhaPosteNoPainel(linhaDe(ie)) ? 'poste' : 'estrutura', linhaDe(ie), 'estrutura ' + codigo);
                    const op = opDaLinha(outros()[linhaDe(ie)]);
                    return el('div', { class: 'vz-no-estrutura' + (OPS_INSTALANDO.includes(op) ? ' instalando' : ''), 'data-estrutura-no': ie }, [
                        el('div', { class: 'vz-no-linha' }, [
                            seloOperacao(op),
                            botao(codigo, () => destacarEstrutura(ie), 'Destacar cabos de ' + codigo),
                            el('span', { class: 'vz-no-ocupacao' + (estruturaPendente(r) ? ' pendente' : ''),
                                text: !e.ocupa ? 'Equipamento' : e.qtd === null ? String(r.tem) : r.tem + '/' + e.qtd,
                                title: r.erros[0] || r.avisos_vinculo?.[0] || (estruturaPendente(r) ? 'Estrutura a revisar' : 'Ocupação') }),
                            editar
                        ]), controlesFicticios
                    ]);
                }).concat(removendo.map(([l, codigo]) => el('div', { class: 'vz-no-estrutura removendo', 'data-estrutura-removendo': l }, [
                    el('div', { class: 'vz-no-linha' }, [seloOperacao(opDaLinha(outros()[l])), el('strong', { text: codigo }), el('span', { text: 'removendo · fora do vínculo' })])
                ])))) : el('p', { class: 'vz-no-ajuda', text: 'Nenhuma estrutura cadastrada neste ponto.' })
            ]);
            // TASK-235: «Editar ponto» abre o formulário do ponto (nome - comentário + linhas); o poste fictício antigo mantém o editor dele
            const editarPoste = ehFicticio ? (no.poste !== null && host.editarNoMapa && controleEdicao('poste-ficticio', no.poste, 'poste fictício'))
                : host.salvarPontoNoMapa ? botao('Editar ponto', () => abrirPonto(no), 'Nome, comentário e linhas do ponto')
                : (no.poste !== null && host.editarNoMapa && botao('Editar poste', () => abrirCriacao('poste', null, no)));
            if (no.poste === null && !estruturas.length) {
                const editar = controleEdicao('estrutura', linhaNo(no), 'linha do ponto ' + (linhaNo(no) + 1));
                if (editar) lateral.append(editar);
            }
            if (host.associarNoMapa) {
                const existeOutroPoste = [...grafo.nos.values()].some(n => n.poste !== null && !linhaFicticia(n) && n.id !== no.id);
                for (const linha of new Set(estruturas.map(linhaDe))) {
                    if (postesDe(linha).length || outros()[linha].entidade === 'POSTE' || outros()[linha].grupoDe) continue;       // TASK-238: a linha do grupo já é do ponto do grupo
                    const ativo = outros()[linha].ativo;
                    const controles = el('div', { class: 'vz-no-editor' }, [el('small', { text: ativo })]);
                    if (existeOutroPoste) controles.append(botao('Associar a poste', () => escolherPosteDaEstrutura(linha, no), 'Associar ' + ativo + ' a poste'));
                    if (outros()[linha]._mapaPosteRef) controles.append(botao('Soltar do poste', () => {
                        const mensagem = trocarPosteDaEstrutura(linha, null);
                        if (mensagem) status.textContent = mensagem;
                    }, 'Soltar ' + ativo + ' do poste'));
                    if (controles.childElementCount > 1) lateral.append(controles);
                }
            }
            let adicionarEstrutura = null;
            if (ehFicticio) {
                adicionarEstrutura = botao('+ Estrutura fictícia', () => pedirEstruturasFicticias('Adicionar estruturas fictícias', (_, codigos) => {
                    const novas = [...new Set(estruturasFicticias.concat(codigos))];
                    if (novas.length > 20) return 'O poste pode ter até 20 estruturas fictícias.';
                    host.antes();
                    row._mapaEstruturasFicticias = novas;
                    atualizar();
                    return '';
                }));
                if (host.removerPosteFicticioNoMapa) lateral.append(botao('Remover poste fictício', () => {
                    if (no.cabos.size) { status.textContent = 'Solte as pontas dos cabos antes de remover o poste.'; return; }
                    if (host.removerPosteFicticioNoMapa(linhaNo(no))) {
                        dialogo.close();
                        atualizar();
                        host.depois?.();
                    }
                }));
            } else if (host.inserirNoMapa && no.poste !== null && !host.salvarPontoNoMapa) adicionarEstrutura = botao('+ Estrutura', () => abrirCriacao('estrutura', no.poste), 'Nova estrutura neste poste');
            const adicionarCabo = host.inserirNoMapa && botao('+ Cabo', () => abrirCriacao('cabo', linhaNo(no)), 'Novo cabo neste ponto');
            const tracarDaqui = host.inserirNoMapa && botao('Traçar cabo daqui', () => { dialogo.close(); alternarVao(true); marcarOrigemVao(no.id); avisos.textContent = 'Origem: ' + nomeDoNo(no) + '. Clique no ponto de destino.'; }, 'Começa o «Traçar cabo» com este ponto como origem');
            const painelNo = el('section', {}, [el('h5', { text: 'Cabos' })]);
            const adicionais = [...grafo.cabosDoMapa.keys()].filter(ic => !pontosDoCabo(ic).includes(no.id) && livres(ic));
            if (adicionais.length) {
                const selecionar = el('select', { class: 'vz-no-seletor', 'aria-label': 'Cabo com ponta livre' }, [
                    el('option', { value: '', text: 'Escolha um cabo com ponta livre…' }),
                    ...adicionais.map(ic => el('option', { value: ic, text: `Cabo ${ic + 1} · ${cabos()[ic].ativo}` }))
                ]);
                const conectar = botao('Conectar ao ponto', () => {
                    if (selecionar.value === '') return;
                    const mensagem = alterarPontoDoCabo(Number(selecionar.value), no, true);
                    if (mensagem) status.textContent = mensagem;
                });
                conectar.disabled = true;
                selecionar.addEventListener('change', () => { conectar.disabled = selecionar.value === ''; });
                painelNo.append(el('details', { class: 'vz-no-mais vz-no-conectar' }, [
                    el('summary', { text: 'Conectar cabo existente' }), selecionar, conectar
                ]));
            }
            const cabosNo = no.cabos;
            if (!cabosNo.size) painelNo.appendChild(el('p', { class: 'vz-no-ajuda', text: 'Nenhum cabo chega a este ponto.' }));
            const conexoes = [];
            const linksComErro = new Map(a.links.filter(l => !l.ok && no.linhas.includes(linhaDe(l.estrutura))).map(l => [l.cabo, l.motivo]));
            const candidatasPorCabo = new Map(), demandaPorEstrutura = new Map();
            if (!ehFicticio) for (const ic of cabosNo) {
                const candidatas = estruturasCompativeis(ic, no);
                candidatasPorCabo.set(ic, candidatas);
                for (const ie of candidatas) demandaPorEstrutura.set(ie, (demandaPorEstrutura.get(ie) || 0) + 1);
            }
            function motivoDaRevisao(ic, candidatas) {
                if (candidatas.length > 1) return candidatas.length + ' estruturas compatíveis.';
                if (!VR.postesAtribuidos(cabos()[ic]).some(l => no.linhas.includes(l))) return 'Chegada detectada pelo CAD. Confirme a estrutura.';
                const ie = candidatas[0], e = infoEst(ie);
                if (e.regra === null || e.qtd === null) return 'A estrutura não tem uma regra de capacidade para vínculo automático.';
                if (ocupantes(ie).length + demandaPorEstrutura.get(ie) > e.qtd) return 'Mais de um cabo disputa as vagas desta estrutura.';
                return 'Uma opção compatível. Confira a vaga e as pontas.';
            }
            let automaticos = 0, pendentes = 0, vinculados = 0;
            cabosNo.forEach(ic => {
                const c = cabos()[ic], r = a.cabos[ic], ancora = ehFicticio ? estruturaDaAncora(ic, no) : null;
                const existentes = (c.vinculoEstruturas || []).filter(ie => no.linhas.includes(linhaDe(ie)));
                const pontos = pontosDoCabo(ic), chega = pontos.includes(no.id);
                const outroNo = chega && pontos.find(id => id !== no.id);
                const origem = outroNo ? nomeDoNo(grafo.nos.get(outroNo)) : (chega ? 'Outra ponta livre' : 'Sem conexão no mapa neste ponto');
                const temVinculo = ehFicticio ? !!ancora : existentes.length > 0;
                const completoSemEsta = !ehFicticio && !temVinculo && chega && r?.vincula && r.tem >= r.precisa;          // TASK-234: ponta extra (cabo M) — opcional, só à mão
                const revisaoPendente = ehFicticio ? chega && !ancora : linksComErro.has(ic) ||
                    (!temVinculo && chega && r?.vincula && !c.congelada && !completoSemEsta);
                const automatico = temVinculo && !revisaoPendente && c.vinculoOrigem === 'mapa' && !ehFicticio;
                if (revisaoPendente) pendentes++;
                else if (automatico) automaticos++;
                else if (temVinculo) vinculados++;

                const nivel = VR.nivelDoCabo(c, st.cfg.dicionario_cabos).nivel;
                const titulo = el('h6', { text: 'Cabo ' + (ic + 1), title: c.ativo });
                const detalhes = el('details', { class: 'vz-no-mais' }, [
                    el('summary', { text: revisaoPendente ? 'Revisar' : 'Opções', 'data-foco': 'c' + ic }),
                    el('div', { class: 'vz-no-acoes' }, [
                        controleEdicao('cabo', ic),
                        chega ? botao('Soltar ponta do mapa', () => {
                            const mensagem = alterarPontoDoCabo(ic, no, false);
                            if (mensagem) status.textContent = mensagem;
                        }, 'Soltar ponta do cabo ' + (ic + 1) + ' do mapa') : null
                    ])
                ]);
                const acoes = detalhes.lastElementChild;
                const cartao = el('section', {
                    class: 'vz-no-cabo' + (ic === caboEmFoco ? ' vz-no-card-foco' : '') + (revisaoPendente ? ' pendente' : temVinculo ? ' vinculado' : ''),
                    'data-cabo-no': ic, 'data-pendencia': String(!!revisaoPendente)
                }, [
                    el('div', { class: 'vz-no-linha' }, [
                        titulo, el('span', { class: 'vz-nivel ' + nivel.toLowerCase(), text: nivel }),
                        el('span', { class: 'vz-op' + opClasse(trim(c.operacao)), text: c.operacao || 'I' }),
                        el('span', { class: 'vz-no-estado', text: revisaoPendente ? 'Pendente' : automatico ? 'Automático' : temVinculo ? 'Vinculado' : completoSemEsta ? 'Opcional' : 'Fora da revisão',
                            ...(completoSemEsta ? { title: 'O cabo já está completo pelas pontas automáticas. Esta ponta é opcional: vincule à mão se precisar.' } : {}) })
                    ]),
                    el('p', { class: 'vz-no-rotulo', text: c.ativo, title: c.ativo }),
                    detalhes
                ]);
                const dica = el('p', { class: 'vz-no-ajuda', text: origem });
                detalhes.insertBefore(dica, acoes);
                if (ehFicticio) {
                    const indice = estruturasFicticias.indexOf(ancora);
                    if (indice >= 0) {
                        conexoes.push({ ic, ancora: indice });
                        acoes.append(botao(ancora + ' ×', () => gravarAncora(ic, no), 'Remover âncora ' + ancora));
                    } else if (chega && estruturasFicticias.length) {
                        const seletor = el('select', { 'aria-label': 'Estrutura fictícia para o cabo ' + (ic + 1) }, opcoes(estruturasFicticias.map(c => [c, c])));
                        acoes.append(seletor, botao('Ancorar', () => gravarAncora(ic, no, seletor.value)));
                    }
                } else {
                    existentes.forEach(ie => {
                        conexoes.push({ ic, ie });
                        acoes.append(botao(codigoDe(ie) + ' ×', () => mudarVinculo(ic, ie, 'remover', 'c' + ic), 'Remover vínculo com ' + codigoDe(ie)));
                    });
                    const candidatas = candidatasPorCabo.get(ic);
                    if (candidatas.length) {
                        const seletor = el('select', { 'aria-label': 'Estrutura para ' + c.ativo }, candidatas.map(ie =>
                            el('option', { value: ie, text: codigoDe(ie) + ' · linha ' + (linhaDe(ie) + 1) })));
                        const vincular = botao('Vincular', () => mudarVinculo(ic, Number(seletor.value), 'ligar', 'c' + ic));
                        vincular.classList.add('p');
                        if (revisaoPendente) {
                            acoes.append(el('span', { class: 'vz-no-ajuda', text: motivoDaRevisao(ic, candidatas) }));
                        }
                        acoes.append(seletor, vincular);
                    } else if (revisaoPendente) {
                        acoes.append(el('span', { class: 'vz-no-ajuda', text: linksComErro.get(ic) || motivoNo(ic, no) || 'Nenhuma estrutura compatível disponível.' }));
                    }
                }
                if (!ehFicticio) {                                                               // TASK-235: clique no cabo + clique na estrutura = vínculo (como na aba Vínculo)
                    const cab = cartao.firstElementChild;
                    cab.classList.add('vz-no-cabo-sel');
                    cab.title = 'Clique para escolher este cabo; depois clique na estrutura';
                    cab.addEventListener('click', e => { if (!e.target.closest('button, select, input, summary')) escolherCaboNo(ic); });
                    if (caboSelNo === ic) cartao.classList.add('escolhido');
                }
                painelNo.append(cartao);
            });
            function escolherCaboNo(ic) {
                if (estruturaEmFoco !== null) { const ie = estruturaEmFoco; estruturaEmFoco = null; caboSelNo = null; mudarVinculo(ic, ie, 'ligar', 'c' + ic); return; }
                caboSelNo = caboSelNo === ic ? null : ic;
                painelNo.querySelectorAll('[data-cabo-no]').forEach(c => c.classList.toggle('escolhido', Number(c.dataset.caboNo) === caboSelNo));
                textoDica();
            }
            const botoes = el('div', { class: 'vz-no-topo-acoes' }, [editarPoste, adicionarEstrutura, adicionarCabo, tracarDaqui, fecharNo].filter(Boolean));
            const filtro = el('input', { type: 'checkbox' });
            filtro.checked = soPendenciasNo;
            const filtros = el('div', { class: 'vz-no-filtros' }, [
                el('label', {}, [filtro, 'Só pendências']), desfazer, consultarNo || null
            ]);
            const semPendencias = el('p', { class: 'vz-no-ajuda', text: 'Nenhum cabo pendente neste ponto.' });
            painelNo.append(semPendencias);
            const ligacoesSvg = doc.createElementNS(ns, 'svg');
            ligacoesSvg.setAttribute('class', 'vz-no-linhas');
            ligacoesSvg.setAttribute('aria-hidden', 'true');
            const mundo = el('div', { class: 'vz-no-rede' }, [painelNo, lateral, ligacoesSvg]);
            const detalhesPonto = host.pontoNoMapa && host.definirPontoNoMapa && el('details', { class: 'vz-no-detalhes' }, [
                el('summary', { text: 'Detalhes do ponto' }),
                el('p', { text: no.pos ? `X ${no.pos[0].toFixed(2)} · Y ${no.pos[1].toFixed(2)}` : 'Sem coordenadas CAD' }),
                caixaDoPonto(no, status, ehFicticio)
            ]);
            const infoPonto = host.pontoNoMapa?.(no.linhas);
            const tipoPonto = ehFicticio || pontoSoFicticio(no) ? 'Fictício' : no.poste !== null ? 'Poste' : 'Estrutura avulsa';
            const tituloNo = infoPonto ? [infoPonto.nome, infoPonto.comentario ? el('small', { text: ' - ' + infoPonto.comentario }) : null, el('span', { class: 'vz-tipo-ponto ' + (tipoPonto === 'Poste' ? 'poste' : tipoPonto === 'Fictício' ? 'fic' : 'avulso'), text: tipoPonto })]
                : [nomeDoNo(no)];
            const antigo = !ehFicticio && host.salvarPontoNoMapa && (!infoPonto || infoPonto.formatoAntigo) && !(st.manterAntigos || new Set()).has(no.id);
            const avisoPonto = antigo ? el('div', { class: 'vz-aviso-antigo' }, [
                el('span', { text: 'Este ponto está no formato antigo. Converter para «@nome - comentário», com a operação em cada linha?' }),
                botao('Converter este ponto', () => {
                    const r = host.salvarPontoNoMapa({ linhas: no.linhas, itens: linhasDoPonto(no).idx.map(l => ({ indice: l, ativo: outros()[l].ativo, op: opDaLinha(outros()[l]) })) }, st.cfg);
                    if (r.erro) { status.textContent = r.erro; return; }
                    depoisDoPonto(r.indice, 'Ponto convertido.');
                }),
                botao('Manter como está', () => { st.manterAntigos = st.manterAntigos || new Set(); st.manterAntigos.add(no.id); preencherNo(); })
            ]) : null;
            if (avisoPonto) avisoPonto.querySelector('button').classList.add('p');
            const linhasPonto = !ehFicticio && host.salvarPontoNoMapa ? secaoLinhasDoPonto(no, status) : null;
            const dicaVinculo = !ehFicticio && no.cabos.size ? el('p', { class: 'vz-no-ajuda vz-dica-vinculo' }) : null;
            const textoDica = () => { if (dicaVinculo) dicaVinculo.textContent = caboSelNo !== null ? 'Cabo ' + (caboSelNo + 1) + ' escolhido: clique na estrutura para vincular (Esc ou outro clique no cabo cancela).'
                : estruturaEmFoco !== null ? codigoDe(estruturaEmFoco) + ' escolhida: clique num cabo para vincular.' : 'Para vincular: clique no cabo e depois na estrutura (ou na estrutura e depois no cabo).'; };
            dialogo.replaceChildren(el('header', { class: 'vz-no-topo' }, [
                el('div', {}, [el('h4', { id: 'vz-no-titulo' }, tituloNo),
                    el('p', { class: completo && !pendentes ? 'vz-no-completo' : '',
                        text: [automaticos && automaticos + ' automáticos', vinculados && vinculados + ' vinculados',
                            pendentes ? pendentes + ' a revisar' : 'Nenhuma ponta a revisar',
                            estruturasPendentes && estruturasPendentes + ' estrutura(s) a revisar'].filter(Boolean).join(' · ') })
                ]), botoes
            ]), el('div', { class: 'vz-no-conteudo' }, [avisoPonto, filtros, status, linhasPonto, dicaVinculo, mundo, detalhesPonto]));
            function filtrarPendencias() {
                soPendenciasNo = filtro.checked;
                painelNo.querySelectorAll('[data-cabo-no]').forEach(card => { card.hidden = soPendenciasNo && card.dataset.pendencia !== 'true'; });
                semPendencias.hidden = !soPendenciasNo || pendentes > 0;
                desenharLigacoes();
            }
            function destacarEstrutura(ie) {
                if (caboSelNo !== null) { const ic = caboSelNo; caboSelNo = null; mudarVinculo(ic, ie, 'ligar', 'c' + ic); return; }
                estruturaEmFoco = estruturaEmFoco === ie ? null : ie;
                atualizarDestaque();
                textoDica();
            }
            function atualizarDestaque() {
                const ligados = new Set(conexoes.filter(l => l.ie === estruturaEmFoco).map(l => l.ic));
                mundo.querySelectorAll('[data-cabo-no]').forEach(el => el.classList.toggle('destaque', ligados.has(Number(el.dataset.caboNo))));
                mundo.querySelectorAll('[data-estrutura-no]').forEach(el => {
                    const ativo = Number(el.dataset.estruturaNo) === estruturaEmFoco;
                    el.classList.toggle('destaque', ativo);
                    el.querySelector('.vz-no-linha > button')?.setAttribute('aria-pressed', String(ativo));
                });
                desenharLigacoes();
            }
            function desenharLigacoes() {
                const base = mundo.getBoundingClientRect();
                const elementos = [];
                for (const {ic, ie, ancora} of conexoes) {
                    const cabo = mundo.querySelector(`[data-cabo-no="${ic}"]`);
                    const destino = mundo.querySelector(ancora === undefined ? `[data-estrutura-no="${ie}"] .vz-no-linha` : `[data-estrutura-ancora="${ancora}"] .vz-no-linha`);
                    if (!cabo || cabo.hidden || !destino) continue;
                    const inicio = cabo.querySelector('.vz-no-linha').getBoundingClientRect(), fim = destino.getBoundingClientRect();
                    const x0 = cabo.getBoundingClientRect().right - base.left, x1 = destino.parentElement.getBoundingClientRect().left - base.left;
                    const y0 = inicio.top + inicio.height / 2 - base.top, y1 = fim.top + fim.height / 2 - base.top;
                    const meio = (x0 + x1) / 2;
                    const linha = doc.createElementNS(ns, 'path');
                    linha.setAttribute('d', `M${x0},${y0} C${meio},${y0} ${meio},${y1} ${x1},${y1}`);
                    const grupo = VR.gruposVisuaisDoVao([ic], cabos(), a, st.cfg.dicionario_cabos)[0];
                    linha.setAttribute('stroke', grupo.tipo === 'MT' ? '#ff6478' : grupo.tipo === 'BT' ? '#54e1a3' : '#bda6ff');
                    linha.setAttribute('stroke-width', estruturaEmFoco === ie && ie !== undefined ? '3' : '2');
                    linha.setAttribute('fill', 'none');
                    if (ancora !== undefined) linha.setAttribute('stroke-dasharray', '4 3');
                    elementos.push(linha);
                }
                ligacoesSvg.replaceChildren(...elementos);
            }
            filtro.addEventListener('change', filtrarPendencias);
            textoDica();
            filtrarPendencias();
            atualizarDestaque();
            observadorNo = new ResizeObserver(desenharLigacoes);
            observadorNo.observe(mundo);
            observadorNo.observe(painelNo);
            observadorNo.observe(lateral);
        }

        dialogo.addEventListener('close', () => {
            observadorNo?.disconnect();
            st.grafoNo = null;
            nosSvg.get(ultimoNo?.dataset.no)?.focus();
        });
        function coordenada(e) {
            const p = svg.createSVGPoint();
            p.x = e.clientX; p.y = e.clientY;
            return p.matrixTransform(svg.getScreenCTM().inverse());
        }
        function pontoDaVista(e) {                                         // posição do mouse em unidades do mapa, pela vista em memória (xMidYMid meet)
            const r = svg.getBoundingClientRect(), s = Math.min(r.width / vb[2], r.height / vb[3]) || 1;
            return { x: vb[0] + (e.clientX - r.left - (r.width - vb[2] * s) / 2) / s, y: vb[1] + (e.clientY - r.top - (r.height - vb[3] * s) / 2) / s };
        }
        svg.addEventListener('wheel', e => { e.preventDefault(); const p = pontoDaVista(e); zoom(e.deltaY > 0 ? 1.15 : 1 / 1.15, p.x, p.y); }, { passive: false });
        const limitar = (n, min, max) => Math.max(min, Math.min(max, n));
        svg.addEventListener('pointerdown', e => {
            if (e.button !== 0 || !e.isPrimary) return;
            mover = false;
            const alvo = e.target.closest('[data-arrastar]');
            const base = { id: e.pointerId, x: e.clientX, y: e.clientY, vb, matriz: svg.getScreenCTM().inverse(), valor: null };
            if (alvo?.dataset.arrastar === 'ponta') {                                   // TASK-233: alça da ponta do cabo → solta sobre um ponto para ligar
                arrasto = Object.assign(base, { ponta: { ic: Number(alvo.dataset.ic), i: Number(alvo.dataset.i), el: alvo, alvo: null, elastico: null } });
                return;
            }
            if (!alvo && e.shiftKey) { arrasto = Object.assign(base, { caixa: { inicio: pontoDaVista(e), el: null } }); return; }     // Shift+arrastar no fundo: seleção por caixa
            const idNo = alvo?.dataset.no;
            const ic = alvo?.dataset.caboMapa !== undefined ? Number(alvo.dataset.caboMapa) : undefined;
            const linha = idNo ? outros()[linhaNo(grafo.nos.get(idNo))] : ic !== undefined ? cabos()[ic] : null;
            const campo = idNo ? campoPonto() : campoCabo();
            const origem = idNo ? posicoes.get(idNo) : linha?.[campo] || [0, 0];
            const grupo = idNo && selecao.has(idNo) && selecao.size > 1 ? [...selecao].filter(id => posicoes.has(id)).map(id => [id, posicoes.get(id).slice()]) : null;
            arrasto = Object.assign(base, { idNo, ic, linha, campo, origem, grupo });
        });
        // ponto mais perto do mouse para o ímã (raio em pixels da tela convertidos para o mapa)
        function noSobMouse(pt, exceto) {
            const raio = 22 * Math.abs(arrasto.matriz.a || 1);
            let melhor = null, dmin = raio;
            nosSvg.forEach((g, id) => { if (id === exceto || !g.isConnected) return; const p = posicoes.get(id), d = Math.hypot(p[0] - pt.x, p[1] - pt.y); if (d <= dmin) { dmin = d; melhor = id; } });
            return melhor;
        }
        svg.addEventListener('pointermove', e => {
            if (!arrasto || e.pointerId !== arrasto.id) return;
            if (!(e.buttons & 1)) { soltar(e, true); return; }
            const dx = e.clientX - arrasto.x, dy = e.clientY - arrasto.y;
            if (!mover && Math.hypot(dx, dy) < 4) return;
            if (!mover) {
                svg.setPointerCapture(e.pointerId);
                svg.classList.add('arrastando');
                mover = true;
            }
            const x = dx * arrasto.matriz.a, y = dy * arrasto.matriz.d;
            if (arrasto.ponta) {
                const pt = pontoDaVista(e), pa = arrasto.ponta, ids = idsDoCabo(pa.ic);
                const alvoId = noSobMouse(pt, ids[1 - pa.i]);
                if (alvoId !== pa.alvo) { if (pa.alvo) nosSvg.get(pa.alvo)?.classList.remove('iman'); pa.alvo = alvoId; if (alvoId) nosSvg.get(alvoId)?.classList.add('iman'); }
                const fim = alvoId ? posicoes.get(alvoId) : [pt.x, pt.y];
                pa.el.setAttribute('cx', fim[0].toFixed(1)); pa.el.setAttribute('cy', fim[1].toFixed(1));
                const d = (desenhosPorCabo.get(pa.ic) || []).find(q => q.caminho), outra = d ? d.caminho[pa.i === 0 ? d.caminho.length - 1 : 0] : fim;
                if (!pa.elastico) { pa.elastico = doc.createElementNS(ns, 'line'); pa.elastico.setAttribute('class', 'vz-elastico'); camadaAlcas.appendChild(pa.elastico); }
                pa.elastico.setAttribute('x1', outra[0]); pa.elastico.setAttribute('y1', outra[1]); pa.elastico.setAttribute('x2', fim[0]); pa.elastico.setAttribute('y2', fim[1]);
                return;
            }
            if (arrasto.caixa) {
                const c = arrasto.caixa, pt = pontoDaVista(e);
                if (!c.el) { c.el = doc.createElementNS(ns, 'rect'); c.el.setAttribute('class', 'vz-retangulo'); camadaAlcas.appendChild(c.el); }
                c.fim = [pt.x, pt.y];
                c.el.setAttribute('x', Math.min(c.inicio.x, pt.x)); c.el.setAttribute('y', Math.min(c.inicio.y, pt.y));
                c.el.setAttribute('width', Math.abs(pt.x - c.inicio.x)); c.el.setAttribute('height', Math.abs(pt.y - c.inicio.y));
                return;
            }
            if (arrasto.grupo) {
                for (const [id, [ox, oy]] of arrasto.grupo) posicoes.set(id, [limitar(ox + x, limites[0], limites[2]), limitar(oy + y, limites[1], limites[3])]);
                arrasto.valor = true;
                if (!quadro) quadro = raiz.requestAnimationFrame(() => { quadro = null; if (arrasto?.grupo) desenharPosicoes({ ids: arrasto.grupo.map(([id]) => id) }); });
                return;
            }
            if (!arrasto.linha) {
                vb = [arrasto.vb[0] - x, arrasto.vb[1] - y, arrasto.vb[2], arrasto.vb[3]];
                aplicarVista();
            } else {
                const [ox, oy] = arrasto.origem;
                if (arrasto.idNo) {
                    const px = limitar(ox + x, limites[0], limites[2]), py = limitar(oy + y, limites[1], limites[3]);
                    const base = bases.get(arrasto.idNo) || [0, 0];
                    arrasto.valor = emEscalaAgora() ? [(px - base[0]) / kLayout, (py - base[1]) / kLayout] : [px / W, py / H];            // 1:500 guarda o deslocamento em metros; Ajustar, a fração da tela
                    posicoes.set(arrasto.idNo, [px, py]);
                } else {
                    arrasto.valor = emEscalaAgora() ? [limitar(ox + x / kLayout, -3000, 3000), limitar(oy + y / kLayout, -3000, 3000)] : [limitar(ox + x / W, -1, 1), limitar(oy + y / H, -1, 1)];
                }
                if (!quadro) quadro = raiz.requestAnimationFrame(() => { quadro = null; desenharPosicoes(arrasto); });
            }
        });
        function soltar(e, cancelar) {
            if (!arrasto || e.pointerId !== arrasto.id) return;
            const alteracao = arrasto;
            const {linha, campo, valor, idNo, origem} = alteracao;
            arrasto = null;
            if (quadro) {
                raiz.cancelAnimationFrame(quadro);
                quadro = null;
            }
            svg.classList.remove('arrastando');
            if (alteracao.ponta) {
                const pa = alteracao.ponta;
                pa.elastico?.remove();
                if (pa.alvo) nosSvg.get(pa.alvo)?.classList.remove('iman');
                if (!mover) { desenharAlcas(); return; }
                if (!cancelar && pa.alvo) { const msg = religarPonta(pa.ic, pa.i, grafo.nos.get(pa.alvo)); if (msg) avisos.textContent = msg; }
                else if (!cancelar) avisos.textContent = 'Solte a ponta sobre um ponto para ligar o cabo.';
                desenharAlcas();
                return;
            }
            if (alteracao.caixa) {
                const c = alteracao.caixa;
                c.el?.remove();
                if (!cancelar && c.fim) {
                    const visiveis = new Map([...posicoes].filter(([id]) => nosSvg.get(id)?.isConnected));
                    MM.dentroDoRetangulo(visiveis, [c.inicio.x, c.inicio.y], c.fim).forEach(id => selecao.add(id));
                    pintarSelecao();
                    avisoSelecao();
                }
                return;
            }
            if (alteracao.grupo) {
                if (valor && !cancelar) { host.antes(); alteracao.grupo.forEach(([id]) => guardarPosicao(id)); }
                else alteracao.grupo.forEach(([id, p]) => posicoes.set(id, p));
                desenharPosicoes({ ids: alteracao.grupo.map(([id]) => id) });
                recortarVista(true);
                return;
            }
            if (valor && !cancelar) {
                host.antes();
                linha[campo] = valor;
            } else if (cancelar && idNo) posicoes.set(idNo, origem);
            if (valor) desenharPosicoes(alteracao);
            recortarVista(true);
        }
        svg.addEventListener('pointerup', e => soltar(e, false));
        svg.addEventListener('pointercancel', e => soltar(e, true));
        svg.addEventListener('lostpointercapture', e => soltar(e, true));
        function consultar(selecao) {
            if (host.consultarAssistente && fechar()) host.consultarAssistente(selecao);
        }
        function consultarCabo(ic) {
            const nos = pontosDoCabo(ic).map(id => grafo.nos.get(id));
            consultar({ tipo: 'vão', rotulo: cabos()[ic].ativo, cabos: [ic], outros: nos.flatMap(n => n.linhas) });
        }
        function ativarMapa(alvo, e) {
            const bloco = alvo.closest('[data-bloco]');
            if (bloco) { transformarBloco(bloco.dataset.bloco); return; }
            const no = alvo.closest('[data-no]');
            if (ferramenta === 'vao') { if (no) cliqueVao(no.dataset.no, !!e?.shiftKey); return; }
            if (no && e?.shiftKey) {                                                         // Shift+clique: entra/sai da seleção
                const id = no.dataset.no;
                if (selecao.has(id)) selecao.delete(id); else selecao.add(id);
                pintarSelecao(); avisoSelecao();
                return;
            }
            const cabo = alvo.closest('[data-cabo-mapa]');
            if (cabo) { destacarCabo(Number(cabo.dataset.caboMapa), true); return; }
            if (!no) { if (!alvo.closest('[data-arrastar]')) limparSelecao(); return; }
            const id = no.dataset.no;
            if (caboEscolhido === null) editarNo(id, no);
            else {
                const mensagem = alterarPontoDoCabo(caboEscolhido, grafo.nos.get(id), true);
                if (mensagem) avisos.textContent = mensagem;
                nosSvg.get(id)?.focus();
            }
        }
        svg.addEventListener('pointermove', e => {                         // ponto em posicionamento acompanha o mouse
            if (!colocando || arrasto) return;
            const pt = pontoDaVista(e);
            alvoPosicionar = [limitar(pt.x, limites[0], limites[2]), limitar(pt.y, limites[1], limites[3])];
            if (!quadroPosicionar) quadroPosicionar = raiz.requestAnimationFrame(() => {
                quadroPosicionar = 0;
                if (!colocando || !alvoPosicionar) return;
                posicoes.set(colocando.id, alvoPosicionar);
                desenharPosicoes({ idNo: colocando.id });
            });
        });
        svg.addEventListener('click', e => {                               // o clique fixa o ponto (um clique depois de arrastar o fundo não conta)
            if (!colocando || mover) return;
            e.stopImmediatePropagation(); e.preventDefault();
            const pt = pontoDaVista(e);
            posicoes.set(colocando.id, [limitar(pt.x, limites[0], limites[2]), limitar(pt.y, limites[1], limites[3])]);
            guardarPosicao(colocando.id);
            desenharPosicoes({ idNo: colocando.id });
            encerrarPosicionamento(true);
        }, true);
        svg.addEventListener('click', e => { if (!mover || e.detail === 0) ativarMapa(e.target, e); });
        svg.addEventListener('keydown', e => {
            if (e.target.closest('[data-no], [data-cabo-mapa]') && (e.key === 'Enter' || e.key === ' ')) {
                e.preventDefault();
                ativarMapa(e.target);
            }
        });

        mapaDesmontar = () => {
            fecharEdicaoAtiva?.(false);
            observadorNo?.disconnect();
            if (quadro) raiz.cancelAnimationFrame(quadro);
            if (quadroVista) { raiz.cancelAnimationFrame(quadroVista); quadroVista = 0; }
            raiz.clearTimeout(fimMovimento);
            encerrarPosicionamento(true, true);
            doc.removeEventListener('keydown', teclaMapa, true);
            cartoes?.observador?.disconnect();
        };
        mapaAtualizar = () => {
            sugestoes = [];
            pintar();
            if (dialogo.open) preencherNo();
        };
        pintar();
        if (!vb) vb = vistaPadrao();
        aplicarVistaJa();
        if (escalar() && !svg.clientWidth) {                              // janela ainda sem tamanho (aba recém-aberta): recentra quando medir
            let tentativas = 0;
            const medir = () => { if (svg.clientWidth) { vb = vistaPadrao(); aplicarVistaJa(); } else if (++tentativas < 20) raiz.requestAnimationFrame(medir); };
            raiz.requestAnimationFrame(medir);
        }
        if (st.grafoNo && grafo.nos.has(st.grafoNo)) editarNo(st.grafoNo, [...svg.querySelectorAll('[data-no]')].find(n => n.dataset.no === st.grafoNo));
    }

    raiz.VinculacaoUI = { abrir, fechar, atualizarObra };
})(typeof window !== 'undefined' ? window : globalThis);
