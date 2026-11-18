/*
 * Teste de desempenho do Resumo e do modal de vinculação com tabelas grandes (TASK-185).
 * Uso: node scripts/perf_resumo.js <url-base> <token> [linhas=450]
 * Cola N linhas em Cabos e N em Outros, mede as ações e sai com código 1 se algum limite (ms) for ultrapassado.
 * Os limites são folgados (navegador sem GPU): servem para pegar regressão grande, não para medir o último milissegundo.
 */
let chromium; try { ({ chromium } = require('playwright')); } catch (e) { ({ chromium } = require('/opt/node-tools/node_modules/playwright')); }
const [, , URL_BASE, TOKEN, NN] = process.argv, N = parseInt(NN || '450');
const LIMITES = { 'modal: selecionar cabo': 60, 'modal: vincular': 150, 'modal: rolar (média/quadro)': 30, 'resumo: Enter': 150, 'resumo: excluir linha': 150, 'resumo: duplicar linha': 150, 'resumo: desfazer': 200, 'resumo: rolar (média/quadro)': 30 };
(async () => {
  if (!URL_BASE || !TOKEN) { console.error('uso: node scripts/perf_resumo.js <url> <token> [linhas]'); process.exit(2); }
  const b = await chromium.launch({ args: ['--no-sandbox'] }), ctx = await b.newContext({ viewport: { width: 1700, height: 1000 } });
  await ctx.addInitScript(t => { localStorage.setItem('auth_token', t); localStorage.setItem('projeto_selecionado', 'PERNAMBUCO'); localStorage.setItem('projeto_selecionado_codigo', '081'); }, TOKEN);
  const p = await ctx.newPage(); await p.goto(URL_BASE + '/resumo'); await p.waitForTimeout(5000);
  const postes = ['11300D', '9200D', '11600D', '12400D', '9400D'], est = ['N1', 'N4', 'N1, N4', 'N3', 'N2'];
  const cab = Array.from({ length: N }, (_, i) => 'I\t' + ['A35P', 'S04P', 'C16', 'A70P'][i % 4] + ' ' + (10 + i % 30) + ' m').join('\n');
  const out = Array.from({ length: N }, (_, i) => 'I\t' + postes[i % 5] + ', ' + est[(i >> 1) % 5]).join('\n');
  for (const [sel, t] of [['#body-cabos-vazio .inp-ativo-vazio', cab], ['#body-outros-vazio .inp-ativo-vazio', out]])
    await p.locator(sel).first().evaluate((el, tx) => { const dt = new DataTransfer(); dt.setData('text', tx); el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); }, t);
  await p.waitForTimeout(3000);
  const R = {};
  Object.assign(R, await p.evaluate(async () => {
    const res = {}, raf = () => new Promise(r => requestAnimationFrame(() => r()));
    const medir = async (nome, fn) => { await raf(); const t0 = performance.now(); fn(); void document.body.offsetHeight; res[nome] = Math.round(performance.now() - t0); await raf(); await new Promise(r => setTimeout(r, 300)); };
    const num = i => document.querySelector(`#body-outros tr[data-index="${i}"] td.col-num`);
    await medir('resumo: Enter', () => { const e = document.querySelector('#body-outros tr[data-index="12"] .inp-ativo'); e.focus(); e.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); });
    await medir('resumo: excluir linha', () => { num(8).dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 100, clientY: 100 })); document.getElementById('ctx-del').click(); });
    await medir('resumo: duplicar linha', () => { num(10).dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 100, clientY: 100 })); document.getElementById('ctx-duplicar').click(); });
    await medir('resumo: desfazer', () => document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true })));
    const w = document.querySelector('#table-outros').closest('.table-wrap'); w.scrollTop = 0; await raf();
    const q = []; let u = performance.now(); for (let i = 0; i < 30; i++) { w.scrollTop += 60; await raf(); const n = performance.now(); q.push(n - u); u = n; }
    res['resumo: rolar (média/quadro)'] = Math.round(q.reduce((a, c) => a + c, 0) / q.length);
    return res;
  }));
  await p.getByText('VINCULAR CABOS', { exact: false }).first().click(); await p.waitForSelector('.vz'); await p.waitForTimeout(1200);
  Object.assign(R, await p.evaluate(async () => {
    const res = {}, raf = () => new Promise(r => requestAnimationFrame(() => r()));
    const medir = async (nome, fn) => { await raf(); const t0 = performance.now(); fn(); void document.body.offsetHeight; res[nome] = Math.round(performance.now() - t0); await raf(); await new Promise(r => setTimeout(r, 300)); };
    await medir('modal: selecionar cabo', () => document.querySelectorAll('.vl-cab[data-c]')[3].click());
    await medir('modal: vincular', () => document.querySelectorAll('.vl-chip[data-e]')[6].click());
    const sc = document.querySelector('.vl-box'), q = []; let u = performance.now();
    for (let i = 0; i < 25; i++) { sc.scrollTop += 120; await raf(); const n = performance.now(); q.push(n - u); u = n; }
    res['modal: rolar (média/quadro)'] = Math.round(q.reduce((a, c) => a + c, 0) / q.length);
    return res;
  }));
  await b.close();
  let falhou = 0;
  Object.entries(LIMITES).forEach(([k, lim]) => { const v = R[k], ruim = !(v <= lim); if (ruim) falhou++; console.log((ruim ? 'FALHOU ' : 'ok     ') + k.padEnd(34) + String(v).padStart(5) + ' ms (limite ' + lim + ')'); });
  process.exit(falhou ? 1 : 0);
})();
