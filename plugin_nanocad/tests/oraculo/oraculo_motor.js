// tests/oraculo/oraculo_motor.js — roda o MOTOR JS DA APLICAÇÃO (static/regras_leitor_engine.js) sobre os casos de
// casos_motor.json e imprime o resultado. É o oráculo da paridade do porte em C# (src/Nucleo/MotorLeitor.cs).
// Uso: node oraculo_motor.js [casos.json] > esperado_motor.json
const fs = require('fs');
const path = require('path');
const raiz = path.resolve(__dirname, '..', '..', '..');
const motor = require(path.join(raiz, 'static', 'regras_leitor_engine.js'));
const ler = f => JSON.parse(fs.readFileSync(path.join(raiz, 'data', f), 'utf8'));
const REGRAS = {
    '081': { proc: ler('regras_leitor_processamento_pernambuco_seed.json'), cls: ler('regras_leitor_classificacao_pernambuco_seed.json') },
    'padrao': { proc: ler('regras_leitor_processamento_seed.json'), cls: ler('regras_leitor_classificacao_seed.json') },
};
const casos = JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, 'casos_motor.json'), 'utf8'));
const saida = casos.map(c => {
    const r = REGRAS[c.regras];
    const itens = c.itens.map(i => ({ texto: i.texto, cor: i.cor || '#000000', layer: i.layer || '', pagina: i.pagina || 1, _x: i._x, _y: i._y }));
    return {
        nome: c.nome,
        resultados: itens.map((it, index) => {
            const res = motor.processarEClassificar(Object.assign({}, it, { index, allItems: itens }), r.proc, r.cls);
            return { operacao: res.operacao, ativo: res.ativo, entidade: res.entidade };
        }),
    };
});
process.stdout.write(JSON.stringify(saida, null, 1) + '\n');
