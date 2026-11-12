// scripts/gerar_icone.js (TASK-152) — gera static/gridr.ico (16…256 px, PNG dentro do .ico) e static/gridr-512.png a partir de static/gridr.svg.
// Só para quem mexe no desenho do ícone (o resultado é versionado; o build do .exe NÃO precisa rodar isto). Uso: node scripts/gerar_icone.js   (precisa do Playwright/Chromium)
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const RAIZ = path.join(__dirname, '..'), SVG = fs.readFileSync(path.join(RAIZ, 'static', 'gridr.svg'), 'utf8');
const TAMANHOS = [16, 24, 32, 48, 64, 128, 256];

function montarIco(pngs) {
    const cab = Buffer.alloc(6 + 16 * pngs.length);
    cab.writeUInt16LE(0, 0); cab.writeUInt16LE(1, 2); cab.writeUInt16LE(pngs.length, 4);
    let off = cab.length;
    pngs.forEach((p, i) => {
        const o = 6 + 16 * i;
        cab[o] = p.tam >= 256 ? 0 : p.tam; cab[o + 1] = p.tam >= 256 ? 0 : p.tam; cab[o + 2] = 0; cab[o + 3] = 0;
        cab.writeUInt16LE(1, o + 4); cab.writeUInt16LE(32, o + 6); cab.writeUInt32LE(p.png.length, o + 8); cab.writeUInt32LE(off, o + 12);
        off += p.png.length;
    });
    return Buffer.concat([cab, ...pngs.map(p => p.png)]);
}

(async () => {
    const b = await chromium.launch({ args: ['--no-sandbox'] });
    const pg = await b.newPage();
    const png = async tam => {
        await pg.setViewportSize({ width: tam, height: tam });
        await pg.setContent(`<body style="margin:0;background:transparent">${SVG.replace('width="112" height="112"', `width="${tam}" height="${tam}"`)}</body>`);
        return pg.screenshot({ omitBackground: true, clip: { x: 0, y: 0, width: tam, height: tam } });
    };
    const lista = [];
    for (const tam of TAMANHOS) lista.push({ tam, png: await png(tam) });
    fs.writeFileSync(path.join(RAIZ, 'static', 'gridr.ico'), montarIco(lista));
    fs.writeFileSync(path.join(RAIZ, 'static', 'gridr-512.png'), await png(512));
    await b.close();
    console.log('gerados: static/gridr.ico (' + TAMANHOS.join(', ') + ' px) e static/gridr-512.png');
})();
