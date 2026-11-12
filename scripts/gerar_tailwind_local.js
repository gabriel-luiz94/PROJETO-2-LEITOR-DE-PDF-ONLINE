// scripts/gerar_tailwind_local.js (TASK-153) — gera static/tailwind_local.css: só as classes do Tailwind usadas no login e no Painel Admin.
// Essas duas telas buscavam o Tailwind em cdn.tailwindcss.com e ficavam sem estilo quando o programa abria sem internet.
// Só para quem mudar as classes dessas telas (o resultado é versionado; o build do .exe NÃO roda isto). Uso: node scripts/gerar_tailwind_local.js   (precisa de internet para o npx)
const fs = require('fs'), os = require('os'), path = require('path');
const { execFileSync } = require('child_process');
const RAIZ = path.join(__dirname, '..');
const FONTES = ['login.html', 'admin.html', 'login.js', 'admin.js', 'campos_segredo.js'].map(f => path.join(RAIZ, 'static', f));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'tw-'));
const entrada = path.join(tmp, 'entrada.css');
fs.writeFileSync(entrada, '@tailwind base;\n@tailwind components;\n@tailwind utilities;\n');
const saida = path.join(RAIZ, 'static', 'tailwind_local.css');
execFileSync('npx', ['--yes', 'tailwindcss@3.4.19', '-i', entrada, '-o', saida, '--content', FONTES.join(','), '--minify'],
    { stdio: 'inherit', cwd: tmp, shell: process.platform === 'win32' });
fs.rmSync(tmp, { recursive: true, force: true });
console.log('Gerado: static/tailwind_local.css (' + fs.statSync(saida).size + ' bytes)');
