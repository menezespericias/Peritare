// inject-env.js — roda durante o build do Netlify
// Substitui os placeholders no HTML pelas variáveis de ambiente reais
const fs = require('fs');
const path = require('path');

const htmlPath = path.join(__dirname, 'public', 'index.html');
let html = fs.readFileSync(htmlPath, 'utf8');

const url = process.env.SUPABASE_URL || '';
const key = process.env.SUPABASE_KEY || '';

html = html.replace('SUPABASE_URL_PLACEHOLDER', url);
html = html.replace('SUPABASE_KEY_PLACEHOLDER', key);

fs.writeFileSync(htmlPath, html);
console.log('✅ Variáveis de ambiente injetadas no HTML');
