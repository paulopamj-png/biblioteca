'use strict';
/* Minha Biblioteca — PWA: audiobooks e PDFs do OneDrive (Microsoft Graph + MSAL) */
const $ = s => document.querySelector(s);
const CFG = window.CONFIG;
const DEMO = !CFG.clientId || /[?&]demo=1/.test(location.search);   // ?demo=1 mostra a lista sem login (teste visual)
const GRAPH = 'https://graph.microsoft.com/v1.0';
const ESCOPOS = ['Files.Read', 'User.Read'];
const VERSAO = '11';
const LOTE = 48;

let catalogo = [];
let aba = 'home', consulta = '', mostrando = LOTE, listaAtual = [];
let clienteMsal = null, conta = null;

/* ---------- ícones ---------- */
const IC = {
  a: '<svg viewBox="0 0 24 24"><path d="M3 14v-2a9 9 0 0 1 18 0v2"/><rect x="3" y="14" width="4" height="7" rx="2"/><rect x="17" y="14" width="4" height="7" rx="2"/></svg>',
  p: '<svg viewBox="0 0 24 24"><path d="M4 4.5A2.5 2.5 0 0 1 6.5 2H20v17H6.5A2.5 2.5 0 0 0 4 21.5z"/><path d="M4 21.5V4.5"/></svg>',
  play: '<svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>',
  pause: '<svg viewBox="0 0 24 24"><rect x="6" y="5" width="4" height="14" rx="1"/><rect x="14" y="5" width="4" height="14" rx="1"/></svg>'
};

/* ---------- utilidades ---------- */
const norm = s => (s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ');
const guarda = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
const le = k => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } };
const esc = s => (s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const vibra = () => { try { navigator.vibrate && navigator.vibrate(8); } catch (e) {} };
const fmtTempo = s => { s = Math.max(0, Math.floor(s || 0)); const h = (s / 3600) | 0, m = ((s % 3600) / 60) | 0, x = s % 60;
  return (h ? h + ':' + String(m).padStart(2, '0') : m) + ':' + String(x).padStart(2, '0'); };
const fmtDur = min => !min ? '' : (min >= 60 ? Math.floor(min / 60) + 'h' + (min % 60 ? String(min % 60).padStart(2, '0') : '') : min + ' min');
function aviso(html) { const a = $('#aviso'); a.innerHTML = html; a.hidden = !html; }

/* ---------- autenticação ---------- */
async function iniciaLogin() {
  if (DEMO) return null;
  clienteMsal = new window.msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: CFG.authority, redirectUri: location.origin + location.pathname },
    cache: { cacheLocation: 'localStorage' }
  });
  await clienteMsal.initialize();
  const r = await clienteMsal.handleRedirectPromise();
  conta = (r && r.account) || clienteMsal.getAllAccounts()[0] || null;
  if (conta) clienteMsal.setActiveAccount(conta);
  return conta;
}
async function token() {
  try { return (await clienteMsal.acquireTokenSilent({ scopes: ESCOPOS, account: conta })).accessToken; }
  catch (e) {
    if (!navigator.onLine) throw new Error('Sem internet');
    await clienteMsal.acquireTokenRedirect({ scopes: ESCOPOS }); throw e;
  }
}
async function graph(caminho, opts = {}) {
  const t = await token();
  const r = await fetch(GRAPH + caminho, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } });
  if (!r.ok) {
    let onde = caminho; try { onde = decodeURIComponent(caminho).replace('/me/drive/root:/', '').replace(/:\/content$/, '').replace(/\?select=.*$/, ''); } catch (x) {}
    const e = new Error('OneDrive respondeu ' + r.status + ' para "' + onde + '"'); e.status = r.status; throw e;
  }
  return r;
}
// CFG.pasta pode ter subpastas ("BOOKS/Biblioteca"); cada parte do caminho é codificada separadamente
const caminhoItem = rel => '/me/drive/root:/' + [...CFG.pasta.split('/'), ...rel.split('/')].map(encodeURIComponent).join('/');

/* ---------- catálogo ---------- */
async function carregaCatalogo() {
  if (DEMO) { catalogo = await (await fetch('catalogo.json')).json(); return; }
  if (!navigator.onLine && le('catalogo')) { catalogo = le('catalogo'); aviso('<b>Sem internet.</b> Você pode ler e ouvir os títulos que baixou para o tablet.'); return; }
  try {
    catalogo = await (await graph(caminhoItem('catalogo.json') + ':/content')).json();
    guarda('catalogo', catalogo);
  } catch (e) {
    const c = le('catalogo');
    if (!c) throw e;
    catalogo = c;
    aviso('<b>Sem conexão com o OneDrive.</b> Mostrando a lista guardada no aparelho.');
  }
}
const porId = id => catalogo.find(i => i.id === id);

/* ---------- peças visuais ---------- */
// Cada título mostra só a capa (a real, ou a imagem padrão de livro), o nome e, abaixo, o autor em fonte menor.
const foto = i => i.c ? '<img class="foto" alt="" loading="lazy" decoding="async" src="capas/' + i.id + '.jpg" onload="this.classList.add(&quot;ok&quot;)" onerror="this.remove()">' : '';
function capa(i, prog, comBotao) {
  return '<div class="capa">' + foto(i) + (prog ? '<i class="prog" style="width:' + prog + '%"></i>' : '') + (comBotao ? btnOff(i) : '') + '</div>';
}
const fmProg = i => i.t === 'a' ? 'a' : (ultFm(i) || (i.fm === 'e' ? 'e' : 'p'));
function progresso(i) {
  if (i.t === 'a') { const p = le('prog:' + i.id); return p && i.d ? Math.min(100, p.t / (i.d * 60) * 100) : 0; }
  if (fmProg(i) === 'e') return le('pct:' + i.id) || 0;
  const n = le('pag:' + i.id); return n > 1 && i.pg ? Math.min(100, n / i.pg * 100) : 0;
}
function cartao(i, atraso) {
  return '<button class="card" data-id="' + i.id + '" style="animation-delay:' + (atraso || 0) + 'ms">' + capa(i, progresso(i), true) +
    '<span class="nm">' + esc(i.ti) + '</span><span class="au">' + esc(i.au || 'Autor não informado') + '</span></button>';
}
function linha(i, atraso, status) {
  const meta = status || (i.t === 'a' ? 'Audiobook' + (i.d ? ' · ' + fmtDur(i.d) : '') : 'Livro' + (i.pg ? ' · ' + i.pg + ' págs' : ''));
  const pr = status ? progresso(i) : 0;
  return '<button class="linha ' + i.t + '" data-id="' + i.id + '" style="animation-delay:' + Math.min(atraso || 0, 300) + 'ms">' + capa(i, 0) +
    '<span class="dd"><b>' + esc(i.ti) + '</b><small>' + esc(i.au || 'Autor não informado') + '</small><span class="mt">' + meta + '</span>' +
    (pr ? '<span class="barra"><i style="width:' + pr + '%"></i></span>' : '') + '</span>' + btnOff(i) + '</button>';
}

const FONE_MINI = '<svg viewBox="0 0 24 24"><path d="M3 14v-2a9 9 0 0 1 18 0v2"/><rect x="3" y="14" width="4" height="7" rx="2"/><rect x="17" y="14" width="4" height="7" rx="2"/></svg>';
// Audiobook: caixa larga e horizontal (capa, se houver, ou a imagem padrão de audiobook), título, autor e duração.
function caixa(i, atraso, status) {
  const pr = progresso(i);
  const meta = status || (i.d ? fmtDur(i.d) : 'Audiobook');
  return '<button class="caixa' + (i.c ? '' : ' sem-capa') + '" data-id="' + i.id + '" style="animation-delay:' + Math.min(atraso || 0, 300) + 'ms">' +
    '<span class="ab-capa">' + foto(i) + '</span>' +
    '<span class="ab-info"><b>' + esc(i.ti) + '</b><small>' + esc(i.au || 'Autor não informado') + '</small>' +
    '<span class="ab-meta">' + FONE_MINI + esc(meta) + '</span>' + (pr ? '<span class="barra"><i style="width:' + pr + '%"></i></span>' : '') + '</span>' +
    btnOff(i) + '<span class="ab-play">' + IC.play + '</span></button>';
}
const item = (i, atraso, status) => i.t === 'a' ? caixa(i, atraso, status) : linha(i, atraso, status);

/* ---------- o que o leitor já começou ---------- */
function iniciados() {
  const out = [], oc = le('ocultos') || {};
  for (const i of catalogo) {
    let q = 0;
    if (i.t === 'a') { const p = le('prog:' + i.id); if (p && p.t > 5) q = p.quando || 1; }
    else if ((le('pct:' + i.id) || 0) > 0 || (le('pag:' + i.id) || 0) > 1) q = le('qdo:' + i.id) || 1;
    if (q && !(oc[i.id] && q <= oc[i.id])) out.push([i, q]);     // removido da lista: só volta se for aberto de novo
  }
  return out.sort((a, b) => b[1] - a[1]).map(x => x[0]);
}
function statusIni(i) {
  const pr = Math.round(progresso(i));
  if (i.t === 'a') return 'Parou em ' + fmtTempo(le('prog:' + i.id).t) + (pr ? ' · ' + pr + '%' : '');
  if (fmProg(i) === 'e') return 'Leitura em ' + pr + '%' + (formatosDe(i).length > 1 ? ' · EPUB' : '');
  return 'Página ' + le('pag:' + i.id) + (i.pg ? ' de ' + i.pg : '') + (pr ? ' · ' + pr + '%' : '') + (formatosDe(i).length > 1 ? ' · PDF' : '');
}

/* ---------- formatos (EPUB/PDF) e downloads para usar sem internet ---------- */
// Os arquivos baixados ficam no Cache Storage do navegador ("bib-offline"); "Remover" apaga de lá e libera o espaço.
const CACHE_OFF = 'bib-offline';
const NOME_FM = { e: 'EPUB', p: 'PDF', a: 'áudio' };
const temEpub = i => i.t === 'p' && (i.fm === 'e' || !!i.ep);
const temPdf = i => i.t === 'p' && i.fm !== 'e';
const formatosDe = i => i.t === 'a' ? ['a'] : [...(temEpub(i) ? ['e'] : []), ...(temPdf(i) ? ['p'] : [])];
const arquivoDe = (i, fm) => (i.t === 'a' || fm === 'p' || (fm === 'e' && i.fm === 'e')) ? i.f : i.f.replace(/\.pdf$/i, '.epub');
const pastaDe = fm => fm === 'a' ? 'audio/' : 'livros/';
const fmPadrao = i => { const f = formatosDe(i); if (f.length === 1) return f[0]; const p = le('fmtpref'); return p && f.includes(p) ? p : 'e'; };
const ultFm = i => { const f = formatosDe(i), u = le('ult:' + i.id); return u && f.includes(u) ? u : (f.length === 1 ? f[0] : null); };
const chaveOff = (i, fm) => 'offline/' + i.id + '.' + fm;
const baixadosMapa = () => le('off') || {};
const estaBaixado = (i, fm) => { const m = baixadosMapa(); return fm ? !!m[i.id + '.' + fm] : formatosDe(i).some(f => !!m[i.id + '.' + f]); };
const bytesBaixados = (i, fm) => { const m = baixadosMapa(); return (fm ? [fm] : formatosDe(i)).reduce((s, f) => s + ((m[i.id + '.' + f] || {}).bytes || 0), 0); };
const emDownload = {};                                  // id -> { fm, pct, ctrl }
const fmtMB = b => !b ? '0 MB' : b >= 1073741824 ? (b / 1073741824).toFixed(1) + ' GB' : b >= 1048576 ? Math.round(b / 1048576) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
const ICO_BAIXAR = '<svg viewBox="0 0 24 24"><path d="M12 4v11M7 11l5 5 5-5M5 20h14"/></svg>';
const ICO_OK = '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>';
function marcaBtn(attr, valor, estado, pct, rotulos) {
  const cls = estado, rot = rotulos[estado];
  return '<span class="btn-off ' + cls + '" role="button" ' + attr + '="' + valor + '" style="--p:' + Math.round(pct) + '%" aria-label="' + rot + '" title="' + rot + '">' +
    (estado === 'baixando' ? '<b>' + Math.round(pct) + '</b>' : estado === 'sim' ? ICO_OK : ICO_BAIXAR) + '</span>';
}
function btnOff(i) {                                   // botao do título (todos os formatos)
  const d = emDownload[i.id], b = estaBaixado(i);
  return marcaBtn('data-off', i.id, d ? 'baixando' : b ? 'sim' : 'nao', d ? d.pct : 0,
    { baixando: 'Cancelar o download', sim: 'Remover do tablet', nao: 'Baixar para ler ou ouvir sem internet' });
}
function btnOffFm(i, fm) {                              // botao de um formato específico
  const d = emDownload[i.id] && emDownload[i.id].fm === fm ? emDownload[i.id] : null, b = estaBaixado(i, fm);
  return marcaBtn('data-off-fm', i.id + '.' + fm, d ? 'baixando' : b ? 'sim' : 'nao', d ? d.pct : 0,
    { baixando: 'Cancelar o download', sim: 'Remover este formato do tablet', nao: 'Baixar ' + NOME_FM[fm] + ' para usar sem internet' });
}
function pintaOff(id) {
  const i = porId(id); if (!i) return;
  document.querySelectorAll('[data-off="' + id + '"]').forEach(el => { el.outerHTML = btnOff(i); });
  document.querySelectorAll('[data-off-fm^="' + id + '."]').forEach(el => { const fm = el.dataset.offFm.split('.').pop(); el.outerHTML = btnOffFm(i, fm); });
  if (atual && atual.id === id) pintaOffPlayer();
  if (escolhaItem && escolhaItem.id === id) renderEscolha();
}
function pintaOffPlayer() {
  if (!atual) return;
  const d = emDownload[atual.id], b = estaBaixado(atual, 'a');
  $('#pOff').textContent = d ? 'Baixando ' + Math.round(d.pct) + '% (tocar cancela)' : b ? '✓ No tablet · Remover' : '⬇ Baixar para ouvir offline';
}
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.hidden = false; requestAnimationFrame(() => t.classList.add('ver'));
  clearTimeout(toast.t); toast.t = setTimeout(() => { t.classList.remove('ver'); setTimeout(() => { t.hidden = true; }, 350); }, 3400);
}
async function baixa(i, fm) {
  fm = fm || fmPadrao(i);
  if (DEMO) return alert('Modo demonstração: não há o que baixar.');
  if (!navigator.onLine) return alert('Sem internet agora. Conecte-se ao Wi-Fi para baixar este título.');
  const rel = pastaDe(fm) + arquivoDe(i, fm);
  const ctrl = new AbortController(); emDownload[i.id] = { fm, pct: 0, ctrl }; pintaOff(i.id);
  const cache = await caches.open(CACHE_OFF);
  try {
    const total = (await (await graph(caminhoItem(rel) + '?select=size')).json()).size || 0;
    let r;
    try { r = await fetch(GRAPH + caminhoItem(rel) + ':/content', { headers: { Authorization: 'Bearer ' + await token() }, signal: ctrl.signal }); }
    catch (e) {                                               // plano B: link temporário já autorizado
      if (ctrl.signal.aborted) throw e;
      r = await fetch(await urlDireta(rel), { signal: ctrl.signal });
    }
    if (!r.ok) throw new Error('OneDrive respondeu ' + r.status);
    const leitor = r.body.getReader(); let feito = 0, ultimo = 0;
    const fluxo = new ReadableStream({
      async pull(c) {
        const { done, value } = await leitor.read();
        if (done) { c.close(); return; }
        feito += value.length; emDownload[i.id].pct = total ? feito / total * 100 : 0;
        if (Date.now() - ultimo > 500) { ultimo = Date.now(); pintaOff(i.id); }
        c.enqueue(value);
      },
      cancel(m) { return leitor.cancel(m); }
    });
    await cache.put(chaveOff(i, fm), new Response(fluxo, { headers: { 'Content-Type': fm === 'a' ? 'audio/mp4' : fm === 'e' ? 'application/epub+zip' : 'application/pdf' } }));
    if (total && feito !== total) throw new Error('o arquivo veio incompleto');
    const m = baixadosMapa(); m[i.id + '.' + fm] = { bytes: feito, quando: Date.now() }; guarda('off', m);
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {}
    toast('Baixado: ' + i.ti + (formatosDe(i).length > 1 ? ' (' + NOME_FM[fm] + ')' : '') + ' · ' + fmtMB(feito));
  } catch (e) {
    try { await cache.delete(chaveOff(i, fm)); } catch (x) {}
    if (ctrl.signal.aborted) toast('Download cancelado');
    else alert('Não consegui baixar "' + i.ti + '".\n' + (e.message === 'Failed to fetch' ? 'O navegador não conseguiu receber o arquivo. Verifique a conexão e tente de novo.' : e.message));
  } finally { delete emDownload[i.id]; pintaOff(i.id); baixadosHome(); }
}
async function urlDireta(rel) {
  const j = await (await graph(caminhoItem(rel) + '?select=id,@microsoft.graph.downloadUrl')).json();
  if (!j['@microsoft.graph.downloadUrl']) throw new Error('Sem URL de download');
  return j['@microsoft.graph.downloadUrl'];
}
async function removeBaixado(i, fm, perguntar = true) {   // fm vazio = remove todos os formatos baixados do título
  const m = baixadosMapa();
  const fs = (fm ? [fm] : formatosDe(i)).filter(f => m[i.id + '.' + f]);
  if (!fs.length) return false;
  const bytes = fs.reduce((s, f) => s + (m[i.id + '.' + f].bytes || 0), 0);
  const quais = fs.length > 1 || formatosDe(i).length > 1 ? ' (' + fs.map(f => NOME_FM[f]).join(' + ') + ')' : '';
  if (perguntar && !confirm('Remover "' + i.ti + '"' + quais + ' do tablet?\n\nIsso libera ' + fmtMB(bytes) + ' de espaço. O título continua na sua biblioteca e você pode baixar de novo quando quiser.')) return false;
  if (atual && atual.id === i.id && fs.includes('a')) encerraAudio();      // libera o arquivo que está em uso
  const cache = await caches.open(CACHE_OFF);
  for (const f of fs) { try { await cache.delete(chaveOff(i, f)); } catch (e) {} delete m[i.id + '.' + f]; }
  guarda('off', m);
  pintaOff(i.id); baixadosHome(); toast('Removido do tablet: ' + i.ti + ' · liberou ' + fmtMB(bytes));
  return true;
}
function cliqueOff(i) {
  if (!i) return;
  if (emDownload[i.id]) { emDownload[i.id].ctrl.abort(); return; }
  if (estaBaixado(i)) removeBaixado(i, null); else baixa(i);
}
function cliqueOffFm(i, fm) {
  if (!i) return;
  if (emDownload[i.id]) { emDownload[i.id].ctrl.abort(); return; }
  if (estaBaixado(i, fm)) removeBaixado(i, fm); else baixa(i, fm);
}
async function doCache(i, fm) { try { return await (await caches.open(CACHE_OFF)).match(chaveOff(i, fm)); } catch (e) { return null; } }
async function dadosLivro(i, fm) {                      // ArrayBuffer do livro: do tablet se baixado; senão, do OneDrive
  if (estaBaixado(i, fm)) {
    const r = await doCache(i, fm);
    if (r) return r.arrayBuffer();
    const m = baixadosMapa(); delete m[i.id + '.' + fm]; guarda('off', m);      // sumiu do armazenamento: corrige o registro
  }
  if (!navigator.onLine) throw new Error('Sem internet. Baixe o título antes (ícone ↓) para ler offline.');
  return (await graph(caminhoItem('livros/' + arquivoDe(i, fm)) + ':/content')).arrayBuffer();
}
async function reconciliaBaixados() {                  // confere o registro de baixados com o armazenamento (e migra o formato antigo)
  try {
    const cache = await caches.open(CACHE_OFF);
    const m = baixadosMapa(); let mudou = false;
    for (const k of Object.keys(m)) {                   // registros antigos "id" -> "id.formato"
      if (k.includes('.')) continue;
      const i = porId(k), fm = i ? formatosDe(i)[0] : null, r = await cache.match('offline/' + k);
      if (i && r) { await cache.put(chaveOff(i, fm), r); m[k + '.' + fm] = m[k]; }
      await cache.delete('offline/' + k); delete m[k]; mudou = true;
    }
    const ks = new Set((await cache.keys()).map(r => new URL(r.url).pathname.split('/').slice(-2).join('/')));
    Object.keys(m).forEach(k => { if (!ks.has('offline/' + k)) { delete m[k]; mudou = true; } });
    if (mudou) guarda('off', m);
  } catch (e) {}
}
function baixadosHome() { if (!$('#home').hidden) desenhaBaixados(); }
async function desenhaBaixados() {
  const m = baixadosMapa(), box = $('#baixadosSec');
  const itens = catalogo.filter(i => estaBaixado(i)).sort((a, b) => Math.max(0, ...formatosDe(b).map(f => (m[b.id + '.' + f] || {}).quando || 0)) - Math.max(0, ...formatosDe(a).map(f => (m[a.id + '.' + f] || {}).quando || 0)));
  const total = itens.reduce((s, i) => s + bytesBaixados(i), 0);
  let uso = '';
  try { if (navigator.storage && navigator.storage.estimate) { const e = await navigator.storage.estimate(); uso = ' · o app ocupa ' + fmtMB(e.usage || 0) + ' no tablet'; } } catch (e) {}
  if (!itens.length) {
    box.innerHTML = '<div class="secao"><h2>Baixados no tablet</h2></div>' +
      '<div class="vazio-ini">Nenhum título baixado ainda. Toque no ícone <b>↓</b> de um livro ou audiobook (na capa, ao lado do play, ou dentro do leitor) para guardá-lo no tablet e usar <b>sem internet</b>. Quando não quiser mais, é só remover aqui e o espaço é liberado.</div>';
    return;
  }
  const status = i => formatosDe(i).filter(f => m[i.id + '.' + f]).map(f => (formatosDe(i).length > 1 ? NOME_FM[f] + ' ' : '') + fmtMB(m[i.id + '.' + f].bytes || 0)).join(' + ') + ' no tablet';
  box.innerHTML = '<div class="secao"><h2>Baixados no tablet</h2><button class="sorteia" data-off-todos="1">Remover todos</button></div>' +
    '<p class="resumo-off">' + itens.length + (itens.length === 1 ? ' título' : ' títulos') + ' · ' + fmtMB(total) + uso + '. Toque no ✓ de um título para removê-lo.</p>' +
    '<div class="continua">' + itens.map((i, n) => item(i, n * 40, status(i))).join('') + '</div>';
}
async function removeTodos() {
  const itens = catalogo.filter(i => estaBaixado(i));
  if (!itens.length) return;
  if (!confirm('Remover todos os ' + itens.length + ' títulos baixados do tablet?\nIsso libera ' + fmtMB(itens.reduce((s, i) => s + bytesBaixados(i), 0)) + '. A biblioteca no OneDrive não é afetada.')) return;
  for (const i of itens) await removeBaixado(i, null, false);
  toast('Todos os downloads foram removidos');
}

/* ---------- escolha entre EPUB e PDF; preferências ---------- */
let escolhaItem = null;
function escolheFormato(i) { escolhaItem = i; $('#escLembrar').checked = false; renderEscolha(); $('#escolhe').hidden = false; }
function renderEscolha() {
  const i = escolhaItem; if (!i) return;
  $('#escCapa').innerHTML = foto(i); $('#escTitulo').textContent = i.ti; $('#escAutor').textContent = i.au || '';
  const info = { e: ['EPUB', 'Texto que se ajusta à tela, com letra e fundo à sua escolha'], p: ['PDF', 'Páginas originais, como no livro impresso'] };
  $('#escOpcoes').innerHTML = formatosDe(i).map(f => '<div class="esc-op" data-escolhe="' + f + '" role="button"><span class="esc-t"><b>' + info[f][0] + '</b><small>' + info[f][1] + '</small>' +
    '<em>' + (estaBaixado(i, f) ? '✓ guardado no tablet' : 'precisa de internet') + '</em></span>' + btnOffFm(i, f) + '</div>').join('');
}
function fechaModais() { $('#escolhe').hidden = true; $('#menu').hidden = true; escolhaItem = null; }
function abreMenu() { renderMenu(); $('#menu').hidden = false; }
function renderMenu() {
  const pref = le('fmtpref') || '';
  const itens = catalogo.filter(i => estaBaixado(i)), total = itens.reduce((s, i) => s + bytesBaixados(i), 0);
  $('#menuCorpo').innerHTML =
    '<h4>Quando o título tiver EPUB e PDF</h4><div class="seg">' + [['', 'Perguntar'], ['e', 'Preferir EPUB'], ['p', 'Preferir PDF']].map(([v, r]) => '<button data-pref="' + v + '" class="' + (pref === v ? 'on' : '') + '">' + r + '</button>').join('') + '</div>' +
    '<h4>Armazenamento do tablet</h4><p>' + itens.length + (itens.length === 1 ? ' título baixado' : ' títulos baixados') + ' · ' + fmtMB(total) + '</p>' +
    '<button class="menu-b" data-menu="baixados">Ver baixados</button>' + (itens.length ? '<button class="menu-b" data-menu="limpar">Remover todos os downloads</button>' : '') +
    (conta ? '<h4>Conta</h4><p>' + esc(conta.username) + '</p><button class="menu-b sair" data-menu="sair">Sair da conta</button>' : '') +
    '<p class="versao">Minha Biblioteca · versão ' + VERSAO + '</p>';
}

/* ---------- telas ---------- */
const saudacao = () => { const h = new Date().getHours(); return h < 5 ? 'Boa madrugada' : h < 12 ? 'Bom dia' : h < 18 ? 'Boa tarde' : 'Boa noite'; };
const cmpTitulo = (x, y) => x.ti.localeCompare(y.ti, 'pt-BR', { sensitivity: 'base' });

function desenha() {
  const buscando = !!consulta;
  const emHome = aba === 'home' && !buscando;
  $('#home').hidden = !emHome; $('#estante').hidden = emHome;
  $('#cats').classList.toggle('compacta', !emHome);
  document.querySelectorAll('.cat').forEach(b => {
    const on = b.dataset.ir === aba;
    b.classList.toggle('ativa', on); b.classList.toggle('apagada', aba !== 'home' && !on);
  });
  $('#nav .ini').classList.toggle('on', emHome);
  $('#limpa').hidden = !buscando;
  $('#infoAudio').textContent = catalogo.filter(i => i.t === 'a').length + ' títulos';
  $('#infoPdf').textContent = catalogo.filter(i => i.t === 'p').length + ' títulos';
  if (emHome) desenhaHome(); else desenhaEstante();
}

function desenhaHome() {
  const ini = iniciados().slice(0, 5);               // os 5 últimos títulos iniciados (livros e audiobooks)
  $('#continuar').innerHTML = '<div class="secao"><h2>Continue de onde parou</h2></div>' +
    (ini.length ? '<div class="continua">' + ini.map((i, n) => item(i, n * 50, statusIni(i)).replace('<span class="btn-off', '<span class="btn-x" role="button" data-ocultar="' + i.id + '" aria-label="Remover desta lista" title="Remover desta lista">✕</span><span class="btn-off')).join('') + '</div>'
                : '<div class="vazio-ini">Você ainda não começou nenhum título. Escolha um nas sugestões abaixo e ele aparecerá aqui, com o ponto em que você parou.</div>');
  desenhaBaixados();
  sugestoes('p'); sugestoes('a');
}
function sorteio(lista, n) {
  const a = lista.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.random() * (i + 1) | 0; [a[i], a[j]] = [a[j], a[i]]; }
  return a.slice(0, n);
}
function sugestoes(t) {
  if (!catalogo.length) return;
  const lista = sorteio(catalogo.filter(i => i.t === t), 10);
  $(t === 'p' ? '#sugLivros' : '#sugAudio').innerHTML =
    '<div class="secao"><h2>' + (t === 'p' ? 'Sugestões de livros' : 'Sugestões de audiobooks') + '</h2>' +
    '<button class="sorteia" data-sort="' + t + '">Surpreenda-me ↻</button></div>' +
    '<div class="carrossel">' + lista.map((i, n) => (t === 'a' ? caixa : cartao)(i, n * 40)).join('') + '</div>';
}

function desenhaEstante() {
  const tipo = aba === 'audio' ? 'a' : aba === 'pdf' ? 'p' : '';
  const termos = norm(consulta).split(' ').filter(Boolean), q0 = termos.join(' ');
  let l = catalogo.filter(i => (!tipo || i.t === tipo) && (!termos.length || termos.every(t => (i._n || (i._n = norm(i.ti + ' ' + i.au))).includes(t))));
  l.sort((x, y) => consulta ? ((norm(x.ti).trim().startsWith(q0) ? 0 : 1) - (norm(y.ti).trim().startsWith(q0) ? 0 : 1)) || cmpTitulo(x, y) : cmpTitulo(x, y));
  listaAtual = l;
  $('#tituloEstante').textContent = consulta ? 'Resultados' : aba === 'audio' ? 'Audiobooks' : 'Livros';
  $('#resumo').textContent = l.length + (l.length === 1 ? ' título' : ' títulos');
  const box = $('#lista'); box.className = consulta ? 'linhas' : aba === 'audio' ? 'caixas' : 'grade';
  box.innerHTML = '';
  mostrando = LOTE; maisLista();
  if (!l.length) box.innerHTML = '<div class="vazio">Nada encontrado para “' + esc(consulta) + '”.<br>Tente outra palavra.</div>';
}
function maisLista() {
  const box = $('#lista'), de = box.querySelectorAll('.card, .linha, .caixa').length;
  if (de >= listaAtual.length) return;
  const f = consulta ? item : aba === 'audio' ? caixa : cartao;
  box.insertAdjacentHTML('beforeend', listaAtual.slice(de, de + LOTE).map((i, n) => f(i, n * 18)).join(''));
}
new IntersectionObserver(es => { if (es[0].isIntersecting && !$('#estante').hidden) maisLista(); }, { rootMargin: '800px' }).observe($('#fim'));

/* ---------- eventos ---------- */
document.addEventListener('click', e => {
  const t = e.target;
  const ox = t.closest('[data-ocultar]');
  if (ox) { e.preventDefault(); e.stopPropagation(); vibra(); const it = porId(ox.dataset.ocultar); if (!confirm('Quer mesmo remover “' + (it ? it.ti : 'este título') + '” de “Continue de onde parou”?' + String.fromCharCode(10, 10) + 'O ponto em que você parou fica guardado.')) return; const oc = le('ocultos') || {}; oc[ox.dataset.ocultar] = Date.now() + 1000; guarda('ocultos', oc); desenhaHome(); toast('Removido de "Continue de onde parou"'); return; }
  const bf = t.closest('[data-off-fm]');
  if (bf) { e.preventDefault(); e.stopPropagation(); vibra(); const k = bf.dataset.offFm, p = k.lastIndexOf('.'); cliqueOffFm(porId(k.slice(0, p)), k.slice(p + 1)); return; }
  const bo = t.closest('[data-off]');
  if (bo) { e.preventDefault(); e.stopPropagation(); vibra(); cliqueOff(porId(bo.dataset.off)); return; }
  if (t.closest('[data-off-todos]')) { vibra(); removeTodos(); return; }
  if (t.closest('[data-fecha]')) { fechaModais(); return; }
  const eo = t.closest('[data-escolhe]');
  if (eo && escolhaItem) { vibra(); const i = escolhaItem, f = eo.dataset.escolhe; if ($('#escLembrar').checked) guarda('fmtpref', f); fechaModais(); abre(i, f); return; }
  const pf = t.closest('[data-pref]');
  if (pf) { guarda('fmtpref', pf.dataset.pref); renderMenu(); toast(pf.dataset.pref ? 'Vou abrir em ' + NOME_FM[pf.dataset.pref] + ' quando houver as duas versões' : 'Vou perguntar quando houver as duas versões'); return; }
  const mn = t.closest('[data-menu]');
  if (mn) {
    const a = mn.dataset.menu;
    if (a === 'baixados') { fechaModais(); vaiPara('home'); setTimeout(() => $('#baixadosSec').scrollIntoView({ behavior: 'smooth', block: 'start' }), 250); }
    else if (a === 'limpar') { fechaModais(); removeTodos(); }
    else if (a === 'sair') { fechaModais(); if (confirm('Sair da conta ' + conta.username + '?')) clienteMsal.logoutRedirect({ account: conta }); }
    return;
  }
  const ir = t.closest('[data-ir]');
  if (ir) {
    vibra(); const dest = ir.dataset.ir;
    ir.classList.add('abre');
    setTimeout(() => { ir.classList.remove('abre'); vaiPara(aba === dest && !consulta ? 'home' : dest, !!consulta); }, 200);
    return;
  }
  if (t.closest('#nav .ini')) { vibra(); vaiPara('home'); return; }
  const so = t.closest('[data-sort]'); if (so) { vibra(); sugestoes(so.dataset.sort); return; }
  const it = t.closest('.card, .linha, .caixa'); if (it) { vibra(); const i = porId(it.dataset.id); abre(i, it.closest('#continuar') ? ultFm(i) : undefined); }
});
function vaiPara(destino, manterBusca) {
  aba = destino;
  if (!manterBusca) { consulta = ''; $('#q').value = ''; }
  $('#q').blur();
  desenha(); scrollTo({ top: 0, behavior: 'instant' in window ? 'instant' : 'auto' });
}
let tq;
$('#q').addEventListener('input', e => { clearTimeout(tq); tq = setTimeout(() => { consulta = e.target.value.trim(); desenha(); }, 140); });
$('#q').addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });   // "Buscar" do teclado do Android só fecha o teclado
$('#limpa').onclick = () => { $('#q').value = ''; consulta = ''; desenha(); };
$('#marca').onclick = () => vaiPara('home');
$('#btnConta').onclick = () => {
  if (!DEMO && !conta) { clienteMsal.loginRedirect({ scopes: ESCOPOS }); return; }
  abreMenu();
};

function abre(i, fm) {
  if (!i) return;
  if (DEMO) { alert('Modo demonstração: os arquivos ainda não estão ligados ao OneDrive.\n\n' + i.ti + (i.au ? '\n' + i.au : '')); return; }
  if (i.t === 'a') {
    if (!navigator.onLine && !estaBaixado(i, 'a')) { alert('Sem internet agora, e este audiobook ainda não foi baixado.\n\nQuando estiver no Wi-Fi, toque no ícone ↓ para baixá-lo e poder ouvir offline.'); return; }
    return tocaAudio(i);
  }
  const fs = formatosDe(i);
  if (!fm) {
    if (fs.length === 1) fm = fs[0];
    else if (!navigator.onLine) {                                     // sem internet: usa a versão que está no tablet
      const gs = fs.filter(f => estaBaixado(i, f)); fm = gs.length ? (gs.includes(fmPadrao(i)) ? fmPadrao(i) : gs[0]) : null;
    } else {
      const pref = le('fmtpref');
      if (pref && fs.includes(pref)) fm = pref; else return escolheFormato(i);
    }
  }
  if (!fm || (!navigator.onLine && !estaBaixado(i, fm))) { alert('Sem internet agora, e este título ainda não foi baixado.\n\nQuando estiver no Wi-Fi, toque no ícone ↓ do título para baixá-lo e poder abrir offline.'); return; }
  guarda('ult:' + i.id, fm);
  if (fm === 'e') abreEpub(i); else abrePdf(i);
}

/* ---------- player de áudio ---------- */
const au = new Audio();
au.preload = 'metadata';
let atual = null, ultimoSalvo = 0, sonoFim = 0, sonoTimer = null;
const VELS = [1, 1.25, 1.5, 1.75, 2, 0.85];
const SONO = [0, 15, 30, 60];

async function urlAudio(i) {
  const j = await (await graph(caminhoItem('audio/' + i.f) + '?select=id,@microsoft.graph.downloadUrl')).json();
  if (!j['@microsoft.graph.downloadUrl']) throw new Error('Sem URL de download');
  return j['@microsoft.graph.downloadUrl'];
}
let blobAtual = null;
async function fonteAudio(i) {                       // do tablet (se baixado) ou streaming do OneDrive
  if (blobAtual) { URL.revokeObjectURL(blobAtual); blobAtual = null; }
  if (estaBaixado(i, 'a')) { const r = await doCache(i, 'a'); if (r) { blobAtual = URL.createObjectURL(await r.blob()); return blobAtual; } }
  if (!navigator.onLine) throw new Error('Sem internet. Baixe o audiobook antes para ouvir offline.');
  return urlAudio(i);
}
function pintaPlayer(i) {
  $('#sCapa').innerHTML = foto(i);
  $('#sTitulo').textContent = i.ti; $('#sAutor').textContent = i.au || '';
  $('#mCapa').innerHTML = foto(i);
  $('#mTitulo').textContent = i.ti; $('#mAutor').textContent = i.au || '';
  pintaOffPlayer();
}
async function tocaAudio(i) {
  if (atual && atual.id === i.id) { abreSheet(true); return; }
  salvaProgresso(true);
  atual = i; pintaPlayer(i);
  $('#mini').hidden = false; document.body.classList.add('com-mini');
  setPlay(false); abreSheet(true);
  try {
    au.src = await fonteAudio(i);
    const p = le('prog:' + i.id);
    au.playbackRate = le('vel') || 1; $('#pVel').textContent = au.playbackRate + '×';
    au.onloadedmetadata = () => { if (p && p.t > 5 && p.t < au.duration - 5) au.currentTime = p.t; au.onloadedmetadata = null; au.play().catch(() => {}); };
    mediaSession(i);
  } catch (e) { aviso('Não consegui abrir o áudio: ' + esc(e.message)); abreSheet(false); }
}
const setPlay = tocando => { const h = tocando ? IC.pause : IC.play; $('#pPlay').innerHTML = h; $('#mPlay').innerHTML = h; };
function abreSheet(v) { const s = $('#sheet'); s.classList.toggle('aberto', v); s.setAttribute('aria-hidden', !v); }

// A URL temporária do OneDrive expira (~1 h): se o streaming falhar, renova e retoma do mesmo ponto.
au.addEventListener('error', async () => {
  if (!atual || blobAtual) return;                       // arquivo do tablet não expira
  const t = au.currentTime;
  try { au.src = await urlAudio(atual); au.onloadedmetadata = () => { au.currentTime = t; au.play().catch(() => {}); au.onloadedmetadata = null; }; } catch (e) {}
});
au.addEventListener('play', () => { setPlay(true); if ('mediaSession' in navigator) navigator.mediaSession.playbackState = 'playing'; });
au.addEventListener('pause', () => { setPlay(false); salvaProgresso(true); if ('mediaSession' in navigator) navigator.mediaSession.playbackState = 'paused'; });
au.addEventListener('timeupdate', () => {
  if (!au.duration) return;
  const f = au.currentTime / au.duration;
  $('#pSeek').value = f * 1000; $('#pSeek').style.setProperty('--v', f * 100 + '%'); $('#mBarra').style.width = f * 100 + '%';
  $('#pAtual').textContent = fmtTempo(au.currentTime); $('#pRest').textContent = '-' + fmtTempo(au.duration - au.currentTime);
  salvaProgresso(false);
});
au.addEventListener('ended', () => {
  if (!atual) return;
  guarda('prog:' + atual.id, { t: 0, quando: Date.now() });
  const i = atual;
  if (estaBaixado(i, 'a')) setTimeout(() => { if (confirm('Você terminou "' + i.ti + '".\n\nRemover do tablet para liberar espaço?')) removeBaixado(i, 'a', false); }, 600);
});
function salvaProgresso(forca) {
  if (!atual || !au.currentTime) return;
  if (!forca && Date.now() - ultimoSalvo < 5000) return;
  ultimoSalvo = Date.now(); guarda('prog:' + atual.id, { t: au.currentTime, quando: ultimoSalvo });
}
const alterna = () => { vibra(); au.paused ? au.play() : au.pause(); };
$('#pPlay').onclick = alterna; $('#mPlay').onclick = alterna;
$('#pVolta').onclick = () => { vibra(); au.currentTime = Math.max(0, au.currentTime - 30); };
$('#pAvanca').onclick = () => { vibra(); au.currentTime = Math.min(au.duration || 1e9, au.currentTime + 30); };
$('#pSeek').oninput = e => { if (au.duration) au.currentTime = e.target.value / 1000 * au.duration; };
$('#pVel').onclick = () => { vibra(); const v = VELS[(VELS.indexOf(au.playbackRate) + 1) % VELS.length]; au.playbackRate = v; guarda('vel', v); $('#pVel').textContent = v + '×'; };
let sonoIdx = 0;
function limpaSono() { clearInterval(sonoTimer); sonoFim = 0; sonoIdx = 0; $('#pSono').textContent = '☾ Dormir'; }
$('#pSono').onclick = () => {
  vibra(); clearInterval(sonoTimer);
  sonoIdx = (sonoIdx + 1) % SONO.length;
  if (!SONO[sonoIdx]) return limpaSono();
  sonoFim = Date.now() + SONO[sonoIdx] * 60000;
  const tic = () => { const r = sonoFim - Date.now(); if (r <= 0) { au.pause(); limpaSono(); } else $('#pSono').textContent = '☾ ' + Math.ceil(r / 60000) + ' min'; };
  tic(); sonoTimer = setInterval(tic, 15000);
};
$('#pOff').onclick = () => { if (atual) cliqueOff(atual); };
$('#sFecha').onclick = () => abreSheet(false);
$('#mAbre').onclick = () => abreSheet(true);
function encerraAudio() {
  salvaProgresso(true); au.pause(); au.removeAttribute('src'); try { au.load(); } catch (e) {}
  atual = null; limpaSono();
  if (blobAtual) { URL.revokeObjectURL(blobAtual); blobAtual = null; }
  if ('mediaSession' in navigator) { try { navigator.mediaSession.metadata = null; navigator.mediaSession.playbackState = 'none'; } catch (e) {} }
  abreSheet(false); $('#mini').hidden = true; document.body.classList.remove('com-mini'); desenha();
}
$('#pFechar').onclick = encerraAudio;
$('#mFecha').onclick = e => { e.stopPropagation(); vibra(); encerraAudio(); };
window.addEventListener('pagehide', () => salvaProgresso(true));
function mediaSession(i) {
  if (!('mediaSession' in navigator)) return;
  navigator.mediaSession.metadata = new MediaMetadata({ title: i.ti, artist: i.au || '', album: 'Minha Biblioteca' });
  navigator.mediaSession.setActionHandler('play', () => au.play());
  navigator.mediaSession.setActionHandler('pause', () => au.pause());
  navigator.mediaSession.setActionHandler('seekbackward', () => $('#pVolta').click());
  navigator.mediaSession.setActionHandler('seekforward', () => $('#pAvanca').click());
}

/* ---------- leitor de PDF ---------- */
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
let pdf = null, pdfItem = null, zoom = 1, observador = null, ratio = 1.41, ultPag = 0;

async function abrePdf(i) {
  pdfItem = i; ultPag = 0; $('#lOffBox').innerHTML = btnOffFm(i, 'p'); $('#leitor').hidden = false; document.body.style.overflow = 'hidden';
  $('#lTitulo').textContent = i.ti; $('#lPag').textContent = i.au || ''; $('#lPaginas').innerHTML = ''; $('#lProg').style.width = '0';
  $('#lCarrega').hidden = false; $('#lCarrega').textContent = 'Carregando…';
  zoom = le('zoom') || 1;
  try {
    pdf = await pdfjsLib.getDocument({ data: await dadosLivro(i, 'p') }).promise;
    const v = (await pdf.getPage(1)).getViewport({ scale: 1 }); ratio = v.height / v.width;
    $('#lCarrega').hidden = true;
    montaPaginas();
    const salva = le('pag:' + i.id);
    if (salva > 1) setTimeout(() => { const el = $('#lPaginas').children[salva - 1]; if (el) el.scrollIntoView(); }, 60);
  } catch (e) { $('#lCarrega').textContent = 'Não consegui abrir este PDF: ' + e.message; }
}
function paginaAtual() {
  const box = $('#lPaginas'), topo = box.getBoundingClientRect().top + 40;
  for (const el of box.children) if (el.getBoundingClientRect().bottom > topo) {
    const n = +el.dataset.n; $('#lPag').textContent = 'Página ' + n + ' de ' + pdf.numPages; $('#lProg').style.width = n / pdf.numPages * 100 + '%'; if (n !== ultPag) { ultPag = n; guarda('pag:' + pdfItem.id, n); guarda('qdo:' + pdfItem.id, Date.now()); } return;
  }
}
function montaPaginas() {
  const box = $('#lPaginas'); box.innerHTML = '';
  if (observador) observador.disconnect();
  const larg = Math.min(box.clientWidth * zoom, 1400);
  observador = new IntersectionObserver(es => es.forEach(en => { const el = en.target; if (en.isIntersecting && !el.dataset.ok) desenhaPagina(el, +el.dataset.n, larg); }),
    { root: box, rootMargin: '1200px 0px' });
  for (let n = 1; n <= pdf.numPages; n++) {
    const d = document.createElement('div'); d.className = 'pg'; d.dataset.n = n;
    d.style.width = larg + 'px'; d.style.height = Math.round(larg * ratio) + 'px';
    box.appendChild(d); observador.observe(d);
  }
  box.onscroll = paginaAtual; paginaAtual();
}
async function desenhaPagina(el, n, larg) {
  el.dataset.ok = '1';
  const p = await pdf.getPage(n);
  const base = p.getViewport({ scale: 1 }), dpr = Math.min(window.devicePixelRatio || 1, 2);
  const v = p.getViewport({ scale: larg / base.width * dpr });
  const c = document.createElement('canvas'); c.width = v.width; c.height = v.height;
  el.style.height = Math.round(v.height / dpr) + 'px';
  await p.render({ canvasContext: c.getContext('2d'), viewport: v }).promise;
  el.innerHTML = ''; el.appendChild(c);
}
$('#lVoltar').onclick = () => { $('#leitor').hidden = true; document.body.style.overflow = ''; if (pdf) { pdf.destroy(); pdf = null; } desenha(); };
$('#lMais').onclick = () => { zoom = Math.min(3, zoom + 0.25); guarda('zoom', zoom); montaPaginas(); };
$('#lMenos').onclick = () => { zoom = Math.max(0.5, zoom - 0.25); guarda('zoom', zoom); montaPaginas(); };


/* ---------- leitor de EPUB (texto que se ajusta à tela) ---------- */
let eLivro = null, eRend = null, eItem = null, eTamanho = le('etam') || 110, eTema = le('etema') || 'papel', eToc = [];
const TEMAS_E = { papel: { bg: '#f6ecd2', fg: '#2b1f16' }, sepia: { bg: '#e7cf9a', fg: '#3a2412' }, noite: { bg: '#17120e', fg: '#e9d9b6' } };

function temaE(nome) {
  eTema = nome; guarda('etema', nome);
  $('#leitorE').dataset.tema = nome;
  if (eRend) eRend.themes.select(nome);
}
// as bibliotecas do leitor de EPUB só são baixadas na primeira vez que um EPUB é aberto
const carregaScript = url => new Promise((ok, falha) => { const s = document.createElement('script'); s.src = url; s.onload = ok; s.onerror = () => falha(new Error('não consegui carregar ' + url)); document.head.appendChild(s); });
async function garanteEpubJs() {
  if (window.ePub) return;
  if (!window.JSZip) await carregaScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
  await carregaScript('https://cdn.jsdelivr.net/npm/epubjs@0.3.93/dist/epub.min.js');
}
async function abreEpub(i, urlLocal) {
  eItem = i; $('#eOffBox').innerHTML = btnOffFm(i, 'e'); $('#leitorE').hidden = false; document.body.style.overflow = 'hidden';
  $('#eTitulo').textContent = i.ti; $('#eCap').textContent = i.au || ''; $('#ePct').textContent = '0%'; $('#eProg').style.width = '0';
  $('#eCarrega').hidden = false; $('#eCarrega').textContent = 'Carregando…'; $('#eIndice').hidden = true;
  $('#leitorE').classList.remove('imersivo'); $('#leitorE').dataset.tema = eTema;
  try {
    await garanteEpubJs();
    const dados = urlLocal ? await (await fetch(urlLocal)).arrayBuffer() : await dadosLivro(i, 'e');
    eLivro = ePub(dados);
    eRend = eLivro.renderTo('eArea', { width: '100%', height: '100%', flow: 'paginated', spread: 'none', minSpreadWidth: 99999 });
    Object.entries(TEMAS_E).forEach(([k, t]) => eRend.themes.register(k, {
      html: { background: t.bg + ' !important' },
      body: { background: t.bg + ' !important', color: t.fg + ' !important', 'line-height': '1.6 !important', 'font-family': 'Georgia,"Times New Roman",serif !important', padding: '0 !important' },
      'p,div,span,li,h1,h2,h3,h4,i,b,em,strong,a,blockquote': { color: t.fg + ' !important' }
    }));
    eRend.themes.select(eTema); eRend.themes.fontSize(eTamanho + '%');
    eRend.on('relocated', loc => aoMudarE(loc));
    await eRend.display(le('epos:' + i.id) || undefined);
    $('#eCarrega').hidden = true;
    eLivro.loaded.navigation.then(nav => { eToc = nav.toc || []; montaIndiceE(); });
    const salvo = le('eloc:' + i.id);
    if (salvo) eLivro.locations.load(salvo);
    else eLivro.ready.then(() => eLivro.locations.generate(1400)).then(() => { guarda('eloc:' + i.id, eLivro.locations.save()); if (eRend && eRend.currentLocation()) aoMudarE(eRend.currentLocation()); }).catch(() => {});
  } catch (e) { $('#eCarrega').textContent = 'Não consegui abrir este livro: ' + e.message; }
}
function aoMudarE(loc) {
  if (!loc || !loc.start || !eItem) return;
  guarda('epos:' + eItem.id, loc.start.cfi);
  let pct = 0;
  try {
    if (eLivro.locations && eLivro.locations.length()) pct = eLivro.locations.percentageFromCfi(loc.start.cfi) * 100;
    else pct = (loc.start.index / Math.max(1, eLivro.spine.length)) * 100;
  } catch (e) {}
  pct = Math.max(0, Math.min(100, pct));
  $('#ePct').textContent = Math.round(pct) + '%'; $('#eProg').style.width = pct + '%';
  if (pct > 0) { guarda('pct:' + eItem.id, Math.round(pct * 10) / 10); guarda('qdo:' + eItem.id, Date.now()); }
  const href = (loc.start.href || '').split('#')[0];
  const cap = eToc.find(t => (t.href || '').split('#')[0] === href);
  if (cap) $('#eCap').textContent = (cap.label || '').trim();
  document.querySelectorAll('#eIndLista button').forEach(b => b.classList.toggle('atual', (b.dataset.href || '').split('#')[0] === href));
}
function montaIndiceE() {
  $('#eIndLista').innerHTML = eToc.map(t => '<button data-href="' + esc(t.href) + '">' + esc((t.label || '').trim()) + '</button>').join('') || '<p class="vazio">Este livro não tem índice.</p>';
}
$('#eIndLista').addEventListener('click', e => { const b = e.target.closest('button'); if (b && eRend) { eRend.display(b.dataset.href); $('#eIndice').hidden = true; } });
$('#eTocBtn').onclick = () => { $('#eIndice').hidden = !$('#eIndice').hidden; };
$('#eIndFecha').onclick = () => { $('#eIndice').hidden = true; };
$('#eTema').onclick = () => { const ks = Object.keys(TEMAS_E); temaE(ks[(ks.indexOf(eTema) + 1) % ks.length]); };
const tamanhoE = d => { eTamanho = Math.max(70, Math.min(220, eTamanho + d)); guarda('etam', eTamanho); if (eRend) eRend.themes.fontSize(eTamanho + '%'); };
$('#eMais').onclick = () => tamanhoE(10); $('#eMenos').onclick = () => tamanhoE(-10);
$('#eVoltar').onclick = () => {
  try { if (eRend) eRend.destroy(); if (eLivro) eLivro.destroy(); } catch (e) {}
  eRend = eLivro = eItem = null; eToc = [];
  $('#leitorE').hidden = true; document.body.style.overflow = ''; desenha();
};
// viragem de pagina: toque nas bordas, deslize, ou setas do teclado; toque no centro esconde/mostra as barras
let eTx = 0, eDeslizou = false;
$('#eToque').addEventListener('touchstart', e => { eTx = e.changedTouches[0].clientX; eDeslizou = false; }, { passive: true });
$('#eToque').addEventListener('touchend', e => {
  const dx = e.changedTouches[0].clientX - eTx;
  if (Math.abs(dx) > 50 && eRend) { eDeslizou = true; dx < 0 ? eRend.next() : eRend.prev(); }
}, { passive: true });
$('#eToque').addEventListener('click', e => {
  if (eDeslizou) { eDeslizou = false; return; }
  if (!eRend) return;
  const r = e.currentTarget.getBoundingClientRect(), x = (e.clientX - r.left) / r.width;
  if (x < 0.3) eRend.prev(); else if (x > 0.7) eRend.next();
  else { $('#leitorE').classList.toggle('imersivo'); setTimeout(() => { try { eRend.resize(); } catch (er) {} }, 80); }
});
document.addEventListener('keydown', e => {
  if ($('#leitorE').hidden || !eRend) return;
  if (e.key === 'ArrowRight' || e.key === 'PageDown') eRend.next(); else if (e.key === 'ArrowLeft' || e.key === 'PageUp') eRend.prev();
});

/* ---------- partida ---------- */
(async function main() {
  $('#versao').textContent = 'Minha Biblioteca · versão ' + VERSAO;
  if ('serviceWorker' in navigator) {
    const tinhaControle = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('sw.js').catch(() => {});
    let recarregou = false;                                   // nova versão assumiu o controle: recarrega uma vez
    navigator.serviceWorker.addEventListener('controllerchange', () => { if (tinhaControle && !recarregou) { recarregou = true; location.reload(); } });
    document.addEventListener('visibilitychange', () => { if (!document.hidden) navigator.serviceWorker.getRegistration().then(r => r && r.update()).catch(() => {}); });
  }
  $('#oi').textContent = saudacao();
  try {
    await iniciaLogin();
    $('#btnConta').textContent = DEMO ? 'D' : conta ? (conta.name || conta.username || 'U')[0].toUpperCase() : '→';
    if (!DEMO && !conta) {
      $('#cats').hidden = true; $('#nav').hidden = true;
      aviso('<b>Bem-vindo!</b> Toque no botão redondo no canto superior para entrar com sua conta Microsoft e abrir a biblioteca.');
      return;
    }
    await carregaCatalogo();
    await reconciliaBaixados();
    if (DEMO) aviso('<b>Modo demonstração.</b> Esta é a lista real dos seus títulos; os arquivos serão ligados ao OneDrive depois do registro no Azure.');
    desenha();
  } catch (e) {
    if (e.status === 404) aviso('<b>Login feito' + (conta ? ' como ' + esc(conta.username) : '') + '!</b><br>Mas ainda não encontrei a pasta <b>' + esc(CFG.pasta) + '</b> (com o arquivo <b>catalogo.json</b>) no seu OneDrive. Envie a pasta e abra o app de novo.');
    else aviso('<b>Erro ao iniciar:</b> ' + esc(e.message));
  }
})();
