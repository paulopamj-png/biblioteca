'use strict';
/* Minha Biblioteca — PWA: audiobooks e PDFs do OneDrive (Microsoft Graph + MSAL) */
const $ = s => document.querySelector(s);
const CFG = window.CONFIG;
const DEMO = !CFG.clientId || /[?&]demo=1/.test(location.search);   // ?demo=1 mostra a lista sem login (teste visual)
const GRAPH = 'https://graph.microsoft.com/v1.0';
const ESCOPOS = ['Files.Read', 'User.Read'];
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
  catch (e) { await clienteMsal.acquireTokenRedirect({ scopes: ESCOPOS }); throw e; }
}
async function graph(caminho, opts = {}) {
  const t = await token();
  const r = await fetch(GRAPH + caminho, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } });
  if (!r.ok) { const e = new Error('OneDrive respondeu ' + r.status); e.status = r.status; throw e; }
  return r;
}
const caminhoItem = rel => '/me/drive/root:/' + [CFG.pasta, ...rel.split('/')].map(encodeURIComponent).join('/');

/* ---------- catálogo ---------- */
async function carregaCatalogo() {
  if (DEMO) { catalogo = await (await fetch('catalogo.json')).json(); return; }
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
function capa(i, prog) {
  return '<div class="capa">' + foto(i) + (prog ? '<i class="prog" style="width:' + prog + '%"></i>' : '') + '</div>';
}
function progresso(i) {
  if (i.t === 'a') { const p = le('prog:' + i.id); return p && i.d ? Math.min(100, p.t / (i.d * 60) * 100) : 0; }
  if (i.fm === 'e') return le('pct:' + i.id) || 0;
  const n = le('pag:' + i.id); return n > 1 && i.pg ? Math.min(100, n / i.pg * 100) : 0;
}
function cartao(i, atraso) {
  return '<button class="card" data-id="' + i.id + '" style="animation-delay:' + (atraso || 0) + 'ms">' + capa(i, progresso(i)) +
    '<span class="nm">' + esc(i.ti) + '</span><span class="au">' + esc(i.au || 'Autor não informado') + '</span></button>';
}
function linha(i, atraso, status) {
  const meta = status || (i.t === 'a' ? 'Audiobook' + (i.d ? ' · ' + fmtDur(i.d) : '') : 'Livro' + (i.pg ? ' · ' + i.pg + ' págs' : ''));
  const pr = status ? progresso(i) : 0;
  return '<button class="linha ' + i.t + '" data-id="' + i.id + '" style="animation-delay:' + Math.min(atraso || 0, 300) + 'ms">' + capa(i, 0) +
    '<span class="dd"><b>' + esc(i.ti) + '</b><small>' + esc(i.au || 'Autor não informado') + '</small><span class="mt">' + meta + '</span>' +
    (pr ? '<span class="barra"><i style="width:' + pr + '%"></i></span>' : '') + '</span></button>';
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
    '<span class="ab-play">' + IC.play + '</span></button>';
}
const item = (i, atraso, status) => i.t === 'a' ? caixa(i, atraso, status) : linha(i, atraso, status);

/* ---------- o que o leitor já começou ---------- */
function iniciados() {
  const out = [];
  for (const i of catalogo) {
    if (i.t === 'a') { const p = le('prog:' + i.id); if (p && p.t > 5) out.push([i, p.quando || 0]); }
    else if (i.fm === 'e') { if ((le('pct:' + i.id) || 0) > 0) out.push([i, le('qdo:' + i.id) || 1]); }
    else { const n = le('pag:' + i.id); if (n > 1) out.push([i, le('qdo:' + i.id) || 1]); }
  }
  return out.sort((a, b) => b[1] - a[1]).map(x => x[0]);
}
function statusIni(i) {
  const pr = Math.round(progresso(i));
  if (i.t === 'a') return 'Parou em ' + fmtTempo(le('prog:' + i.id).t) + (pr ? ' · ' + pr + '%' : '');
  if (i.fm === 'e') return 'Leitura em ' + pr + '%';
  return 'Página ' + le('pag:' + i.id) + (i.pg ? ' de ' + i.pg : '') + (pr ? ' · ' + pr + '%' : '');
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
    (ini.length ? '<div class="continua">' + ini.map((i, n) => item(i, n * 50, statusIni(i))).join('') + '</div>'
                : '<div class="vazio-ini">Você ainda não começou nenhum título. Escolha um nas sugestões abaixo e ele aparecerá aqui, com o ponto em que você parou.</div>');
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
  const ir = t.closest('[data-ir]');
  if (ir) {
    vibra(); const dest = ir.dataset.ir;
    ir.classList.add('abre');
    setTimeout(() => { ir.classList.remove('abre'); vaiPara(aba === dest && !consulta ? 'home' : dest, !!consulta); }, 200);
    return;
  }
  if (t.closest('#nav .ini')) { vibra(); vaiPara('home'); return; }
  const so = t.closest('[data-sort]'); if (so) { vibra(); sugestoes(so.dataset.sort); return; }
  const it = t.closest('.card, .linha, .caixa'); if (it) { vibra(); abre(porId(it.dataset.id)); }
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
  if (DEMO) return alert('Modo demonstração: o login Microsoft será ativado depois do registro no Azure.');
  if (conta) { if (confirm('Sair da conta ' + conta.username + '?')) clienteMsal.logoutRedirect({ account: conta }); }
  else clienteMsal.loginRedirect({ scopes: ESCOPOS });
};

function abre(i) {
  if (!i) return;
  if (DEMO) { alert('Modo demonstração: os arquivos ainda não estão ligados ao OneDrive.\n\n' + i.ti + (i.au ? '\n' + i.au : '')); return; }
  if (i.t === 'a') tocaAudio(i); else if (i.fm === 'e') abreEpub(i); else abrePdf(i);
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
function pintaPlayer(i) {
  $('#sCapa').innerHTML = foto(i);
  $('#sTitulo').textContent = i.ti; $('#sAutor').textContent = i.au || '';
  $('#mCapa').innerHTML = foto(i);
  $('#mTitulo').textContent = i.ti; $('#mAutor').textContent = i.au || '';
}
async function tocaAudio(i) {
  if (atual && atual.id === i.id) { abreSheet(true); return; }
  salvaProgresso(true);
  atual = i; pintaPlayer(i);
  $('#mini').hidden = false; document.body.classList.add('com-mini');
  setPlay(false); abreSheet(true);
  try {
    au.src = await urlAudio(i);
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
  if (!atual) return;
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
au.addEventListener('ended', () => { if (atual) guarda('prog:' + atual.id, { t: 0, quando: Date.now() }); });
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
$('#sFecha').onclick = () => abreSheet(false);
$('#mAbre').onclick = () => abreSheet(true);
$('#pFechar').onclick = () => {
  salvaProgresso(true); au.pause(); au.removeAttribute('src'); atual = null; limpaSono();
  abreSheet(false); $('#mini').hidden = true; document.body.classList.remove('com-mini'); desenha();
};
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
  pdfItem = i; ultPag = 0; $('#leitor').hidden = false; document.body.style.overflow = 'hidden';
  $('#lTitulo').textContent = i.ti; $('#lPag').textContent = i.au || ''; $('#lPaginas').innerHTML = ''; $('#lProg').style.width = '0';
  $('#lCarrega').hidden = false; $('#lCarrega').textContent = 'Carregando…';
  zoom = le('zoom') || 1;
  try {
    const r = await graph(caminhoItem('livros/' + i.f) + ':/content');
    pdf = await pdfjsLib.getDocument({ data: await r.arrayBuffer() }).promise;
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
  eItem = i; $('#leitorE').hidden = false; document.body.style.overflow = 'hidden';
  $('#eTitulo').textContent = i.ti; $('#eCap').textContent = i.au || ''; $('#ePct').textContent = '0%'; $('#eProg').style.width = '0';
  $('#eCarrega').hidden = false; $('#eCarrega').textContent = 'Carregando…'; $('#eIndice').hidden = true;
  $('#leitorE').classList.remove('imersivo'); $('#leitorE').dataset.tema = eTema;
  try {
    await garanteEpubJs();
    const dados = urlLocal ? await (await fetch(urlLocal)).arrayBuffer() : await (await graph(caminhoItem('livros/' + i.f) + ':/content')).arrayBuffer();
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
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
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
    if (DEMO) aviso('<b>Modo demonstração.</b> Esta é a lista real dos seus títulos; os arquivos serão ligados ao OneDrive depois do registro no Azure.');
    desenha();
  } catch (e) {
    if (e.status === 404) aviso('<b>Login feito' + (conta ? ' como ' + esc(conta.username) : '') + '!</b><br>Mas ainda não encontrei a pasta <b>' + esc(CFG.pasta) + '</b> (com o arquivo <b>catalogo.json</b>) no seu OneDrive. Envie a pasta e abra o app de novo.');
    else aviso('<b>Erro ao iniciar:</b> ' + esc(e.message));
  }
})();
