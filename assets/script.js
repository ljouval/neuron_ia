// ═══════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════
let rawText = '', words = [], vocab = [], wordIndex = {}, wordCounts = {};
let embedDim = 8, embeddings = null;
let W1 = null, b1 = null, W2 = null, b2 = null;
let trained = false, lossHistory = [];
let trainWorker = null;

// Options de visualisation
const vizState = { weights: false, embeds: false, signal: false };
// Signal animation state
let signalAnim = null; // { phase, inputIdx, hiddenActs, outputProbs, step }
let signalAnimFrame = null;

// ═══════════════════════════════════════════════════════════════
// NOTIFICATIONS
// ═══════════════════════════════════════════════════════════════
function notify(msg, icon = '✅', duration = 3000) {
  const el = document.getElementById('notif');
  document.getElementById('notifText').textContent = msg;
  document.getElementById('notifIcon').textContent = icon;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), duration);
}

// ═══════════════════════════════════════════════════════════════
// VIZ TOGGLES
// ═══════════════════════════════════════════════════════════════
['togWeights', 'togEmbeds', 'togSignal'].forEach(id => {
  const el = document.getElementById(id);
  if (!el) return;
  el.addEventListener('click', () => {
    const key = id === 'togWeights' ? 'weights' : id === 'togEmbeds' ? 'embeds' : 'signal';
    vizState[key] = !vizState[key];
    document.getElementById(id).classList.toggle('active', vizState[key]);

    if (key === 'weights') {
      document.getElementById('vizWeightsPanel').style.display = vizState.weights ? 'block' : 'none';
      if (vizState.weights && W1) renderWeightMaps();
    }
    if (key === 'embeds') {
      document.getElementById('vizEmbedsPanel').style.display = vizState.embeds ? 'block' : 'none';
      if (vizState.embeds && embeddings) renderLiveEmbedScatter();
    }
    if (key === 'signal') {
      const title = document.getElementById('networkVizTitle');
      title.textContent = vizState.signal ? 'Réseau actif — propagation en cours' : 'Architecture du réseau';
      if (!vizState.signal && signalAnimFrame) { cancelAnimationFrame(signalAnimFrame); signalAnimFrame = null; }
      if (vizState.signal && W1) startSignalAnimation();
    }
  });
});
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', e => { e.preventDefault(); dropzone.classList.remove('dragover'); handleFile(e.dataTransfer.files[0]); });
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

function setStatus(msg, type = '') {
  document.getElementById('fileStatus').style.display = 'flex';
  document.getElementById('statusText').textContent = msg;
  document.getElementById('statusDot').className = 'status-dot ' + type;
}
function setProgress(pct) { document.getElementById('progressFill').style.width = pct + '%'; }

async function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['txt','pdf','docx','pptx','csv','xlsx'].includes(ext)) { notify('Format non supporté', '❌', 4000); return; }
  setStatus('Lecture du fichier…', 'processing'); setProgress(10);
  try {
    if      (ext === 'txt')  rawText = await readTxt(file);
    else if (ext === 'pdf')  rawText = await readPdf(file);
    else if (ext === 'docx') rawText = await readDocx(file);
    else if (ext === 'pptx') rawText = await readPptx(file);
    else if (ext === 'csv')  rawText = await readCsv(file);
    else if (ext === 'xlsx') rawText = await readXlsx(file);
    setProgress(70); setStatus('Traitement…', 'processing');
    processText(); setProgress(100);
    setStatus('✓ ' + file.name + ' · ' + words.length + ' mots · ' + vocab.length + ' mots uniques', '');
    document.getElementById('statusDot').style.background = 'var(--accent3)';
    showStats();
    document.getElementById('sectionEmbed').classList.add('visible');
    setStep(2);
    trained = false;
    document.getElementById('sectionTrain').classList.remove('visible');
    document.getElementById('sectionGenerate').classList.remove('visible');
    notify('Document chargé : ' + vocab.length + ' mots uniques', '📄');
  } catch (err) {
    setStatus('Erreur : ' + err.message, 'error');
    notify('Erreur : ' + err.message, '❌', 6000);
  }
}

// ═══════════════════════════════════════════════════════════════
// PARSEURS
// ═══════════════════════════════════════════════════════════════
function readTxt(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result);
    r.onerror = rej;
    r.readAsText(file, 'UTF-8');
  });
}
async function readPdf(file) {
  if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js non chargé');
  const ab = await file.arrayBuffer();
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(s => s.str).join(' ') + ' ';
    setProgress(10 + (i / pdf.numPages) * 55);
  }
  return text;
}
async function readDocx(file) {
  if (typeof mammoth === 'undefined') throw new Error('Mammoth.js non chargé');
  const ab = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: ab });
  return result.value;
}
async function readPptx(file) {
  if (typeof JSZip === 'undefined') throw new Error('JSZip non chargé');
  const ab = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(ab);
  let text = '';
  const slideFiles = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort();
  if (slideFiles.length === 0) throw new Error('Aucune slide trouvée');
  for (let i = 0; i < slideFiles.length; i++) {
    const xml = await zip.files[slideFiles[i]].async('string');
    const matches = xml.match(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g) || [];
    text += matches.map(m => m.replace(/<[^>]+>/g, '')).join(' ') + ' ';
  }
  return text;
}
async function readCsv(file) {
  if (typeof XLSX === 'undefined') throw new Error('SheetJS non chargé');
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  return data.flat().filter(v => typeof v === 'string' && v.trim()).join(' ');
}
async function readXlsx(file) {
  if (typeof XLSX === 'undefined') throw new Error('SheetJS non chargé');
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: 'array' });
  let text = '';
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    text += data.flat().filter(v => typeof v === 'string' && v.trim()).join(' ') + ' ';
  });
  return text;
}

// ═══════════════════════════════════════════════════════════════
// TEXT PROCESSING
// ═══════════════════════════════════════════════════════════════
function processText() {
  const cleaned = rawText.toLowerCase().replace(/[^\p{L}\p{N}\s'-]/gu, ' ').replace(/\s+/g, ' ').trim();
  words = cleaned.split(' ').filter(w => w.length > 1);
  wordCounts = {};
  words.forEach(w => wordCounts[w] = (wordCounts[w] || 0) + 1);
  // Limiter le vocabulaire aux 500 mots les plus fréquents pour la performance
  const sorted = Object.entries(wordCounts).sort((a, b) => b[1] - a[1]);
  const topVocab = sorted.slice(0, 500).map(([w]) => w);
  vocab = topVocab.sort();
  wordIndex = {};
  vocab.forEach((w, i) => wordIndex[w] = i);
  embeddings = null; W1 = null; b1 = null; W2 = null; b2 = null;
  trained = false; lossHistory = [];
}

// ═══════════════════════════════════════════════════════════════
// STATS
// ═══════════════════════════════════════════════════════════════
function showStats() {
  const sentences = rawText.split(/[.!?]+/).filter(s => s.trim().length > 3).length;
  const avgLen = words.length ? (words.map(w => w.length).reduce((a, b) => a + b, 0) / words.length).toFixed(1) : 0;
  document.getElementById('statsGrid').innerHTML = `
    <div class="stat"><div class="stat-value">${words.length.toLocaleString()}</div><div class="stat-label">Mots totaux</div></div>
    <div class="stat"><div class="stat-value">${vocab.length.toLocaleString()}</div><div class="stat-label">Vocab. réseau</div></div>
    <div class="stat"><div class="stat-value">${sentences.toLocaleString()}</div><div class="stat-label">Phrases ~</div></div>
    <div class="stat"><div class="stat-value">${avgLen}</div><div class="stat-label">Long. moy.</div></div>`;
}

// ═══════════════════════════════════════════════════════════════
// MATHS UTILITAIRES
// ═══════════════════════════════════════════════════════════════
function randn() {
  // Box-Muller
  let u = 0, v = 0;
  while (u === 0) u = Math.random();
  while (v === 0) v = Math.random();
  return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
}

function randMatrix(rows, cols, scale = 0.1) {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => randn() * scale));
}

function randVector(size, scale = 0.1) {
  return Array.from({ length: size }, () => randn() * scale);
}

function matVec(M, v) {
  return M.map(row => row.reduce((s, w, j) => s + w * v[j], 0));
}

function addVec(a, b) { return a.map((x, i) => x + b[i]); }

function tanhVec(v) { return v.map(x => Math.tanh(x)); }

function softmax(v) {
  const max = Math.max(...v);
  const exps = v.map(x => Math.exp(x - max));
  const sum = exps.reduce((a, b) => a + b, 0);
  return exps.map(x => x / sum);
}

// ═══════════════════════════════════════════════════════════════
// EMBEDDINGS — INIT & VISUALISATION
// ═══════════════════════════════════════════════════════════════
document.getElementById('btnInitEmbeds').addEventListener('click', initEmbeddings);

function initEmbeddings() {
  if (!vocab.length) { notify('Importez d\'abord un document', '⚠️'); return; }
  embedDim = parseInt(document.getElementById('embedDim').value);
  const topN = Math.min(parseInt(document.getElementById('embedTopN').value) || 30, vocab.length);

  // Initialisation aléatoire (Xavier)
  const scale = 1 / Math.sqrt(embedDim);
  embeddings = vocab.map(() => randVector(embedDim, scale));

  // Init poids réseau
  const hidden = 32;
  W1 = randMatrix(hidden, embedDim, scale);
  b1 = randVector(hidden, 0);
  W2 = randMatrix(vocab.length, hidden, 0.01);
  b2 = randVector(vocab.length, 0);

  trained = false; lossHistory = [];

  document.getElementById('embedDimLabel').textContent = embedDim;
  document.getElementById('embedSection').style.display = 'block';

  renderEmbedTable(topN);
  renderEmbedScatter(topN);

  document.getElementById('sectionTrain').classList.add('visible');
  document.getElementById('btnReset').disabled = false;
  setStep(3);
  notify('Embeddings initialisés — ' + embedDim + ' dimensions × ' + vocab.length + ' mots', '🧮');
}

function renderEmbedTable(topN) {
  const top = Object.entries(wordCounts).sort((a, b) => b[1] - a[1]).slice(0, topN).map(([w]) => w).filter(w => wordIndex[w] !== undefined);
  let html = '<table class="embed-table"><thead><tr><th>Mot</th>';
  for (let d = 0; d < embedDim; d++) html += `<th>d${d}</th>`;
  html += '</tr></thead><tbody>';
  top.forEach(w => {
    const idx = wordIndex[w];
    const vec = embeddings[idx];
    html += `<tr><td>${w}</td>`;
    vec.forEach(v => {
      const cls = v > 0.05 ? 'embed-val-pos' : v < -0.05 ? 'embed-val-neg' : 'embed-val-zero';
      html += `<td class="${cls}">${v.toFixed(3)}</td>`;
    });
    html += '</tr>';
  });
  document.getElementById('embedTable').innerHTML = html + '</tbody></table>';
}

function renderEmbedScatter(topN) {
  const canvas = document.getElementById('embedCanvas');
  const wrap = canvas.parentElement;
  const W = wrap.clientWidth - 32 || 660;
  const H = 340;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');

  const top = Object.entries(wordCounts).sort((a, b) => b[1] - a[1]).slice(0, topN).map(([w]) => w).filter(w => wordIndex[w] !== undefined);
  if (top.length < 2) return;

  // PCA simplifiée : 2 premières composantes par méthode des puissances
  const vecs = top.map(w => embeddings[wordIndex[w]]);
  const mean = Array(embedDim).fill(0);
  vecs.forEach(v => v.forEach((x, i) => mean[i] += x / vecs.length));
  const centered = vecs.map(v => v.map((x, i) => x - mean[i]));

  // Vecteur propre 1 : itération de puissance
  let pc1 = randVector(embedDim, 1);
  for (let iter = 0; iter < 50; iter++) {
    const proj = centered.map(v => v.reduce((s, x, i) => s + x * pc1[i], 0));
    const newPc = Array(embedDim).fill(0);
    centered.forEach((v, k) => v.forEach((x, i) => newPc[i] += proj[k] * x));
    const norm = Math.sqrt(newPc.reduce((s, x) => s + x * x, 0)) || 1;
    pc1 = newPc.map(x => x / norm);
  }
  // Vecteur propre 2 : orthogonel à pc1
  let pc2 = randVector(embedDim, 1);
  for (let iter = 0; iter < 50; iter++) {
    // déflation
    const dot = pc2.reduce((s, x, i) => s + x * pc1[i], 0);
    pc2 = pc2.map((x, i) => x - dot * pc1[i]);
    const proj = centered.map(v => v.reduce((s, x, i) => s + x * pc2[i], 0));
    const newPc = Array(embedDim).fill(0);
    centered.forEach((v, k) => v.forEach((x, i) => newPc[i] += proj[k] * x));
    const norm = Math.sqrt(newPc.reduce((s, x) => s + x * x, 0)) || 1;
    pc2 = newPc.map(x => x / norm);
  }

  const coords = centered.map(v => ({
    x: v.reduce((s, x, i) => s + x * pc1[i], 0),
    y: v.reduce((s, x, i) => s + x * pc2[i], 0)
  }));

  const xs = coords.map(c => c.x), ys = coords.map(c => c.y);
  const minX = Math.min(...xs), maxX = Math.max(...xs);
  const minY = Math.min(...ys), maxY = Math.max(...ys);
  const PAD = 40;
  const toScreen = c => ({
    sx: PAD + (c.x - minX) / (maxX - minX + 1e-9) * (W - 2 * PAD),
    sy: PAD + (1 - (c.y - minY) / (maxY - minY + 1e-9)) * (H - 2 * PAD)
  });

  ctx.clearRect(0, 0, W, H);
  ctx.fillStyle = 'rgba(28,28,40,0)';
  ctx.fillRect(0, 0, W, H);

  // Axes
  ctx.strokeStyle = 'rgba(42,42,58,0.6)'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(PAD, H / 2); ctx.lineTo(W - PAD, H / 2); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(W / 2, PAD); ctx.lineTo(W / 2, H - PAD); ctx.stroke();

  // Points & labels
  coords.forEach((c, k) => {
    const { sx, sy } = toScreen(c);
    const freq = wordCounts[top[k]] || 1;
    const r = Math.max(3, Math.min(7, 2 + Math.log(freq)));
    ctx.beginPath();
    ctx.arc(sx, sy, r, 0, 2 * Math.PI);
    ctx.fillStyle = 'rgba(109,255,204,0.7)';
    ctx.fill();
    ctx.font = `${Math.round(W * 0.013)}px JetBrains Mono, monospace`;
    ctx.fillStyle = 'rgba(168,168,192,0.9)';
    ctx.fillText(top[k], sx + r + 2, sy + 4);
  });

  ctx.fillStyle = 'rgba(109,255,204,0.4)';
  ctx.font = `${Math.round(W * 0.011)}px JetBrains Mono, monospace`;
  ctx.fillText('PC1 →', W - PAD - 30, H / 2 - 6);
  ctx.fillText('PC2 ↑', W / 2 + 6, PAD + 12);
}

// ═══════════════════════════════════════════════════════════════
// FORWARD PASS
// ═══════════════════════════════════════════════════════════════
function forward(wordIdx) {
  const e = embeddings[wordIdx];
  const h_raw = addVec(matVec(W1, e), b1);
  const h = tanhVec(h_raw);
  const logits = addVec(matVec(W2, h), b2);
  const probs = softmax(logits);
  return { e, h_raw, h, logits, probs };
}

// ═══════════════════════════════════════════════════════════════
// BACKWARD PASS (SGD)
// ═══════════════════════════════════════════════════════════════
function backward(wordIdx, targetIdx, lr) {
  const { e, h_raw, h, probs } = forward(wordIdx);
  const loss = -Math.log(probs[targetIdx] + 1e-12);

  // dL/dlogits
  const dLogits = probs.map((p, i) => i === targetIdx ? p - 1 : p);

  // dL/dW2, dL/db2
  for (let i = 0; i < W2.length; i++) {
    for (let j = 0; j < W2[i].length; j++) W2[i][j] -= lr * dLogits[i] * h[j];
    b2[i] -= lr * dLogits[i];
  }

  // dL/dh
  const dH = Array(h.length).fill(0);
  for (let j = 0; j < h.length; j++)
    for (let i = 0; i < dLogits.length; i++) dH[j] += dLogits[i] * W2[i][j];

  // dL/dh_raw (tanh')
  const dH_raw = dH.map((d, j) => d * (1 - h[j] * h[j]));

  // dL/dW1, dL/db1
  for (let i = 0; i < W1.length; i++) {
    for (let j = 0; j < W1[i].length; j++) W1[i][j] -= lr * dH_raw[i] * e[j];
    b1[i] -= lr * dH_raw[i];
  }

  // dL/dembedding
  for (let j = 0; j < embedDim; j++) {
    let grad = 0;
    for (let i = 0; i < W1.length; i++) grad += dH_raw[i] * W1[i][j];
    embeddings[wordIdx][j] -= lr * grad;
  }

  return loss;
}

// ═══════════════════════════════════════════════════════════════
// ENTRAÎNEMENT
// ═══════════════════════════════════════════════════════════════
document.getElementById('btnTrain').addEventListener('click', startTraining);
document.getElementById('btnReset').addEventListener('click', resetModel);

function startTraining() {
  if (!embeddings) { notify('Initialisez d\'abord les embeddings', '⚠️'); return; }
  const btn = document.getElementById('btnTrain');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Entraînement…';

  const lr = parseFloat(document.getElementById('learningRate').value);
  const hiddenSize = parseInt(document.getElementById('hiddenSize').value);
  const iterations = Math.min(1000, parseInt(document.getElementById('trainIter').value) || 200);

  // Réinitialiser les poids si hiddenSize a changé
  const currentHidden = W1 ? W1.length : 0;
  if (currentHidden !== hiddenSize) {
    const scale = 1 / Math.sqrt(embedDim);
    W1 = randMatrix(hiddenSize, embedDim, scale);
    b1 = randVector(hiddenSize, 0);
    W2 = randMatrix(vocab.length, hiddenSize, 0.01);
    b2 = randVector(vocab.length, 0);
  }

  lossHistory = [];
  document.getElementById('trainSection').style.display = 'block';
  document.getElementById('sectionTrain').classList.add('visible');

  // Construire les paires d'entraînement à partir du texte
  const pairs = [];
  for (let k = 0; k < words.length - 1; k++) {
    const i = wordIndex[words[k]], j = wordIndex[words[k + 1]];
    if (i !== undefined && j !== undefined) pairs.push([i, j]);
  }
  if (pairs.length === 0) { notify('Pas assez de données', '⚠️'); btn.disabled = false; btn.innerHTML = 'Entraîner'; return; }

  let iter = 0;
  const batchSize = Math.min(50, pairs.length);
  const totalSteps = iterations;
  const vizRefreshEl = document.getElementById('vizRefresh');
  const vizRefresh = vizRefreshEl ? (parseInt(vizRefreshEl.value) || 15) : 15;

  function step() {
    if (iter >= totalSteps) {
      trained = true;
      btn.disabled = false; btn.innerHTML = 'Ré-entraîner';
      document.getElementById('sectionGenerate').classList.add('visible');
      setStep(4);
      const topN = Math.min(parseInt(document.getElementById('embedTopN').value) || 30, vocab.length);
      renderEmbedTable(topN);
      renderEmbedScatter(topN);
      renderNetworkViz(hiddenSize);
      if (vizState.weights) renderWeightMaps();
      if (vizState.embeds) renderLiveEmbedScatter();
      notify('Entraînement terminé ! ' + iter + ' itérations', '🎉', 4000);
      return;
    }

    // Mini-batch SGD
    let batchLoss = 0;
    let lastI = 0, lastJ = 0;
    for (let b = 0; b < batchSize; b++) {
      const [i, j] = pairs[Math.floor(Math.random() * pairs.length)];
      batchLoss += backward(i, j, lr);
      lastI = i; lastJ = j;
    }
    lossHistory.push(batchLoss / batchSize);
    iter++;

    if (iter % vizRefresh === 0 || iter === totalSteps) {
      renderLossChart();
      renderTrainStats(iter, totalSteps, batchLoss / batchSize);
      if (vizState.weights) renderWeightMaps();
      if (vizState.embeds) renderLiveEmbedScatter();
      if (vizState.signal) triggerSignalStep(lastI, lastJ);
    }
    if (iter % 5 === 0) setTimeout(step, 0);
    else step();
  }
  step();
}

function resetModel() {
  if (!vocab.length) return;
  const scale = 1 / Math.sqrt(embedDim);
  embeddings = vocab.map(() => randVector(embedDim, scale));
  const hidden = parseInt(document.getElementById('hiddenSize').value) || 32;
  W1 = randMatrix(hidden, embedDim, scale);
  b1 = randVector(hidden, 0);
  W2 = randMatrix(vocab.length, hidden, 0.01);
  b2 = randVector(vocab.length, 0);
  trained = false; lossHistory = [];
  document.getElementById('trainSection').style.display = 'none';
  document.getElementById('sectionGenerate').classList.remove('visible');
  const topN = Math.min(parseInt(document.getElementById('embedTopN').value) || 30, vocab.length);
  renderEmbedTable(topN);
  renderEmbedScatter(topN);
  notify('Modèle réinitialisé', '🔄');
}

// ═══════════════════════════════════════════════════════════════
// LOSS CHART
// ═══════════════════════════════════════════════════════════════
function renderLossChart() {
  const canvas = document.getElementById('lossCanvas');
  const wrap = canvas.parentElement;
  const W = (wrap.clientWidth - 32) || 460;
  const H = 200;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');
  const PAD = { top: 20, right: 20, bottom: 36, left: 52 };
  const cW = W - PAD.left - PAD.right;
  const cH = H - PAD.top - PAD.bottom;

  ctx.clearRect(0, 0, W, H);
  if (lossHistory.length < 2) return;

  const maxL = Math.max(...lossHistory);
  const minL = Math.min(...lossHistory);

  // Grille
  ctx.strokeStyle = 'rgba(42,42,58,0.7)'; ctx.lineWidth = 1;
  for (let g = 0; g <= 4; g++) {
    const y = PAD.top + cH * (g / 4);
    ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(PAD.left + cW, y); ctx.stroke();
    ctx.fillStyle = '#666680'; ctx.font = `${Math.round(W * 0.022)}px JetBrains Mono,monospace`;
    ctx.textAlign = 'right';
    ctx.fillText((maxL - (maxL - minL) * g / 4).toFixed(2), PAD.left - 4, y + 4);
  }

  // Courbe
  ctx.beginPath(); ctx.strokeStyle = 'rgba(109,255,204,0.9)'; ctx.lineWidth = 2;
  lossHistory.forEach((l, i) => {
    const x = PAD.left + (i / (lossHistory.length - 1)) * cW;
    const y = PAD.top + cH - ((l - minL) / (maxL - minL + 1e-9)) * cH;
    i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
  });
  ctx.stroke();

  // Axes labels
  ctx.fillStyle = '#666680'; ctx.font = `${Math.round(W * 0.022)}px JetBrains Mono,monospace`;
  ctx.textAlign = 'center';
  ctx.fillText('Itérations →', PAD.left + cW / 2, H - 6);
  ctx.textAlign = 'left';
  ctx.fillText('0', PAD.left, H - PAD.bottom + 14);
  ctx.fillText('' + lossHistory.length, PAD.left + cW - 10, H - PAD.bottom + 14);

  // Valeur courante
  ctx.fillStyle = 'rgba(109,255,204,0.9)';
  ctx.textAlign = 'right';
  ctx.font = `bold ${Math.round(W * 0.025)}px JetBrains Mono,monospace`;
  ctx.fillText('loss: ' + lossHistory[lossHistory.length - 1].toFixed(4), PAD.left + cW, PAD.top - 4);
}

// ═══════════════════════════════════════════════════════════════
// NETWORK VISUALISATION
// ═══════════════════════════════════════════════════════════════
function renderNetworkViz(hiddenSize) {
  const canvas = document.getElementById('networkCanvas');
  const wrap = canvas.parentElement;
  const W = (wrap.clientWidth - 32) || 300;
  const H = 200;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  const layers = [
    { label: 'Embedding', n: Math.min(embedDim, 6), color: 'rgba(124,109,255,0.8)', full: embedDim },
    { label: 'Cachée', n: Math.min(hiddenSize, 6), color: 'rgba(109,255,204,0.8)', full: hiddenSize },
    { label: 'Sortie', n: Math.min(vocab.length, 6), color: 'rgba(255,109,155,0.8)', full: vocab.length }
  ];

  const xPositions = [W * 0.18, W * 0.5, W * 0.82];
  const nodeR = 7;

  // Connexions (quelques unes)
  layers.forEach((layer, li) => {
    if (li === layers.length - 1) return;
    const nextLayer = layers[li + 1];
    const x1 = xPositions[li], x2 = xPositions[li + 1];
    for (let i = 0; i < layer.n; i++) {
      const y1 = H / 2 + (i - (layer.n - 1) / 2) * 26;
      for (let j = 0; j < nextLayer.n; j++) {
        const y2 = H / 2 + (j - (nextLayer.n - 1) / 2) * 26;
        ctx.beginPath();
        ctx.strokeStyle = 'rgba(42,42,58,0.8)'; ctx.lineWidth = 1;
        ctx.moveTo(x1 + nodeR, y1); ctx.lineTo(x2 - nodeR, y2); ctx.stroke();
      }
    }
  });

  // Noeuds
  layers.forEach((layer, li) => {
    const x = xPositions[li];
    for (let i = 0; i < layer.n; i++) {
      const y = H / 2 + (i - (layer.n - 1) / 2) * 26;
      ctx.beginPath();
      ctx.arc(x, y, nodeR, 0, 2 * Math.PI);
      ctx.fillStyle = layer.color; ctx.fill();
    }
    if (layer.n < layer.full) {
      const yLast = H / 2 + ((layer.n - 1) - (layer.n - 1) / 2) * 26;
      ctx.fillStyle = 'rgba(168,168,192,0.5)'; ctx.font = '11px JetBrains Mono,monospace';
      ctx.textAlign = 'center'; ctx.fillText('…', x, yLast + 20);
    }
    ctx.fillStyle = layer.color; ctx.font = `bold ${Math.round(W * 0.038)}px JetBrains Mono,monospace`;
    ctx.textAlign = 'center';
    ctx.fillText(layer.label, x, H - 8);
    ctx.fillStyle = 'rgba(168,168,192,0.7)'; ctx.font = `${Math.round(W * 0.03)}px JetBrains Mono,monospace`;
    ctx.fillText('×' + layer.full, x, H - 8 - 13);
  });
}

function renderTrainStats(iter, total, lastLoss) {
  document.getElementById('trainStats').innerHTML = `
    <div class="train-stat"><div class="train-stat-val">${iter}</div><div class="train-stat-lbl">Itérations / ${total}</div></div>
    <div class="train-stat"><div class="train-stat-val">${lastLoss.toFixed(4)}</div><div class="train-stat-lbl">Loss actuelle</div></div>
    <div class="train-stat"><div class="train-stat-val">${(lossHistory[0] / (lastLoss + 1e-9)).toFixed(1)}×</div><div class="train-stat-lbl">Amélioration</div></div>
    <div class="train-stat"><div class="train-stat-val">${vocab.length}</div><div class="train-stat-lbl">Mots du vocabulaire</div></div>`;
}

// ═══════════════════════════════════════════════════════════════
// GÉNÉRATION
// ═══════════════════════════════════════════════════════════════
document.getElementById('btnGenerate').addEventListener('click', generateText);
document.getElementById('seedWord').addEventListener('keydown', e => { if (e.key === 'Enter') generateText(); });

function generateText() {
  if (!trained) { notify('Entraînez d\'abord le réseau !', '⚠️'); return; }
  const seedRaw = document.getElementById('seedWord').value.trim().toLowerCase();
  let count = Math.max(1, Math.min(50, parseInt(document.getElementById('repeatCount').value) || 10));
  const temp = parseFloat(document.getElementById('temperature').value);
  document.getElementById('repeatCount').value = count;
  if (!seedRaw) { notify('Entrez un mot de départ', '⚠️'); return; }

  const seed = seedRaw.split(/\s+/).pop();
  if (wordIndex[seed] === undefined) {
    document.getElementById('resultBox').innerHTML =
      `<span class="result-token token-error">${seed}</span> <span style="color:var(--accent2);font-size:.82rem">— Ce mot n'est pas dans le vocabulaire du modèle.</span>`;
    return;
  }

  const tokens = [{ word: seed, type: 'seed' }];
  let current = seed;
  for (let k = 0; k < count; k++) {
    const next = predictNext(current, temp);
    if (!next) { tokens.push({ word: '[fin]', type: 'error' }); break; }
    tokens.push({ word: next, type: 'gen' });
    current = next;
  }
  document.getElementById('resultBox').innerHTML = tokens.map(t =>
    `<span class="result-token token-${t.type}">${t.word}</span>`
  ).join(' ');
}

function predictNext(word, temperature) {
  const i = wordIndex[word];
  if (i === undefined) return null;
  const { probs } = forward(i);

  if (temperature === 0) {
    // Déterministe : argmax (en ignorant le mot lui-même)
    let best = null, bestP = -1;
    probs.forEach((p, j) => { if (j !== i && p > bestP) { bestP = p; best = j; } });
    return best !== null ? vocab[best] : null;
  }

  // Échantillonnage avec température
  const scaled = probs.map((p, j) => j === i ? 0 : Math.pow(p, 1 / temperature));
  const sum = scaled.reduce((a, b) => a + b, 0);
  if (sum === 0) return null;
  const norm = scaled.map(p => p / sum);
  let r = Math.random(), cumul = 0;
  for (let j = 0; j < norm.length; j++) {
    cumul += norm[j];
    if (r <= cumul) return vocab[j];
  }
  return vocab[norm.length - 1];
}

// ═══════════════════════════════════════════════════════════════
// STEPS
// ═══════════════════════════════════════════════════════════════
function setStep(n) {
  for (let i = 1; i <= 4; i++) {
    const el = document.getElementById('step' + i);
    el.className = 'step';
    if (i < n) el.classList.add('done');
    else if (i === n) el.classList.add('active');
  }
}

// ═══════════════════════════════════════════════════════════════
// VIZ 1 — POIDS EN DIRECT (cartes de chaleur W1 et W2)
// ═══════════════════════════════════════════════════════════════
function renderWeightMaps() {
  if (!W1 || !W2) return;
  renderSingleWeightMap('w1Canvas', W1, 'w1Dims');
  renderSingleWeightMap('w2Canvas', W2, 'w2Dims');
}

function renderSingleWeightMap(canvasId, matrix, dimsId) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const rows = matrix.length, cols = matrix[0].length;
  document.getElementById(dimsId).textContent = `(${rows} × ${cols})`;

  const CELL = Math.max(2, Math.min(8, Math.floor(200 / Math.max(rows, cols))));
  const W = cols * CELL, H = rows * CELL;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');

  // Trouver la valeur max absolue pour la normalisation
  let maxAbs = 0;
  matrix.forEach(row => row.forEach(v => { if (Math.abs(v) > maxAbs) maxAbs = Math.abs(v); }));
  maxAbs = maxAbs || 1;

  matrix.forEach((row, r) => {
    row.forEach((v, c) => {
      const norm = v / maxAbs; // -1 à 1
      let R, G, B;
      if (norm < 0) {
        // Négatif : noir → rose
        const t = -norm;
        R = Math.round(255 * t); G = Math.round(109 * t); B = Math.round(155 * t);
      } else {
        // Positif : noir → vert cyan
        const t = norm;
        R = Math.round(109 * t); G = Math.round(255 * t); B = Math.round(204 * t);
      }
      ctx.fillStyle = `rgb(${R},${G},${B})`;
      ctx.fillRect(c * CELL, r * CELL, CELL, CELL);
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// VIZ 2 — EMBEDDINGS ANIMÉS (scatter live)
// ═══════════════════════════════════════════════════════════════
function renderLiveEmbedScatter() {
  const topN = Math.min(parseInt(document.getElementById('embedTopN').value) || 30, vocab.length);
  const canvas = document.getElementById('embedLiveCanvas');
  if (!canvas || !embeddings) return;
  const wrap = canvas.parentElement;
  const W = (wrap.clientWidth - 32) || 660;
  const H = 320;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');

  const top = Object.entries(wordCounts).sort((a, b) => b[1] - a[1]).slice(0, topN).map(([w]) => w).filter(w => wordIndex[w] !== undefined);
  if (top.length < 2) return;

  const vecs = top.map(w => embeddings[wordIndex[w]]);
  const mean = Array(embedDim).fill(0);
  vecs.forEach(v => v.forEach((x, i) => mean[i] += x / vecs.length));
  const centered = vecs.map(v => v.map((x, i) => x - mean[i]));

  // PCA rapide (itération de puissance, moins d'iters pour la performance en live)
  let pc1 = Array(embedDim).fill(0); pc1[0] = 1;
  for (let it = 0; it < 20; it++) {
    const proj = centered.map(v => v.reduce((s, x, i) => s + x * pc1[i], 0));
    const newPc = Array(embedDim).fill(0);
    centered.forEach((v, k) => v.forEach((x, i) => newPc[i] += proj[k] * x));
    const norm = Math.sqrt(newPc.reduce((s, x) => s + x * x, 0)) || 1;
    pc1 = newPc.map(x => x / norm);
  }
  let pc2 = Array(embedDim).fill(0); pc2[Math.min(1, embedDim - 1)] = 1;
  for (let it = 0; it < 20; it++) {
    const dot = pc2.reduce((s, x, i) => s + x * pc1[i], 0);
    pc2 = pc2.map((x, i) => x - dot * pc1[i]);
    const proj = centered.map(v => v.reduce((s, x, i) => s + x * pc2[i], 0));
    const newPc = Array(embedDim).fill(0);
    centered.forEach((v, k) => v.forEach((x, i) => newPc[i] += proj[k] * x));
    const norm = Math.sqrt(newPc.reduce((s, x) => s + x * x, 0)) || 1;
    pc2 = newPc.map(x => x / norm);
  }

  const coords = centered.map(v => ({
    x: v.reduce((s, x, i) => s + x * pc1[i], 0),
    y: v.reduce((s, x, i) => s + x * pc2[i], 0)
  }));

  const xs = coords.map(c => c.x), ys = coords.map(c => c.y);
  const minX = Math.min(...xs), maxX = Math.max(...xs);
  const minY = Math.min(...ys), maxY = Math.max(...ys);
  const PAD = 36;
  const toScreen = c => ({
    sx: PAD + (c.x - minX) / (maxX - minX + 1e-9) * (W - 2 * PAD),
    sy: PAD + (1 - (c.y - minY) / (maxY - minY + 1e-9)) * (H - 2 * PAD)
  });

  ctx.clearRect(0, 0, W, H);
  ctx.strokeStyle = 'rgba(42,42,58,0.5)'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(PAD, H / 2); ctx.lineTo(W - PAD, H / 2); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(W / 2, PAD); ctx.lineTo(W / 2, H - PAD); ctx.stroke();

  // Itération counter
  const iterLabel = lossHistory.length > 0 ? `iter: ${lossHistory.length}  loss: ${lossHistory[lossHistory.length-1].toFixed(4)}` : 'non entraîné';
  ctx.fillStyle = 'rgba(109,255,204,0.5)'; ctx.font = `${Math.round(W*0.022)}px JetBrains Mono,monospace`;
  ctx.textAlign = 'right'; ctx.fillText(iterLabel, W - PAD, PAD - 6);

  coords.forEach((c, k) => {
    const { sx, sy } = toScreen(c);
    const freq = wordCounts[top[k]] || 1;
    const r = Math.max(3, Math.min(6, 2 + Math.log(freq)));
    ctx.beginPath(); ctx.arc(sx, sy, r, 0, 2 * Math.PI);
    ctx.fillStyle = 'rgba(109,255,204,0.75)'; ctx.fill();
    ctx.font = `${Math.round(W * 0.012)}px JetBrains Mono, monospace`;
    ctx.fillStyle = 'rgba(168,168,192,0.85)'; ctx.textAlign = 'left';
    ctx.fillText(top[k], sx + r + 2, sy + 4);
  });
}

// ═══════════════════════════════════════════════════════════════
// VIZ 3 — RÉSEAU ACTIF (signal forward + backward animé)
// ═══════════════════════════════════════════════════════════════
function triggerSignalStep(inputIdx, targetIdx) {
  if (!vizState.signal || !W1) return;
  const { h, probs } = forward(inputIdx);
  signalAnim = { inputIdx, targetIdx, h, probs, phase: 'forward', progress: 0 };
  if (signalAnimFrame) cancelAnimationFrame(signalAnimFrame);
  animateSignal();
}

function startSignalAnimation() {
  if (!W1 || !vocab.length) return;
  const i = Math.floor(Math.random() * vocab.length);
  const j = Math.floor(Math.random() * vocab.length);
  triggerSignalStep(i, j);
}

function animateSignal() {
  if (!vizState.signal || !signalAnim) return;
  signalAnim.progress = Math.min(1, signalAnim.progress + 0.06);
  renderNetworkActive(signalAnim);
  if (signalAnim.progress < 1) {
    signalAnimFrame = requestAnimationFrame(animateSignal);
  } else if (signalAnim.phase === 'forward') {
    signalAnim.phase = 'backward';
    signalAnim.progress = 0;
    signalAnimFrame = requestAnimationFrame(animateSignal);
  }
}

function renderNetworkActive(anim) {
  const canvas = document.getElementById('networkCanvas');
  if (!canvas || !W1) return;
  const wrap = canvas.parentElement;
  const W = (wrap.clientWidth - 32) || 300;
  const H = 220;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  const hiddenSize = W1.length;
  const layers = [
    { label: 'Embedding', n: Math.min(embedDim, 6), color: 'rgba(124,109,255,0.8)', full: embedDim },
    { label: 'Cachée', n: Math.min(hiddenSize, 6), color: 'rgba(109,255,204,0.8)', full: hiddenSize },
    { label: 'Sortie', n: Math.min(vocab.length, 6), color: 'rgba(255,109,155,0.8)', full: vocab.length }
  ];
  const xPositions = [W * 0.18, W * 0.5, W * 0.82];
  const nodeR = 7;
  const t = anim.progress;
  const isForward = anim.phase === 'forward';

  // Connexions animées
  layers.forEach((layer, li) => {
    if (li === layers.length - 1) return;
    const nextLayer = layers[li + 1];
    const x1 = xPositions[li], x2 = xPositions[li + 1];
    const segActive = isForward ? (li === 0 ? t : (t > 0.5 ? (t - 0.5) * 2 : 0))
                                : (li === 1 ? t : (t > 0.5 ? (t - 0.5) * 2 : 0));
    for (let i = 0; i < layer.n; i++) {
      const y1 = H / 2 + (i - (layer.n - 1) / 2) * 26;
      for (let j = 0; j < nextLayer.n; j++) {
        const y2 = H / 2 + (j - (nextLayer.n - 1) / 2) * 26;
        const alpha = 0.08 + segActive * 0.55;
        const col = isForward ? `rgba(109,255,204,${alpha})` : `rgba(255,109,155,${alpha})`;
        ctx.beginPath(); ctx.strokeStyle = col; ctx.lineWidth = 1 + segActive * 1.5;
        ctx.moveTo(x1 + nodeR, y1); ctx.lineTo(x2 - nodeR, y2); ctx.stroke();
      }
    }
  });

  // Noeuds
  layers.forEach((layer, li) => {
    const x = xPositions[li];
    const nodeActive = isForward ? (li === 0 ? t : li === 1 ? Math.max(0, t * 2 - 1) : 0)
                                 : (li === 2 ? t : li === 1 ? Math.max(0, t * 2 - 1) : 0);
    const glowColor = isForward ? `rgba(109,255,204,${nodeActive * 0.4})` : `rgba(255,109,155,${nodeActive * 0.4})`;

    for (let i = 0; i < layer.n; i++) {
      const y = H / 2 + (i - (layer.n - 1) / 2) * 26;
      if (nodeActive > 0.1) {
        ctx.beginPath(); ctx.arc(x, y, nodeR + 4, 0, 2 * Math.PI);
        ctx.fillStyle = glowColor; ctx.fill();
      }
      ctx.beginPath(); ctx.arc(x, y, nodeR, 0, 2 * Math.PI);
      ctx.fillStyle = layer.color; ctx.fill();
    }
    if (layer.n < layer.full) {
      const yLast = H / 2 + ((layer.n - 1) - (layer.n - 1) / 2) * 26;
      ctx.fillStyle = 'rgba(168,168,192,0.5)'; ctx.font = '11px JetBrains Mono,monospace';
      ctx.textAlign = 'center'; ctx.fillText('…', x, yLast + 20);
    }
    ctx.fillStyle = layer.color; ctx.font = `bold ${Math.round(W * 0.038)}px JetBrains Mono,monospace`;
    ctx.textAlign = 'center'; ctx.fillText(layer.label, x, H - 8);
    ctx.fillStyle = 'rgba(168,168,192,0.7)'; ctx.font = `${Math.round(W * 0.03)}px JetBrains Mono,monospace`;
    ctx.fillText('×' + layer.full, x, H - 8 - 13);
  });

  // Label phase
  const phaseLabel = isForward ? '→ forward pass' : '← backward pass';
  const phaseColor = isForward ? 'rgba(109,255,204,0.7)' : 'rgba(255,109,155,0.7)';
  ctx.fillStyle = phaseColor; ctx.font = `${Math.round(W * 0.032)}px JetBrains Mono,monospace`;
  ctx.textAlign = 'center'; ctx.fillText(phaseLabel, W / 2, 14);
}

// ═══════════════════════════════════════════════════════════════
// RESIZE
// ═══════════════════════════════════════════════════════════════
window.addEventListener('resize', () => {
  if (embeddings && document.getElementById('embedSection').style.display !== 'none') {
    const topN = Math.min(parseInt(document.getElementById('embedTopN').value) || 30, vocab.length);
    renderEmbedScatter(topN);
  }
  if (lossHistory.length > 1) renderLossChart();
  if (trained) renderNetworkViz(W1 ? W1.length : 32);
  if (vizState.weights && W1) renderWeightMaps();
  if (vizState.embeds && embeddings) renderLiveEmbedScatter();
});
