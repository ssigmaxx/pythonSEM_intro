const pptxgen = require('pptxgenjs');
const path = require('path');

const ICON = (name) => path.join(__dirname, 'icons', `${name}.png`);

// ---- Palette: Python Blue & Gold (matches the Python intro deck) ----
const NAVY = '132038';
const NAVY2 = '1B2A4A';
const BLUE = '306998';
const BLUE_DK = '1F4B70';
const GOLD = 'FFD43B';
const GOLD_DK = 'C9A227';
const WHITE = 'FFFFFF';
const CARD = 'F4F6F9';
const TEXT = '1F2937';
const MUTED = '5B6472';
const MUTED_LIGHT = 'AEB9CC';
const GREEN = '2C5F2D';
const RED = 'B03A2E';

const HEAD_FONT = 'Arial';
const BODY_FONT = 'Calibri';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
const PW = 13.333, PH = 7.5;

let chapterCounter = 0;
let slideCounter = 0;

function footer(slide, label, dark) {
  slideCounter++;
  slide.addText(label, {
    x: 0.5, y: PH - 0.42, w: 7, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? MUTED_LIGHT : MUTED, isTextBox: true, margin: 0,
  });
  slide.addText(String(slideCounter), {
    x: PW - 1.0, y: PH - 0.42, w: 0.5, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? MUTED_LIGHT : MUTED, align: 'right', isTextBox: true, margin: 0,
  });
}

function iconCircle(slide, iconName, x, y, d, bgColor) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: bgColor }, line: { type: 'none' } });
  const pad = d * 0.26;
  slide.addImage({ path: ICON(iconName), x: x + pad / 2, y: y + pad / 2, w: d - pad, h: d - pad });
}

function kicker(slide, text, color) {
  slide.addText(text.toUpperCase(), {
    x: 0.7, y: 0.42, w: 8, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 13,
    color, charSpacing: 2, isTextBox: true, margin: 0,
  });
}

function pageTitle(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.7, y: opts.y || 0.72, w: opts.w || 11.9, h: opts.h || 0.9, fontFace: HEAD_FONT, bold: true,
    fontSize: opts.size || 32, color: opts.color || TEXT, isTextBox: true, margin: 0,
  });
}

function chapterDivider({ icon, title, subtitle }) {
  chapterCounter++;
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText(`CHAPTER ${chapterCounter}`, {
    x: 0.9, y: 2.55, w: 6, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 16, color: GOLD, charSpacing: 3, isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: 0.9, y: 3.0, w: 9.3, h: 1.6, fontFace: HEAD_FONT, bold: true, fontSize: 44, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText(subtitle, {
    x: 0.9, y: 4.35, w: 8.8, h: 0.8, fontFace: BODY_FONT, fontSize: 16, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  iconCircle(s, icon, PW - 3.1, 2.55, 2.0, BLUE);
  footer(s, `Chapter ${chapterCounter}`, true);
  return s;
}

function contentSlide(chapterLabel) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  footer(s, chapterLabel, false);
  return s;
}

function codeBlock(slide, lines, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  let ly = y + 0.22;
  lines.forEach((ln) => {
    slide.addText(ln, { x: x + 0.35, y: ly, w: w - 0.6, h: 0.4, fontFace: 'Courier New', fontSize: 15, color: GOLD, isTextBox: true, margin: 0 });
    ly += 0.42;
  });
}
function codeH(n) { return n * 0.42 + 0.44; }

// ---------------------------------------------------------------------
// SLIDE 1: TITLE
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'jupyter_logo', PW / 2 - 0.75, 0.9, 1.5, BLUE);
  s.addText('INTRODUCTION TO', {
    x: 0, y: 2.75, w: PW, h: 0.5, align: 'center', fontFace: HEAD_FONT, bold: true,
    fontSize: 18, color: GOLD, charSpacing: 4, isTextBox: true, margin: 0,
  });
  s.addText('Jupyter Notebook', {
    x: 0, y: 3.15, w: PW, h: 1.3, align: 'center', fontFace: HEAD_FONT, bold: true,
    fontSize: 58, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText('A friendly, no-experience-needed course for our team', {
    x: 0, y: 4.5, w: PW, h: 0.5, align: 'center', fontFace: BODY_FONT, fontSize: 18,
    color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', {
    x: PW / 2 - 2.9, y: 5.35, w: 5.8, h: 0.55, rectRadius: 0.28,
    fill: { color: NAVY2 }, line: { color: BLUE, width: 1 },
  });
  s.addText('6 chapters  •  hands-on  •  builds on our Python session', {
    x: PW / 2 - 2.9, y: 5.35, w: 5.8, h: 0.55, align: 'center', valign: 'middle',
    fontFace: BODY_FONT, fontSize: 12.5, color: GOLD, isTextBox: true, margin: 0,
  });
  footer(s, 'Team Python Onboarding', true);
}

// ---------------------------------------------------------------------
// SLIDE 2: WELCOME & ROADMAP
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, 'Welcome', BLUE);
  pageTitle(s, "Where we're headed");
  s.addText('Six short chapters, from installing Jupyter to knowing when and why to reach for it. Bring the Python basics from last session, we will build on them.', {
    x: 0.7, y: 1.55, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  const chapters = [
    { n: '1', t: 'Installing Jupyter', icon: 'download' },
    { n: '2', t: 'What is a notebook?', icon: 'question' },
    { n: '3', t: 'Cells: code & markdown', icon: 'layout' },
    { n: '4', t: 'Interface & the kernel', icon: 'kernel' },
    { n: '5', t: 'Why use Jupyter?', icon: 'compare' },
    { n: '6', t: 'Good habits to know', icon: 'gear' },
  ];
  const cols = 3, cardW = 3.75, cardH = 1.85, gapX = 0.25, gapY = 0.3;
  const startX = 0.7, startY = 2.45;
  chapters.forEach((c, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, c.icon, x + 0.28, y + 0.28, 0.62, BLUE);
    s.addText(`CHAPTER ${c.n}`, { x: x + 1.1, y: y + 0.28, w: cardW - 1.3, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 10.5, color: GOLD_DK, isTextBox: true, margin: 0 });
    s.addText(c.t, { x: x + 1.1, y: y + 0.56, w: cardW - 1.3, h: 0.9, fontFace: HEAD_FONT, bold: true, fontSize: 16, color: TEXT, isTextBox: true, margin: 0 });
  });
}

// ---------------------------------------------------------------------
// SLIDE 3: ICEBREAKER (interactive)
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'comments', 0.7, 0.6, 0.9, BLUE);
  s.addText('LET’S START WITH A QUICK POLL', { x: 1.85, y: 0.68, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Raise your hand: have you opened a notebook before?', { x: 1.85, y: 1.0, w: 10.5, h: 0.7, fontFace: HEAD_FONT, bold: true, fontSize: 25, color: WHITE, isTextBox: true, margin: 0 });

  const options = [
    { l: 'A', t: 'Never, this is brand new to me' },
    { l: 'B', t: 'I have seen one, but never run one myself' },
    { l: 'C', t: 'I have used Jupyter or Google Colab before' },
  ];
  let y = 2.15;
  options.forEach((o) => {
    s.addShape('roundRect', { x: 0.9, y, w: 11.5, h: 1.05, rectRadius: 0.1, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addShape('ellipse', { x: 1.2, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: GOLD }, line: { type: 'none' } });
    s.addText(o.l, { x: 1.2, y: y + 0.25, w: 0.55, h: 0.55, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 18, color: NAVY, isTextBox: true, margin: 0 });
    s.addText(o.t, { x: 2.0, y, w: 10.2, h: 1.05, valign: 'middle', fontFace: BODY_FONT, fontSize: 16.5, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.2;
  });
  s.addText('No wrong answers, we will get everyone to the same starting line.', {
    x: 0.9, y: 6.75, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 0 · Welcome', true);
}

// ---------------------------------------------------------------------
// CHAPTER 1: INSTALLING JUPYTER
// ---------------------------------------------------------------------
chapterDivider({ icon: 'download', title: 'Installing Jupyter', subtitle: 'If Python is already on your machine, Jupyter is one command away.' });

// Before you begin
{
  const s = contentSlide('Chapter 1 · Installing Jupyter');
  kicker(s, 'Chapter 1', BLUE);
  pageTitle(s, 'Before you begin');
  s.addText('If you finished the Python installation session, you already have almost everything you need.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const items = [
    { icon: 'check', t: 'Python installed', d: 'Jupyter runs on top of Python, version 3.8 or newer works well' },
    { icon: 'terminal', t: 'A working pip', d: 'Confirm with pip --version, we used this in the last session' },
    { icon: 'globe', t: 'A modern browser', d: 'Notebooks open and run inside Chrome, Edge, Firefox, or Safari' },
  ];
  const cardW = 3.75, gapX = 0.25, startX = 0.7, y = 2.35, h = 3.4;
  items.forEach((it, i) => {
    const x = startX + i * (cardW + gapX);
    s.addShape('roundRect', { x, y, w: cardW, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, it.icon, x + cardW / 2 - 0.45, y + 0.45, 0.9, BLUE);
    s.addText(it.t, { x: x + 0.3, y: y + 1.55, w: cardW - 0.6, h: 0.5, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 16, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(it.d, { x: x + 0.3, y: y + 2.05, w: cardW - 0.6, h: 1.1, align: 'center', fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Two ways to install
{
  const s = contentSlide('Chapter 1 · Installing Jupyter');
  kicker(s, 'Chapter 1', BLUE);
  pageTitle(s, 'Two ways to get Jupyter');

  const lx = 0.7, lw = 5.6, y = 1.7, h = 4.3;
  s.addShape('roundRect', { x: lx, y, w: lw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'terminal', lx + 0.3, y + 0.3, 0.7, BLUE);
  s.addText('Install with pip', { x: lx + 1.15, y: y + 0.4, w: lw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('Fastest option if Python is already set up. Open a terminal and run:', { x: lx + 0.3, y: y + 1.2, w: lw - 0.6, h: 0.55, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  codeBlock(s, ['pip install notebook'], lx + 0.3, y + 1.85, lw - 0.6, codeH(1));
  s.addText('Prefer the newer interface? Install JupyterLab instead:', { x: lx + 0.3, y: y + 2.75, w: lw - 0.6, h: 0.4, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  codeBlock(s, ['pip install jupyterlab'], lx + 0.3, y + 3.2, lw - 0.6, codeH(1));

  const rx = 6.6, rw = 6.03;
  s.addShape('roundRect', { x: rx, y, w: rw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'download', rx + 0.3, y + 0.3, 0.7, BLUE);
  s.addText('Install with Anaconda', { x: rx + 1.15, y: y + 0.4, w: rw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('A beginner-friendly bundle: Python, Jupyter, and hundreds of data science libraries, all in one installer.', {
    x: rx + 0.3, y: y + 1.2, w: rw - 0.6, h: 1.0, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0,
  });
  const steps = ['Download Anaconda from anaconda.com', 'Run the installer for your OS', 'Open "Anaconda Navigator" and launch Jupyter from there'];
  let sy = y + 2.3;
  steps.forEach((st, i) => {
    s.addShape('ellipse', { x: rx + 0.3, y: sy, w: 0.4, h: 0.4, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: rx + 0.3, y: sy, w: 0.4, h: 0.4, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 12, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: rx + 0.85, y: sy - 0.03, w: rw - 1.15, h: 0.5, valign: 'middle', fontFace: BODY_FONT, fontSize: 12, color: TEXT, isTextBox: true, margin: 0 });
    sy += 0.58;
  });
}

// Launching Jupyter
{
  const s = contentSlide('Chapter 1 · Installing Jupyter');
  kicker(s, 'Chapter 1', BLUE);
  pageTitle(s, 'Launching your first notebook');
  s.addText('However you installed it, starting Jupyter always works the same way.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const steps = [
    'Open a terminal and navigate to a folder for your work',
    'Type jupyter notebook (or jupyter lab) and press Enter',
    'Your browser opens automatically to a file list at localhost:8888',
    'Click "New" in the top right, then choose "Python 3" to create a notebook',
  ];
  let y = 2.5;
  steps.forEach((st, i) => {
    s.addShape('ellipse', { x: 0.7, y, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 0.7, y, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.4, y: y - 0.03, w: 10.8, h: 0.55, valign: 'middle', fontFace: BODY_FONT, fontSize: 14.5, color: TEXT, isTextBox: true, margin: 0 });
    y += 0.68;
  });
  codeBlock(s, ['jupyter notebook'], 0.7, y + 0.15, 6.6, codeH(1));
  s.addShape('roundRect', { x: 7.6, y: y + 0.15, w: 4.6, h: codeH(1), rectRadius: 0.08, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'lightbulb', 7.75, y + 0.15 + (codeH(1) - 0.5) / 2, 0.5, GOLD);
  s.addText('The terminal window must stay open, closing it shuts down the server.', {
    x: 8.35, y: y + 0.15, w: 3.75, h: codeH(1), valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 11, color: MUTED, isTextBox: true, margin: 0,
  });
}

// Try it now (interactive)
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'rocket', 0.7, 0.6, 0.9, GOLD);
  s.addText('TRY IT NOW', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Get a blank notebook on screen', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const steps = [
    'Open a terminal and run jupyter notebook',
    'In the browser tab that opens, click New, then Python 3',
    'You should see a blank notebook with one empty gray box, that is a cell',
    'Rename it: click "Untitled" at the top and type my_first_notebook',
  ];
  let y = 2.1;
  steps.forEach((st, i) => {
    s.addShape('roundRect', { x: 0.9, y, w: 11.5, h: 0.85, rectRadius: 0.1, fill: { color: NAVY2 }, line: { type: 'none' } });
    s.addShape('ellipse', { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.85, y, w: 10.4, h: 0.85, valign: 'middle', fontFace: BODY_FONT, fontSize: 14, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.0;
  });
  s.addText('Stuck? Raise a hand, we will pair up and fix it together before moving on.', {
    x: 0.9, y: 6.15, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 1 · Installing Jupyter', true);
}

// ---------------------------------------------------------------------
// CHAPTER 2: WHAT IS A NOTEBOOK?
// ---------------------------------------------------------------------
chapterDivider({ icon: 'question', title: 'What is a Notebook?', subtitle: 'A little history and the ideas that make notebooks different from a plain script.' });

// What is Jupyter (definition + history)
{
  const s = contentSlide('Chapter 2 · What is a Notebook?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'What is Jupyter Notebook, really?');
  s.addText('A Jupyter Notebook is a document that mixes live, runnable code with text, images, and the output of that code, all in one file you can read top to bottom like a report.', {
    x: 0.7, y: 1.55, w: 6.9, h: 1.5, fontFace: BODY_FONT, fontSize: 15, color: TEXT, isTextBox: true, margin: 0,
  });
  s.addText('It grew out of IPython, an enhanced interactive Python shell. In 2014, that interactive notebook piece split off into its own project named Jupyter, a name that nods to three languages it was built to support: Julia, Python, and R.', {
    x: 0.7, y: 3.2, w: 6.9, h: 2.0, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  const stats = [
    { n: '2014', l: 'Year Project Jupyter split off from IPython' },
    { n: '.ipynb', l: 'The file extension every notebook is saved as' },
    { n: '40+', l: 'Programming languages Jupyter kernels support' },
  ];
  let sy = 1.6;
  stats.forEach((st) => {
    s.addShape('roundRect', { x: 7.95, y: sy, w: 4.65, h: 1.45, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    s.addText(st.n, { x: 8.2, y: sy + 0.15, w: 4.2, h: 0.9, fontFace: HEAD_FONT, bold: true, fontSize: 30, color: BLUE, isTextBox: true, margin: 0 });
    s.addText(st.l, { x: 8.2, y: sy + 0.95, w: 4.2, h: 0.45, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
    sy += 1.65;
  });
}

// Key characteristics
{
  const s = contentSlide('Chapter 2 · What is a Notebook?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'Four traits worth knowing');

  const traits = [
    { icon: 'layout', t: 'Cell-based', d: 'Code is broken into small, runnable chunks called cells, instead of one long file.' },
    { icon: 'restart', t: 'Interactive', d: 'Run a cell, see the result immediately, tweak it, and run it again, without restarting everything.' },
    { icon: 'markdown', t: 'Mixes text and code', d: 'Markdown cells let you explain your thinking right next to the code that does the work.' },
    { icon: 'kernel', t: 'Kernel-powered', d: 'A background process called the kernel actually runs your code and remembers your variables.' },
  ];
  const cols = 2, cardW = 5.75, cardH = 2.1, gapX = 0.3, gapY = 0.3, startX = 0.7, startY = 1.65;
  traits.forEach((t, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, t.icon, x + 0.3, y + 0.3, 0.7, BLUE);
    s.addText(t.t, { x: x + 1.2, y: y + 0.32, w: cardW - 1.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 16, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(t.d, { x: x + 0.3, y: y + 1.05, w: cardW - 0.6, h: 0.95, fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Where notebooks are used
{
  const s = contentSlide('Chapter 2 · What is a Notebook?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'Where notebooks show up at work');
  s.addText('The same cell-by-cell workflow shows up across these everyday tasks.', { x: 0.7, y: 1.55, w: 10, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const domains = [
    { icon: 'chart', t: 'Exploring data', d: 'Poke at a dataset, chart it, and adjust without starting over' },
    { icon: 'robot', t: 'Building AI models', d: 'Train, test, and tweak machine learning models step by step' },
    { icon: 'gradcap', t: 'Teaching and tutorials', d: 'Explain an idea, then let students run the code right there' },
    { icon: 'flask', t: 'Research', d: 'Keep notes, code, and results together for reproducible work' },
    { icon: 'share', t: 'Sharing findings', d: 'Send a colleague one file with your analysis and your reasoning' },
    { icon: 'puzzle', t: 'Prototyping', d: 'Try an idea quickly before turning it into a full application' },
  ];
  const cols = 3, cardW = 3.75, cardH = 2.0, gapX = 0.25, gapY = 0.25, startX = 0.7, startY = 2.15;
  domains.forEach((d, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, d.icon, x + 0.28, y + 0.28, 0.65, BLUE);
    s.addText(d.t, { x: x + 0.28, y: y + 1.05, w: cardW - 0.56, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(d.d, { x: x + 0.28, y: y + 1.42, w: cardW - 0.56, h: 0.55, fontFace: BODY_FONT, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Quiz interactive
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'clipboard', 0.7, 0.6, 0.9, GOLD);
  s.addText('QUICK QUIZ', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Which of these run on notebooks?', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const guesses = ['Google Colab', 'Kaggle', 'JupyterHub', 'VS Code'];
  const cols = 2, cardW = 5.6, cardH = 1.15, gapX = 0.3, gapY = 0.3, startX = 0.9, startY = 2.05;
  guesses.forEach((g, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.1, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addText(g, { x, y, w: cardW, h: cardH, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 19, color: WHITE, isTextBox: true, margin: 0 });
  });
  s.addShape('roundRect', { x: 0.9, y: 4.85, w: 11.0, h: 0.9, rectRadius: 0.1, fill: { color: BLUE_DK }, line: { type: 'none' } });
  s.addText('Answer: all four! Colab and Kaggle run notebooks in the cloud, JupyterHub shares them across a team, and VS Code can open and run .ipynb files too.', {
    x: 0.9, y: 4.85, w: 11.0, h: 0.9, align: 'center', valign: 'middle', fontFace: BODY_FONT, bold: true, fontSize: 13.5, color: GOLD, isTextBox: true, margin: 0,
  });
  s.addText('Anyone here already used one of these without realizing it was Jupyter underneath?', {
    x: 0.9, y: 6.05, w: 11.0, h: 0.4, italic: true, align: 'center', fontFace: BODY_FONT, fontSize: 12, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 2 · What is a Notebook?', true);
}

// ---------------------------------------------------------------------
// CHAPTER 3: CELLS
// ---------------------------------------------------------------------
chapterDivider({ icon: 'layout', title: 'Cells: Code and Markdown', subtitle: 'Everything in a notebook lives inside a cell. There are two kinds worth knowing.' });

// Two types of cells
{
  const s = contentSlide('Chapter 3 · Cells: Code and Markdown');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Two kinds of cells');
  s.addText('A notebook is just a stack of cells, run in whatever order you choose.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  s.addShape('roundRect', { x: 0.7, y: 2.25, w: 5.7, h: 3.6, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'code', 0.95, 2.45, 0.55, BLUE);
  s.addText('Code cell', { x: 1.65, y: 2.53, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: BLUE, isTextBox: true, margin: 0 });
  codeBlock(s, ['total = 4 + 5', 'print(total)'], 0.95, 3.15, 5.2, codeH(2));
  s.addText('Runs as real Python. The result prints directly beneath the cell.', { x: 0.95, y: 3.15 + codeH(2) + 0.15, w: 5.2, h: 0.5, fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });

  s.addShape('roundRect', { x: 6.7, y: 2.25, w: 5.93, h: 3.6, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'markdown', 6.95, 2.45, 0.55, BLUE);
  s.addText('Markdown cell', { x: 7.65, y: 2.53, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: BLUE, isTextBox: true, margin: 0 });
  codeBlock(s, ['# Step 1: Load the data', 'We start by reading the', 'file into memory.'], 6.95, 3.15, 5.4, codeH(3));
  s.addText('Formatted text, headings, and notes. Explains the "why" beside the code.', { x: 6.95, y: 3.15 + codeH(3) + 0.15, w: 5.4, h: 0.5, fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
}

// Running a cell
{
  const s = contentSlide('Chapter 3 · Cells: Code and Markdown');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Running a cell');
  s.addText('Click a cell, then use one of these to run it:', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const rows = [
    { code: 'Shift + Enter', d: 'Run this cell, then jump to the next one' },
    { code: 'Ctrl + Enter', d: 'Run this cell, and stay right here' },
    { code: 'Alt + Enter', d: 'Run this cell, then insert a new one below it' },
  ];
  let y = 2.15;
  s.addText('SHORTCUT', { x: 0.7, y, w: 3.6, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('WHAT IT DOES', { x: 4.6, y, w: 6.0, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  y += 0.45;
  rows.forEach((r, i) => {
    const rh = 0.85;
    if (i % 2 === 0) s.addShape('rect', { x: 0.7, y, w: 11.9, h: rh, fill: { color: CARD }, line: { type: 'none' } });
    s.addShape('roundRect', { x: 0.9, y: y + 0.13, w: 3.4, h: 0.6, rectRadius: 0.06, fill: { color: NAVY }, line: { type: 'none' } });
    s.addText(r.code, { x: 1.05, y: y + 0.13, w: 3.1, h: 0.6, valign: 'middle', fontFace: 'Courier New', fontSize: 12.5, color: GOLD, isTextBox: true, margin: 0 });
    s.addText(r.d, { x: 4.6, y, w: 7.0, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 13.5, color: TEXT, isTextBox: true, margin: 0 });
    y += rh;
  });

  s.addShape('roundRect', { x: 0.7, y: y + 0.2, w: 11.9, h: 0.95, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  iconCircle(s, 'play', 0.9, y + 0.35, 0.65, GOLD);
  s.addText([
    { text: 'Watch the brackets: ', options: { bold: true, color: WHITE } },
    { text: 'In [1]: means this was the first cell run, In [*]: means it is still working. The number tells you the order cells actually ran in, which is not always top to bottom.', options: { color: MUTED_LIGHT } },
  ], { x: 1.7, y: y + 0.2, w: 10.6, h: 0.95, valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, isTextBox: true, margin: 0 });
}

// Markdown cells for documentation
{
  const s = contentSlide('Chapter 3 · Cells: Code and Markdown');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Writing markdown that reads well');
  s.addText('Markdown is a simple way to format text using plain characters, no clicking through menus.', { x: 0.7, y: 1.55, w: 11.5, h: 0.5, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  s.addShape('roundRect', { x: 0.7, y: 2.35, w: 5.7, h: 3.5, rectRadius: 0.09, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText('You type this', { x: 1.0, y: 2.6, w: 5.1, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: MUTED_LIGHT, isTextBox: true, margin: 0 });
  codeBlock(s, [
    '# Data Cleaning',
    '**Step 1**: remove blanks',
    '- drop duplicates',
    '- fix column names',
  ], 1.0, 3.0, 5.1, codeH(4));

  s.addShape('roundRect', { x: 6.7, y: 2.35, w: 5.93, h: 3.5, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  s.addText('You see this', { x: 7.0, y: 2.6, w: 5.3, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('Data Cleaning', { x: 7.0, y: 3.05, w: 5.3, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 22, color: TEXT, isTextBox: true, margin: 0 });
  s.addText([
    { text: 'Step 1', options: { bold: true, breakLine: false } },
    { text: ': remove blanks', options: { breakLine: true } },
  ], { x: 7.0, y: 3.65, w: 5.3, h: 0.6, fontFace: BODY_FONT, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
  s.addText([
    { text: '• drop duplicates', options: { breakLine: true } },
    { text: '• fix column names', options: { breakLine: true } },
  ], { x: 7.0, y: 4.35, w: 5.3, h: 0.8, fontFace: BODY_FONT, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
}

// Try it yourself
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'rocket', 0.7, 0.6, 0.9, GOLD);
  s.addText('TRY IT YOURSELF', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Build a two-cell notebook', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  s.addShape('roundRect', { x: 0.9, y: 2.05, w: 11.5, h: 1.15, rectRadius: 0.08, fill: { color: NAVY2 }, line: { type: 'none' } });
  s.addText([
    { text: '## My First Notebook', options: { color: GOLD, breakLine: true } },
    { text: 'name = "Your Name"', options: { color: WHITE, breakLine: true } },
    { text: 'print("Hello from", name)', options: { color: WHITE, breakLine: false } },
  ], { x: 1.2, y: 2.25, w: 10.9, h: 0.85, fontFace: 'Courier New', fontSize: 15, isTextBox: true, margin: 0 });

  const steps = [
    'In your open notebook, type ## My First Notebook into the first cell',
    'Change that cell to Markdown (press Esc, then M), then run it with Shift + Enter',
    'In the new cell, type the two print lines above, swapping in your name',
    'Run it, and check your name appears in the output below the cell',
  ];
  let y = 3.55;
  steps.forEach((st, i) => {
    s.addShape('ellipse', { x: 0.9, y, w: 0.42, h: 0.42, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 0.9, y, w: 0.42, h: 0.42, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 13, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.48, y: y - 0.04, w: 10.8, h: 0.5, valign: 'middle', fontFace: BODY_FONT, fontSize: 13, color: WHITE, isTextBox: true, margin: 0 });
    y += 0.58;
  });
  footer(s, 'Chapter 3 · Cells: Code and Markdown', true);
}

// ---------------------------------------------------------------------
// CHAPTER 4: INTERFACE & KERNEL
// ---------------------------------------------------------------------
chapterDivider({ icon: 'kernel', title: 'Interface and the Kernel', subtitle: 'A quick tour of the screen in front of you, and the engine running underneath it.' });

// Tour of the interface
{
  const s = contentSlide('Chapter 4 · Interface and the Kernel');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'A quick tour of the toolbar');

  const items = [
    { icon: 'play', t: 'Run', d: 'Executes the current cell and moves to the next' },
    { icon: 'restart', t: 'Restart kernel', d: 'Clears all memory and starts the kernel fresh' },
    { icon: 'layout', t: 'Cell type', d: 'Switches the current cell between Code and Markdown' },
    { icon: 'save', t: 'Save', d: 'Writes your progress to the .ipynb file on disk' },
    { icon: 'export', t: 'Download as', d: 'Exports the notebook to .py, .html, or .pdf' },
    { icon: 'question', t: 'Command palette', d: 'Search every action by name, handy while learning' },
  ];
  const cols = 3, cardW = 3.75, cardH = 2.0, gapX = 0.25, gapY = 0.25, startX = 0.7, startY = 1.7;
  items.forEach((it, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, it.icon, x + 0.28, y + 0.28, 0.65, BLUE);
    s.addText(it.t, { x: x + 0.28, y: y + 1.05, w: cardW - 0.56, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(it.d, { x: x + 0.28, y: y + 1.42, w: cardW - 0.56, h: 0.55, fontFace: BODY_FONT, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// What is a kernel
{
  const s = contentSlide('Chapter 4 · Interface and the Kernel');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'What is "the kernel"?');
  s.addText('The kernel is the separate process that actually runs your code. Your notebook is just the page you type on, the kernel is what reads it and does the work.', {
    x: 0.7, y: 1.55, w: 6.9, h: 1.2, fontFace: BODY_FONT, fontSize: 15, color: TEXT, isTextBox: true, margin: 0,
  });
  s.addText('It keeps every variable you have defined in memory, in the order the cells were run, not the order they appear on the page. That is powerful, and it is also the single most common source of confusion for beginners.', {
    x: 0.7, y: 2.85, w: 6.9, h: 1.5, fontFace: BODY_FONT, fontSize: 13.5, color: MUTED, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: 0.7, y: 4.5, w: 6.9, h: 1.0, rectRadius: 0.08, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'restart', 0.9, 4.67, 0.6, GOLD);
  s.addText('When things feel broken or a variable seems "stuck," restart the kernel and run all cells again from the top.', {
    x: 1.65, y: 4.5, w: 5.8, h: 1.0, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0,
  });

  s.addShape('roundRect', { x: 7.95, y: 1.55, w: 4.65, h: 3.95, rectRadius: 0.09, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText('One kernel, many cells', { x: 8.25, y: 1.8, w: 4.1, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: MUTED_LIGHT, isTextBox: true, margin: 0 });
  codeBlock(s, [
    'x = 10',
    '',
    'y = x * 2',
    '',
    'print(y)',
    '# 20',
  ], 8.25, 2.2, 4.05, codeH(6));
}

// Keyboard shortcuts
{
  const s = contentSlide('Chapter 4 · Interface and the Kernel');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'Handy keyboard shortcuts');
  s.addText('Press Esc to leave a cell (blue border, "command mode"), then use these:', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const shortcuts = [
    { k: 'A', d: 'Insert a new cell above' },
    { k: 'B', d: 'Insert a new cell below' },
    { k: 'D, D', d: 'Delete the selected cell' },
    { k: 'M', d: 'Turn the cell into Markdown' },
    { k: 'Y', d: 'Turn the cell back into Code' },
    { k: 'Enter', d: 'Go back into the cell to edit it' },
  ];
  const cols = 2, cardW = 5.75, cardH = 1.05, gapX = 0.3, gapY = 0.2, startX = 0.7, startY = 2.15;
  shortcuts.forEach((sc, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: CARD }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + 0.2, y: y + (cardH - 0.55) / 2, w: 1.3, h: 0.55, rectRadius: 0.06, fill: { color: NAVY }, line: { type: 'none' } });
    s.addText(sc.k, { x: x + 0.2, y: y + (cardH - 0.55) / 2, w: 1.3, h: 0.55, align: 'center', valign: 'middle', fontFace: 'Courier New', bold: true, fontSize: 14, color: GOLD, isTextBox: true, margin: 0 });
    s.addText(sc.d, { x: x + 1.75, y, w: cardW - 2.0, h: cardH, valign: 'middle', fontFace: BODY_FONT, fontSize: 13, color: TEXT, isTextBox: true, margin: 0 });
  });
}

// Try it now (shortcuts practice)
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'keyboard', 0.7, 0.6, 0.9, GOLD);
  s.addText('TRY IT NOW', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Three shortcuts, thirty seconds', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const steps = [
    'Click a cell, then press Esc so the border turns blue',
    'Press B to add a new cell below it',
    'Press D twice, quickly, to delete that new cell',
    'Press A to add one back above instead',
  ];
  let y = 2.1;
  steps.forEach((st, i) => {
    s.addShape('roundRect', { x: 0.9, y, w: 11.5, h: 0.85, rectRadius: 0.1, fill: { color: NAVY2 }, line: { type: 'none' } });
    s.addShape('ellipse', { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.85, y, w: 10.4, h: 0.85, valign: 'middle', fontFace: BODY_FONT, fontSize: 14.5, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.0;
  });
  s.addText('These same four shortcuts cover most of what you will do all day.', {
    x: 0.9, y: 6.15, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 4 · Interface and the Kernel', true);
}

// ---------------------------------------------------------------------
// CHAPTER 5: WHY USE JUPYTER?
// ---------------------------------------------------------------------
chapterDivider({ icon: 'compare', title: 'Why Use Jupyter?', subtitle: 'A fair, honest look at when a notebook helps, and when a plain script is better.' });

// Notebook vs plain script
{
  const s = contentSlide('Chapter 5 · Why Use Jupyter?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'Notebook vs. plain script');

  const rows = [
    { f: 'Feedback', n: 'See output the moment a cell runs', p: 'Run the whole file to see any output' },
    { f: 'Explaining your work', n: 'Markdown text lives beside the code', p: 'Explanation lives in separate comments or docs' },
    { f: 'Charts and tables', n: 'Rendered inline, right under the cell', p: 'Usually opens in a separate window or file' },
    { f: 'Production and version control', n: 'Harder to diff and to deploy directly', p: 'Built for clean history and deployment' },
  ];
  let y = 1.75;
  s.addText('FACTOR', { x: 0.7, y, w: 2.6, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('NOTEBOOK', { x: 3.5, y, w: 4.3, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: BLUE, isTextBox: true, margin: 0 });
  s.addText('PLAIN SCRIPT (.py)', { x: 8.0, y, w: 4.6, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  y += 0.5;
  rows.forEach((r, i) => {
    const rh = 1.0;
    if (i % 2 === 0) s.addShape('rect', { x: 0.7, y, w: 11.9, h: rh, fill: { color: CARD }, line: { type: 'none' } });
    s.addText(r.f, { x: 0.7, y, w: 2.6, h: rh, valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 13, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(r.n, { x: 3.5, y, w: 4.3, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 12, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(r.p, { x: 8.0, y, w: 4.6, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
    y += rh;
  });
}

// Real strengths grid
{
  const s = contentSlide('Chapter 5 · Why Use Jupyter?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'Where notebooks genuinely shine');

  const strengths = [
    { icon: 'chart', t: 'Inline charts', d: 'Plot a graph and see it appear right below your code, no extra window' },
    { icon: 'restart', t: 'Fast iteration', d: 'Change one line and re-run just that cell, instead of the whole program' },
    { icon: 'bulb_list', t: 'A running record', d: 'Your notebook becomes a readable log of how you reached your answer' },
  ];
  const cardW = 3.75, gapX = 0.25, startX = 0.7, y = 1.75, h = 3.3;
  strengths.forEach((st, i) => {
    const x = startX + i * (cardW + gapX);
    s.addShape('roundRect', { x, y, w: cardW, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, st.icon, x + cardW / 2 - 0.4, y + 0.35, 0.8, BLUE);
    s.addText(st.t, { x: x + 0.3, y: y + 1.35, w: cardW - 0.6, h: 0.55, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(st.d, { x: x + 0.3, y: y + 1.95, w: cardW - 0.6, h: 1.2, align: 'center', fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
  });
  s.addShape('roundRect', { x: 0.7, y: 5.3, w: 11.9, h: 0.85, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  iconCircle(s, 'star', 0.9, 5.45, 0.55, GOLD);
  s.addText('Rule of thumb: explore and experiment in a notebook, then move finished, reusable code into a plain .py file.', {
    x: 1.65, y: 5.3, w: 10.7, h: 0.85, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 13, color: WHITE, isTextBox: true, margin: 0,
  });
}

// When not to use notebooks
{
  const s = contentSlide('Chapter 5 · Why Use Jupyter?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'In fairness: not always the right tool');
  s.addText('Notebooks are wonderful for exploring, less so for a few specific jobs.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const cases = [
    { icon: 'warning', t: 'Production systems', d: 'Software that runs unattended is built and deployed as scripts and applications, not notebooks.' },
    { icon: 'warning', t: 'Large team codebases', d: 'Notebook files are hard to review line by line in a pull request, plain code diffs far better.' },
    { icon: 'warning', t: 'Strict, repeatable order', d: 'Because cells can run out of order, a notebook can quietly hide mistakes a script would not.' },
  ];
  const cardW = 3.75, gapX = 0.25, startX = 0.7, y = 2.3, h = 3.5;
  cases.forEach((c, i) => {
    const x = startX + i * (cardW + gapX);
    s.addShape('roundRect', { x, y, w: cardW, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, c.icon, x + cardW / 2 - 0.4, y + 0.35, 0.8, RED);
    s.addText(c.t, { x: x + 0.3, y: y + 1.35, w: cardW - 0.6, h: 0.55, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(c.d, { x: x + 0.3, y: y + 1.95, w: cardW - 0.6, h: 1.4, align: 'center', fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// ---------------------------------------------------------------------
// CHAPTER 6: GOOD HABITS TO KNOW
// ---------------------------------------------------------------------
chapterDivider({ icon: 'gear', title: 'Good Habits to Know', subtitle: 'The small, practical things that keep a notebook trustworthy and easy to share.' });

// Run cells in order
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Always trust "Restart & Run All"');
  s.addText('Because you can run cells in any order, a notebook can look correct while hiding a mistake. Before you trust or share your results, restart and re-run everything top to bottom.', {
    x: 0.7, y: 1.55, w: 11.5, h: 0.8, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  s.addShape('roundRect', { x: 0.7, y: 2.55, w: 5.7, h: 3.3, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'warning', 0.95, 2.75, 0.55, RED);
  s.addText('The hidden trap', { x: 1.65, y: 2.83, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: RED, isTextBox: true, margin: 0 });
  s.addText('You define x = 5 in cell 8, run it, then delete that cell. The notebook still remembers x, until you restart the kernel. Anyone re-running your file from the top would get an error.', {
    x: 0.95, y: 3.4, w: 5.2, h: 2.3, fontFace: BODY_FONT, fontSize: 13, color: MUTED, isTextBox: true, margin: 0,
  });

  s.addShape('roundRect', { x: 6.7, y: 2.55, w: 5.93, h: 3.3, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'check', 6.95, 2.75, 0.55, GREEN);
  s.addText('The fix', { x: 7.65, y: 2.83, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: GREEN, isTextBox: true, margin: 0 });
  s.addText('From the menu, choose Kernel, then Restart & Run All. If it finishes without errors, top to bottom, you know the notebook actually works as written.', {
    x: 6.95, y: 3.4, w: 5.4, h: 1.6, fontFace: BODY_FONT, fontSize: 13, color: MUTED, isTextBox: true, margin: 0,
  });
  codeBlock(s, ['Kernel > Restart & Run All'], 6.95, 5.05, 5.4, codeH(1));
}

// Saving, organizing, exporting
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Naming, saving, and sharing');

  const habits = [
    { icon: 'save', t: 'Save often', d: 'Ctrl+S saves a checkpoint, Jupyter also autosaves every couple of minutes.' },
    { icon: 'layout', t: 'Name it clearly', d: 'sales_report.ipynb beats Untitled4.ipynb once you have a dozen open.' },
    { icon: 'export', t: 'Export when done', d: 'File > Download as turns your notebook into .py, .html, or .pdf to share.' },
    { icon: 'bulb_list', t: 'Keep it focused', d: 'One notebook, one goal. Split a sprawling notebook into smaller ones.' },
  ];
  const cols = 2, cardW = 5.75, cardH = 1.55, gapX = 0.3, gapY = 0.3, startX = 0.7, startY = 1.7;
  habits.forEach((h, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, h.icon, x + 0.25, y + 0.25, 0.55, BLUE);
    s.addText(h.t, { x: x + 0.95, y: y + 0.22, w: cardW - 1.2, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(h.d, { x: x + 0.3, y: y + 0.9, w: cardW - 0.6, h: 0.6, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Common mistakes
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Mistakes almost everyone makes at first');

  const mistakes = [
    { t: 'Running cells out of order', d: 'Jumping around can leave the notebook in a state no one can reproduce.' },
    { t: 'Never restarting the kernel', d: 'Old, deleted, or renamed variables can quietly linger in memory.' },
    { t: 'One giant notebook', d: 'Dozens of unrelated cells become impossible to navigate or trust.' },
    { t: 'Skipping markdown notes', d: 'Code without explanation is hard for your future self, or a teammate, to follow.' },
  ];
  const cols = 2, cardW = 5.75, cardH = 1.55, gapX = 0.3, gapY = 0.3, startX = 0.7, startY = 1.7;
  mistakes.forEach((m, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, 'warning', x + 0.25, y + 0.25, 0.55, RED);
    s.addText(m.t, { x: x + 0.95, y: y + 0.22, w: cardW - 1.2, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(m.d, { x: x + 0.3, y: y + 0.9, w: cardW - 0.6, h: 0.6, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Where to go next
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Where to keep learning');

  const resources = [
    { icon: 'book', t: 'Official Jupyter docs', d: 'jupyter.org/documentation, the most reliable reference there is' },
    { icon: 'gradcap', t: 'Practice sites', d: 'Google Colab and Kaggle both let you run notebooks free, in the browser' },
    { icon: 'users', t: 'This team', d: 'Our next session, plus a shared channel for questions' },
  ];
  const cardW = 3.75, gapX = 0.25, startX = 0.7, y = 1.75, h = 3.2;
  resources.forEach((r, i) => {
    const x = startX + i * (cardW + gapX);
    s.addShape('roundRect', { x, y, w: cardW, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, r.icon, x + cardW / 2 - 0.4, y + 0.35, 0.8, BLUE);
    s.addText(r.t, { x: x + 0.3, y: y + 1.35, w: cardW - 0.6, h: 0.55, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(r.d, { x: x + 0.3, y: y + 1.95, w: cardW - 0.6, h: 1.1, align: 'center', fontFace: BODY_FONT, fontSize: 12, color: MUTED, isTextBox: true, margin: 0 });
  });
  s.addShape('roundRect', { x: 0.7, y: 5.2, w: 11.9, h: 0.85, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  iconCircle(s, 'star', 0.9, 5.35, 0.55, GOLD);
  s.addText('The fastest way to get comfortable is to open a notebook and explore a small dataset this week.', {
    x: 1.65, y: 5.2, w: 10.7, h: 0.85, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 13, color: WHITE, isTextBox: true, margin: 0,
  });
}

// ---------------------------------------------------------------------
// WRAP-UP
// ---------------------------------------------------------------------

// Recap + discussion combined
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  kicker(s, 'Recap', BLUE);
  pageTitle(s, 'What we covered today');

  const chapters = [
    { t: 'Installed Jupyter', icon: 'download' },
    { t: 'Learned what a notebook is', icon: 'question' },
    { t: 'Used code and markdown cells', icon: 'layout' },
    { t: 'Toured the interface and kernel', icon: 'kernel' },
    { t: 'Compared it to plain scripts', icon: 'compare' },
    { t: 'Picked up good habits', icon: 'gear' },
  ];
  const cols = 3, cardW = 3.75, cardH = 1.5, gapX = 0.25, gapY = 0.25, startX = 0.7, startY = 1.7;
  chapters.forEach((c, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, c.icon, x + 0.25, y + 0.44, 0.6, BLUE);
    s.addText(c.t, { x: x + 0.98, y: y + 0.4, w: cardW - 1.2, h: 0.7, valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 13.5, color: TEXT, isTextBox: true, margin: 0 });
  });

  s.addShape('roundRect', { x: 0.7, y: 5.45, w: 11.9, h: 1.15, rectRadius: 0.09, fill: { color: NAVY }, line: { type: 'none' } });
  iconCircle(s, 'comments', 0.9, 5.63, 0.75, GOLD);
  s.addText('Let’s discuss: what would you want to explore first in a notebook of your own?', {
    x: 1.85, y: 5.45, w: 10.5, h: 1.15, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 15, color: WHITE, isTextBox: true, margin: 0,
  });
  footer(s, 'Recap & Discussion', false);
}

// Thank you
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'jupyter_logo', PW / 2 - 0.65, 0.85, 1.3, BLUE);
  s.addText('Thank you!', {
    x: 0, y: 2.55, w: PW, h: 1.1, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 54, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText('You installed Jupyter, ran real cells, and learned how notebooks fit alongside Python scripts. Great second session.', {
    x: PW / 2 - 4.7, y: 3.75, w: 9.4, h: 0.7, align: 'center', fontFace: BODY_FONT, fontSize: 15, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: PW / 2 - 3.4, y: 4.75, w: 6.8, h: 0.6, rectRadius: 0.3, fill: { color: NAVY2 }, line: { color: BLUE, width: 1 } });
  s.addText('Next session: exploring a real dataset together', {
    x: PW / 2 - 3.4, y: 4.75, w: 6.8, h: 0.6, align: 'center', valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, color: GOLD, isTextBox: true, margin: 0,
  });
  footer(s, 'Team Python Onboarding', true);
}

pres.writeFile({ fileName: path.join(__dirname, 'Introduction_to_Jupyter_Notebook.pptx') }).then(() => console.log('final deck written'));
