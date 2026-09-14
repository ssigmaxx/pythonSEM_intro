const pptxgen = require('pptxgenjs');
const path = require('path');

const ICON = (name) => path.join(__dirname, 'icons', `${name}.png`);

// ---- Palette: Python Blue & Gold ----
const NAVY = '132038';       // dominant dark bg (title/dividers)
const NAVY2 = '1B2A4A';      // slightly lighter navy for cards on dark bg
const BLUE = '306998';       // primary python blue
const BLUE_DK = '1F4B70';
const GOLD = 'FFD43B';       // python gold accent
const GOLD_DK = 'C9A227';
const WHITE = 'FFFFFF';
const CARD = 'F4F6F9';       // light cool card bg (not cream)
const TEXT = '1F2937';       // near-black body text on light bg
const MUTED = '5B6472';
const MUTED_LIGHT = 'AEB9CC';
const GREEN = '2C5F2D';
const RED = 'B03A2E';

const HEAD_FONT = 'Arial';
const BODY_FONT = 'Calibri';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
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

// ---------------------------------------------------------------------
// SLIDE 1: TITLE
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'python_logo', PW / 2 - 0.75, 0.9, 1.5, BLUE);
  s.addText('INTRODUCTION TO', {
    x: 0, y: 2.75, w: PW, h: 0.5, align: 'center', fontFace: HEAD_FONT, bold: true,
    fontSize: 18, color: GOLD, charSpacing: 4, isTextBox: true, margin: 0,
  });
  s.addText('Python', {
    x: 0, y: 3.15, w: PW, h: 1.3, align: 'center', fontFace: HEAD_FONT, bold: true,
    fontSize: 66, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText('A friendly, no-experience-needed course for our team', {
    x: 0, y: 4.5, w: PW, h: 0.5, align: 'center', fontFace: BODY_FONT, fontSize: 18,
    color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', {
    x: PW / 2 - 2.6, y: 5.35, w: 5.2, h: 0.55, rectRadius: 0.28,
    fill: { color: NAVY2 }, line: { color: BLUE, width: 1 },
  });
  s.addText('6 chapters  •  hands-on  •  zero prior coding needed', {
    x: PW / 2 - 2.6, y: 5.35, w: 5.2, h: 0.55, align: 'center', valign: 'middle',
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
  s.addText('Six short chapters, from installing Python to knowing what makes it special. No experience required, just curiosity.', {
    x: 0.7, y: 1.55, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  const chapters = [
    { n: '1', t: 'Installing Python', icon: 'download' },
    { n: '2', t: 'What is Python?', icon: 'question' },
    { n: '3', t: 'Your first program', icon: 'code' },
    { n: '4', t: 'Libraries & modules', icon: 'book' },
    { n: '5', t: 'Why Python wins', icon: 'compare' },
    { n: '6', t: 'Good habits to know', icon: 'gear' },
  ];
  const cols = 3, cardW = 3.75, cardH = 1.85, gapX = 0.25, gapY = 0.3;
  const startX = 0.7, startY = 2.35;
  chapters.forEach((c, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, c.icon === 'question' ? 'question' : c.icon, x + 0.28, y + 0.28, 0.62, BLUE);
    // question icon rendered white on white circle looks fine since bg BLUE
    s.addText(`CHAPTER ${c.n}`, { x: x + 1.1, y: y + 0.28, w: cardW - 1.3, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 10.5, color: GOLD_DK, isTextBox: true, margin: 0 });
    s.addText(c.t, { x: x + 1.1, y: y + 0.56, w: cardW - 1.3, h: 0.9, fontFace: HEAD_FONT, bold: true, fontSize: 16, color: TEXT, isTextBox: true, margin: 0 });
  });
  footer(s, 'Chapter 0 · Welcome', false);
}

// ---------------------------------------------------------------------
// SLIDE 3: ICEBREAKER (interactive)
// ---------------------------------------------------------------------
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'comments', 0.7, 0.6, 0.9, BLUE);
  s.addText('LET’S START WITH A QUICK POLL', { x: 1.85, y: 0.68, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Raise your hand: where are you starting from?', { x: 1.85, y: 1.0, w: 10.5, h: 0.7, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const options = [
    { l: 'A', t: 'I have never written a line of code' },
    { l: 'B', t: 'I have dabbled (Excel formulas, a bit of HTML)' },
    { l: 'C', t: 'I know another language, but not Python' },
  ];
  let y = 2.15;
  options.forEach((o) => {
    s.addShape('roundRect', { x: 0.9, y, w: 11.5, h: 1.05, rectRadius: 0.1, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addShape('ellipse', { x: 1.2, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: GOLD }, line: { type: 'none' } });
    s.addText(o.l, { x: 1.2, y: y + 0.25, w: 0.55, h: 0.55, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 18, color: NAVY, isTextBox: true, margin: 0 });
    s.addText(o.t, { x: 2.0, y, w: 10.2, h: 1.05, valign: 'middle', fontFace: BODY_FONT, fontSize: 16.5, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.2;
  });
  s.addText('There is no wrong answer, this course is built to work for all three.', {
    x: 0.9, y: 6.75, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 0 · Welcome', true);
}

// ---------------------------------------------------------------------
// CHAPTER DIVIDER helper
// ---------------------------------------------------------------------
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

// ---------------------------------------------------------------------
// CHAPTER 1: INSTALLING PYTHON
// ---------------------------------------------------------------------
chapterDivider({ icon: 'download', title: 'Installing Python', subtitle: 'Getting Python onto your machine takes about ten minutes. Let’s do it step by step.' });

// Slide: Before you begin
{
  const s = contentSlide('Chapter 1 · Installing Python');
  kicker(s, 'Chapter 1', BLUE);
  pageTitle(s, 'Before you begin');
  s.addText('You only need three things. No special hardware, no prior setup.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const items = [
    { icon: 'check', t: 'A computer', d: 'Windows, Mac, or Linux, all are fully supported' },
    { icon: 'globe', t: 'An internet connection', d: 'Just for the one-time download, about 30 MB' },
    { icon: 'keyboard', t: '10 to 15 minutes', d: 'Enough time to download, install, and verify it works' },
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

function osInstallSlide({ icon, os, steps, note }) {
  const s = contentSlide('Chapter 1 · Installing Python');
  kicker(s, 'Chapter 1', BLUE);
  iconCircle(s, icon, 0.7, 1.55, 0.9, BLUE);
  s.addText(`Installing on ${os}`, { x: 1.8, y: 1.68, w: 9, h: 0.7, fontFace: HEAD_FONT, bold: true, fontSize: 28, color: TEXT, isTextBox: true, margin: 0 });

  let y = 2.8;
  steps.forEach((st, i) => {
    s.addShape('ellipse', { x: 0.7, y, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 0.7, y, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.4, y: y - 0.03, w: 10.8, h: 0.55, valign: 'middle', fontFace: BODY_FONT, fontSize: 14.5, color: TEXT, isTextBox: true, margin: 0 });
    y += 0.68;
  });

  s.addShape('roundRect', { x: 0.7, y: y + 0.15, w: 11.5, h: 0.75, rectRadius: 0.08, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'lightbulb', 0.85, y + 0.25, 0.5, GOLD);
  s.addText(note, { x: 1.5, y: y + 0.15, w: 10.5, h: 0.75, valign: 'middle', fontFace: BODY_FONT, italic: true, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  return s;
}

osInstallSlide({
  icon: 'windows', os: 'Windows', note: 'The checkbox is easy to miss, if you forget it, you can re-run the installer and select "Modify" to add it.',
  steps: [
    'Go to python.org and open Downloads, then click the yellow "Download Python" button',
    'Run the downloaded .exe file',
    'Important: tick the box "Add python.exe to PATH" at the bottom of the first screen',
    'Click "Install Now" and wait for the confirmation screen',
  ],
});

osInstallSlide({
  icon: 'apple', os: 'Mac', note: 'macOS ships with an old Python 2 for internal use only, always install Python 3 yourself from python.org.',
  steps: [
    'Go to python.org, open Downloads, and choose the macOS installer',
    'Open the downloaded .pkg file and click through the installer',
    'Enter your Mac password when prompted, then finish the install',
    'Open the new "Python 3.x" folder in Applications and double-click "Install Certificates.command"',
  ],
});

osInstallSlide({
  icon: 'linux', os: 'Linux', note: 'Most Linux distributions already include Python 3, this step is only needed to confirm or update it.',
  steps: [
    'Open a terminal window',
    'On Ubuntu or Debian, run: sudo apt update  then  sudo apt install python3 python3-pip',
    'On Fedora, run: sudo dnf install python3 python3-pip',
    'Wait for the install to finish, then move on to verifying below',
  ],
});

// Slide: verify + editor
{
  const s = contentSlide('Chapter 1 · Installing Python');
  kicker(s, 'Chapter 1', BLUE);
  pageTitle(s, 'Verify it, then pick an editor');

  // left column: verify
  const lx = 0.7, lw = 5.6, ly = 1.7;
  s.addShape('roundRect', { x: lx, y: ly, w: lw, h: 4.6, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'terminal', lx + 0.3, ly + 0.3, 0.7, BLUE);
  s.addText('Confirm the install', { x: lx + 1.15, y: ly + 0.38, w: lw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('Open a terminal (Command Prompt, Terminal, or your shell) and type:', { x: lx + 0.3, y: ly + 1.15, w: lw - 0.6, h: 0.5, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  s.addShape('roundRect', { x: lx + 0.3, y: ly + 1.65, w: lw - 0.6, h: 0.55, rectRadius: 0.06, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText('python --version', { x: lx + 0.5, y: ly + 1.65, w: lw - 1.0, h: 0.55, valign: 'middle', fontFace: 'Courier New', fontSize: 14, color: GOLD, isTextBox: true, margin: 0 });
  s.addText('You should see something like Python 3.12.4. Then check pip, the tool that installs libraries:', { x: lx + 0.3, y: ly + 2.35, w: lw - 0.6, h: 0.6, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  s.addShape('roundRect', { x: lx + 0.3, y: ly + 3.0, w: lw - 0.6, h: 0.55, rectRadius: 0.06, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText('pip --version', { x: lx + 0.5, y: ly + 3.0, w: lw - 1.0, h: 0.55, valign: 'middle', fontFace: 'Courier New', fontSize: 14, color: GOLD, isTextBox: true, margin: 0 });
  s.addText('On Mac or Linux, use python3 and pip3 if python alone is not found.', { x: lx + 0.3, y: ly + 3.75, w: lw - 0.6, h: 0.6, italic: true, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });

  // right column: editor
  const rx = 6.6, rw = 6.03;
  s.addShape('roundRect', { x: rx, y: ly, w: rw, h: 4.6, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'code', rx + 0.3, ly + 0.3, 0.7, BLUE);
  s.addText('Pick a place to write code', { x: rx + 1.15, y: ly + 0.38, w: rw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });

  const editors = [
    { t: 'VS Code (recommended)', d: 'Free, friendly, works on every OS. Add the "Python" extension from Microsoft.' },
    { t: 'IDLE', d: 'Installed automatically with Python. No setup, great for your very first script.' },
    { t: 'Jupyter Notebook', d: 'Runs code in blocks with instant output. Popular for data and experimentation.' },
  ];
  let ey = ly + 1.15;
  editors.forEach((e) => {
    s.addShape('ellipse', { x: rx + 0.3, y: ey + 0.05, w: 0.18, h: 0.18, fill: { color: GOLD }, line: { type: 'none' } });
    s.addText(e.t, { x: rx + 0.62, y: ey, w: rw - 1.0, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 13.5, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(e.d, { x: rx + 0.62, y: ey + 0.32, w: rw - 1.0, h: 0.55, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
    ey += 1.05;
  });
}

// Slide: Try it now (interactive)
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'rocket', 0.7, 0.65, 0.9, GOLD);
  s.addText('TRY IT NOW', { x: 1.85, y: 0.75, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Two minutes, on your own laptop', { x: 1.85, y: 1.08, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const steps = [
    'Open a terminal (or Command Prompt on Windows)',
    'Type python and press Enter, you should land inside ">>>"',
    'Type 2 + 2 and press Enter, Python should reply 4',
    'Type exit() to leave, then give a thumbs up when you see the 4',
  ];
  let y = 2.15;
  steps.forEach((st, i) => {
    s.addShape('roundRect', { x: 0.9, y, w: 11.5, h: 0.85, rectRadius: 0.1, fill: { color: NAVY2 }, line: { type: 'none' } });
    s.addShape('ellipse', { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 1.15, y: y + 0.17, w: 0.5, h: 0.5, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.85, y, w: 10.4, h: 0.85, valign: 'middle', fontFace: BODY_FONT, fontSize: 14.5, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.0;
  });
  s.addText('Stuck? Raise a hand, we will pair up and fix it together before moving on.', {
    x: 0.9, y: 6.35, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 1 · Installing Python', true);
}

// ---------------------------------------------------------------------
// CHAPTER 2: WHAT IS PYTHON?
// ---------------------------------------------------------------------
chapterDivider({ icon: 'question', title: 'What is Python?', subtitle: 'A little history and the traits that make Python different from other languages.' });

// Slide: What is Python (definition + history)
{
  const s = contentSlide('Chapter 2 · What is Python?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'What is Python, really?');
  s.addText('Python is a high-level, general-purpose programming language. In plain terms: you write instructions in something close to everyday English, and Python turns them into actions on your computer.', {
    x: 0.7, y: 1.55, w: 6.9, h: 1.6, fontFace: BODY_FONT, fontSize: 15, color: TEXT, isTextBox: true, margin: 0,
  });
  s.addText('It was created by Guido van Rossum and first released in 1991. He named it after the British comedy show "Monty Python\'s Flying Circus," not the snake, which is why you will see jokes and easter eggs about comedy sketches in Python\'s own culture.', {
    x: 0.7, y: 3.35, w: 6.9, h: 1.9, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  // stat callouts on right
  const stats = [
    { n: '1991', l: 'Year it was first released' },
    { n: '#1', l: 'Most popular language, several recent industry surveys' },
    { n: '30+', l: 'Years of active, ongoing development' },
  ];
  let sy = 1.6;
  stats.forEach((st) => {
    s.addShape('roundRect', { x: 7.95, y: sy, w: 4.65, h: 1.45, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    s.addText(st.n, { x: 8.2, y: sy + 0.15, w: 2.0, h: 0.9, fontFace: HEAD_FONT, bold: true, fontSize: 34, color: BLUE, isTextBox: true, margin: 0 });
    s.addText(st.l, { x: 8.2, y: sy + 0.95, w: 4.2, h: 0.45, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
    sy += 1.65;
  });
}

// Slide: key characteristics
{
  const s = contentSlide('Chapter 2 · What is Python?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'Four traits worth knowing');

  const traits = [
    { icon: 'code', t: 'Interpreted', d: 'Python runs your code line by line. No separate compile step, you save the file and run it right away.' },
    { icon: 'gradcap', t: 'High-level', d: 'You describe what you want, not how memory or hardware should handle it. Python manages the details for you.' },
    { icon: 'puzzle', t: 'Dynamically typed', d: 'You don\'t declare a variable\'s type up front. Python figures it out from the value you give it.' },
    { icon: 'globe', t: 'General-purpose', d: 'Web apps, data analysis, automation, AI, games: the same language works across nearly every domain.' },
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

// Slide: where python is used
{
  const s = contentSlide('Chapter 2 · What is Python?');
  kicker(s, 'Chapter 2', BLUE);
  pageTitle(s, 'Where Python shows up at work');
  s.addText('The same language you are learning today powers all of these fields.', { x: 0.7, y: 1.55, w: 10, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const domains = [
    { icon: 'chart', t: 'Data & analytics', d: 'Cleaning spreadsheets, building reports, dashboards' },
    { icon: 'robot', t: 'AI & machine learning', d: 'The leading language for building and training models' },
    { icon: 'web', t: 'Web development', d: 'Backends for sites like Instagram and Pinterest' },
    { icon: 'gear', t: 'Automation & scripting', d: 'Renaming files, scraping data, scheduled tasks' },
    { icon: 'flask', t: 'Scientific computing', d: 'Research, simulations, engineering calculations' },
    { icon: 'database', t: 'Everyday tooling', d: 'Internal tools, quick prototypes, glue between systems' },
  ];
  const cols = 3, cardW = 3.75, cardH = 2.0, gapX = 0.25, gapY = 0.25, startX = 0.7, startY = 2.15;
  domains.forEach((d, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, d.icon, x + 0.28, y + 0.28, 0.65, BLUE);
    s.addText(d.t, { x: x + 0.28, y: y + 1.05, w: cardW - 0.56, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 14.5, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(d.d, { x: x + 0.28, y: y + 1.42, w: cardW - 0.56, h: 0.55, fontFace: BODY_FONT, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Slide: Quiz interactive
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'clipboard', 0.7, 0.6, 0.9, GOLD);
  s.addText('QUICK QUIZ', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Which of these use Python?', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const guesses = ['Instagram', 'Netflix', 'Spotify', 'NASA'];
  const cols = 2, cardW = 5.6, cardH = 1.15, gapX = 0.3, gapY = 0.3, startX = 0.9, startY = 2.05;
  guesses.forEach((g, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.1, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addText(g, { x, y, w: cardW, h: cardH, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 19, color: WHITE, isTextBox: true, margin: 0 });
  });
  s.addShape('roundRect', { x: 0.9, y: 4.85, w: 11.0, h: 0.9, rectRadius: 0.1, fill: { color: BLUE_DK }, line: { type: 'none' } });
  s.addText('Answer: all four! Python shows up almost everywhere large-scale software is built.', {
    x: 0.9, y: 4.85, w: 11.0, h: 0.9, align: 'center', valign: 'middle', fontFace: BODY_FONT, bold: true, fontSize: 15, color: GOLD, isTextBox: true, margin: 0,
  });
  s.addText('Shout out any other companies or apps you think use Python.', {
    x: 0.9, y: 6.05, w: 11.0, h: 0.4, italic: true, align: 'center', fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 2 · What is Python?', true);
}

// ---------------------------------------------------------------------
// CHAPTER 3: YOUR FIRST PROGRAM
// ---------------------------------------------------------------------
chapterDivider({ icon: 'code', title: 'Your First Program', subtitle: 'Time to write real Python. We will keep it short and build confidence fast.' });

function codeBlock(slide, lines, x, y, w, h, startLineNo = 1) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  let ly = y + 0.22;
  lines.forEach((ln) => {
    slide.addText(ln, { x: x + 0.35, y: ly, w: w - 0.6, h: 0.4, fontFace: 'Courier New', fontSize: 15, color: ln.color || GOLD, isTextBox: true, margin: 0 });
    ly += 0.42;
  });
}

// Hello World
{
  const s = contentSlide('Chapter 3 · Your First Program');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Hello, World!');
  s.addText('Every programmer\'s first program prints a greeting. Here is the entire program in Python, one line:', {
    x: 0.7, y: 1.55, w: 6.6, h: 0.9, fontFace: BODY_FONT, fontSize: 14.5, color: MUTED, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: 0.7, y: 2.5, w: 6.6, h: 1.0, rectRadius: 0.08, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText([
    { text: 'print', options: { color: '6EC6FF' } },
    { text: '(', options: { color: WHITE } },
    { text: '"Hello, World!"', options: { color: GOLD } },
    { text: ')', options: { color: WHITE } },
  ], { x: 1.0, y: 2.5, w: 6.0, h: 1.0, valign: 'middle', fontFace: 'Courier New', fontSize: 20, isTextBox: true, margin: 0 });

  s.addText('That is it, no semicolons, no setup boilerplate, no "main function" required. Save this as hello.py, then run it:', {
    x: 0.7, y: 3.75, w: 6.6, h: 0.85, fontFace: BODY_FONT, fontSize: 14.5, color: MUTED, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: 0.7, y: 4.6, w: 6.6, h: 0.65, rectRadius: 0.06, fill: { color: CARD }, line: { type: 'none' } });
  s.addText('python hello.py', { x: 0.95, y: 4.6, w: 6.1, h: 0.65, valign: 'middle', fontFace: 'Courier New', fontSize: 15, color: BLUE_DK, isTextBox: true, margin: 0 });

  // right side: compare with java briefly as a teaser
  s.addShape('roundRect', { x: 7.7, y: 1.55, w: 4.9, h: 3.7, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'lightbulb', 7.95, 1.8, 0.6, GOLD);
  s.addText('Notice what\'s missing', { x: 7.95, y: 2.5, w: 4.4, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('No curly braces, no class wrapper, no import for basic printing, and no semicolon at the end of the line. Python trims away ceremony so you can focus on logic.', {
    x: 7.95, y: 3.0, w: 4.35, h: 2.0, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0,
  });
}

// Variables and data types
{
  const s = contentSlide('Chapter 3 · Your First Program');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Variables and data types');
  s.addText('A variable is a name that holds a value. Python figures out the type automatically:', {
    x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  const rows = [
    { code: 'name = "Amara"', type: 'str (text)', note: 'Letters and words, wrapped in quotes' },
    { code: 'age = 29', type: 'int (whole number)', note: 'A number with no decimal point' },
    { code: 'price = 19.99', type: 'float (decimal)', note: 'A number with a decimal point' },
    { code: 'is_member = True', type: 'bool (true/false)', note: 'Only ever True or False' },
  ];
  let y = 2.15;
  // header row
  s.addText('CODE', { x: 0.7, y, w: 4.0, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('TYPE', { x: 4.9, y, w: 2.8, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('WHAT IT MEANS', { x: 7.9, y, w: 4.6, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  y += 0.45;
  rows.forEach((r, i) => {
    const rh = 0.85;
    if (i % 2 === 0) s.addShape('rect', { x: 0.7, y, w: 11.9, h: rh, fill: { color: CARD }, line: { type: 'none' } });
    s.addShape('roundRect', { x: 0.9, y: y + 0.13, w: 3.6, h: 0.6, rectRadius: 0.06, fill: { color: NAVY }, line: { type: 'none' } });
    s.addText(r.code, { x: 1.05, y: y + 0.13, w: 3.3, h: 0.6, valign: 'middle', fontFace: 'Courier New', fontSize: 12.5, color: GOLD, isTextBox: true, margin: 0 });
    s.addText(r.type, { x: 4.9, y, w: 2.8, h: rh, valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 13, color: BLUE, isTextBox: true, margin: 0 });
    s.addText(r.note, { x: 7.9, y, w: 4.6, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
    y += rh;
  });
}

// Indentation matters
{
  const s = contentSlide('Chapter 3 · Your First Program');
  kicker(s, 'Chapter 3', BLUE);
  pageTitle(s, 'Indentation is not optional');
  s.addText('Most languages use curly braces { } to group code. Python uses indentation (spaces) instead, so spacing is part of the syntax, not just style.', {
    x: 0.7, y: 1.55, w: 11.5, h: 0.6, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0,
  });

  // correct
  s.addShape('roundRect', { x: 0.7, y: 2.35, w: 5.7, h: 3.5, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'check', 0.95, 2.55, 0.55, GREEN);
  s.addText('This works', { x: 1.65, y: 2.63, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: GREEN, isTextBox: true, margin: 0 });
  codeBlock(s, [
    'age = 20',
    'if age >= 18:',
    '    print("You can vote")',
    '    print("Welcome!")',
  ], 0.95, 3.25, 5.2, 2.4);

  // incorrect
  s.addShape('roundRect', { x: 6.7, y: 2.35, w: 5.93, h: 3.5, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'warning', 6.95, 2.55, 0.55, RED);
  s.addText('This breaks', { x: 7.65, y: 2.63, w: 4.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: RED, isTextBox: true, margin: 0 });
  codeBlock(s, [
    'age = 20',
    'if age >= 18:',
    'print("You can vote")',
    '    print("Welcome!")',
  ], 6.95, 3.25, 5.4, 2.4);
}

// Try it yourself
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'rocket', 0.7, 0.6, 0.9, GOLD);
  s.addText('TRY IT YOURSELF', { x: 1.85, y: 0.7, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Make it personal', { x: 1.85, y: 1.02, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  s.addShape('roundRect', { x: 0.9, y: 2.1, w: 11.5, h: 1.4, rectRadius: 0.08, fill: { color: NAVY2 }, line: { type: 'none' } });
  s.addText([
    { text: 'name = ', options: { color: WHITE, breakLine: false } },
    { text: '"Your Name"', options: { color: GOLD, breakLine: true } },
    { text: 'print', options: { color: '6EC6FF', breakLine: false } },
    { text: '(', options: { color: WHITE, breakLine: false } },
    { text: '"Hi, I am learning Python, " ', options: { color: GOLD, breakLine: false } },
    { text: '+ name)', options: { color: WHITE, breakLine: false } },
  ], { x: 1.2, y: 2.3, w: 10.9, h: 1.0, fontFace: 'Courier New', fontSize: 16, isTextBox: true, margin: 0 });

  const steps = [
    'Open your editor and create a new file called intro.py',
    'Type the two lines above, swapping in your own name',
    'Run it: python intro.py, and check the message prints correctly',
  ];
  let y = 3.85;
  steps.forEach((st, i) => {
    s.addShape('ellipse', { x: 0.9, y, w: 0.45, h: 0.45, fill: { color: BLUE }, line: { type: 'none' } });
    s.addText(String(i + 1), { x: 0.9, y, w: 0.45, h: 0.45, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 14, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st, { x: 1.5, y: y - 0.03, w: 10.8, h: 0.5, valign: 'middle', fontFace: BODY_FONT, fontSize: 14, color: WHITE, isTextBox: true, margin: 0 });
    y += 0.62;
  });
  s.addText('Bonus: can you print a second line that shows your favorite number?', {
    x: 0.9, y: 6.1, w: 11.5, h: 0.4, italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 3 · Your First Program', true);
}

// ---------------------------------------------------------------------
// CHAPTER 4: LIBRARIES AND MODULES
// ---------------------------------------------------------------------
chapterDivider({ icon: 'book', title: 'Libraries and Modules', subtitle: 'The real reason Python moves fast: you rarely start from a blank page.' });

// What is a library
{
  const s = contentSlide('Chapter 4 · Libraries and Modules');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'What is a library?');
  s.addText('A library (or module) is a collection of pre-written code that solves a common problem, so you don\'t have to write it yourself.', {
    x: 0.7, y: 1.55, w: 6.6, h: 0.9, fontFace: BODY_FONT, fontSize: 15, color: TEXT, isTextBox: true, margin: 0,
  });
  s.addText('Think of it like a toolbox. If you need to hammer a nail, you don\'t forge your own hammer first, you reach for one that already exists. Libraries are the same idea for code: someone already solved the problem, tested it, and packaged it up for anyone to use.', {
    x: 0.7, y: 2.55, w: 6.6, h: 1.8, fontFace: BODY_FONT, fontSize: 13.5, color: MUTED, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: 0.7, y: 4.55, w: 6.6, h: 0.95, rectRadius: 0.08, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'lightbulb', 0.9, 4.72, 0.6, GOLD);
  s.addText('You import a library once at the top of your file, then use everything inside it.', {
    x: 1.65, y: 4.55, w: 5.5, h: 0.95, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0,
  });

  s.addShape('roundRect', { x: 7.7, y: 1.55, w: 4.9, h: 3.95, rectRadius: 0.09, fill: { color: NAVY }, line: { type: 'none' } });
  s.addText('Without a library', { x: 8.0, y: 1.8, w: 4.3, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: MUTED_LIGHT, isTextBox: true, margin: 0 });
  s.addText('Writing this from scratch takes dozens of lines of math you would need to test yourself.', {
    x: 8.0, y: 2.22, w: 4.3, h: 1.0, fontFace: BODY_FONT, italic: true, fontSize: 12.5, color: 'C9D3E0', isTextBox: true, margin: 0,
  });
  s.addText('With a library', { x: 8.0, y: 3.3, w: 4.3, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, isTextBox: true, margin: 0 });
  codeBlock(s, [
    'import math',
    'math.sqrt(64)',
    '# 8.0, done',
  ], 8.0, 3.7, 4.3, 1.7);
}

// Standard vs third-party
{
  const s = contentSlide('Chapter 4 · Libraries and Modules');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'Built-in vs. installed libraries');

  const lx = 0.7, lw = 5.6, y = 1.7, h = 4.3;
  s.addShape('roundRect', { x: lx, y, w: lw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'check', lx + 0.3, y + 0.3, 0.7, GREEN);
  s.addText('Standard library', { x: lx + 1.15, y: y + 0.4, w: lw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('Comes bundled with Python. No installation needed, just import it.', { x: lx + 0.3, y: y + 1.2, w: lw - 0.6, h: 0.55, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  codeBlock(s, ['import math', 'import random', 'import datetime'], lx + 0.3, y + 1.85, lw - 0.6, 1.5);
  s.addText('Examples: math for calculations, random for randomness, datetime for dates and times.', { x: lx + 0.3, y: y + 3.5, w: lw - 0.6, h: 0.7, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });

  const rx = 6.6, rw = 6.03;
  s.addShape('roundRect', { x: rx, y, w: rw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'download', rx + 0.3, y + 0.3, 0.7, BLUE);
  s.addText('Third-party libraries', { x: rx + 1.15, y: y + 0.4, w: rw - 1.4, h: 0.5, fontFace: HEAD_FONT, bold: true, fontSize: 17, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('Written by the community. Install once with pip, Python\'s built-in package manager:', { x: rx + 0.3, y: y + 1.2, w: rw - 0.6, h: 0.55, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  codeBlock(s, ['pip install requests'], rx + 0.3, y + 1.85, rw - 0.6, 0.65);
  s.addText('Then use it in any script:', { x: rx + 0.3, y: y + 2.65, w: rw - 0.6, h: 0.35, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
  codeBlock(s, ['import requests'], rx + 0.3, y + 3.0, rw - 0.6, 0.65);
}

// Popular libraries grid
{
  const s = contentSlide('Chapter 4 · Libraries and Modules');
  kicker(s, 'Chapter 4', BLUE);
  pageTitle(s, 'A few libraries you will hear about');

  const libs = [
    { icon: 'chart', t: 'NumPy & Pandas', d: 'Work with numbers and spreadsheet-style data' },
    { icon: 'chart', t: 'Matplotlib', d: 'Turn data into charts and graphs' },
    { icon: 'web', t: 'Requests', d: 'Talk to websites and web services' },
    { icon: 'web', t: 'Flask & Django', d: 'Build websites and web applications' },
    { icon: 'robot', t: 'TensorFlow & PyTorch', d: 'Build and train AI models' },
    { icon: 'gear', t: 'openpyxl', d: 'Read and write Excel files automatically' },
  ];
  const cols = 3, cardW = 3.75, cardH = 2.0, gapX = 0.25, gapY = 0.25, startX = 0.7, startY = 1.75;
  libs.forEach((l, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX), y = startY + row * (cardH + gapY);
    s.addShape('roundRect', { x, y, w: cardW, h: cardH, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
    iconCircle(s, l.icon, x + 0.28, y + 0.28, 0.65, BLUE);
    s.addText(l.t, { x: x + 0.28, y: y + 1.05, w: cardW - 0.56, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 14, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(l.d, { x: x + 0.28, y: y + 1.42, w: cardW - 0.56, h: 0.55, fontFace: BODY_FONT, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  });
}

// Match the library interactive
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'puzzle', 0.7, 0.55, 0.9, GOLD);
  s.addText('GROUP ACTIVITY', { x: 1.85, y: 0.65, w: 8, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText('Match the library to its job', { x: 1.85, y: 0.97, w: 10, h: 0.6, fontFace: HEAD_FONT, bold: true, fontSize: 26, color: WHITE, isTextBox: true, margin: 0 });

  const left = ['Pandas', 'Requests', 'Matplotlib', 'Django'];
  const right = ['Build a website', 'Draw a chart', 'Fetch data from a URL', 'Analyze spreadsheet data'];
  let y = 2.05;
  left.forEach((l) => {
    s.addShape('roundRect', { x: 0.9, y, w: 4.4, h: 0.85, rectRadius: 0.09, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addText(l, { x: 0.9, y, w: 4.4, h: 0.85, align: 'center', valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 15.5, color: GOLD, isTextBox: true, margin: 0 });
    y += 1.05;
  });
  y = 2.05;
  right.forEach((r) => {
    s.addShape('roundRect', { x: 6.9, y, w: 5.5, h: 0.85, rectRadius: 0.09, fill: { color: NAVY2 }, line: { color: BLUE_DK, width: 1 } });
    s.addText(r, { x: 6.9, y, w: 5.5, h: 0.85, align: 'center', valign: 'middle', fontFace: BODY_FONT, fontSize: 14.5, color: WHITE, isTextBox: true, margin: 0 });
    y += 1.05;
  });
  s.addText('Call it out loud: which pair goes together? (Answers, in order: Analyze data, Fetch from a URL, Draw a chart, Build a website)', {
    x: 0.9, y: 6.35, w: 11.5, h: 0.5, italic: true, fontFace: BODY_FONT, fontSize: 12, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  footer(s, 'Chapter 4 · Libraries and Modules', true);
}

// ---------------------------------------------------------------------
// CHAPTER 5: WHY PYTHON?
// ---------------------------------------------------------------------
chapterDivider({ icon: 'compare', title: 'Why Python?', subtitle: 'A fair, honest look at what makes Python a great default choice, and when it is not.' });

// Same task less code
{
  const s = contentSlide('Chapter 5 · Why Python?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'Same task, very different amounts of code');
  s.addText('Here is "Hello, World!" in three languages. Same result, very different effort.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const blocks = [
    { lang: 'Java', lines: ['public class Main {', '  public static void', '   main(String[] args) {', '    System.out.println(', '     "Hello, World!");', '  }', '}'] },
    { lang: 'C++', lines: ['#include <iostream>', 'int main() {', '  std::cout <<', '   "Hello, World!";', '  return 0;', '}'] },
    { lang: 'Python', lines: ['print("Hello, World!")'] },
  ];
  const colW = 3.83, gapX = 0.2, startX = 0.7, y = 2.15, h = 3.9;
  blocks.forEach((b, i) => {
    const x = startX + i * (colW + gapX);
    const isPy = b.lang === 'Python';
    s.addShape('roundRect', { x, y, w: colW, h, rectRadius: 0.09, fill: { color: isPy ? BLUE_DK : NAVY }, line: { type: 'none' } });
    s.addText(b.lang, { x: x + 0.25, y: y + 0.18, w: colW - 0.5, h: 0.4, fontFace: HEAD_FONT, bold: true, fontSize: 15, color: isPy ? GOLD : MUTED_LIGHT, isTextBox: true, margin: 0 });
    let ly = y + 0.75;
    b.lines.forEach((ln) => {
      s.addText(ln, { x: x + 0.25, y: ly, w: colW - 0.5, h: 0.35, fontFace: 'Courier New', fontSize: 11.5, color: isPy ? WHITE : 'C9D3E0', isTextBox: true, margin: 0 });
      ly += 0.36;
    });
    s.addText(`${b.lines.length} line${b.lines.length > 1 ? 's' : ''}`, { x: x + 0.25, y: y + h - 0.5, w: colW - 0.5, h: 0.35, italic: true, fontFace: BODY_FONT, fontSize: 11.5, color: isPy ? GOLD : MUTED_LIGHT, isTextBox: true, margin: 0 });
  });
}

// Why teams choose python (comparison table)
{
  const s = contentSlide('Chapter 5 · Why Python?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'Why teams choose Python');

  const rows = [
    { f: 'Readability', p: 'Reads close to plain English', o: 'More symbols and boilerplate' },
    { f: 'Learning curve', p: 'Beginners write real code in days', o: 'Often takes weeks to get productive' },
    { f: 'Libraries', p: 'Huge ecosystem for almost any task', o: 'Varies widely by language' },
    { f: 'Community', p: 'Millions of users, answers everywhere', o: 'Smaller or more specialized' },
  ];
  let y = 1.75;
  s.addText('', { x: 0.7, y, w: 3.6, h: 0.35 });
  s.addText('FACTOR', { x: 0.7, y, w: 3.4, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  s.addText('PYTHON', { x: 4.6, y, w: 4.1, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: BLUE, isTextBox: true, margin: 0 });
  s.addText('MANY OTHER LANGUAGES', { x: 8.9, y, w: 3.7, h: 0.35, fontFace: HEAD_FONT, bold: true, fontSize: 11, color: MUTED, isTextBox: true, margin: 0 });
  y += 0.5;
  rows.forEach((r, i) => {
    const rh = 0.95;
    if (i % 2 === 0) s.addShape('rect', { x: 0.7, y, w: 11.9, h: rh, fill: { color: CARD }, line: { type: 'none' } });
    s.addText(r.f, { x: 0.7, y, w: 3.4, h: rh, valign: 'middle', fontFace: HEAD_FONT, bold: true, fontSize: 13.5, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(r.p, { x: 4.6, y, w: 4.1, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(r.o, { x: 8.9, y, w: 3.7, h: rh, valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0 });
    y += rh;
  });
}

// When python isn't the best fit
{
  const s = contentSlide('Chapter 5 · Why Python?');
  kicker(s, 'Chapter 5', BLUE);
  pageTitle(s, 'In fairness: not always the right tool');
  s.addText('No language is best at everything, and a good developer knows Python\'s limits too.', { x: 0.7, y: 1.55, w: 11, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: MUTED, isTextBox: true, margin: 0 });

  const cases = [
    { icon: 'warning', t: 'Raw speed matters', d: 'Compiled languages like C++ or Rust run faster for heavy, low-level computation such as game engines.' },
    { icon: 'warning', t: 'Mobile apps', d: 'Native iOS and Android apps are usually built in Swift, Kotlin, or a cross-platform framework, not Python.' },
    { icon: 'warning', t: 'Very small footprint devices', d: 'Some embedded or memory-constrained hardware favors lighter languages like C.' },
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
chapterDivider({ icon: 'gear', title: 'Good Habits to Know', subtitle: 'The small, practical things that will save you time and confusion early on.' });

// Running code + pip/venv combined
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Running code, and keeping projects tidy');

  const lx = 0.7, lw = 5.6, y = 1.7, h = 4.3;
  s.addShape('roundRect', { x: lx, y, w: lw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'terminal', lx + 0.3, y + 0.3, 0.7, BLUE);
  s.addText('Three ways to run Python', { x: lx + 1.15, y: y + 0.4, w: lw - 1.4, h: 0.55, fontFace: HEAD_FONT, bold: true, fontSize: 15.5, color: TEXT, isTextBox: true, margin: 0 });
  const ways = [
    { t: 'Terminal', d: 'python filename.py, direct and simple' },
    { t: 'Editor "Run" button', d: 'VS Code and most editors run the file for you' },
    { t: 'Notebook cells', d: 'Jupyter runs one block at a time, great for exploring' },
  ];
  let wy = y + 1.25;
  ways.forEach((w) => {
    s.addShape('ellipse', { x: lx + 0.3, y: wy + 0.05, w: 0.16, h: 0.16, fill: { color: GOLD }, line: { type: 'none' } });
    s.addText(w.t, { x: lx + 0.6, y: wy, w: lw - 0.9, h: 0.3, fontFace: HEAD_FONT, bold: true, fontSize: 13, color: TEXT, isTextBox: true, margin: 0 });
    s.addText(w.d, { x: lx + 0.6, y: wy + 0.3, w: lw - 0.9, h: 0.5, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0 });
    wy += 0.92;
  });

  const rx = 6.6, rw = 6.03;
  s.addShape('roundRect', { x: rx, y, w: rw, h, rectRadius: 0.09, fill: { color: CARD }, line: { type: 'none' } });
  iconCircle(s, 'folder', rx + 0.3, y + 0.3, 0.7, BLUE);
  s.addText('pip and virtual environments', { x: rx + 1.15, y: y + 0.4, w: rw - 1.4, h: 0.55, fontFace: HEAD_FONT, bold: true, fontSize: 15.5, color: TEXT, isTextBox: true, margin: 0 });
  s.addText('pip installs libraries. A virtual environment is a private, isolated copy of Python for one project, so its libraries don\'t clash with another project\'s.', {
    x: rx + 0.3, y: y + 1.2, w: rw - 0.6, h: 1.0, fontFace: BODY_FONT, fontSize: 12.5, color: MUTED, isTextBox: true, margin: 0,
  });
  codeBlock(s, [
    'python -m venv env',
    'source env/bin/activate',
    'pip install requests',
  ], rx + 0.3, y + 2.25, rw - 0.6, 1.6);
  s.addText('(On Windows, activate with: env\\Scripts\\activate)', { x: rx + 0.3, y: y + 3.95, w: rw - 0.6, h: 0.3, italic: true, fontFace: BODY_FONT, fontSize: 10.5, color: MUTED, isTextBox: true, margin: 0 });
}

// Common mistakes
{
  const s = contentSlide('Chapter 6 · Good Habits to Know');
  kicker(s, 'Chapter 6', BLUE);
  pageTitle(s, 'Mistakes almost everyone makes at first');

  const mistakes = [
    { t: 'Mixing tabs and spaces', d: 'Pick spaces (4 is standard) and let your editor handle it consistently.' },
    { t: 'Forgetting the colon', d: 'Lines like if, for, and def all end with a colon before the indented block.' },
    { t: 'Comparing with = instead of ==', d: '= assigns a value, == checks whether two things are equal.' },
    { t: 'Wrong indentation level', d: 'Lines in the same block must line up exactly, Python is strict about this.' },
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
    { icon: 'book', t: 'Official Python docs', d: 'docs.python.org, the most reliable reference there is' },
    { icon: 'gradcap', t: 'Practice sites', d: 'HackerRank, Codewars, and Python.org\'s own beginner tutorial' },
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
  s.addText('The fastest way to learn Python is to write it. Aim for 15 minutes a day over a big weekend push.', {
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
    { n: '1', t: 'Installed Python', icon: 'download' },
    { n: '2', t: 'Learned what Python is', icon: 'question' },
    { n: '3', t: 'Wrote real code', icon: 'code' },
    { n: '4', t: 'Explored libraries', icon: 'book' },
    { n: '5', t: 'Compared it to other languages', icon: 'compare' },
    { n: '6', t: 'Picked up good habits', icon: 'gear' },
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
  s.addText('Let\'s discuss: what is one thing you want to build with Python once you\'re comfortable?', {
    x: 1.85, y: 5.45, w: 10.5, h: 1.15, valign: 'middle', italic: true, fontFace: BODY_FONT, fontSize: 15, color: WHITE, isTextBox: true, margin: 0,
  });
  footer(s, 'Recap & Discussion', false);
}

// Thank you
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, 'python_logo', PW / 2 - 0.65, 0.85, 1.3, BLUE);
  s.addText('Thank you!', {
    x: 0, y: 2.55, w: PW, h: 1.1, align: 'center', fontFace: HEAD_FONT, bold: true, fontSize: 54, color: WHITE, isTextBox: true, margin: 0,
  });
  s.addText('You installed Python, wrote real code, and learned what makes it special. That is a strong first session.', {
    x: PW / 2 - 4.5, y: 3.75, w: 9, h: 0.7, align: 'center', fontFace: BODY_FONT, fontSize: 15, color: MUTED_LIGHT, isTextBox: true, margin: 0,
  });
  s.addShape('roundRect', { x: PW / 2 - 3.0, y: 4.75, w: 6.0, h: 0.6, rectRadius: 0.3, fill: { color: NAVY2 }, line: { color: BLUE, width: 1 } });
  s.addText('Next session: writing your first small project', {
    x: PW / 2 - 3.0, y: 4.75, w: 6.0, h: 0.6, align: 'center', valign: 'middle', fontFace: BODY_FONT, fontSize: 12.5, color: GOLD, isTextBox: true, margin: 0,
  });
  footer(s, 'Team Python Onboarding', true);
}

pres.writeFile({ fileName: path.join(__dirname, 'Introduction_to_Python.pptx') }).then(() => console.log('final deck written'));
