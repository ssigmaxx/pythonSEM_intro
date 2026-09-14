const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const fa = require('react-icons/fa');
const fa6 = require('react-icons/fa6');

// name -> { comp, color }
const icons = {
  python_logo: { comp: fa.FaPython, color: 'FFFFFF' },
  windows: { comp: fa.FaWindows, color: 'FFFFFF' },
  apple: { comp: fa.FaApple, color: 'FFFFFF' },
  linux: { comp: fa.FaLinux, color: 'FFFFFF' },
  download: { comp: fa.FaDownload, color: '306998' },
  terminal: { comp: fa.FaTerminal, color: '306998' },
  check: { comp: fa.FaCheckCircle, color: '2C5F2D' },
  code: { comp: fa.FaCode, color: '306998' },
  book: { comp: fa.FaBook, color: '306998' },
  compare: { comp: fa6.FaScaleBalanced, color: '306998' },
  lightbulb: { comp: fa.FaLightbulb, color: 'B8860B' },
  warning: { comp: fa.FaTriangleExclamation ? fa.FaTriangleExclamation : fa.FaExclamationTriangle, color: 'B03A2E' },
  gradcap: { comp: fa.FaGraduationCap, color: '306998' },
  puzzle: { comp: fa.FaPuzzlePiece, color: '306998' },
  chart: { comp: fa.FaChartColumn ? fa.FaChartColumn : fa.FaChartBar, color: '306998' },
  robot: { comp: fa.FaRobot, color: '306998' },
  gear: { comp: fa.FaGears ? fa.FaGears : fa.FaCogs, color: '306998' },
  globe: { comp: fa.FaGlobe, color: '306998' },
  folder: { comp: fa.FaFolderOpen, color: '306998' },
  question: { comp: fa.FaCircleQuestion ? fa.FaCircleQuestion : fa.FaQuestionCircle, color: 'FFFFFF' },
  rocket: { comp: fa.FaRocket, color: 'FFFFFF' },
  users: { comp: fa.FaUsers, color: '306998' },
  database: { comp: fa.FaDatabase, color: '306998' },
  clipboard: { comp: fa.FaClipboardCheck, color: 'FFFFFF' },
  comments: { comp: fa.FaComments, color: 'FFFFFF' },
  flask: { comp: fa.FaFlask, color: '306998' },
  web: { comp: fa.FaEarthAmericas ? fa.FaEarthAmericas : fa.FaGlobeAmericas, color: '306998' },
  arrow: { comp: fa.FaArrowRight, color: 'FFFFFF' },
  star: { comp: fa.FaStar, color: 'FFD43B' },
  keyboard: { comp: fa.FaKeyboard, color: '306998' },
};

const outDir = path.join(__dirname, 'icons');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

async function run() {
  for (const [name, { comp, color }] of Object.entries(icons)) {
    if (!comp) { console.error('MISSING COMPONENT', name); continue; }
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(comp, { size: 512, color: '#' + color })
    );
    const buf = await sharp(Buffer.from(svg), { density: 300 }).resize(512, 512).png().toBuffer();
    fs.writeFileSync(path.join(outDir, `${name}.png`), buf);
    console.log('wrote', name);
  }
}

run().catch(e => { console.error(e); process.exit(1); });
