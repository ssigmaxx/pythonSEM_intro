const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const fa = require('react-icons/fa');
const fa6 = require('react-icons/fa6');
const si = require('react-icons/si');

const icons = {
  jupyter_logo: { comp: si.SiJupyter, color: 'FFFFFF' },
  markdown: { comp: fa6.FaMarkdown ? fa6.FaMarkdown : fa.FaMarkdown, color: '306998' },
  play: { comp: fa.FaPlay, color: '306998' },
  kernel: { comp: fa.FaMicrochip, color: '306998' },
  restart: { comp: fa.FaRotate ? fa.FaRotate : fa.FaSync, color: '306998' },
  save: { comp: fa.FaFloppyDisk ? fa.FaFloppyDisk : fa.FaSave, color: '306998' },
  layout: { comp: fa.FaTableCells ? fa.FaTableCells : fa.FaThLarge, color: '306998' },
  share: { comp: fa.FaShareNodes ? fa.FaShareNodes : fa.FaShareAlt, color: '306998' },
  export: { comp: fa.FaFileExport, color: '306998' },
  bulb_list: { comp: fa.FaListCheck ? fa.FaListCheck : fa.FaListUl, color: '306998' },
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
