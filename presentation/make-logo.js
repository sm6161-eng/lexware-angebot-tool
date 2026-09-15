/**
 * Erzeugt aus der Original-Logodatei assets/logo.svg die beiden PNGs,
 * die build.js verwendet:
 *   assets/logo-dark.png    Original (schwarz) für helle Hintergründe
 *   assets/logo-light.png   weiß eingefärbt für dunkle Hintergründe
 *
 * Aufruf: npm run logo
 */
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const ASSETS = path.join(__dirname, 'assets');
const WIDTH = 2000; // Pixelbreite der PNGs (transparenter Hintergrund)

async function main() {
  const svg = fs.readFileSync(path.join(ASSETS, 'logo.svg'), 'utf8');
  const white = svg.replace(/fill:\s*#[0-9a-f]{3,6}/gi, 'fill: #FFFFFF').replace(/fill="#[0-9a-f]{3,6}"/gi, 'fill="#FFFFFF"');
  if (white === svg) throw new Error('Keine Füllfarbe in logo.svg gefunden – weiße Variante kann nicht erzeugt werden.');
  const opts = { density: 600 };
  await sharp(Buffer.from(svg), opts).resize({ width: WIDTH }).png().toFile(path.join(ASSETS, 'logo-dark.png'));
  await sharp(Buffer.from(white), opts).resize({ width: WIDTH }).png().toFile(path.join(ASSETS, 'logo-light.png'));
  console.log('geschrieben: assets/logo-dark.png, assets/logo-light.png');
}

main().catch((e) => { console.error(e); process.exit(1); });
