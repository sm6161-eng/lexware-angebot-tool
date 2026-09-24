/**
 * Maiershirts – PowerPoint-Master (Folienvorlage) mit Logo
 *
 * Erzeugt dist/Maiershirts_Master.pptx (Beispieldeck, jedes Layout einmal)
 * und dist/Maiershirts_Master.potx (Vorlage: Doppelklick öffnet eine neue
 * Präsentation mit allen Layouts und Platzhaltern).
 *
 * Logo: assets/logo-dark.png (helle Folien) und assets/logo-light.png (dunkle
 * Folien), erzeugt aus assets/logo.svg mit `npm run logo`.
 *
 * Aufruf: npm run build
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const JSZip = require('jszip');

// ---------- Farben: Maiershirts Brand Colors (Brand Book Mai 2026) ----------
const C = {
  dark: '0F0F0F',       // Schwarz – Text, Struktur, dunkle Blöcke
  green: '819E72',      // Sage – Akzent (Pantone 7492 C)
  greenDark: '677E5B',  // Sage Dark – Akzent-Text, Hover
  greenLight: 'A3C095', // Sage Light – Badges, Pfeile, leichte Akzente
  beige: 'E8DFD0',      // Warm Sand – Sektionsflächen, Karten
  beigeDark: 'D9D2C5',  // Sand Dark – Linien, Tabellenrahmen
  cream: 'FAF7F2',      // Cream – helle Flächen auf Sand
  white: 'FFFFFF',
};
// Gedämpfter Text bleibt in der Palette: Schwarz bzw. Weiß mit Transparenz.
const MUTED_ON_LIGHT = { color: C.dark, transparency: 40 };
const MUTED_ON_DARK = { color: C.white, transparency: 35 };

// ---------- Typografie ----------
const FONT = 'Inter';
const FS = { title: 28, titleBig: 40, section: 36, h: 16, body: 16, sub: 12, label: 10 }; // nichts unter 12 pt außer Label/Fußzeile

const DIST = path.join(__dirname, 'dist');
const ASSETS = path.join(__dirname, 'assets');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'maiershirts-'));

// ---------- Logo ----------
function logoPath(kind) {
  const p = path.join(ASSETS, `logo-${kind}.png`);
  if (!fs.existsSync(p)) throw new Error(`${p} fehlt – bitte "npm run logo" ausführen.`);
  return p;
}
const LOGO_DARK = logoPath('dark');   // schwarz, auf hellem Grund
const LOGO_LIGHT = logoPath('light'); // weiß, auf dunklem Grund
const LOGO = { small: [0.95, 0.5], big: [1.6, 0.88] }; // einheitliche Größen (w, h)

function logo(kind, x, y, size) {
  const [w, h] = LOGO[size];
  return { image: { path: kind === 'dark' ? LOGO_DARK : LOGO_LIGHT, x, y, w, h, sizing: { type: 'contain', w, h } } };
}

// ---------- Hintergründe (Flächen des Layouts als Bild, 240 px = 1 Zoll) ----------
const PX = 240;
const rr = (x, y, w, h, fill, r = 0.12) =>
  `<rect x="${x * PX}" y="${y * PX}" width="${w * PX}" height="${h * PX}" rx="${r * PX}" fill="#${fill}"/>`;
const circle = (cx, cy, r, fill) => `<circle cx="${cx * PX}" cy="${cy * PX}" r="${r * PX}" fill="#${fill}"/>`;
const arrow = (x1, x2, y, color) =>
  `<line x1="${x1 * PX}" y1="${y * PX}" x2="${(x2 - 0.12) * PX}" y2="${y * PX}" stroke="#${color}" stroke-width="5"/>` +
  `<polygon points="${(x2 - 0.14) * PX},${(y - 0.07) * PX} ${x2 * PX},${y * PX} ${(x2 - 0.14) * PX},${(y + 0.07) * PX}" fill="#${color}"/>`;
// Das Motiv der dunklen Layouts: eine Bergflanke aus dem Logo, rechts angeschnitten, deckend in Sage Dark.
const MOTIF = `<polygon points="${8.2 * PX},${5.625 * PX} ${10 * PX},${1.9 * PX} ${10 * PX},${5.625 * PX}" fill="#${C.greenDark}"/>`;

async function background(name, base, inner = '') {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${10 * PX}" height="${5.625 * PX}" viewBox="0 0 ${10 * PX} ${5.625 * PX}">` +
    `<rect width="100%" height="100%" fill="#${base}"/>${inner}</svg>`;
  const out = path.join(TMP, `bg-${name}.png`);
  await sharp(Buffer.from(svg)).png({ compressionLevel: 9 }).toFile(out);
  return out;
}

// ---------- Platzhalterbilder für das Beispieldeck ----------
function qrPattern(size) {
  // deterministisches QR-ähnliches Muster (kein echter Code)
  let seed = 42;
  const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  const n = 21, m = size / n, cells = [];
  const finder = (r, c) => (r < 7 && c < 7) || (r < 7 && c >= n - 7) || (r >= n - 7 && c < 7);
  for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) {
    let on;
    if (finder(r, c)) {
      const rr2 = r < 7 ? r : r - (n - 7), cc = c < 7 ? c : c - (n - 7);
      on = rr2 === 0 || rr2 === 6 || cc === 0 || cc === 6 || (rr2 >= 2 && rr2 <= 4 && cc >= 2 && cc <= 4);
    } else on = rnd() > 0.55;
    if (on) cells.push(`<rect x="${c * m}" y="${r * m}" width="${m}" height="${m}" fill="#${C.dark}"/>`);
  }
  return cells.join('');
}

async function placeholderImage(kind, wIn, hIn) {
  const w = Math.round(wIn * PX), h = Math.round(hIn * PX);
  const font = `font-family="Helvetica, Arial, sans-serif"`;
  let inner = '';
  if (kind === 'foto') {
    const s = Math.min(w, h) * 0.22, cx = w / 2, cy = h / 2 - s * 0.25;
    inner = `<rect width="100%" height="100%" fill="#${C.beige}"/>` +
      `<rect x="${cx - s}" y="${cy - s * 0.7}" width="${2 * s}" height="${1.4 * s}" rx="${s * 0.1}" fill="none" stroke="#${C.greenDark}" stroke-width="${s * 0.08}"/>` +
      `<polygon points="${cx - s * 0.8},${cy + s * 0.5} ${cx - s * 0.25},${cy - s * 0.15} ${cx + s * 0.15},${cy + s * 0.25} ${cx + s * 0.4},${cy} ${cx + s * 0.8},${cy + s * 0.5}" fill="#${C.greenDark}"/>` +
      `<circle cx="${cx + s * 0.45}" cy="${cy - s * 0.35}" r="${s * 0.14}" fill="#${C.greenDark}"/>` +
      `<text x="${cx}" y="${cy + s * 1.35}" ${font} font-size="${Math.max(18, s * 0.42)}" fill="#${C.greenDark}" text-anchor="middle">Foto einfügen</text>`;
  } else if (kind === 'logo') {
    inner = `<rect width="100%" height="100%" fill="#${C.white}"/>` +
      `<text x="${w / 2}" y="${h / 2 + 8}" ${font} font-size="${Math.round(h * 0.18)}" fill="#${C.beigeDark}" text-anchor="middle" font-weight="bold">KUNDENLOGO</text>`;
  } else if (kind === 'qr') {
    const pad = Math.round(w * 0.08);
    inner = `<rect width="100%" height="100%" fill="#${C.white}"/><g transform="translate(${pad},${pad})">${qrPattern(w - 2 * pad)}</g>`;
  } else if (kind === 'portrait') {
    const cx = w / 2;
    inner = `<rect width="100%" height="100%" fill="#${C.beige}"/>` +
      `<circle cx="${cx}" cy="${h * 0.38}" r="${h * 0.17}" fill="#${C.greenDark}"/>` +
      `<path d="M ${cx - h * 0.34} ${h * 0.95} A ${h * 0.34} ${h * 0.34} 0 0 1 ${cx + h * 0.34} ${h * 0.95} Z" fill="#${C.greenDark}"/>`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${inner}</svg>`;
  const out = path.join(TMP, `ph-${kind}-${w}x${h}.png`);
  await sharp(Buffer.from(svg)).png().toFile(out);
  return out;
}

// ---------- Layout-Bausteine ----------
const ph = (name, type, text, o) => ({ placeholder: { options: { name, type, fontFace: FONT, margin: 0, ...o }, text } });
// Geometrie der Bildplatzhalter (x, y, w, h) – wird vom Layout und vom Beispieldeck genutzt
const G = {
  bild: [0.5, 1.35, 4.4, 3.5],
  voll: [0, 0, 10, 4.2],
  textil: (i) => [0.5 + i * 3.05, 1.35, 2.9, 1.9],
  logo: (r, c) => [0.62 + c * 2.3, 1.47 + r * 1.2, 1.86, 0.81],
  logoGrid: (g, r, c) => [g.x0 + c * g.dx, g.y0 + r * g.dy, g.w, g.h],
  portrait: [0.6, 3.45, 1.05, 1.05],
  projekt: [[0.5, 1.35, 4.4, 3.5], [5.1, 1.35, 2.15, 1.65], [7.45, 1.35, 2.05, 1.65], [5.1, 3.2, 4.4, 1.65]],
  kunde: { hero: [4.0, 1.35, 5.5, 2.0], klein: [0, 1, 2].map((i) => [4.0 + i * 1.885, 3.5, 1.73, 1.3]) },
  kollektion: { cols: 6, rows: 2, tile: (i) => [0.5 + (i % 6) * 1.525, 1.35 + Math.floor(i / 6) * 1.72, 1.375, 1.12] },
  qr: [6.45, 3.5, 1.0, 1.0],
};
// Referenz-Raster: 4 x 3 breite Felder (Unternehmen) und 4 x 2 hohe Felder (Vereinswappen)
const REF_GRID = {
  breit: { cols: 4, rows: 4, cellW: 2.1, cellH: 0.78, x0: 0.62, y0: 1.47, dx: 2.3, dy: 0.86, w: 1.86, h: 0.54 },
  hoch: { cols: 4, rows: 2, cellW: 2.1, cellH: 1.6, x0: 0.62, y0: 1.47, dx: 2.3, dy: 1.75, w: 1.86, h: 1.05, caption: { dy: 1.17, h: 0.3 } },
};
const pic = (name, [x, y, w, h], prompt = 'Foto einfügen') => ph(name, 'pic', prompt, { x, y, w, h, fontSize: FS.sub, align: 'center', valign: 'middle', ...MUTED_ON_LIGHT });
const txt = (text, o) => ({ text: { text, options: { fontFace: FONT, margin: 0, ...o } } });
const footer = () => txt('maiershirts.de', { x: 0.5, y: 5.15, w: 3, h: 0.3, fontSize: FS.label, ...MUTED_ON_LIGHT });
const footerDark = () => txt('maiershirts.de', { x: 0.6, y: 5.15, w: 3, h: 0.3, fontSize: FS.label, ...MUTED_ON_DARK });
const slideNumber = { x: 9.0, y: 5.15, w: 0.5, h: 0.3, fontFace: FONT, fontSize: FS.label, color: C.dark, align: 'right', margin: 0 };
const title = (o = {}) => ph('title', 'title', 'Folientitel', { x: 0.5, y: 0.35, w: 7.8, h: 0.7, fontSize: FS.title, bold: true, color: C.dark, align: 'left', valign: 'middle', ...o });
const bodyPh = (x, y, w, h) => ph('body', 'body', 'Text eingeben', { x, y, w, h, fontSize: FS.body, color: C.dark, align: 'left', valign: 'top' });

// Musterinhalte (nur im Beispieldeck, alles Platzhalterwerte)
const KUNDE = 'SV Musterstadt e. V. · Trikotsatz Saison 2027';
// Positionskorrekturen für platzhaltergebundene Bilder (siehe Referenzfolie)
const PIC_FIX = [];
const EMU = (inch) => Math.round(inch * 914400);

async function main() {
  fs.mkdirSync(DIST, { recursive: true });

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9'; // 10" x 5.625"
  pres.author = 'Maiershirts';
  pres.company = 'Maiershirts';
  pres.title = 'Maiershirts – Präsentationsmaster';
  pres.lang = 'de-DE';

  // ---------- Hintergründe ----------
  const BG = {
    dark: await background('dark', C.dark, MOTIF),
    darkPlain: await background('dark-plain', C.dark),
    kennzahlen: await background('kennzahlen', C.beige, [0, 1, 2].map((i) => rr(0.5 + i * 3.05, 1.5, 2.9, 2.4, C.white)).join('')),
    prozess: await background('prozess', C.white,
      [0, 1, 2, 3].map((i) => circle(0.6 + i * 2.3 + 0.5, 2.25, 0.5, C.beige) + (i < 3 ? arrow(0.6 + i * 2.3 + 1.15, 0.6 + (i + 1) * 2.3 - 0.15, 2.25, C.greenLight) : '')).join('')),
    preisstaffel: await background('preisstaffel', C.beige, [0, 1, 2, 3].map((i) => rr(0.5 + i * 2.3, 1.5, 2.15, 2.2, C.white)).join('')),
    verfahren: await background('verfahren', C.white, [0, 1, 2, 3].map((i) => rr(0.5 + i * 2.3, 1.35, 2.15, 3.2, C.beige)).join('')),
    angebot: await background('angebot', C.white, rr(0.5, 4.05, 9, 0.8, C.beige)),
    referenzen: await background('referenzen', C.beige,
      [0, 1, 2, 3].flatMap((r) => [0, 1, 2, 3].map((c) => rr(0.5 + c * 2.3, 1.35 + r * 0.86, 2.1, 0.78, C.white, 0.08))).join('')),
    referenzenHoch: await background('referenzen-hoch', C.beige,
      [0, 1].flatMap((r) => [0, 1, 2, 3].map((c) => rr(0.5 + c * 2.3, 1.35 + r * 1.75, 2.1, 1.6, C.white, 0.08))).join('')),
    abschluss: await background('abschluss', C.dark, MOTIF + rr(6.2, 1.5, 3.3, 3.2, C.beige)),
  };

  // ======================================================================
  // Layouts
  // ======================================================================

  // Titelfolie (dunkel)
  pres.defineSlideMaster({
    title: 'MS_TITEL',
    background: { path: BG.dark },
    objects: [
      logo('light', 0.6, 0.5, 'big'),
      ph('title', 'title', 'Titel der Präsentation', { x: 0.6, y: 1.85, w: 7.2, h: 1.3, fontSize: FS.titleBig, bold: true, color: C.white, align: 'left', valign: 'bottom' }),
      ph('sub', 'body', 'Untertitel', { x: 0.6, y: 3.2, w: 7.2, h: 0.7, fontSize: 18, align: 'left', valign: 'top', ...MUTED_ON_DARK }),
      ph('kunde', 'body', 'Angebot für: Kunde · Projekt', { x: 0.6, y: 4.15, w: 7.2, h: 0.4, fontSize: FS.h, bold: true, color: C.white, align: 'left', valign: 'middle' }),
      ph('meta', 'body', 'Datum · Angebotsnummer', { x: 0.6, y: 4.6, w: 7.2, h: 0.35, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_DARK }),
    ],
  });

  // Abschnittsfolie (dunkel)
  pres.defineSlideMaster({
    title: 'MS_ABSCHNITT',
    background: { path: BG.dark },
    objects: [
      logo('light', 0.6, 0.5, 'big'),
      ph('num', 'body', '01', { x: 0.6, y: 1.5, w: 3, h: 0.95, fontSize: 54, bold: true, color: C.green, align: 'left', valign: 'bottom' }),
      ph('title', 'title', 'Abschnittstitel', { x: 0.6, y: 2.5, w: 7.2, h: 0.9, fontSize: FS.section, bold: true, color: C.white, align: 'left', valign: 'top' }),
      ph('sub', 'body', 'Kurzbeschreibung des Abschnitts', { x: 0.6, y: 3.45, w: 7.2, h: 0.6, fontSize: FS.h, align: 'left', valign: 'top', ...MUTED_ON_DARK }),
    ],
  });

  // Inhaltsfolien (Titel + Textkörper), weiß und Sand
  for (const [name, bg] of [['MS_INHALT', C.white], ['MS_INHALT_SAND', C.beige]]) {
    pres.defineSlideMaster({
      title: name,
      background: { color: bg },
      objects: [logo('dark', 8.55, 0.4, 'small'), title(), bodyPh(0.5, 1.35, 9, 3.5), footer()],
      slideNumber,
    });
  }
  // Freie Folien (nur Titel) für Karten, Tabellen, Diagramme
  for (const [name, bg] of [['MS_FREI', C.white], ['MS_FREI_SAND', C.beige]]) {
    pres.defineSlideMaster({
      title: name,
      background: { color: bg },
      objects: [logo('dark', 8.55, 0.4, 'small'), title(), footer()],
      slideNumber,
    });
  }

  // Bild links, Text rechts
  pres.defineSlideMaster({
    title: 'MS_BILD',
    background: { color: C.white },
    objects: [logo('dark', 8.55, 0.4, 'small'), title(), pic('foto', G.bild), bodyPh(5.3, 1.35, 4.2, 3.5), footer()],
    slideNumber,
  });

  // Vollbild-Foto mit dunkler Textleiste
  pres.defineSlideMaster({
    title: 'MS_BILD_VOLL',
    background: { path: BG.darkPlain },
    objects: [
      pic('foto', G.voll),
      logo('light', 8.55, 4.6, 'small'),
      ph('title', 'title', 'Bildunterschrift', { x: 0.6, y: 4.35, w: 7.6, h: 0.55, fontSize: 24, bold: true, color: C.white, align: 'left', valign: 'middle' }),
      ph('sub', 'body', 'Ergänzung', { x: 0.6, y: 4.9, w: 7.6, h: 0.35, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_DARK }),
    ],
  });

  // Kennzahlen (3 Karten auf Sand)
  pres.defineSlideMaster({
    title: 'MS_KENNZAHLEN',
    background: { path: BG.kennzahlen },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ...[0, 1, 2].flatMap((i) => {
        const x = 0.75 + i * 3.05;
        return [
          ph(`kpi${i + 1}`, 'body', '1.000', { x, y: 1.7, w: 2.4, h: 0.95, fontSize: 44, bold: true, color: C.greenDark, align: 'left', valign: 'middle' }),
          ph(`kpi${i + 1}_label`, 'body', 'Was die Zahl bedeutet', { x, y: 2.7, w: 2.4, h: 0.6, fontSize: FS.h, bold: true, color: C.dark, align: 'left', valign: 'top' }),
          ph(`kpi${i + 1}_note`, 'body', 'Zeitraum / Quelle', { x, y: 3.35, w: 2.4, h: 0.4, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
        ];
      }),
      ph('hinweis', 'body', 'Hinweis', { x: 0.5, y: 4.25, w: 9, h: 0.4, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
      footer(),
    ],
    slideNumber,
  });

  // Prozess (4 Schritte)
  pres.defineSlideMaster({
    title: 'MS_PROZESS',
    background: { path: BG.prozess },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ...[0, 1, 2, 3].flatMap((i) => {
        const x = 0.6 + i * 2.3;
        return [
          txt(String(i + 1), { x, y: 1.75, w: 1.0, h: 1.0, fontSize: 24, bold: true, color: C.greenDark, align: 'center', valign: 'middle' }),
          ph(`schritt${i + 1}`, 'body', `Schritt ${i + 1}`, { x: x - 0.55, y: 3.0, w: 2.1, h: 0.35, fontSize: FS.h, bold: true, color: C.dark, align: 'center', valign: 'top' }),
          ph(`schritt${i + 1}_text`, 'body', 'Was in diesem Schritt passiert', { x: x - 0.55, y: 3.4, w: 2.1, h: 1.0, fontSize: FS.sub, align: 'center', valign: 'top', ...MUTED_ON_LIGHT }),
        ];
      }),
      footer(),
    ],
    slideNumber,
  });

  // Verfahrensvergleich (4 Karten: Digitaldruck, Siebdrucktransfer, Spezialtransfer, Stickerei)
  pres.defineSlideMaster({
    title: 'MS_VERFAHREN',
    background: { path: BG.verfahren },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ph('sub', 'body', 'Welche Technik zu welchem Projekt passt', { x: 0.5, y: 1.0, w: 7.8, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
      ...[0, 1, 2, 3].flatMap((i) => {
        const x = 0.68, cx = 0.5 + i * 2.3 + 0.18, w = 1.8;
        return [
          txt(`0${i + 1}`, { x: cx, y: 1.5, w, h: 0.25, fontSize: FS.label, bold: true, color: C.greenDark, charSpacing: 1 }),
          ph(`v${i + 1}_name`, 'body', 'Verfahren', { x: cx, y: 1.75, w, h: 0.35, fontSize: FS.h, bold: true, color: C.dark, align: 'left', valign: 'middle' }),
          ph(`v${i + 1}_ideal`, 'body', 'Ideal für …', { x: cx, y: 2.12, w, h: 1.2, fontSize: FS.sub, color: C.dark, align: 'left', valign: 'top' }),
          ...['Wäsche', 'Auflage', 'Farbe'].flatMap((label, k) => [
            txt(label, { x: cx, y: 3.38 + k * 0.36, w: 0.7, h: 0.34, fontSize: FS.label, ...MUTED_ON_LIGHT, valign: 'middle' }),
            ph(`v${i + 1}_f${k + 1}`, 'body', '…', { x: cx + 0.7, y: 3.38 + k * 0.36, w: w - 0.7, h: 0.34, fontSize: FS.sub, bold: true, color: C.dark, align: 'left', valign: 'middle' }),
          ]),
        ];
      }),
      ph('fazit', 'body', 'Kurz entschieden: …', { x: 0.5, y: 4.65, w: 9, h: 0.45, fontSize: FS.sub, color: C.dark, align: 'left', valign: 'top' }),
      footer(),
    ],
    slideNumber,
  });

  // Projekte aus der Praxis (1 großes + 3 kleine Fotos, Bildunterschrift im Foto)
  pres.defineSlideMaster({
    title: 'MS_PROJEKTE',
    background: { color: C.white },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ph('sub', 'body', 'Was wir zuletzt umgesetzt haben', { x: 0.5, y: 1.0, w: 7.8, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
      ...G.projekt.flatMap(([x, y, w, h], i) => [
        pic(`projekt${i + 1}`, [x, y, w, h]),
        ph(`projekt${i + 1}_name`, 'body', 'Kunde', { x: x + 0.15, y: y + h - (i === 0 ? 0.62 : 0.52), w: w - 0.3, h: i === 0 ? 0.32 : 0.28, fontSize: i === 0 ? FS.h : 13, bold: true, color: C.white, align: 'left', valign: 'bottom' }),
        ph(`projekt${i + 1}_text`, 'body', 'Was wir gemacht haben', { x: x + 0.15, y: y + h - (i === 0 ? 0.32 : 0.26), w: w - 0.3, h: i === 0 ? 0.25 : 0.22, fontSize: FS.label, color: C.white, align: 'left', valign: 'top' }),
      ]),
      footer(),
    ],
    slideNumber,
  });

  // Projekt im Detail (ein Kunde: Text links, Kollektionsfoto und drei Einzelteile rechts)
  pres.defineSlideMaster({
    title: 'MS_PROJEKT_KUNDE',
    background: { color: C.white },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ph('sub', 'body', 'Branche · Ort', { x: 0.5, y: 1.0, w: 7.8, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
      bodyPh(0.5, 1.35, 3.2, 3.6),
      pic('hero', G.kunde.hero),
      ...G.kunde.klein.flatMap(([x, y, w, h], i) => [
        pic(`teil${i + 1}`, [x, y, w, h]),
        ph(`teil${i + 1}_name`, 'body', 'Teil', { x, y: y + h + 0.03, w, h: 0.25, fontSize: FS.label, bold: true, color: C.dark, align: 'left', valign: 'top' }),
      ]),
      footer(),
    ],
    slideNumber,
  });

  // Kollektionsübersicht (12 Artikel-Kacheln: Mockup, Name, Preis)
  pres.defineSlideMaster({
    title: 'MS_KOLLEKTION',
    background: { color: C.white },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ph('sub', 'body', 'Was die Kollektion ausmacht', { x: 0.5, y: 1.0, w: 7.8, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
      ...Array.from({ length: 12 }, (_, i) => i).flatMap((i) => {
        const [x, y, w, h] = G.kollektion.tile(i);
        return [
          pic(`artikel${i + 1}`, [x, y, w, h]),
          ph(`artikel${i + 1}_name`, 'body', 'Artikel', { x, y: y + h + 0.02, w, h: 0.34, fontSize: FS.label, bold: true, color: C.dark, align: 'left', valign: 'top' }),
          ph(`artikel${i + 1}_preis`, 'body', '0,00 €', { x, y: y + h + 0.36, w, h: 0.2, fontSize: FS.label, bold: true, color: C.greenDark, align: 'left', valign: 'top' }),
        ];
      }),
      footer(),
    ],
    slideNumber,
  });

  // Textil-Auswahl (3 Rohlinge)
  pres.defineSlideMaster({
    title: 'MS_TEXTIL',
    background: { color: C.white },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ...[0, 1, 2].flatMap((i) => {
        const x = 0.5 + i * 3.05;
        return [
          pic(`textil${i + 1}_foto`, G.textil(i)),
          ph(`textil${i + 1}`, 'body', 'Textil', { x, y: 3.35, w: 2.9, h: 0.35, fontSize: FS.h, bold: true, color: C.dark, align: 'left', valign: 'top' }),
          ph(`textil${i + 1}_details`, 'body', 'Material · Grammatur · Farben · Größen', { x, y: 3.7, w: 2.9, h: 0.75, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
          ph(`textil${i + 1}_preis`, 'body', 'ab 0,00 €', { x, y: 4.45, w: 2.9, h: 0.35, fontSize: FS.h, bold: true, color: C.greenDark, align: 'left', valign: 'top' }),
        ];
      }),
      footer(),
    ],
    slideNumber,
  });

  // Preisstaffel (4 Karten auf Sand)
  pres.defineSlideMaster({
    title: 'MS_PREISSTAFFEL',
    background: { path: BG.preisstaffel },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ...[0, 1, 2, 3].flatMap((i) => {
        const x = 0.7 + i * 2.3;
        return [
          ph(`stufe${i + 1}`, 'body', 'ab 10 Stück', { x, y: 1.7, w: 1.75, h: 0.4, fontSize: 14, bold: true, color: C.dark, align: 'left', valign: 'middle' }),
          ph(`stufe${i + 1}_preis`, 'body', '0,00 €', { x, y: 2.15, w: 1.75, h: 0.8, fontSize: 32, bold: true, color: C.greenDark, align: 'left', valign: 'middle' }),
          ph(`stufe${i + 1}_note`, 'body', 'pro Stück, netto', { x, y: 3.0, w: 1.75, h: 0.55, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
        ];
      }),
      ph('hinweis', 'body', 'Hinweise zur Staffel', { x: 0.5, y: 4.0, w: 9, h: 0.8, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
      footer(),
    ],
    slideNumber,
  });

  // Angebot (Tabelle als Inhalt, Konditionen als Platzhalter)
  pres.defineSlideMaster({
    title: 'MS_ANGEBOT',
    background: { path: BG.angebot },
    objects: [
      logo('dark', 8.55, 0.4, 'small'), title(),
      ...['Gültig bis', 'Lieferzeit', 'Zahlung'].flatMap((label, i) => {
        const x = 0.7 + i * 3.0;
        return [
          txt(label.toUpperCase(), { x, y: 4.12, w: 2.7, h: 0.25, fontSize: FS.label, bold: true, color: C.greenDark, charSpacing: 1 }),
          ph(`kond${i + 1}`, 'body', label, { x, y: 4.37, w: 2.7, h: 0.4, fontSize: 14, bold: true, color: C.dark, align: 'left', valign: 'middle' }),
        ];
      }),
      ph('hinweis', 'body', 'Preise netto zzgl. MwSt.', { x: 0.5, y: 4.88, w: 9, h: 0.25, fontSize: FS.label, align: 'left', valign: 'top', ...MUTED_ON_LIGHT }),
      footer(),
    ],
    slideNumber,
  });

  // Referenzen: Logo-Raster 4 x 4 (breite Logos) und 4 x 2 (Vereinswappen), auf Sand
  for (const [name, bg, grid] of [['MS_REFERENZEN', BG.referenzen, REF_GRID.breit], ['MS_REFERENZEN_WAPPEN', BG.referenzenHoch, REF_GRID.hoch]]) {
    pres.defineSlideMaster({
      title: name,
      background: { path: bg },
      objects: [
        logo('dark', 8.55, 0.4, 'small'), title(),
        ph('sub', 'body', 'Eine Auswahl unserer Kunden', { x: 0.5, y: 1.0, w: 7.8, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
        ...Array.from({ length: grid.rows }, (_, r) => r).flatMap((r) => Array.from({ length: grid.cols }, (_, c) => c).flatMap((c) => {
          const n = r * grid.cols + c + 1, [x, y, w] = G.logoGrid(grid, r, c);
          const items = [pic(`logo${n}`, G.logoGrid(grid, r, c), 'Logo einfügen')];
          if (grid.caption) items.push(ph(`name${n}`, 'body', 'Vereinsname', { x, y: y + grid.caption.dy, w, h: grid.caption.h, fontSize: 11, bold: true, color: C.dark, align: 'center', valign: 'middle' }));
          return items;
        })),
        footer(),
      ],
      slideNumber,
    });
  }

  // Abschluss: Ansprechpartner + nächster Schritt (dunkel)
  pres.defineSlideMaster({
    title: 'MS_ABSCHLUSS',
    background: { path: BG.abschluss },
    objects: [
      logo('light', 0.6, 0.5, 'big'),
      ph('title', 'title', 'Lassen Sie uns starten', { x: 0.6, y: 1.5, w: 5.3, h: 1.1, fontSize: 32, bold: true, color: C.white, align: 'left', valign: 'bottom' }),
      ph('sub', 'body', 'Was wir als Nächstes gemeinsam tun', { x: 0.6, y: 2.65, w: 5.3, h: 0.6, fontSize: FS.h, align: 'left', valign: 'top', ...MUTED_ON_DARK }),
      pic('portrait', G.portrait, 'Porträt einfügen'),
      ph('name', 'body', 'Vorname Nachname', { x: 1.85, y: 3.45, w: 4.0, h: 0.35, fontSize: FS.h, bold: true, color: C.white, align: 'left', valign: 'top' }),
      ph('rolle', 'body', 'Ihr Ansprechpartner', { x: 1.85, y: 3.8, w: 4.0, h: 0.3, fontSize: FS.sub, align: 'left', valign: 'top', ...MUTED_ON_DARK }),
      ph('kontakt', 'body', 'Telefon · E-Mail', { x: 1.85, y: 4.1, w: 4.0, h: 0.5, fontSize: FS.sub, color: C.white, align: 'left', valign: 'top' }),
      txt('NÄCHSTER SCHRITT', { x: 6.45, y: 1.7, w: 2.8, h: 0.3, fontSize: FS.label, bold: true, color: C.greenDark, charSpacing: 1 }),
      ph('schritt', 'body', 'Was jetzt zu tun ist, bis wann', { x: 6.45, y: 2.0, w: 2.8, h: 1.35, fontSize: 14, color: C.dark, align: 'left', valign: 'top' }),
      pic('qr', G.qr, 'QR-Code einfügen'),
      ph('qr_text', 'body', 'Anfrage online', { x: 7.6, y: 3.5, w: 1.75, h: 1.0, fontSize: FS.sub, align: 'left', valign: 'middle', ...MUTED_ON_LIGHT }),
      footerDark(),
    ],
  });

  // ======================================================================
  // Beispieldeck
  // ======================================================================
  const T = (s, text) => s.addText(text, { placeholder: 'title', isTextBox: true });
  const P = (s, name, text, o = {}) => s.addText(text, { placeholder: name, isTextBox: true, ...o });
  // Bild in einen Bildplatzhalter: echtes Foto aus assets/fotos/<foto>.jpg (auf das Feld zugeschnitten), sonst Platzhalterbild
  const FOTOS = path.join(ASSETS, 'fotos');
  const IMG = async (s, name, kind, [x, y, w, h], foto, { gradient = false, fit = 'cover' } = {}) => {
    const real = foto && ['jpg', 'jpeg', 'png'].map((e) => path.join(FOTOS, `${foto}.${e}`)).find((f) => fs.existsSync(f));
    let file;
    if (real) {
      file = path.join(TMP, `foto-${foto.replace(/[\/\\]/g, '_')}-${Math.round(w * 100)}x${Math.round(h * 100)}.jpg`);
      const pw = Math.round(w * 300), phh = Math.round(h * 300);
      let img = sharp(real).resize(pw, phh, { fit, position: 'centre', background: '#FFFFFF' });
      if (gradient) { // dunkler Verlauf unten, damit weiße Bildunterschriften lesbar bleiben
        const gh = Math.round(phh * 0.55);
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${pw}" height="${phh}"><defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#0F0F0F" stop-opacity="0"/><stop offset="1" stop-color="#0F0F0F" stop-opacity="0.85"/></linearGradient></defs><rect x="0" y="${phh - gh}" width="${pw}" height="${gh}" fill="url(#g)"/></svg>`;
        img = sharp(await img.jpeg({ quality: 95 }).toBuffer()).composite([{ input: Buffer.from(svg), top: 0, left: 0 }]);
      }
      await img.jpeg({ quality: 88 }).toFile(file);
    } else {
      file = await placeholderImage(kind, w, h);
    }
    s.addImage({ placeholder: name, path: file, x, y, w, h });
  };

  // Aufbau: gemeinsamer Einstieg (neutral) → Abschnitt „Für Vereine“ (du) → Abschnitt „Für Unternehmen“ (Sie) → Abschluss
  const KONTAKT_TEL = '+49 171 9005694', KONTAKT_MAIL = 'office@maiershirts.de';

  // Wiederverwendbare Bausteine
  const projekteSlide = async (file, titel, unterzeile) => {
    const listFile = path.join(FOTOS, file);
    const projekte = fs.existsSync(listFile) ? JSON.parse(fs.readFileSync(listFile, 'utf8')) : [];
    const pages = Math.max(1, Math.ceil(projekte.length / 4));
    for (let page = 0; page < pages; page++) {
      const s = pres.addSlide({ masterName: 'MS_PROJEKTE' });
      T(s, pages > 1 ? `${titel} (${page + 1}/${pages})` : titel);
      P(s, 'sub', unterzeile);
      for (let i = 0; i < 4; i++) {
        const pr = projekte[page * 4 + i];
        if (!pr) continue;
        await IMG(s, `projekt${i + 1}`, 'foto', G.projekt[i], pr.foto, { gradient: true });
        P(s, `projekt${i + 1}_name`, pr.kunde); P(s, `projekt${i + 1}_text`, pr.text);
      }
      s.addNotes(`Layout MS_PROJEKTE: ein großes und drei kleine Fotos mit Kunde und Kurztext im Bild. Einträge in assets/fotos/${file}, Fotos daneben.`);
    }
  };
  const referenzSlides = async (group, titel, layout, grid, unterzeile) => {
    const refDir = path.join(ASSETS, 'referenzen', group);
    const refs = fs.existsSync(refDir) ? fs.readdirSync(refDir).filter((f) => /\.(png|jpe?g|svg)$/i.test(f)).sort() : [];
    const namenFile = path.join(refDir, 'namen.json');
    const namen = fs.existsSync(namenFile) ? JSON.parse(fs.readFileSync(namenFile, 'utf8')) : {};
    const nameOf = (f) => namen[f] || f.replace(/^\d+-/, '').replace(/\.[^.]+$/, '').replace(/-/g, ' ');
    const cells = grid.cols * grid.rows;
    const pages = Math.max(1, Math.ceil(refs.length / cells));
    for (let page = 0; page < pages; page++) {
      const pageRefs = refs.slice(page * cells, (page + 1) * cells);
      const s = pres.addSlide({ masterName: layout });
      T(s, pages > 1 ? `${titel} (${page + 1}/${pages})` : titel);
      P(s, 'sub', unterzeile);
      for (let i = 0; i < cells; i++) {
        if (!pageRefs[i]) continue; // freie Felder bleiben leere Karten
        const [x, y, w, h] = G.logoGrid(grid, Math.floor(i / grid.cols), i % grid.cols);
        // Ränder abschneiden, Logo proportional einpassen und zentrieren (pptxgenjs kennt die Bildmaße nicht)
        const pad = 0.14, boxW = w - 2 * pad, boxH = h - 2 * pad;
        const file = path.join(TMP, `ref-${group}-${page}-${i}.png`);
        await sharp(path.join(refDir, pageRefs[i])).trim({ threshold: 25 }).png().toFile(file);
        const meta = await sharp(file).metadata();
        const scale = Math.min(boxW / meta.width, boxH / meta.height);
        const lw = meta.width * scale, lh = meta.height * scale;
        const lx = x + pad + (boxW - lw) / 2, ly = y + pad + (boxH - lh) / 2;
        s.addImage({ placeholder: `logo${i + 1}`, path: file, x: lx, y: ly, w: lw, h: lh });
        // pptxgenjs setzt bei Platzhalterbildern immer die Platzhalterposition; die zentrierte Position wird im Paket nachgetragen
        const phObj = s._slideLayout._slideObjects.find((o) => o.options && o.options.placeholder === `logo${i + 1}`);
        PIC_FIX.push({ slideNum: s._slideNum, idx: phObj.options._placeholderIdx, x: lx, y: ly });
        if (grid.caption) P(s, `name${i + 1}`, nameOf(pageRefs[i]));
      }
      s.addNotes(`Layout ${layout}: ${cells} Logo-Felder, ${pageRefs.length} belegt aus assets/referenzen/${group}. Logo per Klick auf ein Feld einsetzen; bei Beschnitt: Bildformat → Zuschneiden → Anpassen.`);
    }
  };
  const abschnitt = (num, titel, sub) => {
    const s = pres.addSlide({ masterName: 'MS_ABSCHNITT' });
    P(s, 'num', num); T(s, titel); P(s, 'sub', sub);
    s.addNotes('Layout MS_ABSCHNITT: Nummer, Titel, Kurzbeschreibung.');
  };

  // ---------- Gemeinsamer Einstieg ----------
  {
    const s = pres.addSlide({ masterName: 'MS_TITEL' });
    T(s, 'Textilveredelung aus Ammerbuch');
    P(s, 'sub', 'Druck und Stickerei für Vereine und Unternehmen. Aus einer Hand, aus der Region.');
    P(s, 'kunde', 'Firmenpräsentation');
    P(s, 'meta', 'September 2026 · maiershirts.de');
    s.addNotes('Layout MS_TITEL. Für ein Kundenangebot: Feld „Firmenpräsentation“ durch „Angebot für: Kunde · Projekt“ ersetzen, Datum und Angebotsnummer eintragen.');
  }
  {
    const s = pres.addSlide({ masterName: 'MS_FREI_SAND' });
    T(s, 'Was uns ausmacht');
    s.addText('Maiershirts veredelt Textilien für Vereine, Unternehmen und Events. Vom einzelnen Shirt bis zur kompletten Teamausstattung: Beratung, Textil, Druck oder Stickerei und Versand kommen aus einer Hand.',
      { x: 0.5, y: 1.35, w: 4.1, h: 1.6, fontFace: FONT, fontSize: FS.body, color: C.dark, valign: 'top', margin: 0, isTextBox: true });
    s.addText([
      { text: 'Beratung zu Textil, Verfahren und Motiv', options: { bullet: true, breakLine: true } },
      { text: 'Druck und Stickerei im eigenen Haus', options: { bullet: true, breakLine: true } },
      { text: 'Ab 1 Stück, Serien mit exaktem Farb-Match', options: { bullet: true, breakLine: true } },
      { text: 'Motive und Artikel bleiben hinterlegt', options: { bullet: true } },
    ], { x: 0.5, y: 3.0, w: 4.1, h: 1.8, fontFace: FONT, fontSize: FS.body, color: C.dark, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });
    const cards = [
      ['Textil', 'Markenqualität von Joma, B&C, Stanley/Stella und weiteren Herstellern'],
      ['Veredelung', 'Digitaldruck, Siebdrucktransfer, Spezialtransfer und Stickerei'],
      ['Vereinsshop', 'Eigene Shop-Seite je Verein: online bestellen statt Listen sammeln'],
      ['Nachbestellung', 'Motive, Größen und Artikel bleiben im System, jederzeit abrufbar'],
    ];
    cards.forEach(([head, desc], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 5.0 + col * 2.3, y = 1.35 + row * 1.8;
      s.addShape('roundRect', { x, y, w: 2.15, h: 1.65, fill: { color: C.white }, line: { color: C.white }, rectRadius: 0.12 });
      s.addShape('line', { x: x + 0.25, y: y + 0.3, w: 0.4, h: 0, line: { color: C.green, width: 3 } });
      s.addText(head, { x: x + 0.25, y: y + 0.45, w: 1.7, h: 0.35, fontFace: FONT, fontSize: FS.h, bold: true, color: C.dark, margin: 0, isTextBox: true });
      s.addText(desc, { x: x + 0.25, y: y + 0.82, w: 1.7, h: 0.75, fontFace: FONT, fontSize: FS.sub, ...MUTED_ON_LIGHT, valign: 'top', margin: 0, isTextBox: true });
    });
    s.addNotes('Layout MS_FREI_SAND: Text links, vier Karten rechts.');
  }
  {
    const s = pres.addSlide({ masterName: 'MS_BILD_VOLL' });
    await IMG(s, 'foto', 'foto', G.voll, 'vollbild');
    T(s, 'Ein Motiv, das sitzt.');
    P(s, 'sub', 'Maiershirts-Logo auf Polo-Shirt, aus der eigenen Produktion');
    s.addNotes('Layout MS_BILD_VOLL: Foto über die volle Breite, Bildunterschrift darunter.');
  }
  {
    const s = pres.addSlide({ masterName: 'MS_PROZESS' });
    T(s, 'So läuft ein Auftrag ab');
    [['Anfrage', 'Textil, Menge und Motiv per Mail, Telefon oder über den Shop'], ['Angebot', 'Angebot und Mockup zur Freigabe innerhalb von 48 Stunden'], ['Produktion', 'Druck oder Stickerei nach Freigabe, im eigenen Haus'], ['Lieferung', 'Versand oder Abholung zum vereinbarten Termin']].forEach(([head, desc], i) => {
      P(s, `schritt${i + 1}`, head); P(s, `schritt${i + 1}_text`, desc);
    });
    s.addNotes('Layout MS_PROZESS: vier nummerierte Schritte.');
  }
  {
    const s = pres.addSlide({ masterName: 'MS_VERFAHREN' });
    T(s, 'Vier Verfahren, ein Anspruch');
    P(s, 'sub', 'Ehrlich mit Stärken und Grenzen, damit jedes Motiv da landet, wo es am besten sitzt.');
    const verfahren = [
      ['Digitaldruck', 'Fotos, viele Farben, Namen und Nummern. Kleine Mengen, gemischte Größen, fast jedes Material.', ['bis 60 °C', 'ab 1 Stück', 'Vollfarbe']],
      ['Siebdrucktransfer', 'Serien mit gleichem Motiv: Team-, Vereins- und Workwear. Exaktes Farb-Match, lagerbar für Nachbestellungen.', ['50+ Wäschen', 'ab 25 Stück', 'Sonderfarben']],
      ['Spezialtransfer', 'Berufskleidung mit Industriewäsche: Handwerk, Pflege, Gastro. Übersteht, wo normale Drucke versagen.', ['bis 90 °C', 'ab 50 Stück', 'Vollton']],
      ['Stickerei', 'Polos, Caps, Jacken und Workwear. Edler, erhabener Look für Logos. Schrift ab ca. 6 mm Höhe.', ['bis 95 °C', 'ab 1 Stück', 'Garnfarben']],
    ];
    verfahren.forEach(([name, ideal, facts], i) => {
      P(s, `v${i + 1}_name`, name); P(s, `v${i + 1}_ideal`, ideal);
      facts.forEach((f, k) => P(s, `v${i + 1}_f${k + 1}`, f));
    });
    s.addText([
      { text: 'Kurz entschieden: ', options: { bold: true } },
      { text: 'Foto und viele Farben → Digitaldruck · Exakte Farbe in Serie → Siebdrucktransfer · Industriewäsche → Spezialtransfer · Edler Logo-Look → Stickerei. Auf Anfrage: Sublimation und Siebdruck direkt.' },
    ], { placeholder: 'fazit', isTextBox: true });
    s.addNotes('Layout MS_VERFAHREN: vier Verfahren mit Einsatzgebiet und Kennwerten. Inhalte aus dem Veredelungs-Leitfaden. Handmuster-Karte zum Termin mitnehmen.');
  }

  // ---------- Abschnitt 01: Für Vereine (du) ----------
  abschnitt('01', 'Für Vereine', 'Trikots, Teamwear und ein eigener Vereinsshop. Ohne Sammelliste, ohne Vorkasse durch den Verein.');
  {
    const s = pres.addSlide({ masterName: 'MS_PROZESS' });
    T(s, 'Der Vereinsshop: bestellen statt Listen sammeln');
    [
      ['Kollektion festlegen', 'Ihr wählt Teile, Farben und Logo-Platzierung, wir bauen die Kollektion'],
      ['Shop geht online', 'Eigene Seite auf maiershirts.de mit eurem Vereinslogo, Link und QR-Code'],
      ['Jeder bestellt selbst', 'Größe wählen, bezahlen, fertig. Kein Sammeln, kein Vorstrecken'],
      ['Abholung im Verein', 'Wir produzieren, der Verein holt gesammelt ab oder erhält alles auf einmal. Nachbestellung jederzeit'],
    ].forEach(([head, desc], i) => { P(s, `schritt${i + 1}`, head); P(s, `schritt${i + 1}_text`, desc); });
    s.addNotes('Layout MS_PROZESS als Vereinsshop-Folie. Ansprache „ihr“, passend zum Vereinsteil.');
  }
  {
    // Kollektionsübersichten aus assets/fotos/kollektionen/<name>.json, je Gruppe eine Folie
    const kDir = path.join(FOTOS, 'kollektionen');
    const kFiles = fs.existsSync(kDir) ? fs.readdirSync(kDir).filter((f) => f.endsWith('.json')).sort() : [];
    for (const kf of kFiles) {
      const k = JSON.parse(fs.readFileSync(path.join(kDir, kf), 'utf8'));
      const base = kf.replace(/\.json$/, '');
      const gruppen = [...new Set(k.artikel.map((a) => a.gruppe || ''))];
      for (const gruppe of gruppen) {
        const items = k.artikel.filter((a) => (a.gruppe || '') === gruppe);
        const pages = Math.ceil(items.length / 12);
        for (let page = 0; page < pages; page++) {
          const s = pres.addSlide({ masterName: 'MS_KOLLEKTION' });
          T(s, `${k.titel}${gruppe ? ' · ' + gruppe : ''}${pages > 1 ? ` (${page + 1}/${pages})` : ''}`);
          P(s, 'sub', k.unterzeile || '');
          for (let i = 0; i < 12; i++) {
            const a = items[page * 12 + i];
            if (!a) continue;
            await IMG(s, `artikel${i + 1}`, 'foto', G.kollektion.tile(i), `kollektionen/${base}/${a.foto.replace(/\.[^.]+$/, '')}`, { fit: 'contain' });
            P(s, `artikel${i + 1}_name`, a.name);
            P(s, `artikel${i + 1}_preis`, a.preis || '');
          }
          s.addNotes(`Layout MS_KOLLEKTION: bis zu zwölf Artikel mit Mockup und Name. Quelle: assets/fotos/kollektionen/${kf} (aus dem Shopify-Vereinsshop).`);
        }
      }
    }
  }
  await projekteSlide('projekte-vereine.json', 'Projekte für Vereine', 'Vom Jubiläums-Shirt bis zum Trainingslager: vier Beispiele aus der Werkstatt.');
  await referenzSlides('vereine', 'Vereine, die auf uns setzen', 'MS_REFERENZEN_WAPPEN', REF_GRID.hoch, 'Eine Auswahl. Vereine aus der ganzen Region lassen bei uns ausstatten.');

  // ---------- Abschnitt 02: Für Unternehmen (Sie) ----------
  abschnitt('02', 'Für Unternehmen', 'Workwear, Teamkleidung und Werbetextilien mit Ihrem Logo. Waschbeständig, nachbestellbar, aus einer Hand.');
  await projekteSlide('projekte-unternehmen.json', 'Projekte für Unternehmen', 'Workwear und Teamkleidung: Beispiele aus der Werkstatt.');
  {
    const s = pres.addSlide({ masterName: 'MS_PROJEKT_KUNDE' });
    T(s, 'Im Detail: Emin Isic Montagebau');
    P(s, 'sub', 'Montagebetrieb, Tübingen · Arbeitskleidung für das ganze Team');
    s.addText([
      { text: 'Aufgabe', options: { bold: true, fontSize: FS.h, breakLine: true } },
      { text: 'Einheitliche Arbeitskleidung vom Polo bis zum Sweatshirt, mit Firmenlogo auf der Brust und Handwerker-Motiv auf dem Rücken.', options: { fontSize: FS.sub, breakLine: true, paraSpaceAfter: 12 } },
      { text: 'Umsetzung', options: { bold: true, fontSize: FS.h, breakLine: true } },
      { text: 'Polos, Sweatshirts und Zip-Sweatshirts in Weiß, Logo auf Brust und Ärmel', options: { bullet: true, fontSize: FS.sub, breakLine: true } },
      { text: 'T-Shirts in Weiß und Rot mit Rückenmotiv', options: { bullet: true, fontSize: FS.sub, breakLine: true } },
      { text: 'Einfarbiger Druck, waschbeständig für den Baustellenalltag', options: { bullet: true, fontSize: FS.sub, breakLine: true } },
      { text: 'Artikel sind im System hinterlegt, Nachbestellung jederzeit möglich', options: { bullet: true, fontSize: FS.sub } },
    ], { placeholder: 'body', isTextBox: true, paraSpaceAfter: 4, color: C.dark });
    await IMG(s, 'hero', 'foto', G.kunde.hero, 'isic-kollektion');
    const teile = [['isic-polo', 'Polo, Brust und Ärmel'], ['isic-zip', 'Zip-Sweatshirt, Ärmel'], ['isic-shirt', 'T-Shirt, Rückenmotiv']];
    for (let i = 0; i < 3; i++) { await IMG(s, `teil${i + 1}`, 'foto', G.kunde.klein[i], teile[i][0]); P(s, `teil${i + 1}_name`, teile[i][1]); }
    s.addNotes('Layout MS_PROJEKT_KUNDE: ein Kunde im Detail. Text links, Kollektionsfoto oben rechts, drei Einzelteile darunter.');
  }
  await referenzSlides('unternehmen', 'Unternehmen, die auf uns setzen', 'MS_REFERENZEN', REF_GRID.breit, 'Eine Auswahl. Viele weitere Unternehmen und Teams aus der Region lassen bei uns veredeln.');

  // ---------- Abschluss ----------
  {
    const s = pres.addSlide({ masterName: 'MS_ABSCHLUSS' });
    T(s, 'Der nächste Schritt');
    P(s, 'sub', 'Motiv und Wunschtextil schicken, wir melden uns innerhalb von 48 Stunden mit Angebot und Mockup.');
    await IMG(s, 'portrait', 'portrait', G.portrait);
    P(s, 'name', 'Stefan Maier');
    P(s, 'rolle', 'Inhaber');
    P(s, 'kontakt', `${KONTAKT_TEL}\n${KONTAKT_MAIL}`);
    P(s, 'schritt', 'Anfrage per Mail oder Telefon, Angebot innerhalb von 48 Stunden. Für Vereine: Termin zur Kollektionsplanung, danach geht der Shop online.');
    await IMG(s, 'qr', 'qr', G.qr);
    P(s, 'qr_text', 'Anfrage online:\nmaiershirts.de');
    s.addNotes('Layout MS_ABSCHLUSS. Porträt und QR-Code sind Platzhalterbilder: eigenes Foto und einen echten QR-Code (z. B. auf maiershirts.de) einsetzen.');
  }

  // ---------- Schreiben ----------
  const pptxPath = path.join(DIST, 'Maiershirts_Master.pptx');
  await pres.writeFile({ fileName: pptxPath });

  // ---------- Nachbearbeitung im Paket ----------
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));

  // 1) Theme: Farbpalette und Schriften der Marke (gilt für alles, was in PowerPoint neu eingefügt wird)
  const themeFile = 'ppt/theme/theme1.xml';
  let theme = await zip.file(themeFile).async('string');
  const scheme = `<a:clrScheme name="Maiershirts">` +
    `<a:dk1><a:srgbClr val="${C.dark}"/></a:dk1><a:lt1><a:srgbClr val="${C.white}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${C.greenDark}"/></a:dk2><a:lt2><a:srgbClr val="${C.beige}"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="${C.green}"/></a:accent1><a:accent2><a:srgbClr val="${C.greenLight}"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="${C.greenDark}"/></a:accent3><a:accent4><a:srgbClr val="${C.beigeDark}"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="${C.cream}"/></a:accent5><a:accent6><a:srgbClr val="${C.dark}"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="${C.greenDark}"/></a:hlink><a:folHlink><a:srgbClr val="${C.green}"/></a:folHlink></a:clrScheme>`;
  theme = theme.replace(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/, scheme);
  theme = theme.replace(/<a:fontScheme name="[^"]*">/, '<a:fontScheme name="Maiershirts">');
  theme = theme.replace(/<a:latin typeface="[^"]*"[^/]*\/>/g, `<a:latin typeface="${FONT}"/>`);
  zip.file(themeFile, theme);

  // 2) Bildplatzhalter: pptxgenjs schreibt sie als Textplatzhalter, PowerPoint braucht type="pic"
  const PIC_PROMPT = /(Foto|Logo|QR-Code|Porträt) einfügen/;
  const untyped = /<p:ph\s+idx="(\d+)"\s+(hasCustomPrompt="1")?\s*\/>/g;
  for (const f of Object.keys(zip.files)) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f)) {
      let xml = await zip.file(f).async('string');
      xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (sp) => PIC_PROMPT.test(sp) ? sp.replace(untyped, '<p:ph idx="$1" type="pic" $2/>') : sp);
      zip.file(f, xml);
    } else if (/^ppt\/slides\/slide\d+\.xml$/.test(f)) {
      let xml = await zip.file(f).async('string');
      xml = xml.replace(/<p:pic>[\s\S]*?<\/p:nvPicPr>/g, (blk) => blk.replace(untyped, '<p:ph idx="$1" type="pic" $2/>'));
      const slideNum = Number(f.match(/slide(\d+)\.xml$/)[1]);
      for (const fix of PIC_FIX.filter((p) => p.slideNum === slideNum)) {
        xml = xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, (blk) =>
          new RegExp(`<p:ph idx="${fix.idx}"`).test(blk) ? blk.replace(/<a:off x="\d+" y="\d+"\/>/, `<a:off x="${EMU(fix.x)}" y="${EMU(fix.y)}"/>`) : blk);
      }
      zip.file(f, xml);
    }
  }
  fs.writeFileSync(pptxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('geschrieben:', pptxPath);

  // 3) .potx: gleiches Paket, Content-Type "template"
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
  ));
  const potxPath = path.join(DIST, 'Maiershirts_Master.potx');
  fs.writeFileSync(potxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('geschrieben:', potxPath);

  fs.rmSync(TMP, { recursive: true, force: true });
}

main().catch((e) => { console.error(e); process.exit(1); });
