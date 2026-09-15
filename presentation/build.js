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
  qr: [6.45, 3.5, 1.0, 1.0],
};
// Referenz-Raster: 4 x 3 breite Felder (Unternehmen) und 4 x 2 hohe Felder (Vereinswappen)
const REF_GRID = {
  breit: { cols: 4, rows: 3, cellW: 2.1, cellH: 1.05, x0: 0.62, y0: 1.47, dx: 2.3, dy: 1.2, w: 1.86, h: 0.81 },
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
      [0, 1, 2].flatMap((r) => [0, 1, 2, 3].map((c) => rr(0.5 + c * 2.3, 1.35 + r * 1.2, 2.1, 1.05, C.white, 0.08))).join('')),
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
        ph(`projekt${i + 1}_name`, 'body', 'Kunde', { x: x + 0.15, y: y + h - 0.62, w: w - 0.3, h: 0.32, fontSize: i === 0 ? FS.h : 13, bold: true, color: C.white, align: 'left', valign: 'bottom' }),
        ph(`projekt${i + 1}_text`, 'body', 'Was wir gemacht haben', { x: x + 0.15, y: y + h - 0.32, w: w - 0.3, h: 0.25, fontSize: FS.label, color: C.white, align: 'left', valign: 'top' }),
      ]),
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

  // Referenzen: Logo-Raster 4 x 3 (breite Logos) und 4 x 2 (Vereinswappen), auf Sand
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
  const IMG = async (s, name, kind, [x, y, w, h], foto, { gradient = false } = {}) => {
    const real = foto && ['jpg', 'jpeg', 'png'].map((e) => path.join(FOTOS, `${foto}.${e}`)).find((f) => fs.existsSync(f));
    let file;
    if (real) {
      file = path.join(TMP, `foto-${foto}-${Math.round(w * 100)}x${Math.round(h * 100)}.jpg`);
      const pw = Math.round(w * 300), phh = Math.round(h * 300);
      let img = sharp(real).resize(pw, phh, { fit: 'cover', position: 'centre' });
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

  // 1 Titel
  {
    const s = pres.addSlide({ masterName: 'MS_TITEL' });
    T(s, 'Angebotspräsentation');
    P(s, 'sub', 'Individuell veredelte Textilien für Teams, Vereine und Unternehmen');
    P(s, 'kunde', `Angebot für: ${KUNDE}`);
    P(s, 'meta', 'September 2026 · Angebot Nr. 2026-0815');
    s.addNotes('Layout MS_TITEL. Alle vier Felder sind Platzhalter: Titel, Untertitel, Kunde/Projekt, Datum/Angebotsnummer.');
  }

  // 2 Agenda
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    T(s, 'Agenda');
    const items = [
      ['01', 'Über uns', 'Wer wir sind und wofür wir stehen'],
      ['02', 'Leistungen', 'Verfahren, Textilien und Veredelung'],
      ['03', 'Ihr Angebot', 'Textilien, Staffelpreise und Positionen'],
      ['04', 'Nächste Schritte', 'Freigabe, Produktion und Lieferung'],
    ];
    const runs = [];
    items.forEach(([num, head, desc], i) => {
      runs.push({ text: `${num}   `, options: { fontSize: FS.h, bold: true, color: C.greenDark } });
      runs.push({ text: head, options: { fontSize: FS.h, bold: true, color: C.dark, breakLine: true } });
      runs.push({ text: `        ${desc}`, options: { fontSize: FS.sub, ...MUTED_ON_LIGHT, breakLine: i < items.length - 1, paraSpaceAfter: 14 } });
    });
    s.addText(runs, { placeholder: 'body', isTextBox: true, paraSpaceAfter: 2 });
    s.addNotes('Layout MS_INHALT: Titel und Textkörper. Die Agenda ist normaler Text im Textkörper.');
  }

  // 3 Abschnitt
  {
    const s = pres.addSlide({ masterName: 'MS_ABSCHNITT' });
    P(s, 'num', '01');
    T(s, 'Über uns');
    P(s, 'sub', 'Textilveredelung aus einer Hand: persönlich, schnell und in gleichbleibender Qualität');
    s.addNotes('Layout MS_ABSCHNITT: Nummer, Titel, Kurzbeschreibung.');
  }

  // 4 Was uns ausmacht (freie Folie auf Sand, Karten weiß)
  {
    const s = pres.addSlide({ masterName: 'MS_FREI_SAND' });
    T(s, 'Was uns ausmacht');
    s.addText('Maiershirts veredelt Textilien für Unternehmen, Vereine und Events. Vom einzelnen Shirt bis zur kompletten Teamausstattung begleiten wir jedes Projekt persönlich, von der Motividee bis zum fertigen Paket.',
      { x: 0.5, y: 1.35, w: 4.1, h: 1.6, fontFace: FONT, fontSize: FS.body, color: C.dark, valign: 'top', margin: 0, isTextBox: true });
    s.addText([
      { text: 'Beratung zu Textil, Verfahren und Motiv', options: { bullet: true, breakLine: true } },
      { text: 'Druck und Stickerei im eigenen Haus', options: { bullet: true, breakLine: true } },
      { text: 'Kleine Auflagen ab 1 Stück', options: { bullet: true, breakLine: true } },
      { text: 'Verlässliche Liefertermine', options: { bullet: true } },
    ], { x: 0.5, y: 3.0, w: 4.1, h: 1.8, fontFace: FONT, fontSize: FS.body, color: C.dark, valign: 'top', margin: 0, paraSpaceAfter: 6, isTextBox: true });
    const cards = [
      ['Siebdruck', 'Kräftige Farben, ideal für größere Auflagen'],
      ['Stickerei', 'Hochwertig und langlebig, erste Wahl für Workwear'],
      ['Express', 'Kurze Produktionszeiten auf Anfrage'],
      ['Qualität', 'Geprüfte Textilien namhafter Hersteller'],
    ];
    cards.forEach(([head, desc], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 5.0 + col * 2.3, y = 1.35 + row * 1.8;
      s.addShape('roundRect', { x, y, w: 2.15, h: 1.65, fill: { color: C.white }, line: { color: C.white }, rectRadius: 0.12 });
      s.addShape('line', { x: x + 0.25, y: y + 0.3, w: 0.4, h: 0, line: { color: C.green, width: 3 } });
      s.addText(head, { x: x + 0.25, y: y + 0.45, w: 1.7, h: 0.35, fontFace: FONT, fontSize: FS.h, bold: true, color: C.dark, margin: 0, isTextBox: true });
      s.addText(desc, { x: x + 0.25, y: y + 0.82, w: 1.7, h: 0.7, fontFace: FONT, fontSize: FS.sub, ...MUTED_ON_LIGHT, valign: 'top', margin: 0, isTextBox: true });
    });
    s.addNotes('Layout MS_FREI_SAND: nur Titel, Fläche frei. Hier: Text links, vier weiße Karten rechts.');
  }

  // 5 Vollbild-Foto
  {
    const s = pres.addSlide({ masterName: 'MS_BILD_VOLL' });
    await IMG(s, 'foto', 'foto', G.voll, 'vollbild');
    T(s, 'Ihr Motiv auf dem Textil');
    P(s, 'sub', 'Maiershirts-Logo auf Polo-Shirt, Beispiel aus der Produktion');
    s.addNotes('Layout MS_BILD_VOLL: Foto über die volle Breite, darunter Bildunterschrift. Bild per Klick auf das Platzhalterbild ersetzen.');
  }

  // 6 Kennzahlen
  {
    const s = pres.addSlide({ masterName: 'MS_KENNZAHLEN' });
    T(s, 'Zahlen, die für uns sprechen');
    [['1.200+', 'Projekte pro Jahr', 'Beispielwert'], ['48 h', 'bis zum Angebot', 'Beispielwert'], ['98 %', 'Wiederkehrende Kunden', 'Beispielwert']].forEach(([big, label, note], i) => {
      P(s, `kpi${i + 1}`, big); P(s, `kpi${i + 1}_label`, label); P(s, `kpi${i + 1}_note`, note);
    });
    P(s, 'hinweis', 'Alle Werte sind Platzhalter und werden vor Verwendung durch echte Kennzahlen ersetzt.');
    s.addNotes('Layout MS_KENNZAHLEN: drei Karten mit Zahl, Bedeutung und Quelle, dazu eine Hinweiszeile.');
  }

  // 7 Prozess
  {
    const s = pres.addSlide({ masterName: 'MS_PROZESS' });
    T(s, 'So läuft ein Auftrag ab');
    [['Anfrage', 'Textil, Menge und Motiv per Mail oder Telefon'], ['Angebot', 'Angebot aus Lexware innerhalb von 48 Stunden'], ['Produktion', 'Druck oder Stickerei nach Ihrer Freigabe'], ['Lieferung', 'Versand oder Abholung zum Wunschtermin']].forEach(([head, desc], i) => {
      P(s, `schritt${i + 1}`, head); P(s, `schritt${i + 1}_text`, desc);
    });
    s.addNotes('Layout MS_PROZESS: vier nummerierte Schritte mit Überschrift und Kurztext.');
  }

  // 7b Verfahrensvergleich (Inhalte aus dem Veredelungs-Leitfaden)
  {
    const s = pres.addSlide({ masterName: 'MS_VERFAHREN' });
    T(s, 'Vier Verfahren, ein Anspruch');
    P(s, 'sub', 'Ehrlich mit Stärken und Grenzen, damit Ihr Motiv da landet, wo es am besten sitzt.');
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
    s.addNotes('Layout MS_VERFAHREN: vier Verfahren mit Einsatzgebiet und drei Kennwerten, darunter die Entscheidungshilfe. Inhalte aus dem Veredelungs-Leitfaden (maiershirts-veredelung-leitfaden/leitfaden.html). Handmuster-Karte „Fühl den Unterschied“ zum Termin mitnehmen.');
  }

  // 8b Projekte aus der Praxis: Einträge aus assets/fotos/projekte.json, vier pro Folie
  {
    const listFile = path.join(FOTOS, 'projekte.json');
    const projekte = fs.existsSync(listFile) ? JSON.parse(fs.readFileSync(listFile, 'utf8')) : [];
    const pages = Math.max(1, Math.ceil(projekte.length / 4));
    for (let page = 0; page < pages; page++) {
      const s = pres.addSlide({ masterName: 'MS_PROJEKTE' });
      T(s, pages > 1 ? `Projekte aus der Praxis (${page + 1}/${pages})` : 'Projekte aus der Praxis');
      P(s, 'sub', 'Beispiele aus der Werkstatt, vom Vereinsjubiläum bis zur Workwear.');
      for (let i = 0; i < 4; i++) {
        const pr = projekte[page * 4 + i];
        await IMG(s, `projekt${i + 1}`, 'foto', G.projekt[i], pr && pr.foto, { gradient: !!pr });
        if (pr) { P(s, `projekt${i + 1}_name`, pr.kunde); P(s, `projekt${i + 1}_text`, pr.text); }
      }
      s.addNotes(`Layout MS_PROJEKTE: ein großes und drei kleine Fotos mit Kunde und Kurztext im Bild (Folie ${page + 1} von ${pages}). Einträge in assets/fotos/projekte.json, Fotos daneben; der Build legt einen dunklen Verlauf unter die Bildunterschrift.`);
    }
  }

  // 8 Textil-Auswahl
  {
    const s = pres.addSlide({ masterName: 'MS_TEXTIL' });
    T(s, 'Textilien für Ihr Projekt');
    const textil = [
      ['Polo-Shirt Piqué', 'Kariban Piqué · 220 g/m²\n8 Farben · S bis 4XL', 'ab 17,50 €'],
      ['Cap', 'Beechfield Original · 6 Panel\n12 Farben · Einheitsgröße', 'ab 12,90 €'],
      ['Hoodie Bio', 'Stanley/Stella Cruiser · 350 g/m²\n10 Farben · XS bis 3XL', 'ab 29,90 €'],
    ];
    for (let i = 0; i < 3; i++) {
      await IMG(s, `textil${i + 1}_foto`, 'foto', G.textil(i), `textil-${i + 1}`);
      P(s, `textil${i + 1}`, textil[i][0]); P(s, `textil${i + 1}_details`, textil[i][1]); P(s, `textil${i + 1}_preis`, textil[i][2]);
    }
    s.addNotes('Layout MS_TEXTIL: drei Rohlinge mit Foto, Name, Details und Preis ab. Beispielwerte.');
  }

  // 9 Preisstaffel
  {
    const s = pres.addSlide({ masterName: 'MS_PREISSTAFFEL' });
    T(s, 'Preisstaffel T-Shirt, Brustdruck 1-farbig');
    [['ab 10 Stück', '14,90 €'], ['ab 25 Stück', '12,90 €'], ['ab 50 Stück', '11,50 €'], ['ab 100 Stück', '9,90 €']].forEach(([stufe, preis], i) => {
      P(s, `stufe${i + 1}`, stufe); P(s, `stufe${i + 1}_preis`, preis); P(s, `stufe${i + 1}_note`, 'pro Stück, netto\ninkl. Textil und Druck');
    });
    P(s, 'hinweis', 'Beispielwerte. Staffelpreise gelten je Motiv und Textil; Einrichtungskosten für den Siebdruck fallen einmalig an. Verbindlich ist das Angebot aus Lexware.');
    s.addNotes('Layout MS_PREISSTAFFEL: vier Stufen Menge gegen Stückpreis, darunter Hinweise.');
  }

  // 10 Angebot
  {
    const s = pres.addSlide({ masterName: 'MS_ANGEBOT' });
    T(s, 'Ihr Angebot im Überblick');
    const hdr = (t, align = 'left') => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, align, fontSize: FS.sub } });
    const cell = (t, align = 'left', opts = {}) => ({ text: t, options: { align, fontSize: FS.sub, color: C.dark, ...opts } });
    const sum = (t, align = 'left', opts = {}) => cell(t, align, { bold: true, fill: { color: C.beige }, ...opts });
    const rows = [
      [hdr('Pos.'), hdr('Artikel'), hdr('Menge', 'right'), hdr('Einzelpreis', 'right'), hdr('Gesamt', 'right')],
      [cell('1'), cell('T-Shirt Bio-Baumwolle, Brustdruck 1-farbig'), cell('50', 'right'), cell('12,90 €', 'right'), cell('645,00 €', 'right')],
      [cell('2'), cell('Hoodie, Rückendruck 2-farbig'), cell('25', 'right'), cell('34,50 €', 'right'), cell('862,50 €', 'right')],
      [cell('3'), cell('Polo-Shirt, Logo-Stickerei'), cell('20', 'right'), cell('26,00 €', 'right'), cell('520,00 €', 'right')],
      [cell('4'), cell('Einrichtungskosten Siebdruck'), cell('1', 'right'), cell('45,00 €', 'right'), cell('45,00 €', 'right')],
      [sum(''), sum('Netto gesamt'), sum(''), sum(''), sum('2.072,50 €', 'right', { color: C.greenDark })],
    ];
    s.addTable(rows, {
      x: 0.5, y: 1.35, w: 9.0, colW: [0.6, 4.6, 1.0, 1.4, 1.4],
      fontFace: FONT, rowH: 0.4, border: { type: 'solid', color: C.beigeDark, pt: 0.75 },
      fill: { color: C.white }, valign: 'middle', margin: [0.04, 0.1, 0.04, 0.1],
    });
    P(s, 'kond1', '15. Oktober 2026'); P(s, 'kond2', '10 Arbeitstage nach Freigabe'); P(s, 'kond3', '14 Tage netto');
    P(s, 'hinweis', 'Beispielpositionen · Preise netto zzgl. MwSt. · Verbindlich ist das Angebot Nr. 2026-0815 aus Lexware.');
    s.addNotes('Layout MS_ANGEBOT: Tabelle als Zusammenfassung (nativ, in PowerPoint editierbar), Konditionen als Platzhalter. Das verbindliche Angebot kommt aus Lexware.');
  }

  // 11 Referenzen: eine Folie pro Unterordner in assets/referenzen/ (unternehmen, vereine)
  // Logos alphabetisch (z. B. 01-firma.png), freie Felder bleiben Platzhalter
  const refRoot = path.join(ASSETS, 'referenzen');
  const refGroups = [
    ['unternehmen', 'Unternehmen, die uns vertrauen', 'MS_REFERENZEN', REF_GRID.breit, 'Eine Auswahl. Viele weitere Unternehmen und Teams aus der Region lassen bei uns veredeln.'],
    ['vereine', 'Vereine, die uns vertrauen', 'MS_REFERENZEN_WAPPEN', REF_GRID.hoch, 'Eine Auswahl. Vom Trikotsatz bis zur Fanausstattung, für Vereine aus der ganzen Region.'],
  ];
  for (const [group, titel, layout, grid, unterzeile] of refGroups) {
    const refDir = path.join(refRoot, group);
    const refs = fs.existsSync(refDir) ? fs.readdirSync(refDir).filter((f) => /\.(png|jpe?g|svg)$/i.test(f)).sort() : [];
    const namenFile = path.join(refDir, 'namen.json');
    const namen = fs.existsSync(namenFile) ? JSON.parse(fs.readFileSync(namenFile, 'utf8')) : {};
    const nameOf = (f) => namen[f] || f.replace(/^\d+-/, '').replace(/\.[^.]+$/, '').replace(/-/g, ' ');
    const cells = grid.cols * grid.rows;
    const pages = Math.max(1, Math.ceil(refs.length / cells)); // mehr Logos als Felder: automatisch Fortsetzungsfolie
    for (let page = 0; page < pages; page++) {
    const pageRefs = refs.slice(page * cells, (page + 1) * cells);
    const s = pres.addSlide({ masterName: layout });
    T(s, pages > 1 ? `${titel} (${page + 1}/${pages})` : titel);
    P(s, 'sub', unterzeile);
    for (let i = 0; i < cells; i++) {
      const [x, y, w, h] = G.logoGrid(grid, Math.floor(i / grid.cols), i % grid.cols);
      if (pageRefs[i]) {
        // Luft zum Kartenrand; Logo proportional eingepasst und zentriert (pptxgenjs kennt die Bildmaße nicht)
        const pad = 0.16, boxW = w - 2 * pad, boxH = h - 2 * pad;
        // Weiße bzw. transparente Ränder der Datei abschneiden, damit alle Logos ähnlich groß wirken
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
      } else {
        await IMG(s, `logo${i + 1}`, 'logo', [x, y, w, h]);
      }
    }
    s.addNotes(`Layout ${layout}: ${cells} Logo-Felder. ${pageRefs.length} Logo(s) aus assets/referenzen/${group} eingesetzt (Folie ${page + 1} von ${pages}). Logo per Klick auf das Platzhalterbild einsetzen; bei Beschnitt: Bildformat → Zuschneiden → Anpassen.`);
    }
  }

  // 13 Bild links, Text rechts
  {
    const s = pres.addSlide({ masterName: 'MS_BILD' });
    T(s, 'Stickerei im Detail');
    await IMG(s, 'foto', 'foto', G.bild, 'bild');
    s.addText([
      { text: 'Bis zu 12 Garnfarben pro Motiv', options: { bullet: true, breakLine: true } },
      { text: 'Waschbeständig bis 60 °C', options: { bullet: true, breakLine: true } },
      { text: 'Ideal für Polos, Jacken und Caps', options: { bullet: true, breakLine: true } },
      { text: 'Motivdigitalisierung im Haus', options: { bullet: true } },
    ], { placeholder: 'body', isTextBox: true, paraSpaceAfter: 8 });
    s.addNotes('Layout MS_BILD: Foto links, Textkörper rechts.');
  }

  // 14 Diagramm
  {
    const s = pres.addSlide({ masterName: 'MS_FREI' });
    T(s, 'Auftragsvolumen nach Produktgruppe');
    s.addChart(pres.charts.BAR, [
      { name: 'Aufträge', labels: ['T-Shirts', 'Hoodies', 'Polos', 'Workwear', 'Caps'], values: [420, 260, 180, 150, 90] },
    ], {
      x: 0.5, y: 1.3, w: 5.8, h: 3.6,
      barDir: 'col', chartColors: [C.green],
      showTitle: false, showLegend: false,
      showValue: true, dataLabelPosition: 'outEnd', dataLabelFontFace: FONT, dataLabelFontSize: FS.sub, dataLabelColor: C.dark,
      catAxisLabelFontFace: FONT, catAxisLabelFontSize: FS.sub, catAxisLabelColor: C.dark,
      valAxisLabelFontFace: FONT, valAxisLabelFontSize: FS.label, valAxisLabelColor: C.dark,
      valGridLine: { color: C.beigeDark, size: 0.5 }, catGridLine: { style: 'none' },
      valAxisLineShow: false, catAxisLineShow: false,
    });
    s.addShape('roundRect', { x: 6.6, y: 1.3, w: 2.9, h: 3.6, fill: { color: C.beige }, line: { color: C.beige }, rectRadius: 0.12 });
    s.addText('Beispieldaten', { x: 6.85, y: 1.5, w: 2.4, h: 0.3, fontFace: FONT, fontSize: FS.sub, ...MUTED_ON_LIGHT, margin: 0, isTextBox: true });
    s.addText('T-Shirts machen den größten Anteil aus', { x: 6.85, y: 1.85, w: 2.4, h: 0.7, fontFace: FONT, fontSize: FS.h, bold: true, color: C.dark, margin: 0, isTextBox: true });
    s.addText('Natives Diagramm: Werte per Rechtsklick → „Daten bearbeiten“ ändern.', { x: 6.85, y: 2.65, w: 2.4, h: 1.2, fontFace: FONT, fontSize: FS.sub, ...MUTED_ON_LIGHT, valign: 'top', margin: 0, isTextBox: true });
    s.addNotes('Layout MS_FREI mit nativem Säulendiagramm links und Kernaussage rechts.');
  }

  // 15 Standardfolie
  {
    const s = pres.addSlide({ masterName: 'MS_INHALT' });
    T(s, 'Folientitel');
    s.addText([
      { text: 'Erste Aussage der Folie', options: { bullet: true, breakLine: true } },
      { text: 'Zweite Aussage mit einer kurzen Erläuterung', options: { bullet: true, breakLine: true } },
      { text: 'Dritte Aussage', options: { bullet: true, breakLine: true } },
      { text: 'Unterpunkt zur dritten Aussage', options: { bullet: true, indentLevel: 1, breakLine: true } },
      { text: 'Weiterer Unterpunkt', options: { bullet: true, indentLevel: 1 } },
    ], { placeholder: 'body', isTextBox: true, paraSpaceAfter: 8 });
    s.addNotes('Layout MS_INHALT: Titel und Textkörper mit Aufzählung.');
  }

  // 16 Abschluss
  {
    const s = pres.addSlide({ masterName: 'MS_ABSCHLUSS' });
    T(s, 'Lassen Sie uns Ihr Projekt starten');
    P(s, 'sub', 'Wir freuen uns auf Ihre Freigabe und begleiten Sie bis zur Lieferung.');
    await IMG(s, 'portrait', 'portrait', G.portrait);
    P(s, 'name', 'Stefan Maier');
    P(s, 'rolle', 'Inhaber · Ihr Ansprechpartner');
    P(s, 'kontakt', '+49 (0) 000 000000\ninfo@maiershirts.de');
    P(s, 'schritt', 'Freigabe des Angebots bis 15. Oktober 2026. Danach Musterfreigabe, Produktion und Lieferung zum Saisonstart.');
    await IMG(s, 'qr', 'qr', G.qr);
    P(s, 'qr_text', 'Anfrage online:\nmaiershirts.de/anfrage');
    s.addNotes('Layout MS_ABSCHLUSS: Handlungsaufforderung, Ansprechpartner mit Foto, nächster Schritt mit Termin, QR-Code. Kontaktdaten sind Platzhalter.');
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
