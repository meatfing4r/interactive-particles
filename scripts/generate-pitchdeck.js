#!/usr/bin/env node
/**
 * Pitch Deck Generator
 * Generates a polished PowerPoint pitch deck for the Interactive Particles project.
 *
 * Usage:
 *   node scripts/generate-pitchdeck.js [--output <filename>] [--title <title>]
 *
 * Options:
 *   --output   Output filename (default: pitchdeck.pptx)
 *   --title    Presentation title (default: "Interactive Particles")
 */

'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const args = process.argv.slice(2);
const getArg = (flag) => {
  const idx = args.indexOf(flag);
  return idx !== -1 ? args[idx + 1] : null;
};

const OUTPUT_FILE = getArg('--output') || 'pitchdeck.pptx';
const DECK_TITLE = getArg('--title') || 'Interactive Particles';

// ---------------------------------------------------------------------------
// Design tokens – easy to customise
// ---------------------------------------------------------------------------
const THEME = {
  // Dark, modern palette
  bg:          '0D0D0D',   // near-black background
  surface:     '1A1A2E',   // deep navy surface
  accent:      '00D4FF',   // electric cyan
  accentAlt:   'FF6B6B',   // coral accent
  textPrimary: 'FFFFFF',
  textMuted:   'A0AEC0',

  fontTitle:   'Calibri',
  fontBody:    'Calibri',
};

const SLIDE_W = 10;   // inches (standard widescreen)
const SLIDE_H = 5.63; // inches

// ---------------------------------------------------------------------------
// Helper: full-bleed dark background rect
// ---------------------------------------------------------------------------
function addBg(slide, color) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    fill: { color: color || THEME.bg },
    line: { color: color || THEME.bg },
  });
}

// Helper: top accent bar
function addAccentBar(slide, color) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: 0.06,
    fill: { color: color || THEME.accent },
    line: { color: color || THEME.accent },
  });
}

// Helper: bottom accent line
function addFooterLine(slide) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: SLIDE_H - 0.08, w: SLIDE_W - 1, h: 0.015,
    fill: { color: THEME.surface },
    line: { color: THEME.surface },
  });
}

// Helper: section label (small caps above heading)
function addSectionLabel(slide, text, y) {
  slide.addText(text.toUpperCase(), {
    x: 0.6, y: y, w: SLIDE_W - 1.2, h: 0.3,
    fontSize: 9,
    bold: true,
    color: THEME.accent,
    fontFace: THEME.fontBody,
    charSpacing: 3,
  });
}

// Helper: icon-style bullet (filled circle)
function bulletIcon(slide, x, y, color) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x, y: y + 0.04, w: 0.12, h: 0.12,
    fill: { color: color || THEME.accent },
    line: { color: color || THEME.accent },
  });
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

// 1. COVER
function addCoverSlide() {
  const slide = pptx.addSlide();
  addBg(slide, THEME.bg);

  // Gradient-like layered shapes for visual depth
  slide.addShape(pptx.ShapeType.rect, {
    x: 6.5, y: -0.5, w: 5, h: 7,
    fill: { color: THEME.surface },
    line: { color: THEME.surface },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 7.5, y: 0, w: 0.04, h: SLIDE_H,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
  });

  // Top left accent bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.06, h: 2.5,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
  });

  // Tag
  slide.addText('WEBGL  ·  THREE.JS  ·  INTERACTIVE', {
    x: 0.5, y: 1.1, w: 6.5, h: 0.35,
    fontSize: 9,
    bold: true,
    color: THEME.accent,
    fontFace: THEME.fontBody,
    charSpacing: 3,
  });

  // Title
  slide.addText(DECK_TITLE, {
    x: 0.5, y: 1.5, w: 6.5, h: 1.4,
    fontSize: 44,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  // Subtitle
  slide.addText(
    'High-performance particle systems\nthat react to mouse & touch in real-time',
    {
      x: 0.5, y: 3.0, w: 6.5, h: 1.0,
      fontSize: 16,
      color: THEME.textMuted,
      fontFace: THEME.fontBody,
    }
  );

  // CTA button-style shape
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y: 4.2, w: 2.2, h: 0.5,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
    rectRadius: 0.05,
  });
  slide.addText('View Demo →', {
    x: 0.5, y: 4.2, w: 2.2, h: 0.5,
    fontSize: 13,
    bold: true,
    color: THEME.bg,
    fontFace: THEME.fontBody,
    align: 'center',
  });
}

// 2. THE PROBLEM
function addProblemSlide() {
  const slide = pptx.addSlide();
  addBg(slide);
  addAccentBar(slide);

  addSectionLabel(slide, 'The Challenge', 0.4);

  slide.addText('Engaging users on the web is hard', {
    x: 0.6, y: 0.75, w: SLIDE_W - 1.2, h: 0.8,
    fontSize: 32,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  const problems = [
    ['Static visuals bore visitors',           'Most landing pages use flat images — no depth, no life.'],
    ['Canvas / WebGL is intimidating',         'Setting up performant particle systems from scratch takes weeks.'],
    ['Touch & mouse interactivity is costly',  'Handling pointer events at 60 fps without jank requires expertise.'],
  ];

  problems.forEach(([heading, body], i) => {
    const y = 1.7 + i * 1.1;
    // Card background
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.5, y, w: SLIDE_W - 1, h: 0.95,
      fill: { color: THEME.surface },
      line: { color: '2A2A4A' },
      rectRadius: 0.06,
    });
    // Accent strip
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y, w: 0.04, h: 0.95,
      fill: { color: THEME.accentAlt },
      line: { color: THEME.accentAlt },
    });
    slide.addText(heading, {
      x: 0.8, y: y + 0.1, w: 8.5, h: 0.35,
      fontSize: 14,
      bold: true,
      color: THEME.textPrimary,
      fontFace: THEME.fontBody,
    });
    slide.addText(body, {
      x: 0.8, y: y + 0.48, w: 8.5, h: 0.35,
      fontSize: 11,
      color: THEME.textMuted,
      fontFace: THEME.fontBody,
    });
  });

  addFooterLine(slide);
}

// 3. THE SOLUTION
function addSolutionSlide() {
  const slide = pptx.addSlide();
  addBg(slide);
  addAccentBar(slide, THEME.accentAlt);

  addSectionLabel(slide, 'Our Solution', 0.4);

  slide.addText('Interactive Particles — drop-in WebGL magic', {
    x: 0.6, y: 0.75, w: SLIDE_W - 1.2, h: 0.8,
    fontSize: 28,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  slide.addText(
    'A lightweight, dependency-light Three.js boilerplate that turns any image\n' +
    'into thousands of interactive GPU-accelerated particles — in minutes.',
    {
      x: 0.6, y: 1.55, w: SLIDE_W - 1.2, h: 0.8,
      fontSize: 14,
      color: THEME.textMuted,
      fontFace: THEME.fontBody,
    }
  );

  // Two-column feature summary
  const cols = [
    { x: 0.5,  features: ['GPU-accelerated shaders', 'Off-screen touch texture', 'GSAP-powered transitions'] },
    { x: 5.1,  features: ['Configurable GUI controls', 'Webpack + Babel pipeline', 'MIT licensed'] },
  ];

  cols.forEach(({ x, features }) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 2.5, w: 4.4, h: 2.5,
      fill: { color: THEME.surface },
      line: { color: '2A2A4A' },
      rectRadius: 0.08,
    });

    features.forEach((feat, i) => {
      bulletIcon(slide, x + 0.25, 2.7 + i * 0.65);
      slide.addText(feat, {
        x: x + 0.5, y: 2.65 + i * 0.65, w: 3.7, h: 0.45,
        fontSize: 13,
        color: THEME.textPrimary,
        fontFace: THEME.fontBody,
      });
    });
  });

  addFooterLine(slide);
}

// 4. HOW IT WORKS
function addHowItWorksSlide() {
  const slide = pptx.addSlide();
  addBg(slide);
  addAccentBar(slide);

  addSectionLabel(slide, 'How It Works', 0.4);

  slide.addText('Three steps from image to magic', {
    x: 0.6, y: 0.75, w: SLIDE_W - 1.2, h: 0.7,
    fontSize: 28,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  const steps = [
    {
      num: '01',
      title: 'Image → Particle positions',
      body:  'Pixel colours from a source image are sampled on the CPU and\npassed as vertex attributes to the GPU buffer geometry.',
    },
    {
      num: '02',
      title: 'Touch Texture',
      body:  'An off-screen canvas captures pointer/touch events and paints\na ripple texture that the vertex shader reads each frame.',
    },
    {
      num: '03',
      title: 'GLSL Shader magic',
      body:  'The vertex shader displaces each particle based on the touch\ntexture + simplex noise, creating fluid, organic motion at 60 fps.',
    },
  ];

  steps.forEach(({ num, title, body }, i) => {
    const x = 0.4 + i * 3.2;
    // Step card
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 1.6, w: 3.0, h: 3.5,
      fill: { color: THEME.surface },
      line: { color: '2A2A4A' },
      rectRadius: 0.08,
    });
    // Number badge
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.15, y: 1.75, w: 0.6, h: 0.6,
      fill: { color: THEME.accent },
      line: { color: THEME.accent },
    });
    slide.addText(num, {
      x: x + 0.15, y: 1.75, w: 0.6, h: 0.6,
      fontSize: 11,
      bold: true,
      color: THEME.bg,
      fontFace: THEME.fontBody,
      align: 'center',
      valign: 'middle',
    });
    // Connector arrow (between steps)
    if (i < 2) {
      slide.addText('→', {
        x: x + 3.05, y: 1.95, w: 0.2, h: 0.4,
        fontSize: 18,
        color: THEME.accent,
        fontFace: THEME.fontBody,
        align: 'center',
      });
    }
    slide.addText(title, {
      x: x + 0.2, y: 2.5, w: 2.6, h: 0.6,
      fontSize: 13,
      bold: true,
      color: THEME.textPrimary,
      fontFace: THEME.fontBody,
    });
    slide.addText(body, {
      x: x + 0.2, y: 3.15, w: 2.6, h: 1.8,
      fontSize: 10.5,
      color: THEME.textMuted,
      fontFace: THEME.fontBody,
    });
  });

  addFooterLine(slide);
}

// 5. TECH STACK
function addTechStackSlide() {
  const slide = pptx.addSlide();
  addBg(slide);
  addAccentBar(slide, THEME.accentAlt);

  addSectionLabel(slide, 'Technology', 0.4);

  slide.addText('Built on proven, battle-tested tools', {
    x: 0.6, y: 0.75, w: SLIDE_W - 1.2, h: 0.7,
    fontSize: 28,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  const techs = [
    { name: 'Three.js',   desc: 'WebGL abstraction — renders millions\nof particles with minimal overhead.' },
    { name: 'GLSL',       desc: 'Custom vertex & fragment shaders\nrunning entirely on the GPU.' },
    { name: 'glslify',    desc: 'Module system for GLSL — clean,\nreusable shader code.' },
    { name: 'GSAP',       desc: 'Industry-standard animation platform\nfor silky-smooth transitions.' },
    { name: 'Webpack 4',  desc: 'Zero-config HMR dev server + optimised\nproduction bundles.' },
    { name: 'ControlKit', desc: 'Lightweight GUI panel for live\nparameter tweaking.' },
  ];

  const COLS = 3;
  techs.forEach(({ name, desc }, i) => {
    const col = i % COLS;
    const row = Math.floor(i / COLS);
    const x = 0.4 + col * 3.2;
    const y = 1.7 + row * 1.75;

    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: 3.0, h: 1.55,
      fill: { color: THEME.surface },
      line: { color: '2A2A4A' },
      rectRadius: 0.07,
    });
    // Accent dot
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.2, y: y + 0.22, w: 0.14, h: 0.14,
      fill: { color: i % 2 === 0 ? THEME.accent : THEME.accentAlt },
      line: { color: i % 2 === 0 ? THEME.accent : THEME.accentAlt },
    });
    slide.addText(name, {
      x: x + 0.45, y: y + 0.12, w: 2.4, h: 0.45,
      fontSize: 14,
      bold: true,
      color: THEME.textPrimary,
      fontFace: THEME.fontBody,
    });
    slide.addText(desc, {
      x: x + 0.2, y: y + 0.6, w: 2.65, h: 0.85,
      fontSize: 10,
      color: THEME.textMuted,
      fontFace: THEME.fontBody,
    });
  });

  addFooterLine(slide);
}

// 6. DEMO / GET STARTED
function addDemoSlide() {
  const slide = pptx.addSlide();
  addBg(slide, THEME.bg);

  // Right panel
  slide.addShape(pptx.ShapeType.rect, {
    x: 5.5, y: 0, w: 4.5, h: SLIDE_H,
    fill: { color: THEME.surface },
    line: { color: THEME.surface },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 5.5, y: 0, w: 0.04, h: SLIDE_H,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
  });

  // Left content
  addSectionLabel(slide, 'Get Started', 0.55);

  slide.addText('Try it in 3 commands', {
    x: 0.5, y: 0.9, w: 4.8, h: 0.8,
    fontSize: 28,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  const cmds = ['npm install', 'npm start', 'npm run build'];
  cmds.forEach((cmd, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.5, y: 1.9 + i * 0.75, w: 4.5, h: 0.6,
      fill: { color: '111122' },
      line: { color: THEME.accent },
      rectRadius: 0.05,
    });
    slide.addText('$ ' + cmd, {
      x: 0.7, y: 1.9 + i * 0.75, w: 4.2, h: 0.6,
      fontSize: 14,
      color: THEME.accent,
      fontFace: 'Courier New',
      valign: 'middle',
    });
  });

  // Right panel: links
  slide.addText('Resources', {
    x: 5.8, y: 0.5, w: 4.0, h: 0.5,
    fontSize: 18,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
  });

  const links = [
    ['Live Demo',        'tympanus.net/Tutorials/InteractiveParticles/'],
    ['Codrops Article',  'tympanus.net/codrops/2019/01/17/interactive-particles-with-three-js/'],
    ['GitHub Repo',      'github.com/brunoimbrizi/interactive-particles'],
    ['Three.js Docs',    'threejs.org/docs/'],
  ];

  links.forEach(([label, url], i) => {
    bulletIcon(slide, 5.85, 1.2 + i * 0.9, THEME.accentAlt);
    slide.addText(label, {
      x: 6.1, y: 1.15 + i * 0.9, w: 3.6, h: 0.35,
      fontSize: 13,
      bold: true,
      color: THEME.textPrimary,
      fontFace: THEME.fontBody,
    });
    slide.addText(url, {
      x: 6.1, y: 1.5 + i * 0.9, w: 3.6, h: 0.3,
      fontSize: 9.5,
      color: THEME.accent,
      fontFace: THEME.fontBody,
    });
  });

  // Footer CTA
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.8, y: 4.9, w: 3.8, h: 0.5,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
    rectRadius: 0.05,
  });
  slide.addText('Star on GitHub ★', {
    x: 5.8, y: 4.9, w: 3.8, h: 0.5,
    fontSize: 14,
    bold: true,
    color: THEME.bg,
    fontFace: THEME.fontBody,
    align: 'center',
    valign: 'middle',
  });
}

// 7. CLOSING / THANK YOU
function addClosingSlide() {
  const slide = pptx.addSlide();
  addBg(slide, THEME.bg);

  // Full-width accent stripe at 40% height
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 2.1, w: SLIDE_W, h: 0.06,
    fill: { color: THEME.accent },
    line: { color: THEME.accent },
  });

  slide.addText('Thank You', {
    x: 0, y: 1.0, w: SLIDE_W, h: 1.0,
    fontSize: 54,
    bold: true,
    color: THEME.textPrimary,
    fontFace: THEME.fontTitle,
    align: 'center',
  });

  slide.addText(DECK_TITLE + '  ·  Interactive WebGL Particles for the Modern Web', {
    x: 0, y: 2.3, w: SLIDE_W, h: 0.5,
    fontSize: 13,
    color: THEME.textMuted,
    fontFace: THEME.fontBody,
    align: 'center',
  });

  // Social / link chips
  const chips = [
    'tympanus.net/Tutorials/InteractiveParticles/',
    'github.com/brunoimbrizi/interactive-particles',
  ];
  chips.forEach((text, i) => {
    const chipW = 4.0;
    const chipX = SLIDE_W / 2 - chipW / 2;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: chipX, y: 3.1 + i * 0.75, w: chipW, h: 0.5,
      fill: { color: THEME.surface },
      line: { color: THEME.accent },
      rectRadius: 0.25,
    });
    slide.addText(text, {
      x: chipX, y: 3.1 + i * 0.75, w: chipW, h: 0.5,
      fontSize: 11,
      color: THEME.accent,
      fontFace: THEME.fontBody,
      align: 'center',
      valign: 'middle',
    });
  });

  // Bottom credit
  slide.addText('© Codrops / Bruno Imbrizi — MIT License', {
    x: 0, y: SLIDE_H - 0.45, w: SLIDE_W, h: 0.35,
    fontSize: 9,
    color: THEME.textMuted,
    fontFace: THEME.fontBody,
    align: 'center',
  });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
const pptx = new PptxGenJS();

// Presentation metadata
pptx.layout = 'LAYOUT_WIDE';   // 13.33 × 7.5 — but we override below
pptx.defineLayout({ name: 'CUSTOM_16_9', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'CUSTOM_16_9';

pptx.title  = DECK_TITLE;
pptx.author = 'Bruno Imbrizi / Codrops';
pptx.company = 'Interactive Particles Project';
pptx.subject = 'WebGL Interactive Particle System';

// Build slides
addCoverSlide();
addProblemSlide();
addSolutionSlide();
addHowItWorksSlide();
addTechStackSlide();
addDemoSlide();
addClosingSlide();

// Write file
const outPath = path.resolve(process.cwd(), OUTPUT_FILE);
pptx.writeFile({ fileName: outPath })
  .then(() => {
    console.log(`\n✓ Pitch deck written to: ${outPath}`);
    console.log(`  Slides: 7  |  Theme: Dark / Cyan`);
    console.log('\nCustomise with:');
    console.log('  --output <filename.pptx>   Change output path');
    console.log('  --title  "Your Title"      Change presentation title\n');
  })
  .catch((err) => {
    console.error('Failed to write pitch deck:', err);
    process.exit(1);
  });
