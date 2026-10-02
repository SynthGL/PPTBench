#!/usr/bin/env node
// PPTBench helper for pptx-automizer.
//
// Usage: node pptbench-automizer.js LANE INPUT OUTPUT
//
// pptx-automizer cannot jump to a shape inside an existing deck. Its documented
// way to edit one file ("Loop through the slides of a presentation") loads the
// deck as the truncated root (`removeExistingSlides: true`), loads it a second
// time as a template, and re-adds every slide in presentation order, attaching
// modification callbacks to the slides that change. All three lanes use that
// documented flow; the output is whatever `write()` produces.
'use strict';

const path = require('node:path');
const { default: Automizer, modify } = require('pptx-automizer');

const TEMPLATE = 'deck';

function shapeText(element, paragraphIndex, value) {
  // Documented xmldom callback form: set the text of the first run of the
  // target paragraph and drop its other runs, keeping run properties.
  const paragraphs = element.getElementsByTagName('a:p');
  if (paragraphIndex >= paragraphs.length) {
    throw new Error(`paragraph ${paragraphIndex} is missing`);
  }
  const paragraph = paragraphs.item(paragraphIndex);
  const texts = paragraph.getElementsByTagName('a:t');
  if (texts.length === 0) {
    throw new Error(`paragraph ${paragraphIndex} has no text run`);
  }
  texts.item(0).textContent = value;
  for (let index = texts.length - 1; index > 0; index -= 1) {
    const run = texts.item(index).parentNode;
    run.parentNode.removeChild(run);
  }
}

function tableCell(element, row, column, value) {
  const rows = element.getElementsByTagName('a:tr');
  if (row >= rows.length) {
    throw new Error(`table row ${row} is missing`);
  }
  const cells = rows.item(row).getElementsByTagName('a:tc');
  if (column >= cells.length) {
    throw new Error(`table column ${column} is missing`);
  }
  shapeText(cells.item(column), 0, value);
}

const templateMutation = {
  4: (slide) => {
    slide.modifyElement('PPTBenchTable', [
      (element) => {
        tableCell(element, 1, 1, 'UPDATED-TABLE-A');
        tableCell(element, 2, 2, 'UPDATED-TABLE-B');
      },
    ]);
  },
  6: (slide) => {
    slide.modifyElement('PPTBenchBullets', [
      (element) => shapeText(element, 1, 'UPDATED-BULLET'),
    ]);
  },
};

function pointCategories(points) {
  return points.map((point, index) => ({ label: String(index + 1), values: [point] }));
}

const chartData = {
  0: (slide) => {
    slide.modifyElement('Chart 1', [
      modify.setChartData({
        series: [{ label: 'Revenue' }],
        categories: [
          { label: 'East', values: [42] },
          { label: '42', values: [84] },
          { label: '84', values: [126] },
        ],
      }),
    ]);
  },
  1: (slide) => {
    slide.modifyElement('Chart 1', [
      modify.setChartScatter({
        series: [{ label: 'Trend' }],
        categories: pointCategories([
          { x: 5, y: 13 },
          { x: 21, y: 34 },
          { x: 55, y: 89 },
        ]),
      }),
    ]);
  },
  2: (slide) => {
    slide.modifyElement('Chart 1', [
      modify.setChartBubbles({
        series: [{ label: 'Pipeline' }],
        categories: pointCategories([
          { x: 5, y: 13, size: 21 },
          { x: 34, y: 55, size: 89 },
          { x: 8, y: 13, size: 21 },
        ]),
      }),
    ]);
  },
};

const LANES = {
  'feature-matrix': {},
  'template-mutation': templateMutation,
  'chart-data': chartData,
};

async function main(argv) {
  if (argv.length !== 3) {
    process.stderr.write('usage: pptbench-automizer.js LANE INPUT OUTPUT\n');
    return 2;
  }
  const [lane, input, output] = argv;
  const callbacks = LANES[lane];
  if (callbacks === undefined) {
    process.stderr.write(`unsupported lane: ${lane}\n`);
    return 2;
  }
  const source = path.resolve(input);
  const target = path.resolve(output);
  const automizer = new Automizer({
    templateDir: path.dirname(source),
    outputDir: path.dirname(target),
    removeExistingSlides: true,
    autoImportSlideMasters: false,
    cleanup: true,
    verbosity: 0,
  });
  const pres = automizer
    .loadRoot(path.basename(source))
    .load(path.basename(source), TEMPLATE);
  const info = await pres.getInfo();
  const slides = info.slidesByTemplate(TEMPLATE);
  slides.forEach((slide, position) => {
    pres.addSlide(TEMPLATE, slide.number, callbacks[position]);
  });
  await pres.write(path.basename(target));
  return 0;
}

main(process.argv.slice(2)).then(
  (code) => {
    process.exitCode = code;
  },
  (error) => {
    process.stderr.write(`${error && error.stack ? error.stack : error}\n`);
    process.exitCode = 1;
  },
);
