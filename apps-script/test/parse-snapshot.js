'use strict';
const fs = require('fs');
const path = require('path');

function parseCell(raw, colIdx) {
  const v = raw.trim();
  if (v === '') return '';
  if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
  if (colIdx === 0) {
    const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
    if (m) {
      const yyyy = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
      return new Date(yyyy, Number(m[2]) - 1, Number(m[1]));
    }
  }
  return v;
}

function parseBlock(block) {
  const lines = block.split('\n').map((l) => l.trim()).filter(Boolean);
  const grid = [];
  lines.forEach((line) => {
    if (!line.startsWith('|')) return;
    if (/^\|(\s*:-:\s*\|)+$/.test(line)) return; // fila separadora de markdown
    const cells = line.split('|').slice(1, -1).map((c) => c);
    grid.push(cells.map((c, i) => parseCell(c, i)));
  });
  return grid;
}

function loadSnapshot() {
  const raw = fs.readFileSync(path.join(__dirname, 'real-snapshot.md'), 'utf8');
  const blocks = raw.split(/\n\s*\n/).filter((b) => b.trim());
  if (blocks.length !== 3) throw new Error('Se esperaban 3 tablas, se encontraron ' + blocks.length);
  return {
    Tenants: parseBlock(blocks[0]),
    'To Landlord': parseBlock(blocks[1]),
    Expenses: parseBlock(blocks[2])
  };
}

module.exports = { loadSnapshot, parseCell };
