#!/usr/bin/env node
/**
 * Test script to verify 2026 Energy template integration
 * Creates a simple PPTX using the new layouts
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Import the built library
const PptxGenJS = (await import('../src/bld/pptxgen.es.js')).default;

console.log('🧪 Testing 2026 Energy Template Integration\n');

const pptx = new PptxGenJS();

// Test 1: Verify slide size is correct (2026 is 16:9)
console.log('Test 1: Slide Size');
console.log(`  Slide size: ${pptx.presLayout.width}x${pptx.presLayout.height} inches`);
console.log(`  Expected: 13.333x7.5 inches (16:9)`);
if (Math.abs(pptx.presLayout.width - 13.333) < 0.01 && Math.abs(pptx.presLayout.height - 7.5) < 0.01) {
  console.log('  ✅ PASS: Slide size is correct\n');
} else {
  console.log('  ❌ FAIL: Slide size incorrect\n');
}

// Test 2: Create slides with various layouts from 2026 template
console.log('Test 2: Creating slides with 2026 layouts');

const testLayouts = [
  'Content - no subtitle',
  'Content w/Sub-headline',
  'Two Content',
  'Content 4 Columns',
  'Title Slide',
  'Blank',
];

let slideCount = 0;
for (const layoutName of testLayouts) {
  try {
    const slide = pptx.addSlide({ masterName: layoutName });
    slide.addText(`Test slide using layout: ${layoutName}`, {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 1,
      fontSize: 24,
      color: '000000',
    });
    slideCount++;
    console.log(`  ✅ Created slide with layout: "${layoutName}"`);
  } catch (error) {
    console.log(`  ❌ Failed to create slide with layout: "${layoutName}"`);
    console.log(`     Error: ${error.message}`);
  }
}

console.log(`\n  Created ${slideCount}/${testLayouts.length} test slides\n`);

// Test 3: Write the file
console.log('Test 3: Writing PPTX file');
const outputPath = path.join(__dirname, '..', 'test-output-2026.pptx');

try {
  await pptx.writeFile({ fileName: outputPath });
  
  if (fs.existsSync(outputPath)) {
    const stats = fs.statSync(outputPath);
    console.log(`  ✅ PASS: File created successfully`);
    console.log(`  Output: ${outputPath}`);
    console.log(`  Size: ${(stats.size / 1024).toFixed(1)} KB\n`);
  } else {
    console.log('  ❌ FAIL: File not created\n');
  }
} catch (error) {
  console.log(`  ❌ FAIL: Error writing file: ${error.message}\n`);
}

// Summary
console.log('=' .repeat(50));
console.log('Test complete!');
console.log('Open the generated PPTX file in PowerPoint to verify it works correctly.');
