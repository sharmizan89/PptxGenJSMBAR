#!/usr/bin/env node
/**
 * Test script to verify placeholder inheritance is working correctly
 * Creates a test PPTX with various layouts and placeholder-based content
 */

import path from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';
import fs from 'fs/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Import the built library
const PptxGenJS = (await import('../src/bld/pptxgen.es.js')).default;

async function createTestPresentation() {
  const pptx = new PptxGenJS();
  
  console.log('\n=== Placeholder Inheritance Test ===\n');
  console.log(`Total slide layouts: ${pptx.slideLayouts.length}`);
  
  // Check that layouts have placeholder objects populated
  console.log('\n--- Layout Placeholder Object Check ---');
  const layoutsToTest = [
    'Content - no subtitle',
    'Content w/Sub-headline', 
    'Two Content',
    'Two Content + Subtitles ',  // Note trailing space!
    'Title Only',
  ];
  
  for (const layoutName of layoutsToTest) {
    const layout = pptx.slideLayouts.find(l => l._name === layoutName);
    if (layout) {
      console.log(`\n[${layoutName}]`);
      console.log(`  _slideObjects count: ${layout._slideObjects?.length || 0}`);
      if (layout._slideObjects?.length > 0) {
        layout._slideObjects.forEach((obj, i) => {
          console.log(`    ${i + 1}. placeholder="${obj.options?.placeholder}", type="${obj.options?._placeholderType}", idx=${obj.options?._placeholderIdx}`);
        });
      }
    } else {
      console.log(`\n[${layoutName}] - NOT FOUND!`);
    }
  }
  
  // Create test slides with placeholder content
  console.log('\n--- Creating Test Slides ---');
  
  // Test 1: Content - no subtitle
  const slide1 = pptx.addSlide({ masterName: 'Content - no subtitle' });
  slide1.addText('Test Title using Title placeholder', { placeholder: 'Title 6' });
  slide1.addText('Test content using content placeholder', { placeholder: 'content2' });
  console.log('Created slide 1 with layout "Content - no subtitle"');
  
  // Test 2: Content w/Sub-headline
  const slide2 = pptx.addSlide({ masterName: 'Content w/Sub-headline' });
  slide2.addText('Test Title', { placeholder: 'Title 6' });
  slide2.addText('Test Sub-headline', { placeholder: 'Sub-headline' });
  slide2.addText('Test content', { placeholder: 'content2' });
  console.log('Created slide 2 with layout "Content w/Sub-headline"');
  
  // Test 3: Two Content + Subtitles (WITH trailing space!)
  const slide3 = pptx.addSlide({ masterName: 'Two Content + Subtitles ' }); // Note trailing space
  slide3.addText('Two Content Title', { placeholder: 'Title 6' });
  slide3.addText('Sub-headline text', { placeholder: 'Sub-headline' });
  slide3.addText('Left subtitle', { placeholder: 'subheadline' });
  slide3.addText('Left content', { placeholder: 'content2' });
  console.log('Created slide 3 with layout "Two Content + Subtitles " (with trailing space)');
  
  // Test 4: Title Only
  const slide4 = pptx.addSlide({ masterName: 'Title Only' });
  slide4.addText('Title Only Layout Test', { placeholder: 'Title 6' });
  console.log('Created slide 4 with layout "Title Only"');
  
  // Save the presentation
  const outputPath = path.join(__dirname, '..', 'test_placeholder_output.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log(`\nPresentation saved to: ${outputPath}`);
  
  // Extract and verify the slide XML
  console.log('\n--- Verifying Generated Slide XML ---');
  const pptxBuffer = await fs.readFile(outputPath);
  const zip = await JSZip.loadAsync(pptxBuffer);
  
  // Check slide1.xml for placeholder references
  const slide1Xml = await zip.file('ppt/slides/slide1.xml')?.async('text');
  if (slide1Xml) {
    const hasPlaceholderRef = slide1Xml.includes('<p:ph');
    const hasEmptySpPr = slide1Xml.includes('<p:spPr/>');
    console.log('\nSlide 1 XML Analysis:');
    console.log(`  Has <p:ph> placeholder reference: ${hasPlaceholderRef}`);
    console.log(`  Has empty <p:spPr/>: ${hasEmptySpPr}`);
    
    // Extract p:ph elements
    const phMatches = slide1Xml.match(/<p:ph[^>]*>/g) || [];
    console.log(`  Placeholder elements found: ${phMatches.length}`);
    phMatches.forEach((ph, i) => console.log(`    ${i + 1}. ${ph}`));
  }
  
  // Check slide layout references
  const slide1Rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('text');
  if (slide1Rels) {
    const layoutMatch = slide1Rels.match(/Target="([^"]+slideLayout[^"]+)"/);
    if (layoutMatch) {
      console.log(`\nSlide 1 references layout: ${layoutMatch[1]}`);
    }
  }
  
  console.log('\n=== Test Complete ===\n');
  
  // Check for "Two Content + Subtitles" without trailing space (should NOT exist)
  const layoutWithoutSpace = pptx.slideLayouts.find(l => l._name === 'Two Content + Subtitles');
  const layoutWithSpace = pptx.slideLayouts.find(l => l._name === 'Two Content + Subtitles ');
  
  console.log('--- Layout Name Verification ---');
  console.log(`"Two Content + Subtitles" (no space) exists: ${!!layoutWithoutSpace}`);
  console.log(`"Two Content + Subtitles " (with space) exists: ${!!layoutWithSpace}`);
  
  if (layoutWithSpace && !layoutWithoutSpace) {
    console.log('✓ Correct: Only the version with trailing space exists');
  } else {
    console.log('✗ Warning: Layout name issue detected!');
  }
}

createTestPresentation().catch(console.error);
