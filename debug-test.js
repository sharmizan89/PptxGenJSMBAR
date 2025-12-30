// Debug test to understand placeholder behavior
const PptxGenJS = require('./src/bld/pptxgen.cjs.js');
const fs = require('fs');
const path = require('path');

async function debug() {
    console.log('=== PptxGenJS Placeholder Debug Test ===\n');
    
    const pptx = new PptxGenJS();
    console.log('Version:', pptx.version);
    
    // Create a slide with "Two Content + Subtitles " layout
    const slide = pptx.addSlide({ masterName: "Two Content + Subtitles " });
    
    console.log('\n=== Slide Layout Info ===');
    console.log('Layout name:', slide._slideLayout?._name);
    console.log('Layout _slideObjects count:', slide._slideLayout?._slideObjects?.length || 0);
    
    // Show what's in the layout's _slideObjects
    console.log('\n=== Layout _slideObjects ===');
    (slide._slideLayout?._slideObjects || []).forEach((obj, i) => {
        console.log(`  [${i}] _type=${obj._type}, placeholder=${obj.options?.placeholder}, _placeholderType=${obj.options?._placeholderType}, _placeholderIdx=${obj.options?._placeholderIdx}`);
    });
    
    // Add text to leftSubtitle placeholder
    console.log('\n=== Adding text to leftSubtitle ===');
    slide.addText('[Left subtitle test]', { placeholder: 'leftSubtitle' });
    
    // Examine the slide objects
    console.log('\n=== Slide Objects After addText ===');
    console.log('Total slide objects:', slide._slideObjects.length);
    
    slide._slideObjects.forEach((obj, i) => {
        console.log(`\n--- Object ${i} ---`);
        console.log('  _type:', obj._type);
        console.log('  text:', JSON.stringify(obj.text?.map(t => t.text)).slice(0, 50));
        console.log('  options.placeholder:', obj.options?.placeholder);
        console.log('  options._placeholderType:', obj.options?._placeholderType);
        console.log('  options._placeholderIdx:', obj.options?._placeholderIdx);
        console.log('  options.x:', obj.options?.x);
        console.log('  options.y:', obj.options?.y);
    });
    
    // Test another placeholder
    console.log('\n=== Adding text to leftContent ===');
    slide.addText('[Left content test]', { placeholder: 'leftContent' });
    
    const lastObj = slide._slideObjects[slide._slideObjects.length - 1];
    console.log('  _placeholderType:', lastObj.options?._placeholderType);
    console.log('  _placeholderIdx:', lastObj.options?._placeholderIdx);
    
    // Save the file
    console.log('\n=== Saving test file ===');
    const testPath = path.join(__dirname, 'debug-output.pptx');
    await pptx.writeFile(testPath);
    console.log('Saved to:', testPath);
    
    console.log('\n=== Done ===');
}

debug().catch(console.error);
