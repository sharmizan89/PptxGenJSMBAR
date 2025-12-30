// Test all 12 layouts with placeholder content
const PptxGenJS = require('./dist/pptxgen.cjs.js');

async function testAllLayouts() {
    console.log('=== PptxGenJS Full Layout Test ===\n');
    
    const pptx = new PptxGenJS();
    console.log('Version:', pptx.version);
    
    // Layout 1: Title
    const slide1 = pptx.addSlide({ masterName: "Title" });
    slide1.addText('1', { placeholder: 'slideNumber' });
    slide1.addText('[Title Slide Headline]', { placeholder: 'headline' });
    slide1.addText('[Section name]', { placeholder: 'footer' });
    console.log('Slide 1: Title - objects:', slide1._slideObjects.length);
    
    // Layout 2: Picture w/ Caption
    const slide2 = pptx.addSlide({ masterName: "Picture w/ Caption" });
    slide2.addText('2', { placeholder: 'slideNumber' });
    slide2.addText('[Picture Caption]', { placeholder: 'headline' });
    slide2.addText('[Description text]', { placeholder: 'description' });
    slide2.addText('[Source]', { placeholder: 'source' });
    console.log('Slide 2: Picture w/ Caption - objects:', slide2._slideObjects.length);
    
    // Layout 3: Content w/Sub-headline
    const slide3 = pptx.addSlide({ masterName: "Content w/Sub-headline" });
    slide3.addText('3', { placeholder: 'slideNumber' });
    slide3.addText('[Headline]', { placeholder: 'headline' });
    slide3.addText('[Subheadline]', { placeholder: 'subheadline' });
    slide3.addText('[Content]', { placeholder: 'mainContent' });
    slide3.addText('[Footer]', { placeholder: 'footer' });
    slide3.addText('[Source]', { placeholder: 'source' });
    console.log('Slide 3: Content w/Sub-headline - objects:', slide3._slideObjects.length);
    
    // Layout 4: Two Content
    const slide4 = pptx.addSlide({ masterName: "Two Content" });
    slide4.addText('4', { placeholder: 'slideNumber' });
    slide4.addText('[Headline]', { placeholder: 'headline' });
    slide4.addText('[Subheadline]', { placeholder: 'subheadline' });
    slide4.addText('[Left Content]', { placeholder: 'leftContent' });
    slide4.addText('[Right Content]', { placeholder: 'rightContent' });
    slide4.addText('[Footer]', { placeholder: 'footer' });
    slide4.addText('[Source]', { placeholder: 'source' });
    console.log('Slide 4: Two Content - objects:', slide4._slideObjects.length);
    
    // Layout 5: Two Content + Subtitles (THE CRITICAL TEST!)
    const slide5 = pptx.addSlide({ masterName: "Two Content + Subtitles " });
    slide5.addText('5', { placeholder: 'slideNumber' });
    slide5.addText('[Headline]', { placeholder: 'headline' });
    slide5.addText('[Subheadline]', { placeholder: 'subheadline' });
    slide5.addText('[Left Subtitle]', { placeholder: 'leftSubtitle' });
    slide5.addText('[Left Content]', { placeholder: 'leftContent' });
    slide5.addText('[Right Subtitle]', { placeholder: 'rightSubtitle' });
    slide5.addText('[Right Content]', { placeholder: 'rightContent' });
    slide5.addText('[Footer]', { placeholder: 'footer' });
    slide5.addText('[Source]', { placeholder: 'source' });
    console.log('Slide 5: Two Content + Subtitles - objects:', slide5._slideObjects.length);
    
    // Check for duplicates by counting unique idx values
    const slide5IdxCount = {};
    slide5._slideObjects.forEach(obj => {
        const idx = obj.options?._placeholderIdx;
        const type = obj.options?._placeholderType;
        const key = idx !== undefined ? `idx:${idx}` : `type:${type}`;
        slide5IdxCount[key] = (slide5IdxCount[key] || 0) + 1;
    });
    console.log('  Placeholder idx counts:', slide5IdxCount);
    const duplicates = Object.entries(slide5IdxCount).filter(([k, v]) => v > 1);
    if (duplicates.length > 0) {
        console.log('  ⚠️ DUPLICATES FOUND:', duplicates);
    } else {
        console.log('  ✅ No duplicates!');
    }
    
    // Layout 6-12 (abbreviated tests)
    const layouts = [
        'Content 4 Columns',
        'Content 5 Columns',
        'Content - Sidebar',
        'Content - no subtitle',
        'Full Image w/top bar',
        'Video',
        'Close'
    ];
    
    for (let i = 0; i < layouts.length; i++) {
        const slide = pptx.addSlide({ masterName: layouts[i] });
        slide.addText(`${6 + i}`, { placeholder: 'slideNumber' });
        slide.addText('[Headline]', { placeholder: 'headline' });
        console.log(`Slide ${6 + i}: ${layouts[i]} - objects:`, slide._slideObjects.length);
    }
    
    // Save
    const filename = 'full-layout-test.pptx';
    await pptx.writeFile({ fileName: filename });
    console.log('\nSaved to:', filename);
    console.log('\n=== Test Complete ===');
}

testAllLayouts().catch(console.error);
