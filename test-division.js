// Test division filtering of energy-specific layouts
const PptxGenJS = require('./src/bld/pptxgen.cjs.js');

async function testDivisionFiltering() {
    console.log('=== PptxGenJS Division Filtering Test ===\n');
    
    // Test 1: Default (no division set) - should exclude energy layouts
    console.log('Test 1: Default division (empty)');
    const pptx1 = new PptxGenJS();
    console.log('  Total layouts:', pptx1.slideLayouts.length);
    const energyLayouts1 = pptx1.slideLayouts.filter(l => 
        l._name.startsWith('CERA_') || 
        l._name.startsWith('Horizons_') || 
        l._name.startsWith('Platts_') ||
        l._name === 'Energy' ||
        l._name === 'Energy_Spark' ||
        l._name.includes('Companies')
    );
    console.log('  Energy layouts:', energyLayouts1.length);
    console.log('  Energy layout names:', energyLayouts1.map(l => l._name));
    console.log('  Result:', energyLayouts1.length === 0 ? '✅ PASS (no energy layouts)' : '❌ FAIL (energy layouts present)');
    
    // Test 2: Division = 'energy' - should include energy layouts
    console.log('\nTest 2: Division set to "energy"');
    const pptx2 = new PptxGenJS();
    pptx2.division = 'energy';
    console.log('  Total layouts:', pptx2.slideLayouts.length);
    const energyLayouts2 = pptx2.slideLayouts.filter(l => 
        l._name.startsWith('CERA_') || 
        l._name.startsWith('Horizons_') || 
        l._name.startsWith('Platts_') ||
        l._name === 'Energy' ||
        l._name === 'Energy_Spark' ||
        l._name.includes('Companies')
    );
    console.log('  Energy layouts:', energyLayouts2.length);
    console.log('  Energy layout names:', energyLayouts2.map(l => l._name));
    console.log('  Result:', energyLayouts2.length > 0 ? '✅ PASS (energy layouts present)' : '❌ FAIL (no energy layouts)');
    
    // Test 3: Division = 'other' - should exclude energy layouts
    console.log('\nTest 3: Division set to "other"');
    const pptx3 = new PptxGenJS();
    pptx3.division = 'other';
    console.log('  Total layouts:', pptx3.slideLayouts.length);
    const energyLayouts3 = pptx3.slideLayouts.filter(l => 
        l._name.startsWith('CERA_') || 
        l._name.startsWith('Horizons_') || 
        l._name.startsWith('Platts_') ||
        l._name === 'Energy' ||
        l._name === 'Energy_Spark' ||
        l._name.includes('Companies')
    );
    console.log('  Energy layouts:', energyLayouts3.length);
    console.log('  Result:', energyLayouts3.length === 0 ? '✅ PASS (no energy layouts)' : '❌ FAIL (energy layouts present)');
    
    // Test 4: Can we still use normal layouts after setting division?
    console.log('\nTest 4: Normal layout still works after division set');
    const pptx4 = new PptxGenJS();
    const slide4 = pptx4.addSlide({ masterName: "Two Content + Subtitles " });
    slide4.addText('Test', { placeholder: 'leftSubtitle' });
    console.log('  Slide created:', slide4._slideObjects.length > 0 ? '✅ PASS' : '❌ FAIL');
    
    // Test 5: Energy layout only available when division = 'energy'
    console.log('\nTest 5: Energy layout only available when division is "energy"');
    const pptx5 = new PptxGenJS();
    pptx5.division = 'energy';
    const energyLayoutExists = pptx5.slideLayouts.some(l => l._name === 'Energy');
    console.log('  "Energy" layout exists:', energyLayoutExists ? '✅ YES' : '❌ NO');
    
    console.log('\n=== Test Complete ===');
}

testDivisionFiltering().catch(console.error);
