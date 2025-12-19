import { readFileSync } from 'fs';

console.log('🧪 Testing step_templates.llm.json (v2 - with masterName validation)\n');

const templateFile = readFileSync(new URL('./step_templates.llm.json', import.meta.url), 'utf8');
const templates = JSON.parse(templateFile);

let passed = 0;
let failed = 0;
const issues = [];

console.log('Running validation tests...\n');

// Test 1: Verify 56 layouts exist
console.log('Test 1: Layout Count');
if (templates.layouts.length === 56) {
  console.log('✅ PASS: Found 56 layouts');
  passed++;
} else {
  console.log(`❌ FAIL: Found ${templates.layouts.length} layouts, expected 56`);
  failed++;
  issues.push(`Layout count mismatch: ${templates.layouts.length} != 56`);
}

// Test 2: All layouts have masterName field
console.log('\nTest 2: MasterName Field Exists');
const layoutsWithoutMasterName = templates.layouts.filter(l => !l.masterName);
if (layoutsWithoutMasterName.length === 0) {
  console.log('✅ PASS: All layouts have masterName field');
  passed++;
} else {
  console.log(`❌ FAIL: ${layoutsWithoutMasterName.length} layouts missing masterName`);
  console.log('   Layouts:', layoutsWithoutMasterName.map(l => l.name).join(', '));
  failed++;
  issues.push(`Layouts without masterName: ${layoutsWithoutMasterName.map(l => l.name).join(', ')}`);
}

// Test 3: MasterName matches layout name
console.log('\nTest 3: MasterName Matches Layout Name');
const mismatchedMasterNames = templates.layouts.filter(l => l.masterName !== l.name);
if (mismatchedMasterNames.length === 0) {
  console.log('✅ PASS: All masterName values match layout names');
  passed++;
} else {
  console.log(`❌ FAIL: ${mismatchedMasterNames.length} layouts have mismatched masterName`);
  mismatchedMasterNames.forEach(l => {
    console.log(`   Layout: "${l.name}", masterName: "${l.masterName}"`);
  });
  failed++;
  issues.push(`Mismatched masterNames in ${mismatchedMasterNames.length} layouts`);
}

// Test 4: First code line includes masterName
console.log('\nTest 4: First Code Line Includes MasterName');
let masterNameInCodeIssues = [];
templates.layouts.forEach(layout => {
  const firstLine = layout.code[0];
  if (!firstLine || !firstLine.includes('pptx.addSlide({ masterName:')) {
    masterNameInCodeIssues.push(layout.name);
  } else {
    // Also verify the masterName value in code matches the layout name
    const match = firstLine.match(/masterName: "([^"]+)"/);
    if (match && match[1] !== layout.name) {
      masterNameInCodeIssues.push(`${layout.name} (code has "${match[1]}")`);
    }
  }
});

if (masterNameInCodeIssues.length === 0) {
  console.log('✅ PASS: All layouts have correct masterName in first code line');
  passed++;
} else {
  console.log(`❌ FAIL: ${masterNameInCodeIssues.length} layouts have issues with masterName in code`);
  console.log('   Layouts:', masterNameInCodeIssues.join(', '));
  failed++;
  issues.push(`MasterName code issues: ${masterNameInCodeIssues.join(', ')}`);
}

// Test 5: No placeholder targeting syntax
console.log('\nTest 5: No Placeholder Targeting Syntax');
let placeholderSyntaxFound = false;
templates.layouts.forEach(layout => {
  layout.code.forEach(line => {
    if (line.includes("placeholder:") && !line.includes('// ')) {
      placeholderSyntaxFound = true;
      issues.push(`Layout "${layout.name}" has placeholder: syntax in code`);
    }
  });
});

if (!placeholderSyntaxFound) {
  console.log('✅ PASS: No placeholder: syntax found');
  passed++;
} else {
  console.log('❌ FAIL: Found placeholder: syntax in code');
  failed++;
}

// Test 6: Coordinate-based positioning
console.log('\nTest 6: Coordinate-Based Positioning');
let layoutsWithoutCoords = [];
templates.layouts.forEach(layout => {
  let hasCoords = false;
  layout.code.forEach(line => {
    if (line.includes('x:') && line.includes('y:') && line.includes('w:') && line.includes('h:')) {
      hasCoords = true;
    }
  });
  // Exception for Blank and some special layouts
  if (!hasCoords && !['Blank', 'Content + Photo Black', 'Content + Photo Blue'].includes(layout.name) && layout.code.length > 1) {
    layoutsWithoutCoords.push(layout.name);
  }
});

if (layoutsWithoutCoords.length === 0) {
  console.log('✅ PASS: All layouts use coordinate-based positioning');
  passed++;
} else {
  console.log(`❌ FAIL: ${layoutsWithoutCoords.length} layouts missing coordinates`);
  console.log('   Layouts:', layoutsWithoutCoords.join(', '));
  failed++;
  issues.push(`Layouts without coordinates: ${layoutsWithoutCoords.join(', ')}`);
}

// Test 7: Icon layouts use data: property
console.log('\nTest 7: Icon Layouts Use data: Property');
const iconLayouts = templates.layouts.filter(l => l.name.includes('Icon'));
let iconsCorrect = true;
let iconIssues = [];
iconLayouts.forEach(layout => {
  layout.code.forEach(line => {
    if (line.includes('addImage') && line.includes('icon')) {
      if (!line.includes('data:')) {
        iconsCorrect = false;
        iconIssues.push(layout.name);
      }
    }
  });
});

if (iconsCorrect) {
  console.log(`✅ PASS: All ${iconLayouts.length} icon layouts use data: property`);
  passed++;
} else {
  console.log(`❌ FAIL: Some icon layouts don't use data: property`);
  console.log('   Layouts:', [...new Set(iconIssues)].join(', '));
  failed++;
  issues.push(`Icon layouts without data: property`);
}

// Test 8: Photo layouts use path: property
console.log('\nTest 8: Photo Layouts Use path: Property');
const photoLayouts = templates.layouts.filter(l => l.name.includes('Photo'));
let photosCorrect = true;
let photoIssues = [];
photoLayouts.forEach(layout => {
  layout.code.forEach(line => {
    if (line.includes('addImage') && line.includes('photo')) {
      if (!line.includes('path:')) {
        photosCorrect = false;
        photoIssues.push(layout.name);
      }
    }
  });
});

if (photosCorrect) {
  console.log(`✅ PASS: All ${photoLayouts.length} photo layouts use path: property`);
  passed++;
} else {
  console.log(`❌ FAIL: Some photo layouts don't use path: property`);
  console.log('   Layouts:', [...new Set(photoIssues)].join(', '));
  failed++;
  issues.push(`Photo layouts without path: property`);
}

// Test 9: JSON structure validation
console.log('\nTest 9: JSON Structure Validation');
let structureValid = true;
templates.layouts.forEach((layout, idx) => {
  if (!layout.name || !layout.masterName || !layout.template || !layout.instructions || !Array.isArray(layout.code)) {
    structureValid = false;
    issues.push(`Layout #${idx} missing required fields`);
  }
});

if (structureValid) {
  console.log('✅ PASS: All layouts have valid JSON structure');
  passed++;
} else {
  console.log('❌ FAIL: Some layouts have invalid structure');
  failed++;
}

// Test 10: Units field
console.log('\nTest 10: Units Configuration');
if (templates.units === 'inches') {
  console.log('✅ PASS: Units set to "inches"');
  passed++;
} else {
  console.log(`❌ FAIL: Units is "${templates.units}", expected "inches"`);
  failed++;
  issues.push(`Units field incorrect: ${templates.units}`);
}

// Summary
console.log('\n' + '='.repeat(60));
console.log('📊 Test Summary');
console.log('='.repeat(60));
console.log(`✅ Passed: ${passed}/10`);
console.log(`❌ Failed: ${failed}/10`);

if (failed > 0) {
  console.log('\n⚠️  Issues Found:');
  issues.forEach(issue => console.log(`  - ${issue}`));
  console.log('\n❌ TESTS FAILED');
  process.exit(1);
} else {
  console.log('\n🎉 ALL TESTS PASSED! Templates are ready for use.');
  console.log('\n✨ Key validations completed:');
  console.log('   - All 56 layouts present');
  console.log('   - MasterName field exists and matches layout name');
  console.log('   - First code line: pptx.addSlide({ masterName: "..." })');
  console.log('   - No placeholder targeting syntax');
  console.log('   - Coordinate-based positioning (x, y, w, h)');
  console.log('   - Icons use data: property');
  console.log('   - Photos use path: property');
  console.log('   - Valid JSON structure');
  console.log('   - Units set to "inches"');
  process.exit(0);
}
