import { readFileSync } from 'fs';

console.log('🧪 Testing step_templates.llm.json\n');

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

// Test 2: Check for placeholder: syntax (should be none)
console.log('\nTest 2: No Placeholder Targeting Syntax');
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

// Test 3: Check for x, y, w, h coordinates
console.log('\nTest 3: Coordinate-Based Positioning');
let coordinatesFound = true;
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
    coordinatesFound = false;
    layoutsWithoutCoords.push(layout.name);
  }
});

if (coordinatesFound || layoutsWithoutCoords.length === 0) {
  console.log('✅ PASS: All layouts use coordinate-based positioning');
  passed++;
} else {
  console.log(`❌ FAIL: ${layoutsWithoutCoords.length} layouts missing coordinates`);
  console.log('   Layouts:', layoutsWithoutCoords.join(', '));
  failed++;
  issues.push(`Layouts without coordinates: ${layoutsWithoutCoords.join(', ')}`);
}

// Test 4: Icon layouts use data: property
console.log('\nTest 4: Icon Layouts Use data: Property');
const iconLayouts = templates.layouts.filter(l => l.name.includes('Icon'));
let iconsCorrect = true;
iconLayouts.forEach(layout => {
  layout.code.forEach(line => {
    if (line.includes('addImage') && line.includes('icon')) {
      if (!line.includes('data:')) {
        iconsCorrect = false;
        issues.push(`Layout "${layout.name}" icon should use data: property`);
      }
    }
  });
});

if (iconsCorrect) {
  console.log(`✅ PASS: All ${iconLayouts.length} icon layouts use data: property`);
  passed++;
} else {
  console.log('❌ FAIL: Some icon layouts use wrong property');
  failed++;
}

// Test 5: Photo layouts use path: property
console.log('\nTest 5: Photo Layouts Use path: Property');
const photoLayouts = templates.layouts.filter(l => l.name.includes('Photo'));
let photosCorrect = true;
photoLayouts.forEach(layout => {
  layout.code.forEach(line => {
    if (line.includes('addImage') && line.includes('photo')) {
      if (!line.includes('path:')) {
        photosCorrect = false;
        issues.push(`Layout "${layout.name}" photo should use path: property`);
      }
    }
  });
});

if (photosCorrect) {
  console.log(`✅ PASS: All ${photoLayouts.length} photo layouts use path: property`);
  passed++;
} else {
  console.log('❌ FAIL: Some photo layouts use wrong property');
  failed++;
}

// Test 6: Verify specific test layouts
console.log('\nTest 6: Specific Layout Validation');

const testLayouts = [
  {
    name: 'Content - no subtitle',
    expectedCodeLines: 3,
    shouldHaveCoords: true
  },
  {
    name: 'Icons 3 Columns Vertical',
    expectedCodeLines: 13,
    shouldHaveCoords: true
  },
  {
    name: 'Content + Photo White',
    expectedCodeLines: 6,
    shouldHaveCoords: true
  },
  {
    name: 'Content 4 Columns',
    expectedCodeLines: 11,
    shouldHaveCoords: true
  }
];

let specificTestsPassed = 0;
testLayouts.forEach(test => {
  const layout = templates.layouts.find(l => l.name === test.name);
  if (!layout) {
    console.log(`   ❌ Layout "${test.name}" not found`);
    issues.push(`Missing layout: ${test.name}`);
    return;
  }
  
  if (layout.code.length >= test.expectedCodeLines - 2) { // Allow some variance
    console.log(`   ✅ "${test.name}" has ${layout.code.length} code lines`);
    specificTestsPassed++;
  } else {
    console.log(`   ❌ "${test.name}" has ${layout.code.length} lines, expected ~${test.expectedCodeLines}`);
    issues.push(`${test.name}: Code line count mismatch`);
  }
});

if (specificTestsPassed === testLayouts.length) {
  console.log('✅ PASS: All specific layouts validated');
  passed++;
} else {
  console.log(`❌ FAIL: ${testLayouts.length - specificTestsPassed} specific layouts failed`);
  failed++;
}

// Test 7: Check for proper JSON structure
console.log('\nTest 7: JSON Structure Validation');
let structureValid = true;
templates.layouts.forEach((layout, idx) => {
  if (!layout.name || !layout.template || !layout.instructions || !Array.isArray(layout.code)) {
    structureValid = false;
    issues.push(`Layout ${idx} missing required fields`);
  }
});

if (structureValid) {
  console.log('✅ PASS: All layouts have required fields (name, template, instructions, code)');
  passed++;
} else {
  console.log('❌ FAIL: Some layouts missing required fields');
  failed++;
}

// Test 8: Verify units field
console.log('\nTest 8: Units Configuration');
if (templates.units === 'inches') {
  console.log('✅ PASS: Units set to "inches"');
  passed++;
} else {
  console.log(`❌ FAIL: Units is "${templates.units}", expected "inches"`);
  failed++;
  issues.push('Units field incorrect');
}

// Summary
console.log('\n' + '='.repeat(60));
console.log('TEST SUMMARY');
console.log('='.repeat(60));
console.log(`✅ Passed: ${passed}/8`);
console.log(`❌ Failed: ${failed}/8`);

if (issues.length > 0) {
  console.log('\n⚠️  ISSUES FOUND:');
  issues.forEach((issue, idx) => {
    console.log(`   ${idx + 1}. ${issue}`);
  });
}

console.log('\n' + '='.repeat(60));

if (failed === 0) {
  console.log('🎉 ALL TESTS PASSED! Templates are ready for use.');
  process.exit(0);
} else {
  console.log('⚠️  SOME TESTS FAILED. Please review issues above.');
  process.exit(1);
}
