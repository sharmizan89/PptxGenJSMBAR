#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read both layout files
const layoutsPath = path.join(__dirname, '../src/cust-xml-slide-layouts.ts');
const layout1Path = path.join(__dirname, '../src/cust-xml-slide-layout1.ts');

const content = fs.readFileSync(layoutsPath, 'utf8');
const layout1Content = fs.readFileSync(layout1Path, 'utf8');

function extractPlaceholders(xml) {
  const placeholders = [];
  
  // Match all <p:sp> elements (shapes/placeholders)
  const shapeMatches = xml.matchAll(/<p:sp>.*?<\/p:sp>/gs);
  
  for (const shapeMatch of shapeMatches) {
    const shape = shapeMatch[0];
    
    // Extract name from <p:cNvPr name="...">
    const nameMatch = shape.match(/<p:cNvPr\s+id="\d+"\s+name="([^"]+)"/);
    if (!nameMatch) continue;
    
    const placeholderName = nameMatch[1];
    
    // Skip decorative elements (Rectangle with no placeholder type)
    if (placeholderName.startsWith('Rectangle') && !shape.includes('<p:ph')) continue;
    
    // Extract placeholder type and index
    const phMatch = shape.match(/<p:ph\s+([^>]*?)\/?>/ );
    let type = null;
    let idx = null;
    
    if (phMatch) {
      const typeMatch = phMatch[1].match(/type="([^"]+)"/);
      const idxMatch = phMatch[1].match(/idx="(\d+)"/);
      
      type = typeMatch ? typeMatch[1] : null;
      idx = idxMatch ? parseInt(idxMatch[1]) : null;
    }
    
    // Extract position from <a:xfrm><a:off x="..." y="..."/><a:ext cx="..." cy="..."/>
    const xfrmMatch = shape.match(/<a:xfrm>.*?<a:off\s+x="(\d+)"\s+y="(\d+)".*?<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/s);
    
    if (xfrmMatch) {
      const xEmu = parseInt(xfrmMatch[1]);
      const yEmu = parseInt(xfrmMatch[2]);
      const wEmu = parseInt(xfrmMatch[3]);
      const hEmu = parseInt(xfrmMatch[4]);
      
      placeholders.push({
        name: placeholderName,
        type: type,
        idx: idx,
        xEmu,
        yEmu,
        wEmu,
        hEmu
      });
    } else {
      // Placeholder inherits position from master
      placeholders.push({
        name: placeholderName,
        type: type,
        idx: idx,
        inheritFromMaster: true
      });
    }
  }
  
  return placeholders;
}

// Extract layout from cust-xml-slide-layout1.ts
const layout1Match = layout1Content.match(/CUSTOM_PPT_SLIDE_LAYOUT1_XML\s*=\s*`([^`]+)`/s);
const layouts = [];

if (layout1Match) {
  const xml = layout1Match[1];
  const nameMatch = xml.match(/<p:cSld\s+name="([^"]+)"/);
  const name = nameMatch ? nameMatch[1] : 'Content - no subtitle';
  
  layouts.push({
    id: 0,  // First layout, ID 0
    name,
    placeholders: extractPlaceholders(xml)
  });
}

// Extract layouts from cust-xml-slide-layouts.ts
const layoutMatches = content.matchAll(/\{\s*id:\s*(\d+),\s*name:\s*["']([^"']+)["'],\s*xml:\s*`([^`]+)`\s*\}/gs);

for (const match of layoutMatches) {
  const id = parseInt(match[1]);
  const name = match[2];
  const xml = match[3];
  
  layouts.push({
    id,
    name,
    placeholders: extractPlaceholders(xml)
  });
}

console.log(`Extracted ${layouts.length} layouts`);


// Create output JSON
const output = {
  units: "EMU",
  note: "English Metric Units where 1 inch = 914400 EMU. Coordinates extracted from cust-xml-slide-layout1.ts and cust-xml-slide-layouts.ts. Placeholder names match exact Office Open XML definitions for targeting via addText/addImage/addChart with { placeholder: 'name' }.",
  layouts: layouts.map(layout => ({
    name: layout.name,
    placeholders: layout.placeholders
  }))
};

// Write to file
const outputPath = path.join(__dirname, 'step_templates.data.json');
fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), 'utf8');

console.log(`Wrote ${layouts.length} layouts to ${outputPath}`);
console.log('\nLayout summary:');
layouts.forEach((layout, index) => {
  console.log(`${index + 1}. "${layout.name}" (ID ${layout.id}) - ${layout.placeholders.length} placeholders`);
});
