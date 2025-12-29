#!/usr/bin/env node
/**
 * Extract 2026 Energy Layouts
 * 
 * This script extracts all slide layouts from the 2026 Energy.zip template,
 * removing all picture/image elements as specified (they will be added via testflow).
 * 
 * Output:
 * - cust-xml-slide-layouts-2026.ts: Array of all layout XML definitions
 * - cust-xml-placeholder-registry-2026.ts: Updated placeholder registry
 * - cust-xml-2026.ts: Updated XML templates (theme, content types, etc.)
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const EXTRACTED_PATH = path.join(__dirname, '../2026_Energy_extracted');
const OUTPUT_PATH = path.join(__dirname, '../src');

// Read all slide layouts
function getAllLayoutNames() {
  const layoutsDir = path.join(EXTRACTED_PATH, 'ppt/slideLayouts');
  const files = fs.readdirSync(layoutsDir)
    .filter(f => f.match(/^slideLayout\d+\.xml$/))
    .sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)[0]);
      const numB = parseInt(b.match(/\d+/)[0]);
      return numA - numB;
    });
  return files;
}

// Remove all <p:pic> elements (picture placeholders and embedded images)
function removePictureElements(xml) {
  // Remove entire <p:pic>...</p:pic> blocks
  let result = xml.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, '');
  
  // Also remove <mc:AlternateContent> blocks that contain SVG+PNG fallback images
  result = result.replace(/<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/g, '');
  
  return result;
}

// Remove image relationships from layout rels files
function getCleanLayoutRels(layoutNum) {
  const relsPath = path.join(EXTRACTED_PATH, `ppt/slideLayouts/_rels/slideLayout${layoutNum}.xml.rels`);
  if (!fs.existsSync(relsPath)) {
    return null;
  }
  
  let xml = fs.readFileSync(relsPath, 'utf8');
  
  // Remove image relationships (keep only slideMaster relationship)
  xml = xml.replace(/<Relationship[^>]*Target="[^"]*media\/image[^"]*"[^>]*\/>/g, '');
  xml = xml.replace(/<Relationship[^>]*Target="[^"]*\/media\/[^"]*"[^>]*\/>/g, '');
  
  return xml;
}

// Extract placeholder info from a layout
function extractPlaceholders(xml) {
  const placeholders = [];
  
  // Match all <p:sp> elements (shapes/placeholders)
  const shapeMatches = xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g);
  
  for (const shapeMatch of shapeMatches) {
    const shape = shapeMatch[0];
    
    // Extract name from <p:cNvPr id="..." name="...">
    const nameMatch = shape.match(/<p:cNvPr\s+id="(\d+)"\s+name="([^"]+)"/);
    if (!nameMatch) continue;
    
    const id = parseInt(nameMatch[1]);
    const placeholderName = nameMatch[2];
    
    // Extract placeholder type and index
    const phMatch = shape.match(/<p:ph\s+([^>]*?)\/?>/);
    let type = null;
    let idx = null;
    let sz = null;
    
    if (phMatch) {
      const typeMatch = phMatch[1].match(/type="([^"]+)"/);
      const idxMatch = phMatch[1].match(/idx="(\d+)"/);
      const szMatch = phMatch[1].match(/sz="([^"]+)"/);
      
      type = typeMatch ? typeMatch[1] : 'body'; // default to body if no type specified
      idx = idxMatch ? parseInt(idxMatch[1]) : null;
      sz = szMatch ? szMatch[1] : null;
    }
    
    // Extract position from <a:xfrm>
    const xfrmMatch = shape.match(/<a:xfrm>[\s\S]*?<a:off\s+x="(\d+)"\s+y="(\d+)"[\s\S]*?<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    
    let position = null;
    if (xfrmMatch) {
      position = {
        x: parseInt(xfrmMatch[1]),
        y: parseInt(xfrmMatch[2]),
        cx: parseInt(xfrmMatch[3]),
        cy: parseInt(xfrmMatch[4])
      };
    }
    
    placeholders.push({
      id,
      name: placeholderName,
      type,
      idx,
      sz,
      position,
      hasPosition: !!position
    });
  }
  
  return placeholders;
}

// Map placeholder names to testflow-style semantic names
function mapToSemanticName(name, type) {
  const lowerName = name.toLowerCase();
  
  if (type === 'title' || lowerName.includes('title')) return 'headline';
  if (lowerName.includes('sub-headline') || lowerName.includes('subheadline')) return 'subheadline';
  if (lowerName.includes('subtitle')) {
    // Check for numbered subtitles
    const numMatch = name.match(/(\d+)/);
    if (numMatch) return `subtitle${numMatch[1]}`;
    return 'subheadline';
  }
  if (lowerName.includes('content') && !lowerName.includes('sidebar')) {
    const numMatch = name.match(/(\d+)/);
    if (numMatch) return `content${numMatch[1]}`;
    if (lowerName.includes('left')) return 'contentLeft';
    if (lowerName.includes('right')) return 'contentRight';
    return 'mainContent';
  }
  if (lowerName.includes('sidebar') || lowerName.includes('side content')) return 'sideContent';
  if (lowerName.includes('footnote') || lowerName.includes('source') || lowerName.includes('legal')) return 'source';
  if (type === 'ftr' || lowerName.includes('footer')) return 'footer';
  if (type === 'sldNum' || lowerName.includes('slide number')) return 'slideNumber';
  if (type === 'pic' || lowerName.includes('image') || lowerName.includes('picture') || lowerName.includes('photo')) {
    const numMatch = name.match(/(\d+)/);
    if (numMatch) return `image${numMatch[1]}`;
    return 'image';
  }
  
  return name;
}

// Generate placeholder registry entry for a layout
function generateRegistryEntry(layoutName, placeholders) {
  const entry = {};
  
  for (const ph of placeholders) {
    // Skip pic types as per user request (images will be added via testflow)
    if (ph.type === 'pic') continue;
    
    const semanticName = mapToSemanticName(ph.name, ph.type);
    
    if (ph.type === 'title') {
      entry[semanticName] = { type: 'title' };
    } else if (ph.idx !== null) {
      entry[semanticName] = { type: ph.type || 'body', idx: ph.idx };
    }
  }
  
  return entry;
}

// Main extraction
function extractAllLayouts() {
  const layoutFiles = getAllLayoutNames();
  const layouts = [];
  const registry = {};
  
  console.log(`Found ${layoutFiles.length} layout files`);
  
  for (const file of layoutFiles) {
    const layoutNum = parseInt(file.match(/\d+/)[0]);
    const filePath = path.join(EXTRACTED_PATH, 'ppt/slideLayouts', file);
    let xml = fs.readFileSync(filePath, 'utf8');
    
    // Extract layout name
    const nameMatch = xml.match(/<p:cSld\s+name="([^"]+)"/);
    const layoutName = nameMatch ? nameMatch[1] : `Layout ${layoutNum}`;
    
    // Skip empty/placeholder layouts (names starting with ~)
    const isPlaceholder = layoutName.startsWith('~') || layoutName.trim() === '';
    
    // Extract placeholders BEFORE removing pictures (for registry)
    const placeholders = extractPlaceholders(xml);
    
    // Remove picture elements
    xml = removePictureElements(xml);
    
    // Clean up the XML (remove extra whitespace from removed elements)
    xml = xml.replace(/\n\s*\n/g, '\n');
    
    // Escape backticks for template literal
    xml = xml.replace(/`/g, '\\`');
    
    layouts.push({
      index: layoutNum - 1, // 0-based index
      layoutNum,
      name: layoutName,
      isPlaceholder,
      xml,
      placeholders
    });
    
    // Add to registry
    if (!isPlaceholder) {
      registry[layoutName] = generateRegistryEntry(layoutName, placeholders);
    }
    
    console.log(`  Layout ${layoutNum}: "${layoutName}" - ${placeholders.length} placeholders`);
  }
  
  return { layouts, registry };
}

// Read theme file
function readTheme() {
  const themePath = path.join(EXTRACTED_PATH, 'ppt/theme/theme1.xml');
  let xml = fs.readFileSync(themePath, 'utf8');
  xml = xml.replace(/`/g, '\\`');
  return xml;
}

// Read slide master
function readSlideMaster() {
  const masterPath = path.join(EXTRACTED_PATH, 'ppt/slideMasters/slideMaster1.xml');
  let xml = fs.readFileSync(masterPath, 'utf8');
  
  // Remove picture elements from slide master
  xml = removePictureElements(xml);
  xml = xml.replace(/`/g, '\\`');
  
  return xml;
}

// Read slide master rels
function readSlideMasterRels() {
  const relsPath = path.join(EXTRACTED_PATH, 'ppt/slideMasters/_rels/slideMaster1.xml.rels');
  let xml = fs.readFileSync(relsPath, 'utf8');
  
  // Remove image relationships
  xml = xml.replace(/<Relationship[^>]*Target="[^"]*media\/image[^"]*"[^>]*\/>/g, '');
  xml = xml.replace(/`/g, '\\`');
  
  return xml;
}

// Read Content_Types.xml
function readContentTypes() {
  const ctPath = path.join(EXTRACTED_PATH, '[Content_Types].xml');
  let xml = fs.readFileSync(ctPath, 'utf8');
  xml = xml.replace(/`/g, '\\`');
  return xml;
}

// Read presentation.xml
function readPresentation() {
  const presPath = path.join(EXTRACTED_PATH, 'ppt/presentation.xml');
  let xml = fs.readFileSync(presPath, 'utf8');
  xml = xml.replace(/`/g, '\\`');
  return xml;
}

// Read customXml items
function readCustomXml() {
  const items = [];
  for (let i = 1; i <= 3; i++) {
    const itemPath = path.join(EXTRACTED_PATH, `customXml/item${i}.xml`);
    if (fs.existsSync(itemPath)) {
      let xml = fs.readFileSync(itemPath, 'utf8');
      xml = xml.replace(/`/g, '\\`');
      items.push(xml);
    }
  }
  return items;
}

// Read docProps
function readDocProps() {
  const appPath = path.join(EXTRACTED_PATH, 'docProps/app.xml');
  const corePath = path.join(EXTRACTED_PATH, 'docProps/core.xml');
  
  let appXml = fs.readFileSync(appPath, 'utf8');
  let coreXml = fs.readFileSync(corePath, 'utf8');
  
  appXml = appXml.replace(/`/g, '\\`');
  coreXml = coreXml.replace(/`/g, '\\`');
  
  return { app: appXml, core: coreXml };
}

// Generate the cust-xml-slide-layouts.ts file
function generateLayoutsFile(layouts) {
  let output = `// Auto-generated from 2026 Energy.zip template
// Generated on: ${new Date().toISOString()}
// Total layouts: ${layouts.length}
// Note: All picture/image elements have been removed (will be added via testflow)

export interface SlideLayoutDef {
  id: number;
  name: string;
  xml: string;
}

export const CUSTOM_SLIDE_LAYOUT_DEFS: SlideLayoutDef[] = [
`;

  for (const layout of layouts) {
    output += `  // Layout ${layout.layoutNum}: ${layout.name}
  {
    id: ${layout.index},
    name: "${layout.name.replace(/"/g, '\\"')}",
    xml: \`${layout.xml}\`
  },
`;
  }

  output += `];

// Layout name to index mapping
export const LAYOUT_NAME_TO_INDEX: { [key: string]: number } = {
`;

  for (const layout of layouts) {
    if (!layout.isPlaceholder) {
      output += `  "${layout.name.replace(/"/g, '\\"')}": ${layout.index},
`;
    }
  }

  output += `};

// Total number of layouts
export const TOTAL_LAYOUTS = ${layouts.length};
`;

  return output;
}

// Generate the placeholder registry file
function generateRegistryFile(registry) {
  let output = `// Auto-generated placeholder registry from 2026 Energy.zip template
// Generated on: ${new Date().toISOString()}
// Maps layout name -> placeholder name -> { type, idx }
// Note: Picture (pic) placeholders are excluded (will be added via testflow)

export interface PlaceholderIdxRegistry {
  [layoutName: string]: {
    pic?: number[];
    body?: number[];
    title?: number[];
    ftr?: number[];
    sldNum?: number[];
    dt?: number[];
  };
}

export interface PlaceholderNameRegistry {
  [layoutName: string]: {
    [placeholderName: string]: { type: string; idx?: number };
  };
}

export const PLACEHOLDER_NAME_REGISTRY: PlaceholderNameRegistry = {
`;

  for (const [layoutName, placeholders] of Object.entries(registry)) {
    output += `  "${layoutName.replace(/"/g, '\\"')}": {\n`;
    for (const [phName, phInfo] of Object.entries(placeholders)) {
      if (phInfo.idx !== undefined) {
        output += `    ${phName}: { type: '${phInfo.type}', idx: ${phInfo.idx} },\n`;
      } else {
        output += `    ${phName}: { type: '${phInfo.type}' },\n`;
      }
    }
    output += `  },\n`;
  }

  output += `};
`;

  return output;
}

// Generate the cust-xml.ts updates
function generateCustXmlFile(theme, slideMaster, slideMasterRels, customXml, docProps, layoutCount) {
  let output = `// Auto-generated XML templates from 2026 Energy.zip template
// Generated on: ${new Date().toISOString()}
// Note: All picture/image elements have been removed (will be added via testflow)

`;

  // Generate Content_Types.xml with correct number of layouts
  output += `// [Content_Types].xml
export const CUSTOM_CONTENT_TYPES_XML = \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Default Extension="jpg" ContentType="image/jpeg"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="gif" ContentType="image/gif"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
<Override PartName="/customXml/itemProps2.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
<Override PartName="/customXml/itemProps3.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`;

  // Add all ${layoutCount} slideLayouts
  for (let i = 1; i <= layoutCount; i++) {
    output += `
<Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`;
  }

  output += `
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>\`;

`;

  // Theme
  output += `// ppt/theme/theme1.xml - S-P Global EDP 2026 16-9
export const CUSTOM_PPT_THEME1_XML = \`${theme}\`;

`;

  // Slide Master
  output += `// ppt/slideMasters/slideMaster1.xml
export const CUSTOM_PPT_SLIDE_MASTER_XML = \`${slideMaster}\`;

`;

  // Slide Master Rels
  output += `// ppt/slideMasters/_rels/slideMaster1.xml.rels
export const CUSTOM_PPT_SLIDE_MASTER_REL_XML = \`${slideMasterRels}\`;

`;

  // CustomXml items
  if (customXml.length >= 1) {
    output += `// customXml/item1.xml (SharePoint Form Templates)
export const CUSTOMXML_ITEM1 = \`${customXml[0]}\`;

`;
  }
  if (customXml.length >= 2) {
    output += `// customXml/item2.xml (SharePoint Content Type)
export const CUSTOMXML_ITEM2 = \`${customXml[1]}\`;

`;
  }
  if (customXml.length >= 3) {
    output += `// customXml/item3.xml (SharePoint Properties)
export const CUSTOMXML_ITEM3 = \`${customXml[2]}\`;

`;
  }

  // docProps
  output += `// docProps/app.xml
export const CUSTOM_PROPS_APP_XML = \`${docProps.app}\`;

`;

  return output;
}

// Generate slide layout rels file
function generateLayoutRelsFile(layoutCount) {
  let output = `// Auto-generated slide layout relationships from 2026 Energy.zip template
// Generated on: ${new Date().toISOString()}
// All layouts reference slideMaster1.xml

export interface LayoutRelDef {
  layoutNum: number;
  xml: string;
}

// Base rels for layouts (all reference slideMaster1)
export const CUSTOM_SLIDE_LAYOUT_RELS: LayoutRelDef[] = [
`;

  for (let i = 1; i <= layoutCount; i++) {
    output += `  {
    layoutNum: ${i},
    xml: \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>\`
  },
`;
  }

  output += `];
`;

  return output;
}

// Main
async function main() {
  console.log('=== Extracting 2026 Energy Layouts ===\n');
  
  // Check if extracted folder exists
  if (!fs.existsSync(EXTRACTED_PATH)) {
    console.error(`Error: ${EXTRACTED_PATH} does not exist. Please extract 2026 Energy.zip first.`);
    process.exit(1);
  }
  
  // Extract all layouts
  const { layouts, registry } = extractAllLayouts();
  console.log(`\nExtracted ${layouts.length} layouts`);
  
  // Read supporting files
  console.log('\nReading supporting files...');
  const theme = readTheme();
  console.log('  - Theme');
  const slideMaster = readSlideMaster();
  console.log('  - Slide Master');
  const slideMasterRels = readSlideMasterRels();
  console.log('  - Slide Master Rels');
  const customXml = readCustomXml();
  console.log(`  - Custom XML (${customXml.length} items)`);
  const docProps = readDocProps();
  console.log('  - Doc Props');
  
  // Generate output files
  console.log('\nGenerating output files...');
  
  const layoutsOutput = generateLayoutsFile(layouts);
  fs.writeFileSync(path.join(OUTPUT_PATH, 'cust-xml-slide-layouts-2026.ts'), layoutsOutput);
  console.log('  - cust-xml-slide-layouts-2026.ts');
  
  const registryOutput = generateRegistryFile(registry);
  fs.writeFileSync(path.join(OUTPUT_PATH, 'cust-xml-placeholder-registry-2026.ts'), registryOutput);
  console.log('  - cust-xml-placeholder-registry-2026.ts');
  
  const custXmlOutput = generateCustXmlFile(theme, slideMaster, slideMasterRels, customXml, docProps, layouts.length);
  fs.writeFileSync(path.join(OUTPUT_PATH, 'cust-xml-2026.ts'), custXmlOutput);
  console.log('  - cust-xml-2026.ts');
  
  const layoutRelsOutput = generateLayoutRelsFile(layouts.length);
  fs.writeFileSync(path.join(OUTPUT_PATH, 'cust-xml-slide-layout-rels-2026.ts'), layoutRelsOutput);
  console.log('  - cust-xml-slide-layout-rels-2026.ts');
  
  console.log('\n=== Extraction Complete ===');
  console.log(`\nNext steps:`);
  console.log('1. Review the generated files in src/');
  console.log('2. Replace the existing cust-xml.ts, cust-xml-slide-layouts.ts, etc. with the new versions');
  console.log('3. Run npm run build to compile');
  console.log('4. Test with testflow');
}

main().catch(console.error);
