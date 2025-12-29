#!/usr/bin/env node
/**
 * Regenerate placeholder registry from 2026 Energy template
 * Creates both PLACEHOLDER_NAME_REGISTRY and PLACEHOLDER_IDX_REGISTRY
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const layoutsDir = path.join(__dirname, '..', '2026_Energy_extracted', 'ppt', 'slideLayouts');
const outputFile = path.join(__dirname, '..', 'src', 'cust-xml-placeholder-registry.ts');

// Read all slide layouts
const layoutFiles = fs.readdirSync(layoutsDir)
  .filter(f => f.endsWith('.xml'))
  .sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)[0]);
    const numB = parseInt(b.match(/\d+/)[0]);
    return numA - numB;
  });

const nameRegistry = {};
const idxRegistry = {};

for (const file of layoutFiles) {
  const content = fs.readFileSync(path.join(layoutsDir, file), 'utf8');
  
  // Extract layout name from p:cSld name attribute
  const nameMatch = content.match(/<p:cSld[^>]*name="([^"]+)"/);
  const layoutName = nameMatch ? nameMatch[1] : file;
  
  nameRegistry[layoutName] = {};
  idxRegistry[layoutName] = {};
  
  // Extract all placeholders (sp elements with nvSpPr containing nvPr with ph)
  const spMatches = content.matchAll(/<p:sp[^>]*>([\s\S]*?)<\/p:sp>/g);
  
  for (const match of spMatches) {
    const sp = match[1];
    
    // Extract placeholder info
    const phMatch = sp.match(/<p:ph([^>]*)\/>/);
    if (!phMatch) continue;
    
    const phAttrs = phMatch[1];
    
    // Get type (default to 'body')
    const typeMatch = phAttrs.match(/type="([^"]+)"/);
    const phType = typeMatch ? typeMatch[1] : 'body';
    
    // Skip pic types (pictures)
    if (phType === 'pic') continue;
    
    // Get idx
    const idxMatch = phAttrs.match(/idx="([^"]+)"/);
    const idx = idxMatch ? parseInt(idxMatch[1]) : undefined;
    
    // Get placeholder name from nvPr/nvSpPr descr or text in p:txBody
    let phName = '';
    const descrMatch = sp.match(/<p:cNvPr[^>]*descr="([^"]+)"/);
    if (descrMatch) {
      phName = descrMatch[1];
    } else {
      // Try to get name from nvPr name attribute
      const nvNameMatch = sp.match(/<p:cNvPr[^>]*name="([^"]+)"/);
      if (nvNameMatch) {
        phName = nvNameMatch[1];
      }
    }
    
    // Normalize common names
    if (!phName || phName === '') {
      if (phType === 'title') phName = 'headline';
      else if (phType === 'subTitle') phName = 'subheadline';  
      else if (phType === 'ftr') phName = 'footer';
      else if (phType === 'sldNum') phName = 'slideNumber';
      else if (phType === 'dt') phName = 'date';
      else phName = 'body' + (idx !== undefined ? idx : '');
    } else {
      // Map common variations
      const nameLower = phName.toLowerCase();
      if (nameLower.includes('headline') && !nameLower.includes('sub')) phName = 'headline';
      else if (nameLower.includes('subhead') || nameLower.includes('subtitle')) phName = 'subheadline';
      else if (nameLower.includes('source')) phName = 'source';
      else if (nameLower.includes('footer')) phName = 'footer';
      else if (nameLower.includes('content')) phName = 'content2';
    }
    
    // Add to name registry
    if (!nameRegistry[layoutName][phName]) {
      nameRegistry[layoutName][phName] = { type: phType };
      if (idx !== undefined) nameRegistry[layoutName][phName].idx = idx;
    }
    
    // Add to idx registry
    if (!idxRegistry[layoutName][phType]) {
      idxRegistry[layoutName][phType] = [];
    }
    if (idx !== undefined && !idxRegistry[layoutName][phType].includes(idx)) {
      idxRegistry[layoutName][phType].push(idx);
    }
  }
}

// Generate output
const output = `// Auto-generated placeholder registry from 2026 Energy.zip template
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

export const PLACEHOLDER_NAME_REGISTRY: PlaceholderNameRegistry = ${JSON.stringify(nameRegistry, null, 2)};

export const PLACEHOLDER_IDX_REGISTRY: PlaceholderIdxRegistry = ${JSON.stringify(idxRegistry, null, 2)};
`;

fs.writeFileSync(outputFile, output);
console.log(`Generated placeholder registry with ${Object.keys(nameRegistry).length} layouts`);
console.log(`Output: ${outputFile}`);
