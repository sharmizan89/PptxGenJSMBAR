#!/usr/bin/env node

/**
 * MINIMAL Parser - Only remove image blips, keep everything else
 */

import { readFileSync, writeFileSync } from 'fs';

const LAYOUTS_INPUT = 'src/cust-xml-slide-layouts.txt';
const RELS_INPUT = 'src/cust-xml-slide-layouts.rels.txt';
const LAYOUTS_OUTPUT = 'src/cust-xml-slide-layouts.ts';
const RELS_OUTPUT = 'src/cust-xml-slide-layout-rels.ts';

/**
 * Split XML blocks by XML declaration
 */
function splitXmlBlocks(content) {
	const xmlPattern = /<\?xml[^?]*\?>/g;
	const blocks = [];
	let lastIndex = 0;
	let match;
	
	while ((match = xmlPattern.exec(content)) !== null) {
		if (lastIndex > 0) {
			blocks.push(content.substring(lastIndex, match.index));
		}
		lastIndex = match.index;
	}
	
	if (lastIndex > 0) {
		blocks.push(content.substring(lastIndex));
	}
	
	return blocks;
}

/**
 * Extract layout name from XML
 */
function extractLayoutName(xml) {
	const match = xml.match(/<p:cSld\s+name="([^"]+)"/);
	return match ? match[1] : null;
}

/**
 * MINIMAL: Only remove image blips, keep ALL other attributes and structures
 */
function stripImagesFromLayout(xml) {
	let cleaned = xml;

	// ONLY remove actual image binary references
	cleaned = cleaned.replace(/<a:blip[^>]*\/>/g, '');
	cleaned = cleaned.replace(/<a:blip[^>]*>.*?<\/a:blip>/gs, '');
	cleaned = cleaned.replace(/<a:blipFill>.*?<\/a:blipFill>/gs, '');

	// Make copyright year dynamic
	const currentYear = new Date().getFullYear();
	cleaned = cleaned.replace(/© 2025 S&amp;P Global\.?/g, `© ${currentYear} S&amp;P Global.`);
	cleaned = cleaned.replace(/Copyright © 2025 S&amp;P Global/g, `Copyright © ${currentYear} S&amp;P Global`);
	cleaned = cleaned.replace(/© 2025 by S&amp;P Global/g, `© ${currentYear} by S&amp;P Global`);

	return cleaned;
}

/**
 * Strip image relationships - keep only slideMaster
 */
function stripImagesFromRels(xml) {
	let cleaned = xml;
	// Remove all image relationships
	cleaned = cleaned.replace(/<Relationship[^>]*Type="[^"]*\/image"[^>]*\/>/g, '');
	return cleaned;
}

/**
 * Generate TypeScript file for layouts
 */
function generateLayoutsTS(layouts) {
	const entries = layouts.map((layout, idx) => {
		const escapedXml = layout.xml
			.replace(/\\/g, '\\\\')
			.replace(/`/g, '\\`')
			.replace(/\$/g, '\\$');
		
		return `  { id: ${idx + 1}, name: '${layout.name}', xml: \`${escapedXml}\` }`;
	});

	return `// AUTO-GENERATED - Do not edit manually
export const CUSTOM_SLIDE_LAYOUT_DEFS = [
${entries.join(',\n')}
];
`;
}

/**
 * Generate TypeScript file for rels
 */
function generateRelsTS(rels) {
	const entries = rels.map((rel, idx) => {
		const escapedXml = rel.relsXml
			.replace(/\\/g, '\\\\')
			.replace(/`/g, '\\`')
			.replace(/\$/g, '\\$');
		
		return `  { id: ${idx + 1}, relsXml: \`${escapedXml}\` }`;
	});

	return `// AUTO-GENERATED - Do not edit manually
export const CUSTOM_SLIDE_LAYOUT_RELS = [
${entries.join(',\n')}
];
`;
}

/**
 * Main execution
 */
async function main() {
	console.log('🧪 MINIMAL Parser - Testing with minimal modifications\n');
	console.log('================================================\n');
	
	console.log('🔍 Reading input files...');
	const layoutsContent = readFileSync(LAYOUTS_INPUT, 'utf-8');
	const relsContent = readFileSync(RELS_INPUT, 'utf-8');

	console.log('📦 Splitting XML blocks...');
	const layoutBlocks = splitXmlBlocks(layoutsContent);
	const relsBlocks = splitXmlBlocks(relsContent);
	console.log(`   Found ${layoutBlocks.length} layout blocks`);
	console.log(`   Found ${relsBlocks.length} rels blocks`);

	console.log('🧹 Processing layouts (minimal changes)...');
	const layouts = [];
	for (const block of layoutBlocks) {
		if (!block.trim() || block.trim().length < 100) continue;
		
		const name = extractLayoutName(block);
		if (!name) continue;
		
		const xml = stripImagesFromLayout(block);
		layouts.push({ name, xml });
	}

	console.log('🧹 Processing rels (remove image refs only)...');
	const rels = [];
	for (const block of relsBlocks) {
		if (!block.trim() || block.trim().length < 50) continue;
		const relsXml = stripImagesFromRels(block);
		rels.push({ relsXml });
	}

	console.log('📝 Generating TypeScript files...');
	const layoutsTS = generateLayoutsTS(layouts);
	const relsTS = generateRelsTS(rels);

	console.log('💾 Writing output files...');
	writeFileSync(LAYOUTS_OUTPUT, layoutsTS, 'utf-8');
	writeFileSync(RELS_OUTPUT, relsTS, 'utf-8');

	console.log('\n✨ MINIMAL parsing complete!');
	console.log(`   Generated ${layouts.length} layouts with MINIMAL modifications`);
	console.log(`   - Removed: image blips only`);
	console.log(`   - Kept: ALL attributes, extLst, picture placeholders, etc.`);
	console.log(`\n📋 Files written:`);
	console.log(`      ${LAYOUTS_OUTPUT}`);
	console.log(`      ${RELS_OUTPUT}`);
}

main().catch(console.error);
