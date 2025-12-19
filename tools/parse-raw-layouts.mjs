#!/usr/bin/env node
/**
 * Parse raw XML layout and rels files (without delimiters)
 * Strip image references and generate TypeScript registry files
 */

import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';

// File paths
const LAYOUTS_TXT = resolve('src/cust-xml-slide-layouts.txt');
const RELS_TXT = resolve('src/cust-xml-slide-layouts.rels.txt');
const LAYOUTS_TS = resolve('src/cust-xml-slide-layouts.ts');
const RELS_TS = resolve('src/cust-xml-slide-layout-rels.ts');

/**
 * Split file content by XML declarations
 */
function splitXmlBlocks(content) {
	const blocks = [];
	// Split by XML declarations and filter empty strings
	const parts = content.split(/(?=<\?xml)/);
	
	for (const part of parts) {
		const trimmed = part.trim();
		if (trimmed && trimmed.startsWith('<?xml')) {
			blocks.push(trimmed);
		}
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
 * Strip image references from layout XML
 * Converts picture placeholders to body placeholders and removes image fills
 */
function stripImagesFromLayout(xml) {
	let cleaned = xml;

	// Convert picture placeholders to body placeholders to preserve layout structure
	cleaned = cleaned.replace(/<p:ph([^>]*?)type="pic"([^>]*)\/>/g, '<p:ph$1type="body"$2/>' );

	// Rename placeholder names from "Picture Placeholder" to "Content Placeholder" to avoid misleading labels
	cleaned = cleaned.replace(/<p:cNvPr([^>]*?)name="Picture Placeholder ([^"]+)"/g, '<p:cNvPr$1name="Content Placeholder $2"');

	// Remove standalone <p:pic> elements entirely (rare in layouts)
	cleaned = cleaned.replace(/<p:pic>.*?<\/p:pic>/gs, '');

	// Remove any remaining image blips and fills within shapes
	cleaned = cleaned.replace(/<a:blip[^>]*\/>/g, '');
	cleaned = cleaned.replace(/<a:blip[^>]*>.*?<\/a:blip>/gs, '');
	cleaned = cleaned.replace(/<a:blipFill>.*?<\/a:blipFill>/gs, '');

	// Remove common placeholder instructional text for photos
	cleaned = cleaned.replace(/\[Insert photo\]/g, '');
	cleaned = cleaned.replace(/\[Click to insert photo[^\]]*\]/g, '');

	// Remove generic instructional text that implies image insertion
	cleaned = cleaned.replace(/\[Image\] or remove and place icon here/g, '');

	// Strip risky extension lists and attributes that some PowerPoint builds "repair"
	// Remove adec:decorative ext and any extLst entries we don't need in layouts
	cleaned = cleaned.replace(/<a:ext\s+uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}"[^>]*>.*?<\/a:ext>/gs, '');
	// Remove a16:creationId entries
	cleaned = cleaned.replace(/<a:ext\s+uri="\{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236\}"[^>]*>.*?<\/a:ext>/gs, '');
	// Remove entire a:extLst if it becomes empty
	cleaned = cleaned.replace(/<a:extLst>\s*<\/a:extLst>/g, '');

	// Remove p:extLst blocks like p14:creationId and p15:sldGuideLst
	cleaned = cleaned.replace(/<p:extLst>.*?<\/p:extLst>/gs, '');

    // Remove attributes that can trigger repair in some builds
    cleaned = cleaned.replace(/\s+hasCustomPrompt="[01]"/g, '');
    cleaned = cleaned.replace(/\s+userDrawn="[01]"/g, '');

    // Remove p:bgRef idx attribute anomalies (normalize by dropping idx in layouts)
    cleaned = cleaned.replace(/<p:bgRef[^>]*>/g, '<p:bgRef>');	// Make copyright year dynamic - replace "2025" with current year variants
	const currentYear = new Date().getFullYear();
	cleaned = cleaned.replace(/© 2025 S&amp;P Global\.?/g, `© ${currentYear} S&amp;P Global.`);
	cleaned = cleaned.replace(/Copyright © 2025 S&amp;P Global/g, `Copyright © ${currentYear} S&amp;P Global`);
	cleaned = cleaned.replace(/© 2025 by S&amp;P Global/g, `© ${currentYear} by S&amp;P Global`);

	// Ensure required attributes on p:sldLayout root for PowerPoint compatibility
	// Note: NOT adding type="cust" as it may cause issues in some PowerPoint builds
	
	// Add preserve="1" back with proper context - required by some builds
	if (!cleaned.includes('preserve=')) {
		cleaned = cleaned.replace(
			/<p:sldLayout([^>]*?)>/,
			'<p:sldLayout$1 preserve="1">'
		);
	}

	// Per OOXML schema (ECMA-376), elements must appear in this order:
	// p:cSld, p:clrMapOvr, p:transition (opt), p:timing (opt), p:hf (opt), p:extLst (opt)
	// Insert timing and hf BEFORE clrMapOvr
	
	if (!cleaned.includes('<p:timing>') && cleaned.includes('<p:clrMapOvr>')) {
		cleaned = cleaned.replace(
			'<p:clrMapOvr>',
			'<p:timing><p:tnLst/></p:timing><p:clrMapOvr>'
		);
	}

	if (!cleaned.includes('<p:hf') && cleaned.includes('<p:clrMapOvr>')) {
		cleaned = cleaned.replace(
			'<p:clrMapOvr>',
			'<p:hf sldNum="0" ftr="0" dt="0"/><p:clrMapOvr>'
		);
	}

	return cleaned;
}

/**
 * Strip image relationships from rels XML
 * Keeps only slideMaster relationship
 */
function stripImagesFromRels(xml) {
	// Parse and rebuild relationships, keeping only non-image ones
	let cleaned = xml;
	
	// Remove individual image relationship elements
	cleaned = cleaned.replace(/<Relationship[^>]+Type="[^"]*\/image"[^>]*\/>/g, '');
	
	// Clean up any extra whitespace
	cleaned = cleaned.replace(/>\s+</g, '><');
	
	return cleaned;
}

/**
 * Validate layout data
 */
function validateLayouts(layouts) {
	const names = new Set();
	const errors = [];
	
	layouts.forEach((layout, idx) => {
		// Check for name
		if (!layout.name || layout.name.trim() === '') {
			errors.push(`Layout ${idx + 1}: Missing or empty name`);
		}
		
		// Check for duplicates
		if (names.has(layout.name)) {
			errors.push(`Layout ${idx + 1}: Duplicate name "${layout.name}"`);
		}
		names.add(layout.name);
		
		// Check for required XML elements
		if (!layout.xml.includes('<p:sldLayout')) {
			errors.push(`Layout ${idx + 1} (${layout.name}): Missing <p:sldLayout> element`);
		}
	});
	
	return errors;
}

/**
 * Validate rels data
 */
function validateRels(rels) {
	const errors = [];
	
	rels.forEach((rel, idx) => {
		// Check for slideMaster relationship
		if (!rel.relsXml.includes('slideMaster')) {
			errors.push(`Rels ${idx + 1}: Missing slideMaster relationship`);
		}
		
		// Check for image references (should be removed)
		if (rel.relsXml.includes('/image')) {
			errors.push(`Rels ${idx + 1}: Still contains image references!`);
		}
	});
	
	return errors;
}

/**
 * Generate TypeScript file content
 */
function generateLayoutsTS(layouts) {
	const imports = `/**
 * Custom Slide Layout Definitions
 * Auto-generated by parse-raw-layouts.mjs
 * Contains ${layouts.length} custom slide layouts
 */

export interface CustomSlideLayoutDef {
	id: number;
	name: string;
	xml: string;
}
`;

	const arrayStart = '\nexport const CUSTOM_SLIDE_LAYOUT_DEFS: CustomSlideLayoutDef[] = [\n';
	
	const items = layouts.map((layout, idx) => {
		const escapedXml = layout.xml.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$/g, '\\$');
		return `\t{
\t\tid: ${layout.id},
\t\tname: ${JSON.stringify(layout.name)},
\t\txml: \`${escapedXml}\`,
\t}`;
	});
	
	const arrayEnd = '\n];\n';
	
	return imports + arrayStart + items.join(',\n') + arrayEnd;
}

/**
 * Generate TypeScript rels file content
 */
function generateRelsTS(rels) {
	const imports = `/**
 * Custom Slide Layout Relationship Definitions
 * Auto-generated by parse-raw-layouts.mjs
 * Contains ${rels.length} relationship definitions
 */

export interface CustomSlideLayoutRelDef {
	id: number;
	relsXml: string;
}
`;

	const arrayStart = '\nexport const CUSTOM_SLIDE_LAYOUT_RELS: CustomSlideLayoutRelDef[] = [\n';
	
	const items = rels.map((rel) => {
		const escapedXml = rel.relsXml.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$/g, '\\$');
		return `\t{
\t\tid: ${rel.id},
\t\trelsXml: \`${escapedXml}\`,
\t}`;
	});
	
	const arrayEnd = '\n];\n';
	
	return imports + arrayStart + items.join(',\n') + arrayEnd;
}

/**
 * Main processing function
 */
function main() {
	console.log('🔍 Reading input files...');
	const layoutsContent = readFileSync(LAYOUTS_TXT, 'utf-8');
	const relsContent = readFileSync(RELS_TXT, 'utf-8');
	
	console.log('📦 Splitting XML blocks...');
	const layoutBlocks = splitXmlBlocks(layoutsContent);
	const relsBlocks = splitXmlBlocks(relsContent);
	
	console.log(`   Found ${layoutBlocks.length} layout blocks`);
	console.log(`   Found ${relsBlocks.length} rels blocks`);
	
	if (layoutBlocks.length !== relsBlocks.length) {
		console.error(`❌ Error: Mismatch between layouts (${layoutBlocks.length}) and rels (${relsBlocks.length})`);
		process.exit(1);
	}
	
	console.log('🧹 Stripping images from layouts...');
	const layouts = layoutBlocks.map((xml, idx) => {
		const name = extractLayoutName(xml);
		const cleanedXml = stripImagesFromLayout(xml);
		return {
			id: idx + 1,
			name: name || `Layout ${idx + 1}`,
			xml: cleanedXml
		};
	});
	
	console.log('🧹 Stripping images from rels...');
	const rels = relsBlocks.map((xml, idx) => {
		const cleanedXml = stripImagesFromRels(xml);
		return {
			id: idx + 1,
			relsXml: cleanedXml
		};
	});
	
	console.log('✅ Validating layouts...');
	const layoutErrors = validateLayouts(layouts);
	if (layoutErrors.length > 0) {
		console.error('❌ Layout validation errors:');
		layoutErrors.forEach(err => console.error(`   - ${err}`));
		process.exit(1);
	}
	
	console.log('✅ Validating rels...');
	const relsErrors = validateRels(rels);
	if (relsErrors.length > 0) {
		console.error('❌ Rels validation errors:');
		relsErrors.forEach(err => console.error(`   - ${err}`));
		process.exit(1);
	}
	
	console.log('📝 Generating TypeScript files...');
	const layoutsTS = generateLayoutsTS(layouts);
	const relsTS = generateRelsTS(rels);
	
	console.log('💾 Writing output files...');
	writeFileSync(LAYOUTS_TS, layoutsTS, 'utf-8');
	writeFileSync(RELS_TS, relsTS, 'utf-8');
	
	console.log('✨ Success!');
	console.log(`   Generated ${layouts.length} slide layouts`);
	console.log(`   Layout names:`);
	layouts.forEach(l => console.log(`      ${l.id}. ${l.name}`));
	console.log(`   Files written:`);
	console.log(`      ${LAYOUTS_TS}`);
	console.log(`      ${RELS_TS}`);
}

// Run if called directly
if (import.meta.url === `file://${process.argv[1]}`) {
	try {
		main();
	} catch (error) {
		console.error('❌ Fatal error:', error.message);
		console.error(error.stack);
		process.exit(1);
	}
}
