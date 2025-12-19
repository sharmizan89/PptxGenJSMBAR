import { readFileSync, writeFileSync } from 'fs';

const dataFile = readFileSync(new URL('./step_templates.data.json', import.meta.url), 'utf8');
const data = JSON.parse(dataFile);

function generateLayoutCode(layout) {
  const { name, placeholders } = layout;
  const code = [];
  
  // CRITICAL: First line must specify the masterName to select correct layout
  code.push(`const slide = pptx.addSlide({ masterName: "${name}" });`);
  
  // Filter placeholders that should have code generated
  const contentPlaceholders = placeholders.filter(ph => !ph.inheritFromMaster);
  
  for (const ph of contentPlaceholders) {
    const { name: phName, type, x, y, w, h } = ph;
    
    // Determine the context variable name based on placeholder name and type
    let ctxVar = 'ctx.content';
    let method = 'addText';
    
    // Map placeholder names to context variables
    if (phName.includes('Title') && !phName.includes('Subtitle')) {
      ctxVar = 'ctx.title';
    } else if (phName.includes('Sub-headline')) {
      ctxVar = 'ctx.subHeadline';
    } else if (phName.includes('Subtitle Left') || phName.includes('Subtitle Top Left')) {
      ctxVar = 'ctx.leftSubtitle';
    } else if (phName.includes('Subtitle Right') || phName.includes('Subtitle Top Right')) {
      ctxVar = 'ctx.rightSubtitle';
    } else if (phName.includes('Subtitle Middle')) {
      ctxVar = 'ctx.middleSubtitle';
    } else if (phName.includes('Subtitle Left Middle')) {
      ctxVar = 'ctx.leftMiddleSubtitle';
    } else if (phName.includes('Subtitle Right Middle')) {
      ctxVar = 'ctx.rightMiddleSubtitle';
    } else if (phName.includes('Subtitle Top')) {
      ctxVar = 'ctx.topSubtitle';
    } else if (phName.includes('Subtitle Bottom')) {
      ctxVar = 'ctx.bottomSubtitle';
    } else if (phName.includes('Footnote')) {
      ctxVar = 'ctx.footnote';
    } else if (phName.includes('Content Placeholder 2 Left') || phName.includes('Left Placeholder')) {
      ctxVar = 'ctx.leftContent';
    } else if (phName.includes('Content Placeholder 2 Right') || phName.includes('Right Placeholder')) {
      ctxVar = 'ctx.rightContent';
    } else if (phName.includes('Content Placeholder 3')) {
      ctxVar = 'ctx.mainContent';
    } else if (phName.includes('Content Placeholder 2')) {
      ctxVar = 'ctx.body';
    } else if (phName.includes('Text Placeholder 3')) {
      ctxVar = 'ctx.sidebarContent';
    } else if (phName.includes('Text Placeholder 7 Left Middle')) {
      ctxVar = 'ctx.col2Content';
    } else if (phName.includes('Text Placeholder 7 Left')) {
      ctxVar = 'ctx.col1Content';
    } else if (phName.includes('Text Placeholder 7 Middle')) {
      ctxVar = 'ctx.col3Content';
    } else if (phName.includes('Text Placeholder 7 Right Middle')) {
      ctxVar = 'ctx.col4Content';
    } else if (phName.includes('Text Placeholder 7 Right')) {
      ctxVar = 'ctx.col5Content';
    } else if (phName.includes('Text Placeholder 15')) {
      const idx = contentPlaceholders.filter(p => p.name.includes('Text Placeholder 15')).indexOf(ph) + 1;
      ctxVar = `ctx.text${idx}`;
    } else if (phName.includes('Text Placeholder 20')) {
      ctxVar = 'ctx.text';
    } else if (phName.includes('Text Placeholder 22')) {
      ctxVar = 'ctx.middleText';
    } else if (phName.includes('Text Placeholder 24')) {
      ctxVar = 'ctx.rightText';
    } else if (phName.includes('Text Placeholder 27')) {
      ctxVar = 'ctx.text3';
    } else if (phName.includes('Text Placeholder 29')) {
      ctxVar = 'ctx.text4';
    } else if (phName.includes('Picture Placeholder') || phName.includes('Photo')) {
      method = 'addImage';
      if (phName.includes('Left')) {
        ctxVar = 'ctx.iconLeft';
      } else if (phName.includes('Middle') && phName.includes('Top')) {
        ctxVar = 'ctx.icon2';
      } else if (phName.includes('Middle') && phName.includes('Bottom')) {
        ctxVar = 'ctx.icon3';
      } else if (phName.includes('Middle')) {
        ctxVar = 'ctx.iconMiddle';
      } else if (phName.includes('Right')) {
        ctxVar = 'ctx.iconRight';
      } else if (phName.includes('Top')) {
        ctxVar = 'ctx.icon1';
      } else if (phName.includes('Bottom')) {
        ctxVar = 'ctx.icon4';
      } else if (phName.includes('6')) {
        ctxVar = 'ctx.imageUrl';
      } else if (phName.includes('7')) {
        ctxVar = 'ctx.photoUrl';
      } else {
        ctxVar = 'ctx.imageUrl';
      }
    } else if (phName.includes('Chart') || phName.includes('Table')) {
      continue; // Charts handled separately
    } else if (phName.includes('Statement')) {
      ctxVar = 'ctx.statement';
    } else if (phName.includes('Section')) {
      ctxVar = 'ctx.sectionTitle';
    } else if (phName.includes('Divider')) {
      ctxVar = 'ctx.title';
    } else if (phName.includes('Agenda')) {
      ctxVar = 'ctx.agendaItems';
    } else if (phName.includes('TOC')) {
      ctxVar = 'ctx.tocItems';
    } else if (phName.includes('Author')) {
      ctxVar = 'ctx.authorName';
    } else if (phName.includes('Contact')) {
      ctxVar = 'ctx.contact1Name';
    }
    
    // Generate code line with coordinates
    if (x !== undefined && y !== undefined && w !== undefined && h !== undefined) {
      // Round all coordinates to 2 decimal places
      const xRound = Math.round(x * 100) / 100;
      const yRound = Math.round(y * 100) / 100;
      const wRound = Math.round(w * 100) / 100;
      const hRound = Math.round(h * 100) / 100;
      
      const coords = `{ x: ${xRound}, y: ${yRound}, w: ${wRound}, h: ${hRound} }`;
      
      if (method === 'addImage') {
        // Icons use data:, photos use path:
        const isIcon = name.includes('Icon') && !name.includes('Image');
        const prop = isIcon ? 'data' : 'path';
        code.push(`if (${ctxVar}) slide.${method}({ ${prop}: ${ctxVar}, x: ${xRound}, y: ${yRound}, w: ${wRound}, h: ${hRound} })`);
      } else {
        code.push(`slide.${method}(${ctxVar}, ${coords})`);
      }
    }
  }
  
  // Add custom code for special layouts
  if (name === 'Blank') {
    code.push('// Add custom content using slide.addText(), slide.addImage(), etc.');
    code.push('// Example: slide.addText("Custom text", { x: 1, y: 2, w: 5, h: 1 })');
  }
  
  if (name.includes('Chart')) {
    const chartPh = placeholders.find(p => p.name.includes('Chart') || p.name.includes('Content Placeholder 10'));
    if (chartPh && chartPh.x !== undefined) {
      // Round chart coordinates to 2 decimal places
      const xRound = Math.round(chartPh.x * 100) / 100;
      const yRound = Math.round(chartPh.y * 100) / 100;
      const wRound = Math.round(chartPh.w * 100) / 100;
      const hRound = Math.round(chartPh.h * 100) / 100;
      code.push(`if (ctx.chartData) slide.addChart(ctx.chartType, ctx.chartData, { x: ${xRound}, y: ${yRound}, w: ${wRound}, h: ${hRound} })`);
    }
  }
  
  return code.length > 0 ? code : ['// No content placeholders for this layout'];
}

function generateTemplate(layout) {
  const { name } = layout;
  
  // Generate human-readable template descriptions
  const templates = {
    'Content - no subtitle': 'Standard content slide with title, body text, and footnote',
    'Content w 2 Line Title and Sub-headline': 'Content slide with two-line title, sub-headline, body, and footnote',
    'Two Content': 'Side-by-side comparison with title, sub-headline, left/right content',
    'Two Content + Subtitles': 'Side-by-side with individual subtitles for each column',
    'Content 4 Columns': 'Four-column layout with individual subtitles',
    'Content 5 Columns': 'Five-column layout for detailed frameworks',
    'Content with Sidebar': 'Main content area with sidebar for key notes',
    'Title Only': 'Slide with only title and optional sub-headline',
    'Blank': 'Completely blank slide for custom layouts',
    'Content + Image/Icon': 'Content with accompanying image or icon',
    'Content + Photo White': 'Content with full-height photo on white background',
    'Content + Photo Black': 'Content with full-height photo on black background',
    'Content + Photo Blue': 'Content with full-height photo on blue background',
    'Icons 3 Columns Vertical': 'Three columns with icons, subtitles, and text arranged vertically',
    'Icons 3 Columns Horizontal': 'Three columns with icons, subtitles, and text arranged horizontally',
    'Icons 4 Columns + Content': 'Four icon columns with additional content area',
    'Icons 4 Columns + Content Black': 'Four icon columns with content area on black background',
    'Icons 4 Columns + Content Blue': 'Four icon columns with content area on blue background',
    'Icons 2 x 3 Columns': 'Six icons arranged in 2 rows of 3 columns',
    'Content + Chart/Table 1': 'Content area with chart or table',
    'Chart - Horizontal 2': 'Full-width horizontal chart layout',
    'Chart + Statement 2': 'Chart with key statement or insight',
    'Chart + Statement 3': 'Chart with extended statement area',
    'Statement Photo': 'Large statement with background photo',
    'Statement Black': 'Large statement on black background',
    'Statement White': 'Large statement on white background',
    'Section Header': 'Section divider with title on dark background',
    'Divider 4 Photo': 'Section divider with 4 photos',
    'Divider 1': 'Simple section divider with title',
    'Divider 2': 'Section divider with blue accent',
    'Divider Photo 2': 'Section divider with 2 photos',
    'Two Placeholders': 'Two equal content areas side by side',
    'Three Placeholders 1': 'Three equal content areas',
    'Three Placeholders 2': 'Three content areas with alternate sizing',
    'Three Placeholders 3': 'Three content areas with alternate arrangement',
    'Four Placeholders': 'Four equal content areas',
    'Single Author': 'Author bio with photo and details',
    '2 Authors': 'Two author bios side by side',
    '3 Authors': 'Three author bios',
    '4 Authors': 'Four author bios',
    'Agenda - presentations': 'Presentation agenda with numbered items',
    'TOC - reports': 'Table of contents for reports',
    'Title White - reports and presentations (hIHS)': 'White title slide with subtitle and right banner image',
    'Title Image Bottom': 'Title slide with image at bottom on black background',
    'Energy': 'Energy division title slide',
    'Companies & Transactions': 'Companies & Transactions title slide',
    'Contact us': 'Contact information slide',
    'Content w/Sub-headline': 'Content slide with title, sub-headline, body, and footnote'
  };
  
  return templates[name] || `Layout: ${name}`;
}

function generateInstructions(layout) {
  const { name } = layout;
  
  const instructions = {
    'Content - no subtitle': 'Use for main content slides. Keep title concise (1-2 lines). Use 3-6 bullet points maximum. Add source citation in footnote.',
    'Content w 2 Line Title and Sub-headline': 'Use when you need both a title and explanatory sub-headline. Title can span 2 lines. Sub-headline should be 1 sentence. Body: 3-5 bullets.',
    'Two Content': 'Use for comparisons or parallel content. Keep each column to 3-5 bullets. Use parallel structure between columns.',
    'Two Content + Subtitles': 'Use when each column needs its own subtitle (2-4 words). Keep content brief in each column (2-4 bullets).',
    'Content 4 Columns': 'Use for frameworks, processes, or categorical breakdowns. Each column: subtitle (2-3 words) + 2-4 bullets.',
    'Content 5 Columns': 'Use for complex processes or 5-part frameworks. Each column: subtitle (1-2 words) + 2-3 bullets. Keep text concise.',
    'Content with Sidebar': 'Use sidebar for definitions, key metrics, or callouts. Main content: 3-6 bullets. Sidebar: 1-3 short points.',
    'Title Only': 'Use for section breaks or when content will be added manually. Title can be 1-2 lines. Optional sub-headline for context.',
    'Blank': 'Use when you need complete design freedom. Add shapes, images, or text manually using x/y coordinates.',
    'Content + Image/Icon': 'Use for concept explanation with visual. Content: 3-5 bullets. Provide image description for selection.',
    'Content + Photo White': 'Use for visual storytelling. Photo on right, content on left. Content: 3-5 bullets. Provide photo description.',
    'Content + Photo Black': 'Use for dramatic visual impact. Same as white variant but with dark theme. Provide photo description.',
    'Content + Photo Blue': 'Use for brand-aligned visual storytelling. Same layout as white/black variants. Provide photo description.',
    'Icons 3 Columns Vertical': 'Use for 3-part concepts. Each column: icon + subtitle (2-3 words) + description (1-3 lines). Request 3 icons.',
    'Icons 3 Columns Horizontal': 'Similar to vertical but different icon sizing. Use for 3-part concepts with shorter descriptions. Request 3 icons.',
    'Icons 4 Columns + Content': 'Use for frameworks with explanation. 4 icons with subtitles + main content area. Request 4 icons. Keep text brief.',
    'Icons 4 Columns + Content Black': 'Black variant of 4-column icon layout. Use for dramatic frameworks. Request 4 icons.',
    'Icons 4 Columns + Content Blue': 'Blue variant of 4-column icon layout. Use for brand-aligned frameworks. Request 4 icons.',
    'Icons 2 x 3 Columns': 'Use for displaying 6 related concepts or categories. Each icon with brief label. Request 6 icons.',
    'Content + Chart/Table 1': 'Use for data visualization with explanatory text. Content: 2-4 bullets. Provide chart data or table structure.',
    'Chart - Horizontal 2': 'Use for prominent data display. Provide complete chart data with labels and values.',
    'Chart + Statement 2': 'Use to highlight a key data point. Statement: 1-2 sentences. Provide chart data.',
    'Chart + Statement 3': 'Similar to Chart + Statement 2 but with more room for statement. Use for complex insights. Provide chart data.',
    'Statement Photo': 'Use for impactful quotes or key messages. Statement: 1-3 sentences. Provide photo description.',
    'Statement Black': 'Use for dramatic emphasis. Statement: 1-3 sentences maximum.',
    'Statement White': 'Use for clean, professional emphasis. Statement: 1-3 sentences maximum.',
    'Section Header': 'Use to introduce new presentation sections. Title: 3-7 words. Optional subtitle for context.',
    'Divider 4 Photo': 'Visual section break with imagery. Provide 4 photo descriptions. Optional title text.',
    'Divider 1': 'Clean section break. Title: 3-7 words.',
    'Divider 2': 'Brand-aligned section break. Title: 3-7 words. Optional subtitle.',
    'Divider Photo 2': 'Visual section break with paired imagery. Provide 2 photo descriptions.',
    'Two Placeholders': 'Use for balanced comparisons. Keep each area to 3-5 bullets.',
    'Three Placeholders 1': 'Use for 3-part structures. Keep each area to 2-4 bullets.',
    'Three Placeholders 2': 'Similar to Three Placeholders 1 with different proportions. Use for emphasis on specific areas.',
    'Three Placeholders 3': 'Third variant of 3-placeholder layout. Use based on content hierarchy.',
    'Four Placeholders': 'Use for 4-part structures. Keep each area to 2-3 bullets.',
    'Single Author': 'Use for speaker introduction. Include name, title, contact, bio (2-3 sentences). Provide headshot photo.',
    '2 Authors': 'Use for co-presenters. Each author: name, title, contact, bio (1-2 sentences). Provide 2 headshot photos.',
    '3 Authors': 'Use for panel or team. Each author: name, title, contact. Brief bios (1 sentence each). Provide 3 headshot photos.',
    '4 Authors': 'Use for larger panels. Each author: name, title, contact. Keep bios very brief. Provide 4 headshot photos.',
    'Agenda - presentations': 'Use for agenda slide. List 3-7 agenda items. Keep each item to 3-5 words.',
    'TOC - reports': 'Use for report TOC. List chapter/section titles with page numbers.',
    'Title White - reports and presentations (hIHS)': 'Use as first slide. Title: 5-10 words. Subtitle: company/event/date. Provide image description for right banner.',
    'Title Image Bottom': 'Use as dramatic opening slide. Title: 5-10 words. Subtitle optional. Provide image description.',
    'Energy': 'Energy-branded title slide. Title: 5-10 words. Subtitle: event/date info.',
    'Companies & Transactions': 'Title slide for M&A/transaction content. Title: 5-10 words. Subtitle: deal/date info.',
    'Contact us': 'Pre-formatted contact slide. Modify contact names and emails as needed.',
    'Content w/Sub-headline': 'Use when you need a sub-headline for context. Title: 1-2 lines. Sub-headline: 1 sentence. Body: 3-6 bullets.'
  };
  
  return instructions[name] || `Use the ${name} layout for appropriate content.`;
}

// Generate the complete template file
const layouts = data.layouts.map(layout => ({
  name: layout.name,
  masterName: layout.name,  // Required for pptx.addSlide({ masterName: "..." })
  template: generateTemplate(layout),
  instructions: generateInstructions(layout),
  placeholders: layout.placeholders.filter(p => !p.inheritFromMaster).map(p => p.name),
  code: generateLayoutCode(layout)
}));

const output = {
  units: 'inches',
  note: 'LLM-ready step templates for all 56 slide layouts. Each layout entry provides: template (human-readable), instructions (concise guidance), and code (PptxGenJS snippets with x,y,w,h coordinates). All positioning uses explicit coordinates from step_templates.data.json.',
  layouts
};

writeFileSync(
  new URL('./step_templates.llm.json', import.meta.url),
  JSON.stringify(output, null, 2)
);

console.log(`✅ Generated step_templates.llm.json with ${layouts.length} layouts`);
console.log(`✅ All code uses x, y, w, h coordinates (no placeholder targeting)`);
console.log(`✅ Icons use data: property, images use path: property`);
