/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */

/**
 * Units of Measure used in PowerPoint documents
 *
 * PowerPoint units are in `DXA` (except for font sizing)
 * - 1 inch is 1440 DXA
 * - 1 inch is 72 points
 * -  1 DXA is 1/20th's of a point
 * - 20 DXA is 1 point
 *
 * Another form of measurement using is an `EMU`
 * - 914400 EMUs is 1 inch
 * -  12700 EMUs is 1 point
 *
 * @see https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 */

/**
 * Object Layouts
 *
 * - 16x9 (10" x 5.625")
 * - 16x10 (10" x 6.25")
 * - 4x3 (10" x 7.5")
 * - Wide (13.33" x 7.5")
 * - [custom] (any size)
 *
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)
 */

import JSZip from 'jszip'
import Slide from './slide'
import {
	AlignH,
	AlignV,
	CHART_TYPE,
	ChartType,
	DEF_PRES_LAYOUT,
	DEF_PRES_LAYOUT_NAME,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	OutputType,
	SCHEME_COLOR_NAMES,
	SHAPE_TYPE,
	SchemeColor,
	ShapeType,
	WRITE_OUTPUT_TYPE,
} from './core-enums'
import {
	AddSlideProps,
	IPresentationProps,
	PresLayout,
	PresSlide,
	SectionProps,
	SlideLayout,
	SlideMasterProps,
	SlideNumberProps,
	TableToSlidesProps,
	ThemeProps,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './core-interfaces'
import {
	CUSTOM_CONTENT_TYPES_XML,
	CUSTOM_RELS_XML,
	CUSTOM_PROPS_APP_XML,
	CUSTOM_PROPS_CORE_XML,
	CUSTOM_PROPS_CUSTOM_XML,
	CUSTOM_PPT_PRESENTATION_XML,
	CUSTOM_PPT_PRESENTATION_RELS_XML,
	CUSTOM_PPT_THEME1_XML,
	CUSTOM_PPT_SLIDE_MASTER_XML,
	CUSTOM_PPT_SLIDE_MASTER_REL_XML,
	CUSTOMXML_ITEM1,
	CUSTOMXML_ITEM2,
	CUSTOMXML_ITEM3,
	CUSTOMXML_ITEMPROPS1,
	CUSTOMXML_ITEMPROPS2,
	CUSTOMXML_ITEMPROPS3,
	CUSTOMXML_ITEM1_RELS,
	CUSTOMXML_ITEM2_RELS,
	CUSTOMXML_ITEM3_RELS,
} from './cust-xml'
import { CUSTOM_SLIDE_LAYOUT_DEFS } from './cust-xml-slide-layouts'
import { CUSTOM_SLIDE_LAYOUT_RELS } from './cust-xml-slide-layout-rels'
import * as genCharts from './gen-charts'
import * as genObj from './gen-objects'
import * as genMedia from './gen-media'
import * as genTable from './gen-tables'
import * as genXml from './gen-xml'

const VERSION = '4.0.1'

export default class PptxGenJS implements IPresentationProps {
	// Property getters/setters

	/**
	 * Remove only broken images from layout/notes XML:
	 * - Strips <p:pic> blocks that either:
	 *   a) contain one or more <a:blip r:embed="rIdX"> where rIdX is missing in rels, or
	 *   b) contain no <a:blip> at all (an empty picture placeholder that renders "This picture can't be displayed").
	 * - Strips <p:sp> blocks only when they contain blip embeds with missing rels.
	 * - Preserves true placeholders (e.g., <p:ph type="pic"/>) in all cases.
	 */
	private removeBrokenPics(xml: string, relsXml?: string): string {
		if (!xml || typeof xml !== 'string') return xml
		// Collect known relationship Ids from relsXml (if provided)
		const knownIds = new Set<string>()
		if (relsXml && typeof relsXml === 'string') {
			const relRe = /<Relationship[^>]*\bId="([^"]+)"[^>]*>/g
			let m: RegExpExecArray
			while ((m = relRe.exec(relsXml)) !== null) {
				knownIds.add(m[1])
			}
		}

		let out = xml

		// Helper to decide if a block has an embedded blip without a valid rel
		const hasBrokenBlip = (block: string): boolean => {
			const embeds = [...block.matchAll(/\br:embed="([^"]+)"/g)].map(mm => mm[1])
			if (embeds.length === 0) return false
			// If no relsXml is provided, any embed is considered broken in these layout/notes contexts
			if (knownIds.size === 0) return true
			return embeds.some(id => !knownIds.has(id))
		}

		// 1) Remove <p:pic> blocks with broken blips OR with no <a:blip> at all
		out = out.replace(/<p:pic\b[\s\S]*?<\/p:pic>/g, (seg: string) => {
			const hasAnyBlip = /<a:blip\b/i.test(seg)
			if (!hasAnyBlip) return ''
			return hasBrokenBlip(seg) ? '' : seg
		})

		// 2) Remove <p:sp> blocks that include blip fills with broken embeds (but keep pure placeholders)
		out = out.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (seg: string) => {
			// Keep if it's a placeholder-only shape, or has no blip embeds
			if (seg.includes('<p:ph')) return seg
			return hasBrokenBlip(seg) ? '' : seg
		})

		return out
	}


	/**
	 * Presentation layout name
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10"    x 7.5")
	 * - 'LAYOUT_16x9'  (10"    x 5.625")
	 * - 'LAYOUT_16x10' (10"    x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 * Custom layouts:
	 * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	private _layout: string
	public set layout(value: string) {
		const newLayout: PresLayout = this.LAYOUTS[value]

		if (newLayout) {
			this._layout = value
			this._presLayout = newLayout
		} else {
			throw new Error('UNKNOWN-LAYOUT')
		}
	}

	public get layout(): string {
		return this._layout
	}

	/**
	 * PptxGenJS Library Version
	 */
	private readonly _version: string = VERSION
	public get version(): string {
		return this._version
	}

	/**
	 * @type {string}
	 */
	private _author: string
	public set author(value: string) {
		this._author = value
	}

	public get author(): string {
		return this._author
	}

	/**
	 * @type {string}
	 */
	private _company: string
	public set company(value: string) {
		this._company = value
	}

	public get company(): string {
		return this._company
	}

	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	private _revision: string
	public set revision(value: string) {
		this._revision = value
	}

	public get revision(): string {
		return this._revision
	}

	/**
	 * @type {string}
	 */
	private _subject: string
	public set subject(value: string) {
		this._subject = value
	}

	public get subject(): string {
		return this._subject
	}

	/**
	 * @type {ThemeProps}
	 */
	private _theme: ThemeProps
	public set theme(value: ThemeProps) {
		this._theme = value
	}

	public get theme(): ThemeProps {
		return this._theme
	}

	/**
	 * @type {string}
	 */
	private _title: string
	public set title(value: string) {
		this._title = value
	}

	public get title(): string {
		return this._title
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	private _rtlMode: boolean
	public set rtlMode(value: boolean) {
		this._rtlMode = value
	}

	public get rtlMode(): boolean {
		return this._rtlMode
	}

	/** master slide layout object */
	private readonly _masterSlide: PresSlide
	public get masterSlide(): PresSlide {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private readonly _slides: PresSlide[]
	public get slides(): PresSlide[] {
		return this._slides
	}

	/** this Presentation's sections */
	private readonly _sections: SectionProps[]
	public get sections(): SectionProps[] {
		return this._sections
	}

	/** slide layout definition objects, used for generating slide layout files */
	private _slideLayouts: SlideLayout[]
	public get slideLayouts(): SlideLayout[] {
		return this._slideLayouts
	}

	private LAYOUTS: { [key: string]: PresLayout }

	// Exposed class props
	private readonly _alignH = AlignH
	public get AlignH(): typeof AlignH {
		return this._alignH
	}

	private readonly _alignV = AlignV
	public get AlignV(): typeof AlignV {
		return this._alignV
	}

	private readonly _chartType = ChartType
	public get ChartType(): typeof ChartType {
		return this._chartType
	}

	private readonly _outputType = OutputType
	public get OutputType(): typeof OutputType {
		return this._outputType
	}

	private _presLayout: PresLayout
	public get presLayout(): PresLayout {
		return this._presLayout
	}

	private readonly _schemeColor = SchemeColor
	public get SchemeColor(): typeof SchemeColor {
		return this._schemeColor
	}

	private readonly _shapeType = ShapeType
	public get ShapeType(): typeof ShapeType {
		return this._shapeType
	}

	/**
	 * @depricated use `ChartType`
	 */
	private readonly _charts = CHART_TYPE
	public get charts(): typeof CHART_TYPE {
		return this._charts
	}

	/**
	 * @depricated use `SchemeColor`
	 */
	private readonly _colors = SCHEME_COLOR_NAMES
	public get colors(): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}

	/**
	 * @depricated use `ShapeType`
	 */
	private readonly _shapes = SHAPE_TYPE
	public get shapes(): typeof SHAPE_TYPE {
		return this._shapes
	}

	constructor() {
		const layout4x3: PresLayout = { name: 'screen4x3', width: 9144000, height: 6858000 }
		const layout16x9: PresLayout = { name: 'screen16x9', width: 9144000, height: 5143500 }
		const layout16x10: PresLayout = { name: 'screen16x10', width: 9144000, height: 5715000 }
		const layoutWide: PresLayout = { name: 'custom', width: 12192000, height: 6858000 }
		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: layout4x3,
			LAYOUT_16x9: layout16x9,
			LAYOUT_16x10: layout16x10,
			LAYOUT_WIDE: layoutWide,
		}

		// Core
		this._author = 'PptxGenJS'
		this._company = 'PptxGenJS'
		this._revision = '1' // Note: Must be a whole number
		this._subject = 'PptxGenJS Presentation'
		this._title = 'PptxGenJS Presentation'
		// PptxGenJS props
		this._presLayout = {
			name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			_sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
			width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
		}
		this._rtlMode = false
		//
		// Initialize slide layouts from consolidated registry names (order defines slideLayoutN)
		this._slideLayouts = (CUSTOM_SLIDE_LAYOUT_DEFS && CUSTOM_SLIDE_LAYOUT_DEFS.length > 0)
			? CUSTOM_SLIDE_LAYOUT_DEFS.map((def, idx) => ({
				_margin: DEF_SLIDE_MARGIN_IN,
				_name: def.name,
				_presLayout: this._presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
				_slide: null,
				_slideNum: 1000 + idx + 1,
				_slideNumberProps: null,
				_slideObjects: [],
			}))
			: [
				{
					_margin: DEF_SLIDE_MARGIN_IN,
					_name: DEF_PRES_LAYOUT_NAME,
					_presLayout: this._presLayout,
					_rels: [],
					_relsChart: [],
					_relsMedia: [],
					_slide: null,
					_slideNum: 1000,
					_slideNumberProps: null,
					_slideObjects: [],
				},
			]
		this._slides = []
		this._sections = []
		this._masterSlide = {
			addChart: null,
			addImage: null,
			addMedia: null,
			addNotes: null,
			addShape: null,
			addTable: null,
			addText: null,
			//
			_name: null,
			_presLayout: this._presLayout,
			_rId: null,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideId: null,
			_slideLayout: null,
			_slideNum: null,
			_slideNumberProps: null,
			_slideObjects: [],
		}
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {AddSlideProps} options - slide masterName and/or sectionTitle
	 * @return {PresSlide} new Slide
	 */
	private readonly addNewSlide = (options?: AddSlideProps): PresSlide => {
		// Continue using sections if the first slide using auto-paging has a Section
		const sectAlreadyInUse =
			this.sections.length > 0 &&
			this.sections[this.sections.length - 1]._slides.filter(slide => slide._slideNum === this.slides[this.slides.length - 1]._slideNum).length > 0

		options.sectionTitle = sectAlreadyInUse ? this.sections[this.sections.length - 1].title : null

		return this.addSlide(options)
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {PresSlide} Slide
	 * @since 3.0.0
	 */
	private readonly getSlide = (slideNum: number): PresSlide => this.slides.filter(slide => slide._slideNum === slideNum)[0]

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {SlideNumberProps} slideNum - slide number config
	 */
	private readonly setSlideNumber = (slideNum: SlideNumberProps): void => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide._slideNumberProps = slideNum

		// 2: Add slideNumber to first defined layout (was DEF_PRES_LAYOUT_NAME)
		if (this.slideLayouts && this.slideLayouts.length > 0) {
			this.slideLayouts[0]._slideNumberProps = slideNum
		}
	}

	/**
	 * Create all chart and media rels for this Presentation
	 * @param {PresSlide | SlideLayout} slide - slide with rels
	 * @param {JSZip} zip - JSZip instance
	 * @param {Promise<string>[]} chartPromises - promise array
	 */
	private readonly createChartMediaRels = (slide: PresSlide | SlideLayout, zip: JSZip, chartPromises: Promise<string>[]): void => {
		slide._relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide._relsMedia.forEach(rel => {
			if (rel.type !== 'online' && rel.type !== 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (!data.includes(',') && !data.includes(';')) data = 'image/png;base64,' + data
				else if (!data.includes(',')) data = 'image/png;base64,' + data
				else if (!data.includes(';')) data = 'image/png;' + data

				// C: Add media
				zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true })
			}
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {string} exportName - output file type
	 * @param {Blob} blobContent - Blob content
	 * @return {Promise<string>} Promise with file name
	 */
	private readonly writeFileToBrowser = async (exportName: string, blobContent: Blob): Promise<string> => {
		// STEP 1: Create element
		const eleLink = document.createElement('a')
		eleLink.setAttribute('style', 'display:none;')
		eleLink.dataset.interception = 'off' // @see https://docs.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
		document.body.appendChild(eleLink)

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` to D/L files in client browsers (FYI: synchronously executed)
		if (window.URL.createObjectURL) {
			const url = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
			eleLink.href = url
			eleLink.download = exportName
			eleLink.click()

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(() => {
				window.URL.revokeObjectURL(url)
				document.body.removeChild(eleLink)
			}, 100)

			// Done
			return await Promise.resolve(exportName)
		}
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	private readonly exportPresentation = async (props: WriteProps): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> => {
		const arrChartPromises: Promise<string>[] = []
		let arrMediaPromises: Promise<string>[] = []
		const zip = new JSZip()

		// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
		this.slides.forEach(slide => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(slide))
		})
		this.slideLayouts.forEach(layout => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(layout))
		})
		arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(this.masterSlide))

		// STEP 2: Wait for Promises (if any) then generate the PPTX file
		return await Promise.all(arrMediaPromises).then(async () => {
			// A: Add empty placeholder objects to slides that don't already have them
			this.slides.forEach(slide => {
				if (slide._slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
			})

			// B: Add all required folders and files
			zip.folder('_rels')
			zip.folder('docProps')
			zip.folder('ppt').folder('_rels')
			zip.folder('ppt/charts').folder('_rels')
			zip.folder('ppt/embeddings')
			zip.folder('ppt/media')
			zip.folder('ppt/slideLayouts').folder('_rels')
			zip.folder('ppt/slideMasters').folder('_rels')
			zip.folder('ppt/slides').folder('_rels')
			zip.folder('ppt/theme')
			zip.folder('ppt/notesMasters').folder('_rels')
			zip.folder('ppt/notesSlides').folder('_rels')
			zip.file('[Content_Types].xml', genXml.makeXmlContentTypes(this)) // dynamically generate content-types
			zip.file('_rels/.rels', genXml.makeXmlRootRels())
			zip.file('docProps/app.xml', CUSTOM_PROPS_APP_XML) // TODO: pass only `this` like below! 20200206
			zip.file('docProps/core.xml', CUSTOM_PROPS_CORE_XML) // TODO: pass only `this` like below! 20200206
			zip.file('docProps/custom.xml', CUSTOM_PROPS_CUSTOM_XML) // TODO: pass only `this` like below! 20200206

			// Add customXml files (SharePoint metadata)
			zip.folder('customXml').folder('_rels')
			zip.file('customXml/item1.xml', CUSTOMXML_ITEM1)
			zip.file('customXml/item2.xml', CUSTOMXML_ITEM2)
			zip.file('customXml/item3.xml', CUSTOMXML_ITEM3)
			zip.file('customXml/itemProps1.xml', CUSTOMXML_ITEMPROPS1)
			zip.file('customXml/itemProps2.xml', CUSTOMXML_ITEMPROPS2)
			zip.file('customXml/itemProps3.xml', CUSTOMXML_ITEMPROPS3)
			zip.file('customXml/_rels/item1.xml.rels', CUSTOMXML_ITEM1_RELS)
			zip.file('customXml/_rels/item2.xml.rels', CUSTOMXML_ITEM2_RELS)
			zip.file('customXml/_rels/item3.xml.rels', CUSTOMXML_ITEM3_RELS)

		zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this.slides))
		zip.file('ppt/theme/theme1.xml', CUSTOM_PPT_THEME1_XML)
		// Add theme2 identical to theme1 for notes master to avoid repair prompt
		zip.file('ppt/theme/theme2.xml', CUSTOM_PPT_THEME1_XML)
		zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this))
			zip.file('ppt/presProps.xml', genXml.makeXmlPresProps())
			zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles())
			zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps())

			// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
			// If a custom registry entry exists for a given index, use that exact XML; otherwise generate dynamically
			this.slideLayouts.forEach((_layout, idx) => {
				const layout = this.slideLayouts[idx]
				const def = CUSTOM_SLIDE_LAYOUT_DEFS[idx]
				const relDef = CUSTOM_SLIDE_LAYOUT_RELS[idx]
				if (def && def.xml) {
					// Check if defineSlideMaster() added objects to this layout (like logos)
					let layoutXml = def.xml
					if (layout._slideObjects && layout._slideObjects.length > 0) {
						// Generate XML for the additional objects (logos, images, etc.)
						const extraObjectsXml = genXml.generateLayoutObjectsXml(layout)
						// Inject the extra objects before </p:spTree>
						layoutXml = layoutXml.replace('</p:spTree>', extraObjectsXml + '</p:spTree>')
					}
					zip.file(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, layoutXml)
					// Use dynamic rels to include any images added via defineSlideMaster()
					const layoutRels = genXml.makeXmlSlideLayoutRel(idx + 1, this.slideLayouts)
					zip.file(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, layoutRels)
				} else {
					const genLayoutXml = genXml.makeXmlLayout(this.slideLayouts[idx])
					const genLayoutRels = genXml.makeXmlSlideLayoutRel(idx + 1, this.slideLayouts)
					const sanitizedLayout = this.removeBrokenPics(genLayoutXml, genLayoutRels)
					zip.file(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, sanitizedLayout)
					zip.file(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, genLayoutRels)
				}
			})
			this.slides.forEach((slide, idx) => {
				const slideXml = genXml.makeXmlSlide(slide)
				const slideRels = genXml.makeXmlSlideRel(this.slides, this.slideLayouts, idx + 1)
				// Do NOT sanitize slide XML: slides can legitimately embed images with valid rels
				zip.file(`ppt/slides/slide${idx + 1}.xml`, slideXml)
				zip.file(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, slideRels)
				// Create all slide notes related items. For notes, strip only broken images if any
				const notesRel = genXml.makeXmlNotesSlideRel(idx + 1)
				const notesXml = this.removeBrokenPics(genXml.makeXmlNotesSlide(slide), notesRel)
				zip.file(`ppt/notesSlides/notesSlide${idx + 1}.xml`, notesXml)
				zip.file(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, notesRel)
			})
			zip.file('ppt/slideMasters/slideMaster1.xml', genXml.makeXmlMaster(this.masterSlide, this.slideLayouts))
			zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', genXml.makeXmlSlideMasterRels(this.slideLayouts))
			const notesMasterRel = genXml.makeXmlNotesMasterRel()
			const notesMasterXml = this.removeBrokenPics(genXml.makeXmlNotesMaster(), notesMasterRel)
			zip.file('ppt/notesMasters/notesMaster1.xml', notesMasterXml)
			zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', notesMasterRel)

			// D: Create all Rels (images, media, chart data)
			this.slideLayouts.forEach(layout => {
				this.createChartMediaRels(layout, zip, arrChartPromises)
			})
			this.slides.forEach(slide => {
				this.createChartMediaRels(slide, zip, arrChartPromises)
			})
			this.createChartMediaRels(this.masterSlide, zip, arrChartPromises)

			// E: Wait for Promises (if any) then generate the PPTX file
			return await Promise.all(arrChartPromises).then(async () => {
				if (props.outputType === 'STREAM') {
					// A: stream file
					return await zip.generateAsync({ type: 'nodebuffer', compression: props.compression ? 'DEFLATE' : 'STORE' })
				} else if (props.outputType) {
					// B: Node [fs]: Output type user option or default
					return await zip.generateAsync({ type: props.outputType })
				} else {
					// C: Browser: Output blob as app/ms-pptx
					return await zip.generateAsync({ type: 'blob', compression: props.compression ? 'DEFLATE' : 'STORE' })
				}
			})
		})
	}

	// EXPORT METHODS

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props - output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
	 */
	async stream(props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return await this.exportPresentation({
			compression: props?.compression,
			outputType: 'STREAM',
		})
	}

	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
	 */
	async write(props?: WriteProps | WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		// DEPRECATED: @deprecated v3.5.0 - outputType - [[remove in v4.0.0]]
		const propsOutpType = typeof props === 'object' && props?.outputType ? props.outputType : props ? (props as WRITE_OUTPUT_TYPE) : null
		const propsCompress = typeof props === 'object' && props?.compression ? props.compression : false

		return await this.exportPresentation({
			compression: propsCompress,
			outputType: propsOutpType,
		})
	}

	/**
	 * Export the current Presentation.
	 * Write the generated presentation to disk (Node) or trigger a download (browser).
	 * @param {WriteFileProps} props - output file properties
	 * @returns {Promise<string>} the presentation name
	 */
	async writeFile(props?: WriteFileProps | string): Promise<string> {
		// STEP 1: Figure out where we are running
		const isNode = typeof process !== 'undefined' && !!process.versions?.node && process.release?.name === 'node'

		// STEP 2: Normalise the user arguments
		if (typeof props === 'string') {
			// DEPRECATED: @deprecated v3.5.0 - fileName - [[remove in v4.0.0]]
			console.warn('[WARNING] writeFile(string) is deprecated - pass { fileName } instead.')
			props = { fileName: props }
		}
		const { fileName: rawName = 'Presentation.pptx', compression = false } = props as WriteFileProps
		const fileName = rawName.toLowerCase().endsWith('.pptx') ? rawName : `${rawName}.pptx`

		// STEP 3: Get the binary/Blob from exportPresentation()
		const outputType = isNode ? ('nodebuffer' as const) : null
		const data = await this.exportPresentation({ compression, outputType })

		// STEP 4: Write the file out
		if (isNode) {
			// Dynamically import to avoid bundling fs in the browser build
			const { promises: fs } = await import('node:fs')
			const { writeFile } = fs
			await writeFile(fileName, data as Buffer)
			return fileName
		}

		// Browser branch - push a download
		await this.writeFileToBrowser(fileName, data as Blob)
		return fileName
	}

	// PRESENTATION METHODS

	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: SectionProps): void {
		if (!section) console.warn('addSection requires an argument')
		else if (!section.title) console.warn('addSection requires a title')

		const newSection: SectionProps = {
			_type: 'user',
			_slides: [],
			title: section.title,
		}

		if (section.order) this.sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} options - slide options
	 * @returns {PresSlide} the new Slide
	 */
	addSlide(options?: AddSlideProps): PresSlide {
		// TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
		const masterSlideName = typeof options === 'string' ? options : options?.masterName ? options.masterName : ''
		// Default to using the first defined slide layout where possible (ensures slideLayout1 is used)
		let slideLayout: SlideLayout = this.slideLayouts && this.slideLayouts.length > 0
			? this.slideLayouts[0]
			: {
				_name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
				_presLayout: this.presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
			_slideNum: this.slides.length + 1,
		}

		if (masterSlideName) {
			const tmpLayout = this.slideLayouts.filter(layout => layout._name === masterSlideName)[0]
			if (tmpLayout) slideLayout = tmpLayout
		}		const newSlide = new Slide({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
			presLayout: this.presLayout,
			setSlideNum: this.setSlideNumber,
			slideId: this.slides.length + 256,
			slideRId: this.slides.length + 2,
			slideNumber: this.slides.length + 1,
			slideLayout,
		})

		// A: Add slide to pres
		this._slides.push(newSlide)

		// B: Sections
		// B-1: Add slide to section (if any provided)
		// B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
		if (options?.sectionTitle) {
			const sect = this.sections.filter(section => section.title === options.sectionTitle)[0]
			if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else sect._slides.push(newSlide)
		} else if (this.sections && this.sections.length > 0 && (!options?.sectionTitle)) {
			const lastSect = this._sections[this.sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect._type === 'default') lastSect._slides.push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else {
				this._sections.push({
					title: `Default-${this.sections.filter(sect => sect._type === 'default').length + 1}`,
					_type: 'default',
					_slides: [newSlide],
				})
			}
		}

		return newSlide
	}

	/**
	 * Create a custom Slide Layout in any size
	 * @param {PresLayout} layout - layout properties
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: PresLayout): void {
		// @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
		if (!layout) console.warn('defineLayout requires `{name, width, height}`')
		else if (!layout.name) console.warn('defineLayout requires `name`')
		else if (!layout.width) console.warn('defineLayout requires `width`')
		else if (!layout.height) console.warn('defineLayout requires `height`')
		else if (typeof layout.height !== 'number') console.warn('defineLayout `height` should be a number (inches)')
		else if (typeof layout.width !== 'number') console.warn('defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = {
			name: layout.name,
			_sizeW: Math.round(Number(layout.width) * EMU),
			_sizeH: Math.round(Number(layout.height) * EMU),
			width: Math.round(Number(layout.width) * EMU),
			height: Math.round(Number(layout.height) * EMU),
		}
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props - layout properties
	 */
	defineSlideMaster(props: SlideMasterProps): void {
		// (ISSUE#406;PULL#1176) deep clone the props object to avoid mutating the original object
		const propsClone = JSON.parse(JSON.stringify(props))
		if (!propsClone.title) throw new Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		// Check if a layout with this name already exists
		const existingLayout = this.slideLayouts.find(layout => layout._name === propsClone.title)
		
		let newLayout: SlideLayout
		
		if (existingLayout) {
			// Update existing layout instead of creating a duplicate
			newLayout = existingLayout
			newLayout._margin = propsClone.margin || DEF_SLIDE_MARGIN_IN
			newLayout._slideNumberProps = propsClone.slideNumber || null
			newLayout.background = propsClone.background || null
			newLayout.bkgd = propsClone.bkgd || null
			// Clear existing objects before adding new ones
			newLayout._slideObjects = []
		} else {
			// Create new layout
			newLayout = {
				_margin: propsClone.margin || DEF_SLIDE_MARGIN_IN,
				_name: propsClone.title,
				_presLayout: this.presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
				_slide: null,
				_slideNum: 1000 + this.slideLayouts.length + 1,
				_slideNumberProps: propsClone.slideNumber || null,
				_slideObjects: [],
				background: propsClone.background || null,
				bkgd: propsClone.bkgd || null,
			}
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideMaster(propsClone, newLayout)

		// STEP 2: Add it to layout defs (only if it's a new layout)
		if (!existingLayout) {
			this.slideLayouts.push(newLayout)
		}

		// STEP 3: Add background (image data/path must be captured before `exportPresentation()` is called)
		if (propsClone.background || propsClone.bkgd) genObj.addBackgroundDefinition(propsClone.background, newLayout)

		// STEP 4: Add slideNumber to master slide (if any)
		if (newLayout._slideNumberProps && !this.masterSlide._slideNumberProps) this.masterSlide._slideNumberProps = newLayout._slideNumberProps
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		// @note `verbose` option is undocumented; used for verbose output of layout process
		genTable.genTableToSlides(
			this,
			eleId,
			options,
			options?.masterSlideName ? this.slideLayouts.filter(layout => layout._name === options.masterSlideName)[0] : null
		)
	}
}
