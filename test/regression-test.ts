/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, describe, it */

'use strict';

export {};

const { ExcelTemplate } = require('../build');
const JSZip = require('jszip');
const etree = require('elementtree');
const path = require('path');
const fs = require('fs');
const { spawnSync } = require('child_process');

function makeInlineStringWorkbook(
	sheetData = '<row r="1"><c r="A1" t="inlineStr"><is><t>${name}</t></is></c></row>',
	dimension = 'A1'
) {
	const zip = new JSZip();

	zip.file('_rels/.rels',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
		'</Relationships>');

	zip.file('xl/workbook.xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
		'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
		'</workbook>');

	zip.file('xl/_rels/workbook.xml.rels',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
		'</Relationships>');

	zip.file('xl/worksheets/sheet1.xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
		'<dimension ref="' + dimension + '"/><sheetData>' + sheetData + '</sheetData>' +
		'</worksheet>');

	zip.file('[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
		'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
		'</Types>');

	return zip.generateAsync({ type: 'nodebuffer' });
}

async function makeTwoSheetWorkbook(sheet1Data = '<row r="1"><c r="A1" t="inlineStr"><is><t>Привет</t></is></c></row>', sheet2Data = '<row r="1"><c r="A1" t="inlineStr"><is><t>Второй лист</t></is></c></row>') {
	const zip = await JSZip.loadAsync(await makeInlineStringWorkbook(sheet1Data));
	const workbookPath = 'xl/workbook.xml';
	const workbook = await zip.file(workbookPath).async('string');
	zip.file(workbookPath, workbook
		.replace('<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>',
			'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>')
		.replace('</workbook>',
			'<definedNames>' +
			'<definedName name="Removed" localSheetId="0">Sheet1!$A$1</definedName>' +
			'<definedName name="Kept" localSheetId="1">Sheet2!$A$1</definedName>' +
			'</definedNames></workbook>'));

	const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
	zip.file(workbookRelsPath, (await zip.file(workbookRelsPath).async('string')).replace('</Relationships>',
		'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'));

	zip.file('xl/worksheets/sheet2.xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
		'<dimension ref="A1"/><sheetData>' + sheet2Data + '</sheetData>' +
		'</worksheet>');

	zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>',
		'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'));

	return zip.generateAsync({ type: 'nodebuffer' });
}

async function readOutput(template) {
	const zip = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
	const sheet = etree.parse(await zip.file('xl/worksheets/sheet1.xml').async('string')).getroot();
	return { zip, sheet };
}

async function addSharedStrings(zip, texts: string[]) {
	const filename = 'xl/_rels/workbook.xml.rels';
	zip.file(filename, (await zip.file(filename).async('string')).replace('</Relationships>',
		'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
	));
	const contentTypes = '[Content_Types].xml';
	zip.file(contentTypes, (await zip.file(contentTypes).async('string')).replace('</Types>',
		'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
	));
	const strings = etree.Element('sst', { xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' });
	for (const text of texts) {
		etree.SubElement(etree.SubElement(strings, 'si'), 't').text = text;
	}
	zip.file('xl/sharedStrings.xml', etree.tostring(strings));
}

function cellAt(sheet, ref: string) {
	const cell = sheet.find("./sheetData/row/c[@r='" + ref + "']");
	expect(cell).not.toBeNull();
	return cell;
}

function expectInlineText(sheet, ref: string, text: string) {
	const cell = cellAt(sheet, ref);
	expect(cell.attrib.t).toBe('inlineStr');
	expect(cell.find('v')).toBeNull();
	expect(cell.find('f')).toBeNull();
	expect(cell.find('is/t')).not.toBeNull();
	expect(cell.find('is/t').text || '').toBe(text);
}

describe('Regression coverage', () => {
	it('resets the shared string lookup when loading another workbook', async () => {
		const template = new ExcelTemplate();

		await template.load(path.join(__dirname, 'templates', 't1.xlsx'));
		template.stringIndex('Only in the previous workbook');
		await template.load(path.join(__dirname, 'templates', 't2.xlsx'));

		const expectedLookup = {};
		template.sharedStrings.forEach((string, index) => {
			expectedLookup[string] = index;
		});

		expect(template.sharedStringsLookup).toEqual(expectedLookup);
		await template.process(1, { revision: 'Only in the previous workbook' });
		const { zip, sheet } = await readOutput(template);
		const strings = etree.parse(await zip.file('xl/sharedStrings.xml').async('string')).getroot().findall('si');
		const values = sheet.findall('./sheetData/row/c').filter(cell => cell.attrib.t === 's').map(cell => {
			const index = Number(cell.find('v').text);
			expect(index).toBeGreaterThanOrEqual(0);
			expect(index).toBeLessThan(strings.length);
			return strings[index].find('t').text;
		});
		expect(values).toContain('Only in the previous workbook');
	});

	it('loads and processes an inline-string workbook without sharedStrings.xml', async () => {
		const template = new ExcelTemplate();

		await template.load(await makeInlineStringWorkbook());
		expect(template.sharedStrings).toEqual([]);

		await template.process('Sheet1', { name: 'Alice & Bob' });

		const { zip, sheet } = await readOutput(template);
		expectInlineText(sheet, 'A1', 'Alice & Bob');
		expect(zip.file('xl/sharedStrings.xml')).toBeNull();

		const rebuilt = await template.build({ type: 'nodebuffer' });
		const reloaded = new ExcelTemplate();
		await reloaded.load(rebuilt);
		await reloaded.process(1, {});
		expectInlineText((await readOutput(reloaded)).sheet, 'A1', 'Alice & Bob');
	});

	it('processes every worksheet through processAll', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>First: ${first}</t></is></c></row>',
			'<row r="1"><c r="A1" t="inlineStr"><is><t>Second: ${second}</t></is></c></row>'
		));

		await template.processAll({ first: 'Alice', second: 'Bob' });

		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const firstSheet = etree.parse(await output.file('xl/worksheets/sheet1.xml').async('string')).getroot();
		const secondSheet = etree.parse(await output.file('xl/worksheets/sheet2.xml').async('string')).getroot();
		expectInlineText(firstSheet, 'A1', 'First: Alice');
		expectInlineText(secondSheet, 'A1', 'Second: Bob');
	});

	it('expands a single-cell dimension across Z without changing its starting cell', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="5"><c r="Z5" t="inlineStr"><is><t>${values}</t></is></c></row>',
			'Z5'
		));

		await template.process(1, { values: [10, 20] });

		const { sheet } = await readOutput(template);
		expect(sheet.find('dimension').attrib.ref).toBe('Z5:AA5');
		expect(sheet.findall('sheetData/row').map(row => row.attrib.r)).toEqual(['5']);
		expect(sheet.findall('sheetData/row/c').map(cell => cell.attrib.r)).toEqual(['Z5', 'AA5']);
		expect(cellAt(sheet, 'Z5').find('v').text).toBe('10');
		expect(cellAt(sheet, 'AA5').find('v').text).toBe('20');
	});

	it('preserves cell styles and mixed value types in a shared-string table', async () => {
		const template = new ExcelTemplate();
		await template.load(path.join(__dirname, 'templates', 't1.xlsx'));
		const originalSheet = await template.loadSheet(1);
		const style = cellAt(originalSheet.root, 'B7').attrib.s;
		const originalStyles = await template.archive.file('xl/styles.xml').async('string');
		const styles = etree.parse(originalStyles).getroot().findall('cellXfs/xf');
		expect(style).toBeDefined();
		expect(styles[Number(style)]).toBeDefined();

		await template.process(1, { planData: [{ name: 123 }, { name: 'text' }, { name: false }] });

		const { zip, sheet } = await readOutput(template);
		const strings = etree.parse(await zip.file('xl/sharedStrings.xml').async('string')).getroot().findall('si');
		const first = cellAt(sheet, 'B7');
		const second = cellAt(sheet, 'B8');
		const third = cellAt(sheet, 'B9');
		for (const cell of [first, second, third]) {
			expect(cell.getchildren().map(child => child.tag)).toEqual(['v']);
			expect(cell.attrib.s).toBe(style);
		}
		expect(first.attrib.t).toBeUndefined();
		expect(first.find('v').text).toBe('123');
		expect(second.attrib.t).toBe('s');
		expect(strings[Number(second.find('v').text)].find('t').text).toBe('text');
		expect(third.attrib.t).toBe('b');
		expect(third.find('v').text).toBe('0');
		expect(await zip.file('xl/styles.xml').async('string')).toBe(originalStyles);
	});

	it.each([
		{ label: 'null', substitutions: { name: null }, type: 'inlineStr', tag: 'is/t', text: '' },
		{ label: 'undefined', substitutions: { name: undefined }, type: 'inlineStr', tag: 'is/t', text: '' },
		{ label: 'missing value', substitutions: {}, type: 'inlineStr', tag: 'is/t', text: '' },
		{ label: 'empty string', substitutions: { name: '' }, type: 'inlineStr', tag: 'is/t', text: '' },
		{ label: 'zero', substitutions: { name: 0 }, type: undefined, tag: 'v', text: '0' },
		{ label: 'decimal', substitutions: { name: -1.25 }, type: undefined, tag: 'v', text: '-1.25' },
		{ label: 'true', substitutions: { name: true }, type: 'b', tag: 'v', text: '1' },
		{ label: 'false', substitutions: { name: false }, type: 'b', tag: 'v', text: '0' },
		{ label: 'date', substitutions: { name: new Date('2013-01-01T00:00:00Z') }, type: undefined, tag: 'v', text: '41275' },
		{ label: 'formula', substitutions: { name: '=SUM(1,2)' }, type: undefined, tag: 'f', text: 'SUM(1,2)' },
	])('serializes $label with a valid cell type', async ({ substitutions, type, tag, text }) => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook());
		await template.process(1, substitutions);
		const { zip, sheet } = await readOutput(template);
		const cell = cellAt(sheet, 'A1');
		expect(cell.attrib.t).toBe(type);
		expect(cell.getchildren().map(child => child.tag)).toEqual([tag === 'is/t' ? 'is' : tag]);
		expect(cell.find(tag).text || '').toBe(text);
		expect(zip.file('xl/sharedStrings.xml')).toBeNull();
	});

	it('preserves whitespace and XML-sensitive Unicode text in inline substitutions', async () => {
		const template = new ExcelTemplate();
		const text = '  Привет & <мир>\n  ';
		await template.load(await makeInlineStringWorkbook());
		await template.process(1, { name: text });
		const { sheet } = await readOutput(template);
		expectInlineText(sheet, 'A1', text);
		expect(cellAt(sheet, 'A1').find('is/t').attrib['xml:space']).toBe('preserve');
	});

	it('substitutes placeholders split across rich-text runs, including empty runs', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is>' +
			'<r><t>Hello ${na</t></r><r><t/></r><r><t>me}: ${count}</t></r>' +
			'</is></c></row>'
		));
		await template.process(1, { name: 'Alice', count: 0 });
		expectInlineText((await readOutput(template)).sheet, 'A1', 'Hello Alice: 0');
	});

	it('preserves empty cells and untouched rich text', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"/><c r="B1" t="inlineStr"><is><t/></is></c>' +
			'<c r="C1" t="inlineStr"><is><r><rPr><b/></rPr><t>Unchanged</t></r></is></c></row>',
			'A1:C1'
		));
		const before = (await readOutput(template)).sheet;
		await template.process(1, {});
		const { sheet } = await readOutput(template);
		for (const ref of ['A1', 'B1', 'C1']) {
			expect(etree.tostring(cellAt(sheet, ref))).toBe(etree.tostring(cellAt(before, ref)));
		}
	});

	it('expands inline arrays without dangling shared-string references', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook());
		await template.process(1, { name: ['Alice', 0, null, false, '=1+1'] });
		const { zip, sheet } = await readOutput(template);
		expectInlineText(sheet, 'A1', 'Alice');
		expect(cellAt(sheet, 'B1').find('v').text).toBe('0');
		expectInlineText(sheet, 'C1', '');
		expect(cellAt(sheet, 'D1').attrib.t).toBe('b');
		expect(cellAt(sheet, 'D1').find('v').text).toBe('0');
		expect(cellAt(sheet, 'E1').find('f').text).toBe('1+1');
		expect(cellAt(sheet, 'E1').find('v')).toBeNull();
		expect(sheet.find('dimension').attrib.ref).toBe('A1:E1');
		expect(zip.file('xl/sharedStrings.xml')).toBeNull();
	});

	it.each([0, false, '=1+1'])('keeps table row types independent when the first value is %p', async first => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>${table:rows.value}</t></is></c></row>'
		));
		await template.process(1, { rows: [{ value: first }, { value: 'Alice' }, { value: null }] });
		const { zip, sheet } = await readOutput(template);
		expectInlineText(sheet, 'A2', 'Alice');
		expectInlineText(sheet, 'A3', '');
		const firstCell = cellAt(sheet, 'A1');
		expect(firstCell.find('is')).toBeNull();
		if (typeof first === 'string') {
			expect(firstCell.find('f').text).toBe('1+1');
			expect(firstCell.find('v')).toBeNull();
		} else {
			expect(firstCell.attrib.t).toBe(typeof first === 'boolean' ? 'b' : undefined);
			expect(firstCell.find('v').text).toBe('0');
		}
		expect(sheet.find('dimension').attrib.ref).toBe('A1:A3');
		expect(zip.file('xl/sharedStrings.xml')).toBeNull();
	});

	it('handles inline and shared strings in the same workbook', async () => {
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>${name}</t></is></c>' +
			'<c r="B1" t="s"><v>0</v></c></row>', 'A1:B1'
		));
		await addSharedStrings(zip, ['${name}']);
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.processAll({ name: 'Alice & Bob' });
		const output = await readOutput(template);
		expectInlineText(output.sheet, 'A1', 'Alice & Bob');
		const cell = cellAt(output.sheet, 'B1');
		expect(cell.attrib.t).toBe('s');
		const strings = etree.parse(await output.zip.file('xl/sharedStrings.xml').async('string')).getroot().findall('si');
		expect(strings[Number(cell.find('v').text)].find('t').text).toBe('Alice & Bob');
	});

	it('can reload shared, inline, and shared workbooks on the same instance', async () => {
		const sharedZip = await JSZip.loadAsync(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
		));
		await addSharedStrings(sharedZip, ['${name}']);
		const sharedWorkbook = await sharedZip.generateAsync({ type: 'nodebuffer' });
		const template = new ExcelTemplate();
		await template.load(sharedWorkbook);
		await template.process(1, { name: 'Old value' });
		await template.load(await makeInlineStringWorkbook());
		await template.process(1, { name: 'Inline value' });
		const inlineOutput = await readOutput(template);
		expectInlineText(inlineOutput.sheet, 'A1', 'Inline value');
		expect(inlineOutput.zip.file('xl/sharedStrings.xml')).toBeNull();
		await template.load(sharedWorkbook);
		await template.process(1, { name: 'Old value' });
		const output = await readOutput(template);
		const strings = etree.parse(await output.zip.file('xl/sharedStrings.xml').async('string')).getroot().findall('si');
		const cell = cellAt(output.sheet, 'A1');
		expect(strings[Number(cell.find('v').text)].find('t').text).toBe('Old value');
	});

	it('substitutes hyperlinks without a shared-strings part', async () => {
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook());
		const sheetPath = 'xl/worksheets/sheet1.xml';
		zip.file(sheetPath, (await zip.file(sheetPath).async('string')).replace('</worksheet>',
			'<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks></worksheet>'
		));
		const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
		zip.file(relsPath, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External" Target="https://example.test/$%257Bname%257D"/></Relationships>'
		);
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.process(1, { name: 'Alice Smith' });
		const output = await readOutput(template);
		const rels = etree.parse(await output.zip.file(relsPath).async('string')).getroot();
		expect(rels.find('Relationship').attrib.Target).toBe('https://example.test/Alice%20Smith');
		expectInlineText(output.sheet, 'A1', 'Alice Smith');
	});

	it('rejects a workbook that references a missing shared-strings part', async () => {
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook());
		await addSharedStrings(zip, []);
		zip.remove('xl/sharedStrings.xml');
		const template = new ExcelTemplate();
		await expect(template.load(await zip.generateAsync({ type: 'nodebuffer' })))
			.rejects.toThrow('Shared strings file not found: xl/sharedStrings.xml');
	});

	it('deletes and reindexes local defined names with the sheet', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		await template.deleteSheet('Sheet1');

		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const workbook = etree.parse(await output.file('xl/workbook.xml').async('string')).getroot();
		const sheets = workbook.findall('sheets/sheet');
		const definedNames = workbook.findall('definedNames/definedName');

		expect(sheets).toHaveLength(1);
		expect(sheets[0].attrib.name).toBe('Sheet2');
		expect(sheets[0].attrib.sheetId).toBe('1');
		expect(definedNames).toHaveLength(1);
		expect(definedNames[0].attrib.name).toBe('Kept');
		expect(definedNames[0].attrib.localSheetId).toBe('0');
	});

	it('copies the original sheet part and registers its content type', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		await template.copySheet('Sheet1');

		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const copiedSheet = etree.parse(await output.file('xl/worksheets/sheet3.xml').async('string')).getroot();
		const workbook = etree.parse(await output.file('xl/workbook.xml').async('string')).getroot();
		const contentTypes = etree.parse(await output.file('[Content_Types].xml').async('string')).getroot();

		expect(await output.file('xl/worksheets/sheet3.xml').async('nodebuffer'))
			.toEqual(await output.file('xl/worksheets/sheet1.xml').async('nodebuffer'));
		expect(copiedSheet.find('sheetData/row/c/is/t').text).toBe('Привет');
		expect(workbook.find("sheets/sheet[@name='Sheet3']")).not.toBeNull();
		expect(contentTypes.find("Override[@PartName='/xl/worksheets/sheet3.xml']")).not.toBeNull();
	});

	it('copies comment relationship parts when copying a sheet', async () => {
		const zip = await JSZip.loadAsync(await makeTwoSheetWorkbook());
		zip.file('xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
			'</Relationships>');
		zip.file('xl/comments1.xml', '<comments><authors><author>tc={11111111-1111-1111-1111-111111111111}</author></authors></comments>');
		zip.file('xl/drawings/vmlDrawing1.vml', '<xml data="1"/>');

		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.copySheet('Sheet1');

		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const copiedComments = await output.file('xl/comments3.xml').async('string');
		const copiedVml = await output.file('xl/drawings/vmlDrawing3.vml').async('string');
		const copiedRels = etree.parse(await output.file('xl/worksheets/_rels/sheet3.xml.rels').async('string')).getroot();

		expect(copiedComments).toContain('tc={');
		expect(copiedComments).not.toContain('11111111-1111-1111-1111-111111111111');
		expect(copiedVml).toContain('data="3"');
		expect(copiedRels.find("Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']").attrib.Target).toBe('../comments3.xml');
	});

	it('moves hyperlink anchors when rows and columns are inserted', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row>' +
			'<row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>',
			'A1:B2'
		));
		const sheet = await template.loadSheet('Sheet1');
		const hyperlinks = etree.SubElement(sheet.root, 'hyperlinks');
		etree.SubElement(hyperlinks, 'hyperlink', { ref: 'B2' });
		etree.SubElement(hyperlinks, 'hyperlink', { ref: 'A1' });

		template.pushRight(template.workbook, sheet.root, 'A1', 1);
		template.pushDown(template.workbook, sheet.root, [], 1, 1);

		expect(sheet.root.find('hyperlinks/hyperlink').attrib.ref).toBe('B3');
		expect(sheet.root.findall('hyperlinks/hyperlink')[1].attrib.ref).toBe('A1');
	});

	it('hides zero-based column indexes', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row>', 'A1:C1'
		));
		const sheet = await template.loadSheet('Sheet1');
		const cols = etree.Element('cols');
		sheet.root.getchildren().splice(sheet.root.getchildren().indexOf(sheet.root.find('sheetData')), 0, cols);
		etree.SubElement(cols, 'col', { min: '1', max: '2', width: '10' });
		etree.SubElement(cols, 'col', { min: '3', max: '3', width: '10' });
		template.archive.file(sheet.filename, etree.tostring(sheet.root));

		await template.hideCols('Sheet1', [1, 2]);
		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const outputSheet = etree.parse(await output.file('xl/worksheets/sheet1.xml').async('string')).getroot();

		expect(outputSheet.findall('cols/col').map(col => col.attrib)).toEqual([
			{ min: '1', max: '1', width: '10' },
			{ min: '2', max: '2', width: '10', hidden: '1' },
			{ min: '3', max: '3', width: '10', hidden: '1' }
		]);
	});

	it('loads drawings whose drawing relationships file is absent', async () => {
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>image</t></is></c></row>'
		));
		const sheetPath = 'xl/worksheets/sheet1.xml';
		zip.file(sheetPath, (await zip.file(sheetPath).async('string')).replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'));
		zip.file('xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>' +
			'</Relationships>');
		zip.file('xl/drawings/drawing1.xml',
			'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>');

		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		const sheet = await template.loadSheet('Sheet1');
		const rels = await template.loadSheetRels(sheet.filename);
		const drawing = await template.loadDrawing(sheet.root, sheet.filename, rels.root);

		expect(drawing.root.tag).toBe('xdr:wsDr');
		expect(drawing.relRoot.tag).toBe('Relationships');
	});

	it('awaits drawing loading when substituting a normal image', async () => {
		const template = new ExcelTemplate();
		const onePixelPng = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>${image:photo}</t></is></c></row>'
		));

		await template.process(1, { photo: onePixelPng });
		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));

		expect(output.file('xl/media/image1.jpg')).not.toBeNull();
		expect(output.file('xl/drawings/drawing1.xml')).not.toBeNull();
		expect(output.file('xl/drawings/_rels/drawing1.xml.rels')).not.toBeNull();
	});

	it('hides columns without existing column definitions and preserves workbook relationships', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		const workbookBefore = await template.archive.file('xl/workbook.xml').async('string');
		const relsBefore = await template.archive.file('xl/_rels/workbook.xml.rels').async('string');
		await template.hideCols(1, [3, 1, 3]);
		await template.hideCols(1, [1]);
		const { zip, sheet } = await readOutput(template);
		expect(sheet.findall('cols/col').map(col => col.attrib)).toEqual([
			{ min: '2', max: '2', hidden: '1' },
			{ min: '4', max: '4', hidden: '1' }
		]);
		expect(sheet.getchildren().indexOf(sheet.find('cols')))
			.toBeLessThan(sheet.getchildren().indexOf(sheet.find('sheetData')));
		expect(await zip.file('xl/workbook.xml').async('string')).toBe(workbookBefore);
		expect(await zip.file('xl/_rels/workbook.xml.rels').async('string')).toBe(relsBefore);
	});

	it.each([-1, 1.5, 16384, '1'])('rejects invalid column index %p without modifying the sheet', async index => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook());
		const before = await template.archive.file('xl/worksheets/sheet1.xml').async('string');
		await expect(template.hideCols(1, [index])).rejects.toThrow();
		expect(await template.archive.file('xl/worksheets/sheet1.xml').async('string')).toBe(before);
	});

	it('copies a sheet after deletion without overwriting any existing worksheet', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		await template.deleteSheet('Sheet1');
		const before = await template.archive.file('xl/worksheets/sheet2.xml').async('nodebuffer');
		await template.copySheet('Sheet2', 'Копия');
		expect(template.sheets[0].filename).not.toBe(template.sheets[1].filename);
		expect(await template.archive.file('xl/worksheets/sheet2.xml').async('nodebuffer')).toEqual(before);
		await template.process('Копия', {});
		expect(await template.archive.file('xl/worksheets/sheet2.xml').async('nodebuffer')).toEqual(before);
		const reloaded = new ExcelTemplate();
		await reloaded.load(await template.build({ type: 'nodebuffer' }));
		expect(reloaded.sheets.map(sheet => sheet.name)).toEqual(['Sheet2', 'Копия']);
	});

	it('preserves sheet-to-part mappings with reordered relationship IDs', async () => {
		const zip = await JSZip.loadAsync(await makeTwoSheetWorkbook());
		const filename = 'xl/workbook.xml';
		zip.file(filename, (await zip.file(filename).async('string'))
			.replace('r:id="rId1"', 'r:id="rId20"'));
		const rels = 'xl/_rels/workbook.xml.rels';
		zip.file(rels, (await zip.file(rels).async('string'))
			.replace('Id="rId1"', 'Id="rId20"')
			.replace('</Relationships>', '<Relationship Id="rId7" Type="custom" Target="custom.xml"/></Relationships>'));
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.copySheet(1, 'Копия');
		expect((await template.loadSheet('Sheet1')).root.find('sheetData/row/c/is/t').text).toBe('Привет');
		expect((await template.loadSheet('Sheet2')).root.find('sheetData/row/c/is/t').text).toBe('Второй лист');
		expect(template.workbookRels.find("Relationship[@Id='rId7']").attrib.Target).toBe('custom.xml');
		await template.deleteSheet('Копия');
		expect((await template.loadSheet('Sheet1')).root.find('sheetData/row/c/is/t').text).toBe('Привет');
	});

	it('copies local names by sheet position, including quoted names and constants', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		const name = "Bob's report";
		template.workbook.find('sheets/sheet').attrib.name = name;
		const names = template.workbook.find('definedNames');
		names.find('definedName').text = "'Bob''s report'!$A$1,'Bob''s report'!$C$1";
		etree.SubElement(names, 'definedName', { name: 'Constant', localSheetId: '0' }).text = '42';
		etree.SubElement(names, 'definedName', { name: 'Global' }).text = "'Bob''s report'!$B$1";
		template._rebuild();
		await template.copySheet(1, "Copy's report");
		const copiedNames = template.workbook.findall('definedNames/definedName')
			.filter(item => item.attrib.localSheetId === '2');
		expect(copiedNames.map(item => item.text)).toEqual([
			"'Copy''s report'!$A$1,'Copy''s report'!$C$1", '42'
		]);
		expect(names.findall('definedName').filter(item => item.attrib.name === 'Global')).toHaveLength(1);
		expect(names.find('definedName').text).toBe("'Bob''s report'!$A$1,'Bob''s report'!$C$1");
	});

	it('copies threaded comments with distinct IDs and preserves reply and author links', async () => {
		const zip = await JSZip.loadAsync(await makeTwoSheetWorkbook());
		const first = '{11111111-1111-4111-8111-111111111111}';
		const reply = '{22222222-2222-4222-8222-222222222222}';
		const other = '{33333333-3333-4333-8333-333333333333}';
		const person = '{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}';
		zip.file('xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
			'<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>' +
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
			'</Relationships>');
		const comments = '<comments xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
			'<authors><author>tc=' + first + '</author><author>tc=' + other + '</author></authors>' +
			'<commentList><comment ref="A1" authorId="0" xr:uid="' + first + '"><text><t>Привет ' + first + '</t></text></comment>' +
			'<comment ref="B1" authorId="1" xr:uid="' + other + '"><text><t>Другой</t></text></comment></commentList></comments>';
		const threaded = '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
			'<threadedComment ref="A1" id="' + first + '" personId="' + person + '"><text>Root</text></threadedComment>' +
			'<threadedComment ref="A1" id="' + reply + '" parentId="' + first + '" personId="' + person + '"><text>Reply</text></threadedComment>' +
			'<threadedComment ref="B1" id="' + other + '" personId="' + person + '"><text>Other</text></threadedComment></ThreadedComments>';
		zip.file('xl/comments1.xml', comments);
		zip.file('xl/threadedComments/threadedComment1.xml', threaded);
		zip.file('xl/drawings/vmlDrawing1.vml', '<xml data="1"/>');
		zip.file('xl/drawings/_rels/vmlDrawing1.vml.rels', '<Relationships><Relationship Id="rId1" Target="../media/unchanged.png"/></Relationships>');
		zip.file('xl/comments3.xml', 'existing comments');
		zip.file('xl/drawings/vmlDrawing3.vml', 'existing VML');
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.copySheet(1, 'Копия');
		const output = await JSZip.loadAsync(await template.build({ type: 'nodebuffer' }));
		const rels = etree.parse(await output.file('xl/worksheets/_rels/sheet3.xml.rels').async('string')).getroot();
		const targets = rels.findall('Relationship').map(rel =>
			path.posix.join('xl/worksheets', rel.attrib.Target));
		const copiedComments = etree.parse(await output.file(targets[0]).async('string')).getroot();
		const copiedThreads = etree.parse(await output.file(targets[1]).async('string')).getroot().findall('threadedComment');
		const ids = copiedThreads.map(comment => comment.attrib.id);
		expect(new Set(ids).size).toBe(3);
		expect(ids).not.toContain(first);
		expect(ids).not.toContain(reply);
		expect(ids).not.toContain(other);
		expect(copiedThreads[1].attrib.parentId).toBe(ids[0]);
		expect(copiedThreads.map(comment => comment.attrib.personId)).toEqual([person, person, person]);
		expect(copiedComments.findall('authors/author').map(author => author.text)).toEqual(['tc=' + ids[0], 'tc=' + ids[2]]);
		expect(copiedComments.findall('commentList/comment').map(comment => comment.attrib['xr:uid'])).toEqual([ids[0], ids[2]]);
		expect(copiedComments.find('commentList/comment/text/t').text).toBe('Привет ' + first);
		expect(await output.file('xl/comments1.xml').async('string')).toBe(comments);
		expect(await output.file('xl/threadedComments/threadedComment1.xml').async('string')).toBe(threaded);
		expect(await output.file('xl/comments3.xml').async('string')).toBe('existing comments');
		expect(await output.file('xl/drawings/vmlDrawing3.vml').async('string')).toBe('existing VML');
		const vmlRels = path.posix.join(path.posix.dirname(targets[2]), '_rels', path.posix.basename(targets[2]) + '.rels');
		expect(await output.file(vmlRels).async('string'))
			.toBe(await output.file('xl/drawings/_rels/vmlDrawing1.vml.rels').async('string'));
	});

	it('moves hyperlink ranges with their cells during array and table substitution', async () => {
		const arrayToken = '$' + '{values}';
		const tableToken = '$' + '{table:rows.value}';
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>' + arrayToken + '</t></is></c><c r="C1" t="inlineStr"><is><t>Link</t></is></c></row>' +
			'<row r="2"><c r="A2" t="inlineStr"><is><t>' + tableToken + '</t></is></c></row>' +
			'<row r="3"><c r="C3" t="inlineStr"><is><t>Below</t></is></c></row>', 'A1:D4'));
		const sheetPath = 'xl/worksheets/sheet1.xml';
		zip.file(sheetPath, (await zip.file(sheetPath).async('string')).replace('</worksheet>',
			'<hyperlinks><hyperlink ref="C1:D1" location="Sheet1!A1"/>' +
			'<hyperlink ref="C3:D4" location="Sheet1!A1"/></hyperlinks></worksheet>'));
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.process(1, { values: ['a', 'b'], rows: [{ value: 'a' }, { value: 'b' }] });
		const { sheet } = await readOutput(template);
		expect(sheet.findall('hyperlinks/hyperlink').map(link => link.attrib.ref)).toEqual(['D1:E1', 'C4:D5']);
		expectInlineText(sheet, 'D1', 'Link');
		expectInlineText(sheet, 'C4', 'Below');
	});

	it('splits multi-row hyperlink ranges when only one row expands', async () => {
		const template = new ExcelTemplate();
		const sheet = etree.parse('<worksheet><hyperlinks>' +
			'<hyperlink ref="C1:D3" r:id="rId1"/><hyperlink ref="A2:B2" location="Sheet1!A1"/>' +
			'</hyperlinks></worksheet>').getroot();
		template.pushRight(etree.Element('workbook'), sheet, 'A2', 2);
		expect(sheet.findall('hyperlinks/hyperlink').map(link => link.attrib)).toEqual([
			{ ref: 'C1:D1', 'r:id': 'rId1' },
			{ ref: 'E2:F2', 'r:id': 'rId1' },
			{ ref: 'C3:D3', 'r:id': 'rId1' },
			{ ref: 'A2:D2', location: 'Sheet1!A1' }
		]);
		template.pushDown(etree.Element('workbook'), sheet, [], 1, 2);
		expect(sheet.findall('hyperlinks/hyperlink').map(link => link.attrib.ref))
			.toEqual(['C1:D1', 'E4:F4', 'C5:D5', 'A4:D4']);
	});

	it('leaves the workbook intact if a related comment part is missing', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		template.archive.file('xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../missing.xml"/></Relationships>');
		const files = Object.keys(template.archive.files);
		const workbookBefore = etree.tostring(template.workbook);
		const typesBefore = etree.tostring(template.contentTypes);
		await expect(template.copySheet(1, 'Copy')).rejects.toThrow('Related sheet part not found');
		expect(Object.keys(template.archive.files)).toEqual(files);
		expect(etree.tostring(template.workbook)).toBe(workbookBefore);
		expect(etree.tostring(template.contentTypes)).toBe(typesBefore);
	});

	it('rejects a duplicate sheet name and picks a free default name', async () => {
		const template = new ExcelTemplate();
		await template.load(await makeTwoSheetWorkbook());
		await expect(template.copySheet(1, 'sheet2')).rejects.toThrow('Sheet already exists');
		await template.copySheet(1, 'Sheet4');
		await template.copySheet(1);
		expect(template.sheets.map(sheet => sheet.name)).toEqual(['Sheet1', 'Sheet2', 'Sheet4', 'Sheet5']);
	});

	it.each([false, true])('inserts an image into an existing drawing with rels=%p', async withRels => {
		const token = '$' + '{image:photo}';
		const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64');
		const zip = await JSZip.loadAsync(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>' + token + '</t></is></c></row>'));
		const sheetPath = 'xl/worksheets/sheet1.xml';
		zip.file(sheetPath, (await zip.file(sheetPath).async('string')).replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'));
		zip.file('xl/worksheets/_rels/sheet1.xml.rels',
			'<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/drawing1.xml"/></Relationships>');
		zip.file('xl/drawings/drawing1.xml',
			'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
			'<xdr:twoCellAnchor><xdr:from><xdr:row>3</xdr:row></xdr:from><xdr:to><xdr:row>4</xdr:row></xdr:to></xdr:twoCellAnchor></xdr:wsDr>');
		const relsPath = 'xl/drawings/_rels/drawing1.xml.rels';
		if (withRels)
			zip.file(relsPath, '<Relationships><Relationship Id="rId5" Type="custom" Target="../media/old.png"/></Relationships>');
		const template = new ExcelTemplate();
		await template.load(await zip.generateAsync({ type: 'nodebuffer' }));
		await template.process(1, { photo: png });
		const { zip: output } = await readOutput(template);
		const drawing = etree.parse(await output.file('xl/drawings/drawing1.xml').async('string')).getroot();
		const rels = etree.parse(await output.file(relsPath).async('string')).getroot();
		expect(drawing.findall('xdr:twoCellAnchor')).toHaveLength(1);
		expect(drawing.findall('xdr:oneCellAnchor')).toHaveLength(1);
		expect(rels.findall('Relationship')).toHaveLength(withRels ? 2 : 1);
		const blip = drawing.find('xdr:oneCellAnchor/xdr:pic/xdr:blipFill/a:blip');
		const imageRel = rels.findall('Relationship').find(rel => rel.attrib.Id === blip.attrib['r:embed']);
		expect(await output.file(path.posix.join('xl/drawings', imageRel.attrib.Target)).async('nodebuffer')).toEqual(png);
	});

	it.each(['buffer', 'base64', 'absolute path', 'relative path', 'imageRootPath'])('preserves image bytes and scaled dimensions from %s', async inputType => {
		const filename = path.join(__dirname, 'templates', 'dataset', 'large.png');
		const png = fs.readFileSync(filename); // Fixture dimensions: 276 x 85 pixels.
		const template = new ExcelTemplate({
			imageRatio: 75.4,
			imageRootPath: inputType === 'imageRootPath' ? path.dirname(filename) : null
		});
		const inputs = {
			buffer: png,
			base64: png.toString('base64'),
			'absolute path': filename,
			'relative path': path.relative(process.cwd(), filename),
			imageRootPath: path.basename(filename)
		};
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>${image:photo}</t></is></c></row>'));
		await template.process(1, { photo: inputs[inputType] });
		const { zip } = await readOutput(template);
		const drawing = etree.parse(await zip.file('xl/drawings/drawing1.xml').async('string')).getroot();
		const anchors = drawing.findall('xdr:oneCellAnchor');
		expect(anchors).toHaveLength(1);
		const expectedSize = {
			cx: String(Math.floor(276 * 9525 * 75.4 / 100)),
			cy: String(Math.floor(85 * 9525 * 75.4 / 100))
		};
		expect(anchors[0].find('xdr:ext').attrib).toEqual(expectedSize);
		expect(anchors[0].find('xdr:pic/xdr:spPr/a:xfrm/a:ext').attrib).toEqual(expectedSize);
		const rels = etree.parse(await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string')).getroot();
		const embed = anchors[0].find('xdr:pic/xdr:blipFill/a:blip').attrib['r:embed'];
		const relationship = rels.findall('Relationship').find(rel => rel.attrib.Id === embed);
		expect(relationship).toBeDefined();
		expect(await zip.file(path.posix.join('xl/drawings', relationship.attrib.Target)).async('nodebuffer')).toEqual(png);
	});

	it.each([
		['unrecognized buffer', Buffer.from('not an image')],
		['invalid base64', 'not an image'],
		['truncated PNG', Buffer.from('89504e470d0a1a0a', 'hex')]
	])('rejects %s with the library image error', async (_name, photo) => {
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>${image:photo}</t></is></c></row>'));
		await expect(template.process(1, { photo })).rejects.toThrow(/imageObj cannot be parse as a buffer/);
		expect(Object.keys(template.archive.files).filter(name => name.startsWith('xl/media/'))).toEqual([]);
	});

	it('rejects an ICNS entry with zero length without hanging', () => {
		// A separate process bounds runtime even if a future dependency regresses to an infinite loop.
		const result = spawnSync(process.execPath, ['-e',
			"const assert = require('assert');" +
			"const { ExcelTemplate } = require(process.argv[1]);" +
			"const invalid = Buffer.from('69636e73000000106963703400000000', 'hex');" +
			"assert.throws(() => new ExcelTemplate().imageToBuffer(invalid), /imageObj cannot be parse as a buffer image/);",
			require.resolve('../build')
		], { encoding: 'utf8', timeout: 2000 });
		expect(result.error).toBeUndefined();
		expect(result.stderr).toBe('');
		expect(result.status).toBe(0);
	}, 5000);

	it('keeps cells and styles for all image table rows', async () => {
		const token = '$' + '{table:rows.photo:image}';
		const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
		const template = new ExcelTemplate();
		await template.load(await makeInlineStringWorkbook(
			'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>' + token + '</t></is></c></row>'));
		await template.process(1, { rows: [{ photo: png }, { photo: png }] });
		const { zip, sheet } = await readOutput(template);
		expect(cellAt(sheet, 'A1').attrib.s).toBe('1');
		expect(cellAt(sheet, 'A2').attrib.s).toBe('1');
		const drawing = etree.parse(await zip.file('xl/drawings/drawing1.xml').async('string')).getroot();
		expect(drawing.findall('xdr:oneCellAnchor/xdr:from/xdr:row').map(row => row.text)).toEqual(['0', '1']);
	});
});
