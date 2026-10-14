/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, describe, it */

'use strict';

export {};

const { ExcelTemplate } = require('../build');
const JSZip = require('jszip');
const etree = require('elementtree');
const path = require('path');

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
});
