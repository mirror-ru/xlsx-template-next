/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, module, Buffer */

const path = require('path');
const sizeOf = require('image-size').imageSize;
const fs = require('fs');
const etree = require('elementtree');
const JSZip = require('jszip');
const { randomBytes } = require('crypto');

// Package paths use forward slashes on every platform; targets may be root-relative.
function resolvePartPath(source, target) {
	return path.posix.normalize(target.startsWith('/') ? target.slice(1) : path.posix.join(path.posix.dirname(source), target));
}

const DOCUMENT_RELATIONSHIP       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const CALC_CHAIN_RELATIONSHIP     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const SHARED_STRINGS_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const HYPERLINK_RELATIONSHIP      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

var _get_simple = function (obj, desc) {
	if (desc.indexOf('[') >= 0) {
		var specification = desc.split(/[[[\]]/);
		var property = specification[0];
		var index = specification[1];

		return obj[property][index];
	}

	return obj[desc];
}

/**
 * Based on http://stackoverflow.com/questions/8051975
 * Mimic https://lodash.com/docs#get
 */
var _get = function(obj, desc, defaultValue) {
	var arr = desc.split('.');
	try {
		while (arr.length) {
			obj = _get_simple(obj, arr.shift());
		}
	} catch(ex) {
		/* invalid chain */
		obj = undefined;
	}
	
	return obj === undefined ? defaultValue : obj;
}

class ExcelTemplate {
	/**
	 * Create a new workbook. Either pass the raw data of a .xlsx file,
	 * and call `loadTemplate()` later.
	 */
	constructor(option = {}) {
		this.archive = null;
		this.sharedStrings = [];
		this.sharedStringsLookup = {};
		this.option = {
			moveImages: false,
			subsituteAllTableRow: false,
			moveSameLineImages: false,
			imageRatio: 100,
			pushDownPageBreakOnTableSubstitution: false,
			imageRootPath: null,
			handleImageError: null,
		};
		Object.assign(this.option, option);
		this.sharedStringsPath = "";
		this.sheets = [];
		this.sheet = null;
		this.workbook = null;
		this.workbookPath = null;
		this.contentTypes = null;
		this.prefix = null;
		this.workbookRels = null;
		this.calChainRel = null;
		this.calcChainPath = "";
	}

	/**
	 * Delete unused sheets if needed
	 */
	async deleteSheet(sheetName) {
		var sheet = await this.loadSheet(sheetName);

		var sh = this.workbook.find("sheets/sheet[@sheetId='" + sheet.id + "']");
		var sheets = this.workbook.findall('sheets/sheet');
		var sheetIndex = sheets.indexOf(sh);
		var definedNamesParent = this.workbook.find('definedNames');

		// A local defined name is tied to the zero-based position of a sheet.
		// Remove names belonging to the deleted sheet and shift the following ones.
		if (definedNamesParent !== null) {
			var definedNamesToRemove = [];

			this.workbook.findall('definedNames/definedName').forEach(element => {
				if (element.attrib.localSheetId === undefined)
					return;

				var localSheetId = parseInt(element.attrib.localSheetId, 10);
				if (localSheetId === sheetIndex) {
					definedNamesToRemove.push(element);
				} else if (localSheetId > sheetIndex) {
					element.attrib.localSheetId = (localSheetId - 1).toString();
				}
			});

			definedNamesToRemove.forEach(element => definedNamesParent.remove(element));
		}

		this.workbook.find('sheets').remove(sh);

		var rel = this.workbookRels.find("Relationship[@Id='" + sh.attrib['r:id'] + "']");
		this.workbookRels.remove(rel);

		this._rebuild();

		return this;
	}

	/**
	 * Clone sheets in current workbook template
	 * @param {string|number} sheetName
	 * @param {string} [copyName]
	 */
	async copySheet(sheetName, copyName) {
		const sheet = await this.loadSheet(sheetName);
		const sheets = this.workbook.findall('sheets/sheet');
		const sourceIndex = sheets.findIndex(item => Number(item.attrib.sheetId) === sheet.id);
		const newSheetIndex = sheets.length + 1;
		const existingNames = new Set(sheets.map(item => item.attrib.name.toLowerCase()));
		let finalSheetName = copyName;
		if (!finalSheetName) {
			let suffix = newSheetIndex;
			while (existingNames.has(('Sheet' + suffix).toLowerCase()))
				suffix++;
			finalSheetName = 'Sheet' + suffix;
		}
		if (typeof finalSheetName !== 'string' || finalSheetName.length > 31 ||
			/[:\\/?*\[\]]/.test(finalSheetName) || /^'|'$/.test(finalSheetName))
			throw new Error('Invalid sheet name: ' + finalSheetName);
		if (existingNames.has(finalSheetName.toLowerCase()))
			throw new Error('Sheet already exists: ' + finalSheetName);

		// Prepare all parts before modifying the archive, so a failed read leaves it intact.
		const parts = new Map();
		const contentTypes = [];
		const relsPath = filename => path.posix.join(path.posix.dirname(filename), '_rels', path.posix.basename(filename) + '.rels');
		const unusedPartPath = filename => {
			const extension = path.posix.extname(filename);
			const base = filename.slice(0, -extension.length).replace(/\d+$/, '');
			let suffix = newSheetIndex;
			let candidate;
			do {
				candidate = base + suffix++ + extension;
			} while (this.archive.files[candidate] || this.archive.files[relsPath(candidate)] || parts.has(candidate) ||
				this.contentTypes.findall('Override').some(item => item.attrib.PartName === '/' + candidate));
			return candidate;
		};
		const arcName = unusedPartPath(path.posix.join(this.prefix, 'worksheets/sheet.xml'));
		const sheetDirectory = path.posix.dirname(arcName);
		parts.set(arcName, await this.archive.file(sheet.filename).async('nodebuffer'));
		contentTypes.push([arcName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']);

		const sourceRels = await this.loadSheetRels(sheet.filename);
		const newRelsRoot = this.cloneElement(sourceRels.root);
		const copiedTypes = new Map([
			['http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'],
			['http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', 'application/vnd.ms-excel.threadedcomments+xml'],
			['http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing', 'application/vnd.openxmlformats-officedocument.vmlDrawing']
		]);
		const commentIds = new Map();
		const remapId = id => {
			const key = id.replace(/[{}]/g, '').toLowerCase();
			if (!commentIds.has(key))
				commentIds.set(key, this.generateUUID());
			const value = commentIds.get(key);
			return id.startsWith('{') ? '{' + value + '}' : value;
		};
		const rewriteCommentIds = node => {
			const tag = node.tag.split(':').pop();
			if (tag === 'threadedComment') {
				for (const attr of ['id', 'parentId']) {
					if (node.attrib[attr])
						node.attrib[attr] = remapId(node.attrib[attr]);
				}
			} else if (tag === 'comment') {
				for (const attr of Object.keys(node.attrib)) {
					if (attr.endsWith(':uid'))
						node.attrib[attr] = remapId(node.attrib[attr]);
				}
			} else if (tag === 'author' && /^tc=\{[0-9a-f-]{36}\}$/i.test(node.text || '')) {
				node.text = 'tc=' + remapId(node.text.slice(3));
			}
			node.getchildren().forEach(rewriteCommentIds);
		};
		for (const relationship of newRelsRoot.findall('Relationship')) {
			const target = relationship.attrib.Target;
			if (!target || relationship.attrib.TargetMode === 'External')
				continue;
			const sourcePath = resolvePartPath(sheet.filename, target);
			let copiedPath = sourcePath;
			if (copiedTypes.has(relationship.attrib.Type)) {
				const sourceFile = this.archive.file(sourcePath);
				if (sourceFile === null)
					throw new Error('Related sheet part not found: ' + sourcePath);
				copiedPath = unusedPartPath(sourcePath);
				let content = await sourceFile.async('string');
				if (relationship.attrib.Type.endsWith('/vmlDrawing')) {
					content = content.replace(/data="\d+"/, 'data="' + newSheetIndex + '"');
				} else {
					const root = etree.parse(content).getroot();
					rewriteCommentIds(root);
					content = etree.tostring(root);
				}
				parts.set(copiedPath, content);
				contentTypes.push([copiedPath, copiedTypes.get(relationship.attrib.Type)]);
				// VML can have its own image relationships. The copy stays in the same directory.
				const sourcePartRels = this.archive.file(relsPath(sourcePath));
				if (sourcePartRels !== null)
					parts.set(relsPath(copiedPath), await sourcePartRels.async('nodebuffer'));
			}
			relationship.attrib.Target = path.posix.relative(sheetDirectory, copiedPath);
		}
		if (newRelsRoot.findall('Relationship').length)
			parts.set(relsPath(arcName), etree.tostring(newRelsRoot));

		const newDefinedNames = this.workbook.findall('definedNames/definedName')
			.filter(element => Number(element.attrib.localSheetId) === sourceIndex && element.attrib.localSheetId !== undefined)
			.map(element => {
				const copy = this.cloneElement(element);
				copy.attrib.localSheetId = String(newSheetIndex - 1);
				// Skip formula string literals; quote the new name and escape apostrophes.
				copy.text = (copy.text || '').replace(/"(?:[^"]|"")*"|('(?:[^']|'')+'|(?:\[[^\]]+\])?[A-Za-z_\u0080-\uFFFF][A-Za-z0-9_.\u0080-\uFFFF]*)!/g,
					(match, name) => {
						if (!name)
							return match;
						const unquoted = name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
						return unquoted.toLowerCase() === sheet.name.toLowerCase()
							? "'" + finalSheetName.replace(/'/g, "''") + "'!"
							: match;
					});
				return copy;
			});
		const usedRelIds = new Set(this.workbookRels.findall('Relationship').map(item => item.attrib.Id));
		let relIndex = 1;
		while (usedRelIds.has('rId' + relIndex))
			relIndex++;
		const relId = 'rId' + relIndex;

		for (const [filename, content] of parts)
			this.archive.file(filename, content);
		for (const [filename, contentType] of contentTypes)
			this.addContentType('/' + filename, contentType);
		etree.SubElement(this.workbook.find('sheets'), 'sheet', {
			name: finalSheetName, sheetId: String(newSheetIndex), 'r:id': relId
		});
		etree.SubElement(this.workbookRels, 'Relationship', {
			Id: relId,
			Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
			Target: path.posix.relative(this.prefix, arcName)
		});
		for (const definedName of newDefinedNames)
			this.workbook.find('definedNames').append(definedName);
		this.archive.file('[Content_Types].xml', etree.tostring(this.contentTypes));
		this._rebuild();
		return this;
	}

	/**
	 * Persist workbook changes without changing relationship IDs or their targets.
	 */
	_rebuild() {
		this.workbook.findall('sheets/sheet').forEach((item, index) => {
			item.attrib.sheetId = (index + 1).toString();
		});
		this.archive.file(this.prefix + '/_rels/' + path.basename(this.workbookPath) + '.rels', etree.tostring(this.workbookRels));
		this.archive.file(this.workbookPath, etree.tostring(this.workbook));
		this.sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);
	}

	/**
	 * Load document from file or from buffer (auto-detection)
	 */
	async load(value) {
		if(typeof value === 'string') 
		{
			await this.loadFile(value);
		}
		else
		{
			await this.loadTemplate(value);
		}
	}

	/**
	 * Load a .xlsx file from filename
	 */
	async loadFile(filename) {
		const buffer = fs.readFileSync(filename);

		await this.loadTemplate(buffer);
	}

	/**
	 * Load a .xlsx file from a byte array.
	 */
	async loadTemplate(data) {
		if (Buffer.isBuffer(data))
			data = data.toString('binary');
		
		this.archive = await JSZip.loadAsync(data, { base64: false, checkCRC32: true });

		// Load relationships
		let parseTextRels = await this.archive.file('_rels/.rels').async('string');
		var rels = etree.parse(parseTextRels).getroot(), workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

		this.workbookPath = workbookPath;
		this.prefix = path.dirname(workbookPath);

		let parseTextWB = await this.archive.file(workbookPath).async('string');
		this.workbook = etree.parse(parseTextWB).getroot();

		let parseTextWBRels = await this.archive.file(this.prefix + "/" + '_rels' + "/" + path.basename(workbookPath) + '.rels').async('string');
		this.workbookRels = etree.parse(parseTextWBRels).getroot();
		this.sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);
		this.calChainRel = this.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']");

		if (this.calChainRel)
			this.calcChainPath = this.prefix + "/" + this.calChainRel.attrib.Target;

		this.sharedStringsPath = "";
		this.sharedStrings = [];
		this.sharedStringsLookup = {};

		var sharedStringsRel = this.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']");
		if (sharedStringsRel !== null) {
			this.sharedStringsPath = this.prefix + "/" + sharedStringsRel.attrib.Target;
			let sharedStringsFile = this.archive.file(this.sharedStringsPath);

			if (sharedStringsFile === null)
				throw new Error('Shared strings file not found: ' + this.sharedStringsPath);

			let parseTextSharedStr = await sharedStringsFile.async('string');

			etree.parse(parseTextSharedStr).getroot().findall('si').forEach(si => {
				var t = { text: '' };

				si.findall('t').forEach(tmp => {
					t.text += tmp.text;
				});

				si.findall('r/t').forEach(tmp => {
					t.text += tmp.text;
				});

				this.sharedStrings.push(t.text);
				this.sharedStringsLookup[t.text] = this.sharedStrings.length - 1;
			});
		}

		let parseTextContentTypes = await this.archive.file('[Content_Types].xml').async('string');

		this.contentTypes = etree.parse(parseTextContentTypes).getroot();

		var jpgType = this.contentTypes.find('Default[@Extension="jpg"]');

		if (jpgType === null)
			etree.SubElement(this.contentTypes, 'Default', { 'ContentType': 'image/png', 'Extension': 'jpg' });
	}

	/**
	 * Interpolate values for all the sheets using the given substitutions
	 * (an object).
	 */
	async processAll(substitutions) {
		await this.substituteAll(substitutions);
	}

	/**
	 * Interpolate values for all the sheets using the given substitutions
	 * (an object).
	 */
	async substituteAll(substitutions) {
		const sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);

		for(const sheet of sheets) {
			await this.substitute(sheet.id, substitutions);
		}
	}

	/**
	 * Interpolate values for the sheet with the given number (1-based) or
	 * name (if a string) using the given substitutions (an object).
	 */
	async process(sheetName, substitutions) {
		await this.substitute(sheetName, substitutions);
	}

	/**
	 * Interpolate values for the sheet with the given number (1-based) or
	 * name (if a string) using the given substitutions (an object).
	 */
	async substitute(sheetName, substitutions) {
		var sheet = await this.loadSheet(sheetName);
		this.sheet = sheet;

		var dimension = sheet.root.find('dimension');
		var sheetData = sheet.root.find('sheetData');
		var currentRow = null;
		var totalRowsInserted = 0;
		var totalColumnsInserted = 0;
		var namedTables = await this.loadTables(sheet.root, sheet.filename);
		var rows = [];
		var drawing = null;

		var rels = await this.loadSheetRels(sheet.filename);

		for (const row of sheetData.findall('row')) {
			row.attrib.r = currentRow = this.getCurrentRow(row, totalRowsInserted);
			rows.push(row);

			var cells = [], cellsInserted = 0, newTableRows = [], cellsForsubstituteTable = []; // Contains all the row cells when substitute tables

			for (const cell of row.findall('c')) {
				var appendCell = true;

				cell.attrib.r = this.getCurrentCell(cell, currentRow, cellsInserted);

				// Look for shared and inline strings that may contain placeholders.
				if (cell.attrib.t === 's' || cell.attrib.t === 'inlineStr') {
					var string;

					if (cell.attrib.t === 's') {
						var cellValue = cell.find('v'), stringIndex = parseInt(cellValue.text, 10);
						string = this.sharedStrings[stringIndex];
					} else {
						string = this.getInlineString(cell);
					}

					if (string === undefined)
						continue;

					// Loop over placeholders
					for (const placeholder of this.extractPlaceholders(string)) {
						// Only substitute things for which we have a substitution
						var substitution = _get(substitutions, placeholder.name, ''), newCellsInserted = 0;
						
						if (placeholder.full && placeholder.type === 'table' && Array.isArray(substitution)) {
							if (placeholder.subType === 'image' && drawing == null) {
								if (rels) {
									drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
								} else {
									console.log('Need to implement initRels. Or init this with Excel');
								}
							}

							cellsForsubstituteTable.push(cell); // When substitute table, push (all) the cell 
							
							newCellsInserted = this.substituteTable(
								row, newTableRows,
								cells, cell,
								namedTables, substitution, placeholder.key,
								placeholder, drawing
							);

							// don't double-insert cells
							// this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
							if (newCellsInserted !== 0 || substitution.length) {
								if (substitution.length === 1)
									appendCell = true;
								
								if (Array.isArray(substitution[0][placeholder.key]))
									appendCell = false;
							}

							// Did we insert new columns (array values)?
							if (newCellsInserted !== 0) {
								cellsInserted += newCellsInserted;
								this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
							}
						} else if (placeholder.full && placeholder.type === 'normal' && Array.isArray(substitution)) {
							appendCell = false; // don't double-insert cells

							newCellsInserted = this.substituteArray(cells, cell, substitution);

							if (newCellsInserted !== 0) {
								cellsInserted += newCellsInserted;
								this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
							}
						} else if (placeholder.full && placeholder.type === 'image') {
							if (rels != null) {
								if (drawing == null)
									drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
								
								string = this.substituteImage(cell, string, placeholder, substitution, drawing);
							} else {
								console.log('Need to implement initRels. Or init this with Excel');
							}
						} else {
							if (placeholder.key)
								substitution = _get(substitutions, placeholder.name + '.' + placeholder.key);

							string = this.substituteScalar(cell, string, placeholder, substitution);
						}
					}
				}

				// if we are inserting columns, we may not want to keep the original cell anymore
				if (appendCell)
					cells.push(cell);
			}

			// We may have inserted columns, so re-build the children of the row
			this.replaceChildren(row, cells);

			// Update row spans attribute
			if (cellsInserted !== 0) {
				this.updateRowSpan(row, cellsInserted);

				if (cellsInserted > totalColumnsInserted)
					totalColumnsInserted = cellsInserted;
			}

			// Add newly inserted rows
			if (newTableRows.length > 0) {
				// Move images for each subsitute array if option is active
				if (this.option['moveImages'] && rels) {
					if (drawing == null) {
						// Maybe we can load drawing at the begining of function and remove all the this.loadDrawing() along the function ?
						// If we make this, we create all the time the drawing file (like rels file at this moment)
						drawing = await this.loadDrawing(sheet.root, sheet.filename, rels.root);
					}

					if (drawing != null)
						this.moveAllImages(drawing, row.attrib.r, newTableRows.length);
				}

				// Filter all the cellsForsubstituteTable cell with the 'row' cell
				var cellsOverTable = row.findall('c').filter(cell => !cellsForsubstituteTable.includes(cell));

				newTableRows.forEach(row => {
					if (this.option && this.option.subsituteAllTableRow) {
						// I happend the other cell in substitute new table rows
						cellsOverTable.forEach(cellOverTable => {
							var newCell = this.cloneElement(cellOverTable);

							newCell.attrib.r = this.joinRef({
								row: row.attrib.r,
								col: this.splitRef(newCell.attrib.r).col
							});
							
							row.append(newCell);
						});

						// I sort the cell in the new row
						var newSortRow = row.findall('c').sort((a, b) => {
							var colA = this.splitRef(a.attrib.r).col;
							var colB = this.splitRef(b.attrib.r).col;

							return this.charToNum(colA) - this.charToNum(colB);
						});

						// And I replace the cell
						this.replaceChildren(row, newSortRow);
					}

					rows.push(row);

					++totalRowsInserted;
				});

				this.pushDown(this.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
			}
		}

		// We may have inserted rows, so re-build the children of the sheetData
		this.replaceChildren(sheetData, rows);

		// Update placeholders in table column headers
		this.substituteTableColumnHeaders(namedTables, substitutions);

		// Update placeholders in hyperlinks
		await this.substituteHyperlinks(rels, substitutions);

		// Update <dimension /> if we added rows or columns
		if (dimension) {
			if (totalRowsInserted > 0 || totalColumnsInserted > 0) {
				var dimensionRange = this.splitRange(dimension.attrib.ref), dimensionEndRef = this.splitRef(dimensionRange.end || dimensionRange.start);

				dimensionEndRef.row += totalRowsInserted;
				dimensionEndRef.col = this.numToChar(this.charToNum(dimensionEndRef.col) + totalColumnsInserted);
				dimensionRange.end = this.joinRef(dimensionEndRef);

				dimension.attrib.ref = this.joinRange(dimensionRange);
			}
		}

		//Here we are forcing the values in formulas to be recalculated
		// existing as well as just substituted
		sheetData.findall('row').forEach(row => {
			row.findall('c').forEach(cell => {
				var formulas = cell.findall('f');

				if (formulas && formulas.length > 0) {
					cell.findall('v').forEach(v => {
						cell.remove(v);
					});
				}
			});
		});

		// Write back the modified XML trees
		this.archive.file(sheet.filename, etree.tostring(sheet.root));
		this.archive.file(this.workbookPath, etree.tostring(this.workbook));

		if (rels)
			this.archive.file(rels.filename, etree.tostring(rels.root));

		this.archive.file('[Content_Types].xml', etree.tostring(this.contentTypes));

		// Remove calc chain - Excel will re-build, and we may have moved some formulae
		if (this.calcChainPath && this.archive.file(this.calcChainPath))
			this.archive.remove(this.calcChainPath);

		await this.writeSharedStrings();

		this.writeTables(namedTables);
		this.writeDrawing(drawing);
	}

	/**
	 * Build a new binary .xlsx file
	 * @param {import('jszip').JSZipGeneratorOptions} [options]
	 */
	async build(options) {
		if (!options)
			options = { type: 'uint8array' };

		return await this.archive.generateAsync(options);
	}

	// Helpers
	// Write back the new shared strings list
	async writeSharedStrings() {
		if (!this.sharedStringsPath)
			return;

		let sharedStringsFile = this.archive.file(this.sharedStringsPath);
		if (sharedStringsFile === null)
			throw new Error('Shared strings file not found: ' + this.sharedStringsPath);

		let parseTextSharedStrPath = await sharedStringsFile.async('string');
		var root = etree.parse(parseTextSharedStrPath).getroot(), children = root.getchildren();

		root.delSlice(0, children.length);

		this.sharedStrings.forEach(string => {
			var si = new etree.Element('si'), t = new etree.Element('t');

			t.text = string;
			si.append(t);
			root.append(si);
		});

		root.attrib.count = this.sharedStrings.length;
		root.attrib.uniqueCount = this.sharedStrings.length;

		this.archive.file(this.sharedStringsPath, etree.tostring(root));
	}
	
	// Add a new shared string
	addSharedString(s) {
		var idx = this.sharedStrings.length;

		this.sharedStrings.push(s);
		this.sharedStringsLookup[s] = idx;

		return idx;
	}

	// Get the text from an inline string cell.
	getInlineString(cell) {
		var inlineString = cell.find('is');

		if (inlineString === null)
			return '';

		var string = '';

		inlineString.findall('t').forEach(tmp => {
			string += tmp.text || '';
		});

		inlineString.findall('r/t').forEach(tmp => {
			string += tmp.text || '';
		});

		return string;
	}

	// Get the number of a shared string, adding a new one if necessary.
	stringIndex(s) {
		let idx = this.sharedStringsLookup[s];

		if (idx === undefined)
			idx = this.addSharedString(s);

		return idx;
	}

	// Replace a shared string with a new one at the same index. Return the
	// index.
	replaceString(oldString, newString) {
		var idx = this.sharedStringsLookup[oldString];

		if (idx === undefined) {
			idx = this.addSharedString(newString);
		} else {
			this.sharedStrings[idx] = newString;
			delete this.sharedStringsLookup[oldString];
			this.sharedStringsLookup[newString] = idx;
		}

		return idx;
	}

	// Get a list of sheet ids, names and filenames
	loadSheets(prefix, workbook, workbookRels) {
		var sheets = [];

		workbook.findall('sheets/sheet').forEach(sheet => {
			var sheetId = sheet.attrib.sheetId, relId = sheet.attrib['r:id'], relationship = workbookRels.find("Relationship[@Id='" + relId + "']"), filename = resolvePartPath(prefix + '/workbook.xml', relationship.attrib.Target);

			sheets.push({
				id: parseInt(sheetId, 10),
				name: sheet.attrib.name,
				filename: filename
			});
		});

		return sheets;
	}

	// Get sheet a sheet, including filename and name
	async loadSheet(sheet) {
		var info = null;

		for (var i = 0; i < this.sheets.length; ++i) {
			if ((typeof (sheet) === 'number' && this.sheets[i].id === sheet) || (this.sheets[i].name === sheet)) {
				info = this.sheets[i];
				break;
			}
		}

		if (info === null && (typeof (sheet) === 'number')) {
			//Get the sheet that corresponds to the 0 based index if the id does not work
			info = this.sheets[sheet - 1];
		}

		if (info == null)
			throw new Error('Sheet ' + sheet + ' not found');

		const sheetFile = this.archive.file(info.filename);
		if (sheetFile === null)
			throw new Error('Sheet file not found: ' + info.filename);
		let parseTextFileName = await sheetFile.async('string');

		return {
			filename: info.filename,
			name: info.name,
			id: info.id,
			root: etree.parse(parseTextFileName).getroot()
		};
	}

	//Load rels for a sheetName
	async loadSheetRels(sheetFilename) {
		var sheetDirectory = path.dirname(sheetFilename), sheetName = path.basename(sheetFilename), relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/'), relsFile = this.archive.file(relsFilename);
		
		if (relsFile === null)
			return this.initSheetRels(sheetFilename);

		let parseTextRels = await relsFile.async('string');

		var rels = { filename: relsFilename, root: etree.parse(parseTextRels).getroot() };

		return rels;
	}

	initSheetRels(sheetFilename) {
		var sheetDirectory = path.dirname(sheetFilename), sheetName = path.basename(sheetFilename), relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
		var element = etree.Element;
		var ElementTree = etree.ElementTree;
		var root = element('Relationships');
		root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
		var relsEtree = new ElementTree(root);
		var rels = { filename: relsFilename, root: relsEtree.getroot() };
		return rels;
	}
	
	// Load Drawing file
	async loadDrawing(sheet, sheetFilename, rels) {
		var drawing = { filename: '', root: null, relFilename: '', relRoot: null };
		var drawingPart = sheet.find('drawing');
		
		if (drawingPart === null) {
			drawing = this.initDrawing(sheet, rels);
			return drawing;
		}

		var relationshipId = drawingPart.attrib['r:id'];
		var drawingRelationship = rels.find("Relationship[@Id='" + relationshipId + "']");
		if (drawingRelationship === null)
			throw new Error('Drawing relationship not found: ' + relationshipId);

		var drawingFilename = resolvePartPath(sheetFilename, drawingRelationship.attrib.Target);
		var drawingFile = this.archive.file(drawingFilename);
		if (drawingFile === null)
			throw new Error('Drawing file not found: ' + drawingFilename);

		let parseTextDrawingFilename = await drawingFile.async('string');
		var drawingTree = etree.parse(parseTextDrawingFilename);
		
		drawing.filename = drawingFilename;
		drawing.root = drawingTree.getroot();
		drawing.relFilename = path.posix.join(path.posix.dirname(drawingFilename), '_rels', path.posix.basename(drawingFilename) + '.rels');
		
		var drawingRelFile = this.archive.file(drawing.relFilename);
		if (drawingRelFile === null) {
			drawing.relRoot = etree.Element('Relationships');
			drawing.relRoot.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
		} else {
			let parseTextRelFilename = await drawingRelFile.async('string');
			drawing.relRoot = etree.parse(parseTextRelFilename).getroot();
		}
		
		return drawing;
	}

	addContentType(partName, contentType) {
		etree.SubElement(this.contentTypes, 'Override', { 'ContentType': contentType, 'PartName': partName });
	}

	initDrawing(sheet, rels) {
		var maxId = this.findMaxId(rels, 'Relationship', 'Id', /rId(\d*)/);
		var rel = etree.SubElement(rels, 'Relationship');

		sheet.insert(sheet._children.length, etree.Element('drawing', { 'r:id': 'rId' + maxId }));

		rel.set('Id', 'rId' + maxId);
		rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');

		var drawing = {};
		var drawingFilename = 'drawing' + this.findMaxFileId(/xl\/drawings\/drawing\d*\.xml/, /drawing(\d*)\.xml/) + '.xml';

		rel.set('Target', '../drawings/' + drawingFilename);

		drawing.root = etree.Element('xdr:wsDr');
		drawing.root.set('xmlns:xdr', "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		drawing.root.set('xmlns:a', "http://schemas.openxmlformats.org/drawingml/2006/main");
		drawing.filename = 'xl/drawings/' + drawingFilename;
		drawing.relFilename = 'xl/drawings/_rels/' + drawingFilename + '.rels';
		drawing.relRoot = etree.Element('Relationships');
		drawing.relRoot.set('xmlns', "http://schemas.openxmlformats.org/package/2006/relationships");

		this.addContentType('/' + drawing.filename, 'application/vnd.openxmlformats-officedocument.drawing+xml');

		return drawing;
	}

	// Write Drawing file
	writeDrawing(drawing) {
		if (drawing !== null) {
			this.archive.file(drawing.filename, etree.tostring(drawing.root));
			this.archive.file(drawing.relFilename, etree.tostring(drawing.relRoot));
		}
	}

	// Move all images after fromRow of nbRow row
	moveAllImages(drawing, fromRow, nbRow) {
		drawing.root.getchildren().forEach(drawElement => {
			if (drawElement.tag == 'xdr:twoCellAnchor')
				this._moveTwoCellAnchor(drawElement, fromRow, nbRow);
		
			// TODO : make the other tags image
		});
	}

	// Move TwoCellAnchor tag images after fromRow of nbRow row
	_moveTwoCellAnchor(drawingElement, fromRow, nbRow) {
		var _moveImage = (drawingElement, fromRow, nbRow) => {
			var from = Number.parseInt(drawingElement.find('xdr:from').find('xdr:row').text, 10) + Number.parseInt(nbRow, 10);
			drawingElement.find('xdr:from').find('xdr:row').text = from;

			var to = Number.parseInt(drawingElement.find('xdr:to').find('xdr:row').text, 10) + Number.parseInt(nbRow, 10);
			drawingElement.find('xdr:to').find('xdr:row').text = to;
		};

		if (this.option['moveSameLineImages']) {
			if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 >= fromRow)
				_moveImage(drawingElement, fromRow, nbRow);
		} else {
			if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text) + 1 > fromRow) 
				_moveImage(drawingElement, fromRow, nbRow);
		}
	}

	// Load tables for a given sheet
	async loadTables(sheet, sheetFilename) {
		let sheetDirectory = path.dirname(sheetFilename); 
		let sheetName = path.basename(sheetFilename);
		let relsFilename = sheetDirectory + '/' + '_rels' + '/' + sheetName + '.rels'; 
		let relsFile = this.archive.file(relsFilename); 

		let tables = [];

		if (relsFile === null)
			return tables;
	
		let parseTextRelsFile = await relsFile.async('string');
		let rels = etree.parse(parseTextRelsFile).getroot();

		for(const tablePart of sheet.findall('tableParts/tablePart'))
		{
			let relationshipId = tablePart.attrib['r:id'];
			let target = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target;
			let tableFilename = target.replace('..', this.prefix);

			let parseTextTableFilename = await this.archive.file(tableFilename).async('string');
			let tableTree = etree.parse(parseTextTableFilename);

			tables.push({ filename: tableFilename, root: tableTree.getroot() });
		}

		return tables;
	}

	// Write back possibly-modified tables
	writeTables(tables) {
		tables.forEach(namedTable => {
			this.archive.file(namedTable.filename, etree.tostring(namedTable.root));
		});
	}

	//Perform substitution in hyperlinks
	async substituteHyperlinks(rels, substitutions) {
		if (rels === null)
			return;
		
		const relationships = rels.root._children;

		relationships.forEach(relationship => {
			if (relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {

				let target = relationship.attrib.Target;

				//Double-decode due to excel double encoding url placeholders
				target = decodeURI(decodeURI(target));

				this.extractPlaceholders(target).forEach(placeholder => {
					const substitution = substitutions[placeholder.name];

					if (substitution === undefined)
						return;
					
					target = target.replace(placeholder.placeholder, this.stringify(substitution));

					relationship.attrib.Target = encodeURI(target);
				});
			}
		});
	}

	// Perform substitution in table headers
	substituteTableColumnHeaders(tables, substitutions) {
		tables.forEach(table => {
			var root = table.root, columns = root.find('tableColumns'), autoFilter = root.find('autoFilter'), tableRange = this.splitRange(root.attrib.ref), idx = 0, inserted = 0, newColumns = [];

			columns.findall('tableColumn').forEach(col => {
				++idx;
				col.attrib.id = Number(idx).toString();
				
				newColumns.push(col);

				var name = col.attrib.name;

				this.extractPlaceholders(name).forEach(placeholder => {
					var substitution = substitutions[placeholder.name];

					if (substitution === undefined)
						return;

					// Array -> new columns
					if (placeholder.full && placeholder.type === 'normal' && Array.isArray(substitution)) {
						substitution.forEach((element, i) => {
							var newCol = col;

							if (i > 0) {
								newCol = this.cloneElement(newCol);
								newCol.attrib.id = Number(++idx).toString();

								newColumns.push(newCol);

								++inserted;

								tableRange.end = this.nextCol(tableRange.end);
							}

							newCol.attrib.name = this.stringify(element);
						});
						// Normal placeholder
					} else {
						name = name.replace(placeholder.placeholder, this.stringify(substitution));
						col.attrib.name = name;
					}
				});
			});

			this.replaceChildren(columns, newColumns);

			// Update range if we inserted columns
			if (inserted > 0) {
				columns.attrib.count = Number(idx).toString();
				root.attrib.ref = this.joinRange(tableRange);

				if (autoFilter !== null) {
					// XXX: This is a simplification that may stomp on some configurations
					autoFilter.attrib.ref = this.joinRange(tableRange);
				}
			}

			//update ranges for totalsRowCount
			var tableRoot = table.root, tableRange = this.splitRange(tableRoot.attrib.ref), tableStart = this.splitRef(tableRange.start), tableEnd = this.splitRef(tableRange.end);

			if (tableRoot.attrib.totalsRowCount) {
				var autoFilter = tableRoot.find('autoFilter');

				if (autoFilter !== null) {
					autoFilter.attrib.ref = this.joinRange({
						start: this.joinRef(tableStart),
						end: this.joinRef(tableEnd),
					});
				}

				++tableEnd.row;

				tableRoot.attrib.ref = this.joinRange({
					start: this.joinRef(tableStart),
					end: this.joinRef(tableEnd),
				});
			}
		});
	}

	// Return a list of tokens that may exist in the string.
	// Keys are: `placeholder` (the full placeholder, including the `${}`
	// delineators), `name` (the name part of the token), `key` (the object key
	// for `table` tokens), `full` (boolean indicating whether this placeholder
	// is the entirety of the string) and `type` (one of `table` or `cell`)
	extractPlaceholders(text) {
		const re = /\${(?:(.+?):)?([^{}]+?)(?:\.(.+?))?(?::(.+?))??}/g;

		let match = null, matches = [];

		while ((match = re.exec(text)) !== null) {
			matches.push({
				placeholder: match[0],
				type: match[1] || 'normal',
				name: match[2],
				key: match[3],
				subType: match[4],
				full: match[0].length === text.length
			});
		}

		return matches;
	}

	// Split a reference into an object with keys `row` and `col` and,
	// optionally, `table`, `rowAbsolute` and `colAbsolute`.
	splitRef(ref) {
		const match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)?(\$)?([0-9]+)/);

		return {
			table: match && match[1] || null,
			colAbsolute: Boolean(match && match[2]),
			col: match && match[3] || '',
			rowAbsolute: Boolean(match && match[4]),
			row: parseInt(match && match[5], 10)
		};
	}

	// Join an object with keys `row` and `col` into a single reference string
	joinRef(ref) {
		return (ref.table ? ref.table + '!' : '') +
			(ref.colAbsolute ? '$' : '') +
			ref.col.toUpperCase() +
			(ref.rowAbsolute ? '$' : '') +
			Number(ref.row).toString();
	}

	// Get the next column's cell reference given a reference like "B2".
	nextCol(ref) {
		ref = ref.toUpperCase();

		return ref.replace(/[A-Z]+/, match => {
			return this.numToChar(this.charToNum(match) + 1);
		});
	}

	// Get the next row's cell reference given a reference like "B2".
	nextRow(ref) {
		ref = ref.toUpperCase();

		return ref.replace(/[0-9]+/, match => {
			return (parseInt(match, 10) + 1).toString();
		});
	}

	// Turn a reference like "AA" into a number like 27
	charToNum(str) {
		var num = 0;

		for (var idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
			var thisChar = str.charCodeAt(idx) - 64, // A -> 1; B -> 2; ... Z->26
				multiplier = Math.pow(26, iteration);

			num += multiplier * thisChar;
		}

		return num;
	}

	// Turn a number like 27 into a reference like "AA"
	numToChar(num) {
		var str = '';

		for (var i = 0; num > 0; ++i) {
			var remainder = num % 26, charCode = remainder + 64;

			num = (num - remainder) / 26;

			// Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
			if (remainder === 0) { // 26 -> Z
				charCode = 90;
				--num;
			}

			str = String.fromCharCode(charCode) + str;
		}

		return str;
	}

	// Is ref a range?
	isRange(ref) {
		return ref.indexOf(':') !== -1;
	}

	// Is ref inside the table defined by startRef and endRef?
	isWithin(ref, startRef, endRef) {
		var start = this.splitRef(startRef), end = this.splitRef(endRef), target = this.splitRef(ref);

		start.col = this.charToNum(start.col);
		end.col = this.charToNum(end.col);
		target.col = this.charToNum(target.col);

		return (
			start.row <= target.row && target.row <= end.row &&
			start.col <= target.col && target.col <= end.col
		);
	}

	// Turn a value of any type into a string
	stringify(value) {
		if (value instanceof Date) {
			//In Excel date is a number of days since 01/01/1900
			//timestamp in ms    to days      + number of days from 1900 to 1970
			return Number((value.getTime() / (1000 * 60 * 60 * 24)) + 25569);
		} else if (typeof (value) === 'number' || typeof (value) === 'boolean') {
			return Number(value).toString();
		} else if (typeof (value) === 'string') {
			return String(value).toString();
		}

		return '';
	}

	// Insert a substitution value into a cell (c tag)
	insertCellValue(cell, substitution) {
		var cellValue = cell.find('v'), stringified = this.stringify(substitution), isInlineString = cell.attrib.t === 'inlineStr';
		var isFormula = typeof substitution === 'string' && substitution[0] === '=';

		if (isInlineString) {
			var inlineString = cell.find('is');

			if (inlineString !== null)
				cell.remove(inlineString);

			// Null and missing values stringify to empty text, just like shared strings.
			if (!isFormula && typeof substitution !== 'number' && typeof substitution !== 'boolean' && !(substitution instanceof Date)) {
				if (cellValue !== null)
					cell.remove(cellValue);

				inlineString = etree.SubElement(cell, 'is');
				var text = etree.SubElement(inlineString, 't', { 'xml:space': 'preserve' });
				text.text = stringified;

				return stringified;
			}
		}

		if (cellValue === null) {
			cellValue = new etree.Element('v');
			cell.append(cellValue);
		}

		if (isFormula) {
			//substitution, started with '=' is a formula substitution
			var formula = new etree.Element('f');
			formula.text = substitution.substr(1);
			cell.insert(1, formula);
			delete cell.attrib.t; //cellValue will be deleted later

			return formula.text;
		}

		if (typeof (substitution) === 'number' || substitution instanceof Date) {
			delete cell.attrib.t;
			cellValue.text = stringified;
		} else if (typeof (substitution) === 'boolean') {
			cell.attrib.t = 'b';
			cellValue.text = stringified;
		} else {
			cell.attrib.t = 's';
			cellValue.text = Number(this.stringIndex(stringified)).toString();
		}

		return stringified;
	}

	// Perform substitution of a single value
	substituteScalar(cell, string, placeholder, substitution) {
		if (placeholder.full)
			return this.insertCellValue(cell, substitution);
		
		let newString = string.replace(placeholder.placeholder, this.stringify(substitution));

		if (cell.attrib.t !== 'inlineStr')
			cell.attrib.t = 's';
		
		return this.insertCellValue(cell, newString);
	}

	// Perform a columns substitution from an array
	substituteArray(cells, cell, substitution) {
		let newCellsInserted = -1; // we technically delete one before we start adding back
		let	currentCell = cell.attrib.r;

		// add a cell for each element in the list
		substitution.forEach(element => {
			++newCellsInserted;

			if (newCellsInserted > 0)
				currentCell = this.nextCol(currentCell);

			let newCell = this.cloneElement(cell);

			this.insertCellValue(newCell, element);

			newCell.attrib.r = currentCell;

			cells.push(newCell);
		});

		return newCellsInserted;
	}

	// Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
	// Returns total number of new cells inserted on the original row.
	substituteTable(row, newTableRows, cells, cell, namedTables, substitution, key, placeholder, drawing) {
		var newCellsInserted = 0; // on the original row

		// if no elements, blank the cell, but don't delete it
		if (substitution.length === 0) {
			delete cell.attrib.t;

			this.replaceChildren(cell, []);
		} else {
			var parentTables = namedTables.filter(namedTable => {
				var range = this.splitRange(namedTable.root.attrib.ref);

				return this.isWithin(cell.attrib.r, range.start, range.end);
			});
		
			// Credit: kennycreeper - Fix merge cells style
			const mergeCell = this.sheet.root.findall('mergeCells/mergeCell').find(c => this.splitRange(c.attrib.ref).start === cell.attrib.r);

			// Later rows must start from the template, before the first value changes its type.
			const templateCell = this.cloneElement(cell);

			substitution.forEach((element, idx) => {
				var newRow, newCell, newCellsInsertedOnNewRow = 0, newCells = [], value = _get(element, key, '');

				if (idx === 0) { // insert in the row where the placeholders are
					if (Array.isArray(value)) {
						newCellsInserted = this.substituteArray(cells, cell, value);
					} else if (placeholder.subType == 'image' && value != "") {
						this.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
					} else {
						this.insertCellValue(cell, value);
					}

				} else { // insert new rows (or reuse rows just inserted)
					// Do we have an existing row to use? If not, create one.
					if ((idx - 1) < newTableRows.length) {
						newRow = newTableRows[idx - 1];
					} else {
						newRow = this.cloneElement(row, false);
						newRow.attrib.r = this.getCurrentRow(row, newTableRows.length + 1);
						
						newTableRows.push(newRow);
					}

					// Create a new cell
					newCell = this.cloneElement(templateCell);

					newCell.attrib.r = this.joinRef({
						row: newRow.attrib.r,
						col: this.splitRef(newCell.attrib.r).col
					});

					if (Array.isArray(value)) {
						newCellsInsertedOnNewRow = this.substituteArray(newCells, newCell, value);

						// Add each of the new cells created by substituteArray()
						newCells.forEach(newCell => {
							newRow.append(newCell);
						});

						this.updateRowSpan(newRow, newCellsInsertedOnNewRow);
					} else if (placeholder.subType == 'image' && value != '') {
						this.substituteImage(newCell, placeholder.placeholder, placeholder, value, drawing);
						newRow.append(newCell);
					} else {
						this.insertCellValue(newCell, value);

						// Add the cell that previously held the placeholder
						newRow.append(newCell);

						// Credit: kennycreeper - Fix merge cells style, mirror-ru: some fixes
						if (mergeCell) {
							let mergeRange  = this.splitRange(mergeCell.attrib.ref);
							let mergeStart  = this.splitRef(mergeRange.start);
							let mergeEnd    = this.splitRef(mergeRange.end);
							let templateRow = this.sheet.root.findall('sheetData/row').find(r => r.attrib.r === mergeStart.row);	

							for (let colNum = this.charToNum(mergeStart.col); colNum < this.charToNum(mergeEnd.col); colNum++) {
								const templateCell = templateRow.find(`c[@r="${this.numToChar(colNum + 1)}${mergeStart.row}"]`);

								if(templateCell) {
									const cell = this.cloneElement(templateCell);
									
									cell.attrib.r = this.joinRef({ row: newRow.attrib.r, col: this.numToChar(colNum + 1) });
								
								 	newRow.append(cell);
								}
							}
						}
					}

					// expand named table range if necessary
					parentTables.forEach(namedTable => {
						var tableRoot = namedTable.root; 
						var autoFilter = tableRoot.find('autoFilter'); 
						var range = this.splitRange(tableRoot.attrib.ref);

						if (!this.isWithin(newCell.attrib.r, range.start, range.end)) {
							range.end = this.nextRow(range.end);
							tableRoot.attrib.ref = this.joinRange(range);

							if (autoFilter !== null) {
								// XXX: This is a simplification that may stomp on some configurations
								autoFilter.attrib.ref = tableRoot.attrib.ref;
							}
						}
					});
				}
			});
		}

		return newCellsInserted;
	}

	substituteImage(cell, string, placeholder, substitution, drawing) {
		this.substituteScalar(cell, string, placeholder, '');

		if (substitution == null || substitution == "") {
			// TODO : @kant2002 if image is null or empty string in user substitution data, throw an error or not ?
			// If yes, remove this test.
			return true;
		}

		// get max refid
		// update rel file.
		var maxId = this.findMaxId(drawing.relRoot, 'Relationship', 'Id', /rId(\d*)/);
		var maxFildId = this.findMaxFileId(/xl\/media\/image\d*.jpg/, /image(\d*)\.jpg/);
		var rel = etree.SubElement(drawing.relRoot, 'Relationship');

		rel.set('Id', 'rId' + maxId);
		rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
		rel.set('Target', '../media/image' + maxFildId + '.jpg');

		function toArrayBuffer(buffer) {
			var ab = new ArrayBuffer(buffer.length);
			var view = new Uint8Array(ab);

			for (var i = 0; i < buffer.length; ++i) {
				view[i] = buffer[i];
			}

			return ab;
		};

		try {
			substitution = this.imageToBuffer(substitution);
		} catch (error) {
			if (this.option && this.option.handleImageError && typeof this.option.handleImageError === 'function') {
				this.option.handleImageError(substitution, error);
			} else {
				throw error;
			}
		}

		// put image to media.
		this.archive.file('xl/media/image' + maxFildId + '.jpg', toArrayBuffer(substitution), { binary: true, base64: false });
		var dimension = sizeOf(substitution);
		var imageWidth = this.pixelsToEMUs(dimension.width);
		var imageHeight = this.pixelsToEMUs(dimension.height);
		// var sheet = await this.loadSheet(this.substitueSheetName);
		var imageInMergeCell = false;

		this.sheet.root.findall('mergeCells/mergeCell').forEach(mergeCell => {
			// If image is in merge cell, fit the image
			if (this.cellInMergeCells(cell, mergeCell)) {
				var mergeCellWidth = this.getWidthMergeCell(mergeCell, this.sheet);
				var mergeCellHeight = this.getHeightMergeCell(mergeCell, this.sheet);
				var mergeWidthEmus = this.columnWidthToEMUs(mergeCellWidth);
				var mergeHeightEmus = this.rowHeightToEMUs(mergeCellHeight);
				// Maybe we can add an option for fit image to mergecell if image is more little. Not by default
				/*if (imageWidth <= mergeWidthEmus && imageHeight <= mergeHeightEmus) {
					// Image as more little than the merge cell
					imageWidth = mergeWidthEmus;
					imageHeight = mergeHeightEmus;
				}*/
				var widthRate = imageWidth / mergeWidthEmus;
				var heightRate = imageHeight / mergeHeightEmus;
				if (widthRate > heightRate) {
					imageWidth = Math.floor(imageWidth / widthRate);
					imageHeight = Math.floor(imageHeight / widthRate);
				} else {
					imageWidth = Math.floor(imageWidth / heightRate);
					imageHeight = Math.floor(imageHeight / heightRate);
				}
				imageInMergeCell = true;
			}
		});

		if (imageInMergeCell == false) {
			var ratio = 100;
			if (this.option && this.option.imageRatio) {
				ratio = this.option.imageRatio;
			}
			if (ratio <= 0) {
				ratio = 100;
			}
			imageWidth = Math.floor(imageWidth * ratio / 100);
			imageHeight = Math.floor(imageHeight * ratio / 100);
		}

		var imagePart = etree.SubElement(drawing.root, 'xdr:oneCellAnchor');
		var fromPart = etree.SubElement(imagePart, 'xdr:from');
		var fromCol = etree.SubElement(fromPart, 'xdr:col');

		fromCol.text = (this.charToNum(this.splitRef(cell.attrib.r).col) - 1).toString();

		var fromColOff = etree.SubElement(fromPart, 'xdr:colOff');
		fromColOff.text = '0';

		var fromRow = etree.SubElement(fromPart, 'xdr:row');
		fromRow.text = (this.splitRef(cell.attrib.r).row - 1).toString();

		var fromRowOff = etree.SubElement(fromPart, 'xdr:rowOff');
		fromRowOff.text = '0';
		var extImagePart = etree.SubElement(imagePart, 'xdr:ext', { cx: imageWidth, cy: imageHeight });
		var picNode = etree.SubElement(imagePart, 'xdr:pic');
		var nvPicPr = etree.SubElement(picNode, 'xdr:nvPicPr');
		var cNvPr = etree.SubElement(nvPicPr, 'xdr:cNvPr', { id: maxId, name: 'image_' + maxId, descr: '' });
		var cNvPicPr = etree.SubElement(nvPicPr, 'xdr:cNvPicPr');
		var picLocks = etree.SubElement(cNvPicPr, 'a:picLocks', { noChangeAspect: '1' });
		var blipFill = etree.SubElement(picNode, 'xdr:blipFill');
		var blip = etree.SubElement(blipFill, 'a:blip', {
			"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			"r:embed": "rId" + maxId
		});
		var stretch = etree.SubElement(blipFill, 'a:stretch');
		var fillRect = etree.SubElement(stretch, 'a:fillRect');
		var spPr = etree.SubElement(picNode, 'xdr:spPr');
		var xfrm = etree.SubElement(spPr, 'a:xfrm');
		var off = etree.SubElement(xfrm, 'a:off', { x: "0", y: "0" });
		var ext = etree.SubElement(xfrm, 'a:ext', { cx: imageWidth, cy: imageHeight });
		var prstGeom = etree.SubElement(spPr, 'a:prstGeom', { 'prst': 'rect' });
		var avLst = etree.SubElement(prstGeom, 'a:avLst');
		var clientData = etree.SubElement(imagePart, 'xdr:clientData');

		return true;
	}

	// Clone an element. If `deep` is true, recursively clone children
	cloneElement(element, deep) {
		let newElement = etree.Element(element.tag, element.attrib);

		newElement.text = element.text;
		newElement.tail = element.tail;

		if (deep !== false) {
			element.getchildren().forEach(child => {
				newElement.append(this.cloneElement(child, deep));
			});
		}

		return newElement;
	}

	// Replace all children of `parent` with the nodes in the list `children`
	replaceChildren(parent, children) {
		parent.delSlice(0, parent.len());

		children.forEach(child => {
			parent.append(child);
		});
	}

	// Calculate the current row based on a source row and a number of new rows
	// that have been inserted above
	getCurrentRow(row, rowsInserted) {
		return parseInt(row.attrib.r, 10) + rowsInserted;
	}

	// Calculate the current cell based on asource cell, the current row index,
	// and a number of new cells that have been inserted so far
	getCurrentCell(cell, currentRow, cellsInserted) {
		const colRef = this.splitRef(cell.attrib.r).col; 
		const colNum = this.charToNum(colRef);

		return this.joinRef({ row: currentRow, col: this.numToChar(colNum + cellsInserted) });
	}

	// Adjust the row `spans` attribute by `cellsInserted`
	updateRowSpan(row, cellsInserted) {
		if (cellsInserted !== 0 && row.attrib.spans) {
			var rowSpan = row.attrib.spans.split(':').map(f => { 
				return parseInt(f, 10); 
			});

			rowSpan[1] += cellsInserted;

			row.attrib.spans = rowSpan.join(":");
		}
	}

	// Split a range like "A1:B1" into {start: "A1", end: "B1"}
	splitRange(range) {
		const [ start, end ] = range.split(':');

		return { start, end };
	}

	// Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}
	joinRange(range) {
		return range.start + ':' + range.end;
	}

	// Look for any merged cell or named range definitions to the right of
	// `currentCell` and push right by `numCols`.
	pushRight(workbook, sheet, currentCell, numCols) {
		var cellRef = this.splitRef(currentCell), currentRow = cellRef.row, currentCol = this.charToNum(cellRef.col);

		// Update merged cells on the same row, at a higher column
		sheet.findall('mergeCells/mergeCell').forEach(mergeCell => {
			var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStart = this.splitRef(mergeRange.start), mergeStartCol = this.charToNum(mergeStart.col), mergeEnd = this.splitRef(mergeRange.end), mergeEndCol = this.charToNum(mergeEnd.col);

			if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
				mergeStart.col = this.numToChar(mergeStartCol + numCols);
				mergeEnd.col = this.numToChar(mergeEndCol + numCols);

				mergeCell.attrib.ref = this.joinRange({
					start: this.joinRef(mergeStart),
					end: this.joinRef(mergeEnd),
				});
			}
		});

		// Named cells/ranges
		workbook.findall('definedNames/definedName').forEach(name => {
			var ref = name.text;

			if (this.isRange(ref)) {
				var namedRange = this.splitRange(ref), namedStart = this.splitRef(namedRange.start), namedStartCol = this.charToNum(namedStart.col), namedEnd = this.splitRef(namedRange.end), namedEndCol = this.charToNum(namedEnd.col);

				if (namedStart.row === currentRow && currentCol < namedStartCol) {
					namedStart.col = this.numToChar(namedStartCol + numCols);
					namedEnd.col = this.numToChar(namedEndCol + numCols);

					name.text = this.joinRange({
						start: this.joinRef(namedStart),
						end: this.joinRef(namedEnd),
					});
				}
			} else {
				var namedRef = this.splitRef(ref), namedCol = this.charToNum(namedRef.col);

				if (namedRef.row === currentRow && currentCol < namedCol) {
					namedRef.col = this.numToChar(namedCol + numCols);

					name.text = this.joinRef(namedRef);
				}
			}
		});

		// Hyperlinks are anchored to cells too, so move their refs with inserted columns.
		sheet.findall('hyperlinks/hyperlink').forEach(hyperlink => {
			const range = this.splitRange(hyperlink.attrib.ref);
			const start = this.splitRef(range.start);
			const end = this.splitRef(range.end || range.start);
			if (!numCols || currentRow < start.row || currentRow > end.row ||
				this.charToNum(end.col) <= currentCol)
				return;
			const parent = sheet.find('hyperlinks');
			let position = parent.getchildren().indexOf(hyperlink);
			const appendRange = (from, to) => {
				const copy = this.cloneElement(hyperlink);
				const first = this.joinRef(from), last = this.joinRef(to);
				copy.attrib.ref = first === last ? first : first + ':' + last;
				parent.getchildren().splice(position++, 0, copy);
			};
			parent.remove(hyperlink);
			// Array expansion moves cells in this row only. Split a multi-row anchor
			// so the unaffected rows keep their links at the same coordinates.
			if (start.row < currentRow)
				appendRange(start, { ...end, row: currentRow - 1 });
			const shiftedStart = { ...start, row: currentRow };
			const shiftedEnd = { ...end, row: currentRow };
			for (const ref of [shiftedStart, shiftedEnd]) {
				const col = this.charToNum(ref.col);
				if (col > currentCol)
					ref.col = this.numToChar(col + numCols);
			}
			appendRange(shiftedStart, shiftedEnd);
			if (end.row > currentRow)
				appendRange({ ...start, row: currentRow + 1 }, end);
		});
	}

	// Look for any merged cell, named table or named range definitions below
	// `currentRow` and push down by `numRows` (used when rows are inserted).
	pushDown(workbook, sheet, tables, currentRow, numRows) {
		var mergeCells = sheet.find('mergeCells');

		// Update merged cells below this row
		sheet.findall('mergeCells/mergeCell').forEach(mergeCell => {
			var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStart = this.splitRef(mergeRange.start), mergeEnd = this.splitRef(mergeRange.end);

			if (mergeStart.row > currentRow) {
				mergeStart.row += numRows;
				mergeEnd.row += numRows;

				mergeCell.attrib.ref = this.joinRange({ start: this.joinRef(mergeStart), end: this.joinRef(mergeEnd) });
			}

			//add new merge cell
			if (mergeStart.row == currentRow) {
				for (var i = 1; i <= numRows; i++) {
					var newMergeCell = this.cloneElement(mergeCell);

					mergeStart.row += 1;
					mergeEnd.row += 1;

					newMergeCell.attrib.ref = this.joinRange({ start: this.joinRef(mergeStart), end: this.joinRef(mergeEnd) });
					mergeCells.attrib.count += 1;
					mergeCells._children.push(newMergeCell);
				}
			}
		});

		// Update named tables below this row
		tables.forEach(table => {
			var tableRoot = table.root, tableRange = this.splitRange(tableRoot.attrib.ref), tableStart = this.splitRef(tableRange.start), tableEnd = this.splitRef(tableRange.end);

			if (tableStart.row > currentRow) {
				tableStart.row += numRows;
				tableEnd.row += numRows;

				tableRoot.attrib.ref = this.joinRange({
					start: this.joinRef(tableStart),
					end: this.joinRef(tableEnd),
				});

				var autoFilter = tableRoot.find('autoFilter');

				if (autoFilter !== null) {
					// XXX: This is a simplification that may stomp on some configurations
					autoFilter.attrib.ref = tableRoot.attrib.ref;
				}
			}
		});

		// Named cells/ranges
		workbook.findall('definedNames/definedName').forEach(name => {
			var ref = name.text;
			if (this.isRange(ref)) {
				var namedRange = this.splitRange(ref), //TODO : I think is there a bug, the ref is equal to [sheetName]![startRange]:[endRange]
					namedStart = this.splitRef(namedRange.start), // here, namedRange.start is [sheetName]![startRange] ?
					namedEnd = this.splitRef(namedRange.end);
				if (namedStart) {
					if (namedStart.row > currentRow) {
						namedStart.row += numRows;
						namedEnd.row += numRows;

						name.text = this.joinRange({
							start: this.joinRef(namedStart),
							end: this.joinRef(namedEnd),
						});

					}
				}
				if (this.option && this.option.pushDownPageBreakOnTableSubstitution) {
					if (this.sheet.name == name.text.split("!")[0].replace(/'/gi, "") && namedEnd) {
						if (namedEnd.row > currentRow) {
							namedEnd.row += numRows;
							name.text = this.joinRange({
								start: this.joinRef(namedStart),
								end: this.joinRef(namedEnd),
							});
						}
					}
				}
			} else {
				var namedRef = this.splitRef(ref);

				if (namedRef.row > currentRow) {
					namedRef.row += numRows;
					name.text = this.joinRef(namedRef);
				}
			}
		});

		// Hyperlinks below inserted rows must keep pointing at their original cells.
		sheet.findall('hyperlinks/hyperlink').forEach(hyperlink => {
			hyperlink.attrib.ref = hyperlink.attrib.ref.split(':').map(value => {
				const ref = this.splitRef(value);
				if (ref.row > currentRow)
					ref.row += numRows;
				return this.joinRef(ref);
			}).join(':');
		});
	}

	/**
	 * Hide columns by their zero-based indexes, as in the upstream API.
	 * @param {string|number} sheetName
	 * @param {number[]} hideItemIndexes
	 */
	async hideCols(sheetName, hideItemIndexes) {
		if (!Array.isArray(hideItemIndexes))
			throw new TypeError('Column indexes must be an array');
		if (hideItemIndexes.some(index => !Number.isInteger(index) || index < 0 || index >= 16384))
			throw new RangeError('Column indexes must be integers between 0 and 16383');
		const sheet = await this.loadSheet(sheetName);
		if (!hideItemIndexes.length)
			return this;
		let cols = sheet.root.find('cols');
		if (cols === null) {
			cols = etree.Element('cols');
			sheet.root.getchildren().splice(sheet.root.getchildren().indexOf(sheet.root.find('sheetData')), 0, cols);
		}
		let definitions = cols.findall('col');
		for (const index of new Set(hideItemIndexes)) {
			const colIndex = index + 1;
			const position = definitions.findIndex(col => colIndex >= Number(col.attrib.min) && colIndex <= Number(col.attrib.max));
			if (position < 0) {
				definitions.push(etree.Element('col', { min: String(colIndex), max: String(colIndex), hidden: '1' }));
				continue;
			}
			const col = definitions[position];
			const min = Number(col.attrib.min), max = Number(col.attrib.max);
			const segments = [];
			const segment = (from, to, hidden) => etree.Element('col', {
				...col.attrib, min: String(from), max: String(to), ...hidden
			});
			if (min < colIndex)
				segments.push(segment(min, colIndex - 1, {}));
			segments.push(segment(colIndex, colIndex, { hidden: '1' }));
			if (colIndex < max)
				segments.push(segment(colIndex + 1, max, {}));
			definitions.splice(position, 1, ...segments);
		}
		definitions.sort((a, b) => Number(a.attrib.min) - Number(b.attrib.min));
		this.replaceChildren(cols, definitions);
		this.archive.file(sheet.filename, etree.tostring(sheet.root));
		this.sheet = sheet;
		return this;
	}

	getWidthCell(numCol, sheet) {
		var defaultWidth = sheet.root.find('sheetFormatPr').attrib['defaultColWidth'];

		if (!defaultWidth) {
			// TODO : Check why defaultColWidth is not set ? 
			defaultWidth = 11.42578125;
		}

		var finalWidth = defaultWidth;

		sheet.root.findall('cols/col').forEach(col => {
			if (numCol >= col.attrib['min'] && numCol <= col.attrib['max']) {
				if (col.attrib['width'] != undefined) {
					finalWidth = col.attrib['width'];
				}
			}
		});

		return Number.parseFloat(finalWidth);
	}

	getWidthMergeCell(mergeCell, sheet) {
		var mergeWidth = 0;
		var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col), mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col);
		
		for (let i = mergeStartCol; i < mergeEndCol + 1; i++) {
			mergeWidth += this.getWidthCell(i, sheet);
		}
		
		return mergeWidth;
	}

	getHeightCell(numRow, sheet) {
		var defaultHight = sheet.root.find('sheetFormatPr').attrib['defaultRowHeight'];
		var finalHeight = defaultHight;
		
		sheet.root.findall('sheetData/row').forEach(row => {
			if (numRow == row.attrib['r']) {
				if (row.attrib['ht'] != undefined)
					finalHeight = row.attrib['ht'];	
			}
		});

		return Number.parseFloat(finalHeight);
	}

	getHeightMergeCell(mergeCell, sheet) {
		var mergeHeight = 0;
		var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStartRow = this.splitRef(mergeRange.start).row, mergeEndRow = this.splitRef(mergeRange.end).row;
		
		for (let i = mergeStartRow; i < mergeEndRow + 1; i++) {
			mergeHeight += this.getHeightCell(i, sheet);
		}
		
		return mergeHeight;
	}

	/**
	 * @param {{ attrib: { ref: any; }; }} mergeCell
	 */
	getNbRowOfMergeCell(mergeCell) {
		var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStartRow = this.splitRef(mergeRange.start).row, mergeEndRow = this.splitRef(mergeRange.end).row;
		
		return mergeEndRow - mergeStartRow + 1;
	}

	/**
	 * @param {number} pixels
	 */
	pixelsToEMUs(pixels) {
		return Math.round(pixels * 914400 / 96);
	}

	/**
	 * @param {number} width
	 */
	columnWidthToEMUs(width) {
		// TODO : This is not the true. Change with true calcul
		// can find help here : 
		// https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
		// https://stackoverflow.com/questions/58021996/how-to-set-the-fixed-column-width-values-in-inches-apache-poi
		// https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Sheet.html#setColumnWidth-int-int-
		// https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
		// https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
		// http://lcorneliussen.de/raw/dashboards/ooxml/
		return this.pixelsToEMUs(width * 7.625579987895905);
	}

	/**
	 * @param {number} height
	 */
	rowHeightToEMUs(height) {
		// TODO : need to be verify
		return Math.round(height / 72 * 914400);
	}

	/**
	 * Find max file id.
	 * @param {RegExp} fileNameRegex 
	 * @param {RegExp} idRegex 
	 * @returns {number} 
	 */
	findMaxFileId(fileNameRegex, idRegex) {
		var files = this.archive.file(fileNameRegex);

		var maxid = files.reduce((p, c) => {
			const num = parseInt(idRegex.exec(c.name)[1]);

			if (p == null)
				return num;

			return p > num ? p : num;
		}, 0);

		maxid++;

		return maxid;
	}

	generateUUID() {
		const bytes = randomBytes(16);
		bytes[6] = (bytes[6] & 0x0f) | 0x40;
		bytes[8] = (bytes[8] & 0x3f) | 0x80;
		const hex = bytes.toString('hex');
		return [hex.slice(0, 8), hex.slice(8, 12), hex.slice(12, 16), hex.slice(16, 20), hex.slice(20)].join('-');
	}

	cellInMergeCells(cell, mergeCell) {
		var cellCol = this.charToNum(this.splitRef(cell.attrib.r).col);
		var cellRow = this.splitRef(cell.attrib.r).row;
		var mergeRange = this.splitRange(mergeCell.attrib.ref), mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col), mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col), mergeStartRow = this.splitRef(mergeRange.start).row, mergeEndRow = this.splitRef(mergeRange.end).row;
		
		if (cellCol >= mergeStartCol && cellCol <= mergeEndCol) {
			if (cellRow >= mergeStartRow && cellRow <= mergeEndRow)
				return true;	
		}

		return false;
	}

	/**
	 * @param {string} str
	 */
	isUrl(str) {
		var pattern = new RegExp('^(https?:\\/\\/)?' + // protocol
			'((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain name
			'((\\d{1,3}\\.){3}\\d{1,3}))' + // OR ip (v4) address
			'(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // port and path
			'(\\?[;&a-z\\d%_.~+=-]*)?' + // query string
			'(\\#[-a-z\\d_]*)?$', 'i'); // fragment locator

		return !!pattern.test(str);
	}

	toArrayBuffer(buffer) {
		var ab = new ArrayBuffer(buffer.length);
		var view = new Uint8Array(ab);

		for (var i = 0; i < buffer.length; ++i) {
			view[i] = buffer[i];
		}

		return ab;
	}

	imageToBuffer(imageObj) {
		/**
		 * Check if the buffer image is supported by the library before return it
		 * @param {Buffer} buffer the final buffer image
		 */
		function checkImage(buffer) {
			try {
				sizeOf(buffer);
				return buffer;
			} catch (error) {
				throw new TypeError('imageObj cannot be parse as a buffer image');
			}
		}

		if (!imageObj)
			throw new TypeError('imageObj cannot be null');

		if (imageObj instanceof Buffer) {
			return checkImage(imageObj);
		}
		else 
		{
			if (typeof (imageObj) === 'string' || imageObj instanceof String) {
				imageObj = imageObj.toString();
				//if(this.isUrl(imageObj)){
				// TODO
				//}else{
				var imagePath = this.option && this.option.imageRootPath ? `${this.option.imageRootPath}/${imageObj}` : imageObj;
				if (fs.existsSync(imagePath)) {
					return checkImage(Buffer.from(fs.readFileSync(imagePath, { encoding: 'base64' }), 'base64'));
				}
				//}
				try {
					return checkImage(Buffer.from(imageObj, 'base64'));
				} catch (error) {
					throw new TypeError('imageObj cannot be parse as a buffer');
				}
			}
		}

		throw new TypeError(`imageObj type is not supported : ${typeof (imageObj)}`);
	}

	findMaxId(element, tag, attr, idRegex) {
		let maxId = 0;

		element.findall(tag).forEach(element => {
			const match = idRegex.exec(element.attrib[attr]);
			
			if (match == null) {
				throw new Error('Can not find the id!');
			}

			const cid = parseInt(match[1]);

			if (cid > maxId)
				maxId = cid;
		});

		return ++maxId;
	}
}

module.exports = { ExcelTemplate };
