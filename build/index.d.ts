export class ExcelTemplate {
    /**
     * Create a new workbook. Either pass the raw data of a .xlsx file,
     * and call `loadTemplate()` later.
     */
    constructor(option?: {});
    archive: JSZip;
    sharedStrings: any[];
    sharedStringsLookup: {};
    option: {
        moveImages: boolean;
        subsituteAllTableRow: boolean;
        moveSameLineImages: boolean;
        imageRatio: number;
        pushDownPageBreakOnTableSubstitution: boolean;
        imageRootPath: any;
        handleImageError: any;
    };
    sharedStringsPath: string;
    sheets: any[];
    sheet: {
        filename: any;
        name: any;
        id: any;
        root: any;
    };
    workbook: any;
    workbookPath: any;
    contentTypes: any;
    prefix: string;
    workbookRels: any;
    calChainRel: any;
    calcChainPath: string;
    /**
     * Delete unused sheets if needed
     */
    deleteSheet(sheetName: any): Promise<this>;
    /**
     * Clone sheets in current workbook template
     */
    copySheet(sheetName: any, copyName: any): Promise<this>;
    /**
     *  Partially rebuild after copy/delete sheets
     */
    _rebuild(): void;
    /**
     * Load document from file or from buffer (auto-detection)
     */
    load(value: any): Promise<void>;
    /**
     * Load a .xlsx file from filename
     */
    loadFile(filename: any): Promise<void>;
    /**
     * Load a .xlsx file from a byte array.
     */
    loadTemplate(data: any): Promise<void>;
    /**
     * Interpolate values for all the sheets using the given substitutions
     * (an object).
     */
    processAll(substitutions: any): Promise<void>;
    /**
     * Interpolate values for all the sheets using the given substitutions
     * (an object).
     */
    substituteAll(substitutions: any): Promise<void>;
    /**
     * Interpolate values for the sheet with the given number (1-based) or
     * name (if a string) using the given substitutions (an object).
     */
    process(sheetName: any, substitutions: any): Promise<void>;
    /**
     * Interpolate values for the sheet with the given number (1-based) or
     * name (if a string) using the given substitutions (an object).
     */
    substitute(sheetName: any, substitutions: any): Promise<void>;
    /**
     * Build a new binary .xlsx file
     */
    build(options: any): Promise<string | number[] | ArrayBuffer | Uint8Array<ArrayBufferLike> | Blob | Buffer<ArrayBufferLike>>;
    writeSharedStrings(): Promise<void>;
    addSharedString(s: any): number;
    getInlineString(cell: any): string;
    stringIndex(s: any): any;
    replaceString(oldString: any, newString: any): any;
    loadSheets(prefix: any, workbook: any, workbookRels: any): any[];
    loadSheet(sheet: any): Promise<{
        filename: any;
        name: any;
        id: any;
        root: any;
    }>;
    loadSheetRels(sheetFilename: any): Promise<{
        filename: string;
        root: any;
    }>;
    initSheetRels(sheetFilename: any): {
        filename: string;
        root: any;
    };
    loadDrawing(sheet: any, sheetFilename: any, rels: any): Promise<{
        filename: string;
        root: any;
    }>;
    addContentType(partName: any, contentType: any): void;
    initDrawing(sheet: any, rels: any): {
        root: any;
        filename: string;
        relFilename: string;
        relRoot: any;
    };
    writeDrawing(drawing: any): void;
    moveAllImages(drawing: any, fromRow: any, nbRow: any): void;
    _moveTwoCellAnchor(drawingElement: any, fromRow: any, nbRow: any): void;
    loadTables(sheet: any, sheetFilename: any): Promise<any[]>;
    writeTables(tables: any): void;
    substituteHyperlinks(rels: any, substitutions: any): Promise<void>;
    substituteTableColumnHeaders(tables: any, substitutions: any): void;
    extractPlaceholders(text: any): {
        placeholder: string;
        type: string;
        name: string;
        key: string;
        subType: string;
        full: boolean;
    }[];
    splitRef(ref: any): {
        table: any;
        colAbsolute: boolean;
        col: any;
        rowAbsolute: boolean;
        row: number;
    };
    joinRef(ref: any): string;
    nextCol(ref: any): any;
    nextRow(ref: any): any;
    charToNum(str: any): number;
    numToChar(num: any): string;
    isRange(ref: any): boolean;
    isWithin(ref: any, startRef: any, endRef: any): boolean;
    stringify(value: any): string | number;
    insertCellValue(cell: any, substitution: any): any;
    substituteScalar(cell: any, string: any, placeholder: any, substitution: any): any;
    substituteArray(cells: any, cell: any, substitution: any): number;
    substituteTable(row: any, newTableRows: any, cells: any, cell: any, namedTables: any, substitution: any, key: any, placeholder: any, drawing: any): number;
    substituteImage(cell: any, string: any, placeholder: any, substitution: any, drawing: any): boolean;
    cloneElement(element: any, deep: any): any;
    replaceChildren(parent: any, children: any): void;
    getCurrentRow(row: any, rowsInserted: any): any;
    getCurrentCell(cell: any, currentRow: any, cellsInserted: any): string;
    updateRowSpan(row: any, cellsInserted: any): void;
    splitRange(range: any): {
        start: any;
        end: any;
    };
    joinRange(range: any): string;
    pushRight(workbook: any, sheet: any, currentCell: any, numCols: any): void;
    pushDown(workbook: any, sheet: any, tables: any, currentRow: any, numRows: any): void;
    getWidthCell(numCol: any, sheet: any): number;
    getWidthMergeCell(mergeCell: any, sheet: any): number;
    getHeightCell(numRow: any, sheet: any): number;
    getHeightMergeCell(mergeCell: any, sheet: any): number;
    /**
     * @param {{ attrib: { ref: any; }; }} mergeCell
     */
    getNbRowOfMergeCell(mergeCell: {
        attrib: {
            ref: any;
        };
    }): number;
    /**
     * @param {number} pixels
     */
    pixelsToEMUs(pixels: number): number;
    /**
     * @param {number} width
     */
    columnWidthToEMUs(width: number): number;
    /**
     * @param {number} height
     */
    rowHeightToEMUs(height: number): number;
    /**
     * Find max file id.
     * @param {RegExp} fileNameRegex
     * @param {RegExp} idRegex
     * @returns {number}
     */
    findMaxFileId(fileNameRegex: RegExp, idRegex: RegExp): number;
    cellInMergeCells(cell: any, mergeCell: any): boolean;
    /**
     * @param {string} str
     */
    isUrl(str: string): boolean;
    toArrayBuffer(buffer: any): ArrayBuffer;
    imageToBuffer(imageObj: any): Buffer<ArrayBufferLike>;
    findMaxId(element: any, tag: any, attr: any, idRegex: any): number;
}
import JSZip = require("jszip");
