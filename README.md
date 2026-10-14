# XLSX Template Next

[![npm version](https://badge.fury.io/js/xlsx-template-next.svg)](https://badge.fury.io/js/xlsx-template-next) ![Downloads](https://img.shields.io/npm/dm/xlsx-template-next.svg)

Generate Excel (`.xlsx`) reports from templates created in Excel.

## Installation

```sh
npm i xlsx-template-next
```

Node.js 18 or newer is required.

## Basic template

From:

![Template xlsx](https://raw.githubusercontent.com/mirror-ru/xlsx-template-next/master/images/before.png)

To:

![Result process](https://raw.githubusercontent.com/mirror-ru/xlsx-template-next/master/images/after.png)

## Basic usage

Put tags in the cells of `template.xlsx`, for example:

    | Date | ${extractDate} |
    | Dates | ${dates} |
    | ${table:people.name} | ${table:people.age} |

Save this as `main.mjs` beside the template and run `node main.mjs`:

```js
import { ExcelTemplate } from 'xlsx-template-next';
import fs from 'node:fs';

async function main() {
	const template = new ExcelTemplate();

	await template.load('./template.xlsx');

	await template.process('Sheet1', {
		extractDate: new Date(),
		dates: [
			new Date('2023-06-01'),
			new Date('2023-06-02'),
			new Date('2023-06-03')
		],
		people: [
			{ name: 'John Smith', age: 20 },
			{ name: 'Bob Johnson', age: 22 }
		]
	});

	fs.writeFileSync('./output.xlsx', await template.build());
}

main();
```

Replace `Sheet1` with the name of the sheet tab in your workbook.

### Loading from a buffer

If the template is already in a `Buffer`, pass the buffer to `load()`:

```js
import { ExcelTemplate } from 'xlsx-template-next';
import fs from 'node:fs';

async function main() {
	const template = new ExcelTemplate();
	const buffer = fs.readFileSync('./template.xlsx');

	await template.load(buffer);

	await template.process('Sheet1', {
		extractDate: new Date(),
		dates: [
			new Date('2023-06-01'),
			new Date('2023-06-02'),
			new Date('2023-06-03')
		],
		people: [
			{ name: 'John Smith', age: 20 },
			{ name: 'Bob Johnson', age: 22 }
		]
	});

	fs.writeFileSync('./output.xlsx', await template.build());
}

main();
```

### One sheet or every sheet

For a sheet tab named `Invoice` with a cell containing `${invoiceNumber}`:

```js
await template.process('Invoice', {
	invoiceNumber: 'INV-001'
});
```

If the same tags appear on several sheets, fill all of them at once:

```js
await template.processAll({
	invoiceNumber: 'INV-001'
});
```

## Tags

### Scalars

Use `${name}` for a single value. The placeholder may occupy a whole cell or be
part of its text. For example:

    | Extracted on: | ${extractDate} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2023 |

For dates, keep the placeholder alone in the cell and set its date format in
Excel.

Array indexes are supported:

Given data

```js
const data = {
	extractDates: [ 'Jun-01-2022', 'Jun-01-2023' ]
};
```

Use it in a template as:

    | Extracted on: | ${extractDates[0]} |

Result:

    | Extracted on: | Jun-01-2022 |

### Formulas

Put `${total}` in cell B11, then pass a formula as its value:

```js
await template.process('Report', {
	total: '=SUM(B2:B10)'
});
```

Excel calculates the formula when the generated file is opened.

### Columns

An array in a placeholder creates one cell per item across columns. The
placeholder must be the only content of its cell. For example:

    | ${dates} |

might result in:

    | Jun-01-2023 | Jun-02-2023 | Jun-03-2023 |

### Tables

Put table tags in one template row:


    | Name                 | Age                 |
    | ${table:people.name} | ${table:people.age} |

Pass an array:

```js
await template.process('Sheet1', {
	people: [
		{ name: 'John Smith', age: 20 },
		{ name: 'Bob Johnson', age: 22 }
	]
});
```

The result:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

An array field in a table row is repeated across columns too.

### Images

Put `${image:imageName}` in a cell, then pass an image path:

```js
await template.process('Sheet1', {
	imageName: './logo.png'
});
```

For an image in every row of a table, use this tag:

    | My images | ${table:images.name:image} |

Given data

```js
const data = {
	images: [
		{ name: 'helloImage1.jpg' },
		{ name: 'helloImage2.jpg' }
	]
};
```

Paths, a `Buffer`, and raw base64 strings are supported. For a `Buffer`:

```js
await template.process('Report', {
	imageName: fs.readFileSync('./logo.png')
});
```

If all images are in one directory, configure it once and use filenames in
your data:

```js
const template = new ExcelTemplate({
	imageRootPath: './images',
	imageRatio: 75
});

await template.load('./template.xlsx');
await template.process('Sheet1', {
	imageName: 'logo.png'
});
```

`imageRatio` is a percentage and applies to ordinary cells. An image in a
merged cell is fitted to that merged area.

## Sheet helpers

Copy a sheet, hide columns B and D, then remove the copy:

```js
await template.copySheet('Invoice', 'Invoice copy');
await template.hideCols('Invoice', [1, 3]);
await template.deleteSheet('Invoice copy');
```

Column indexes start at `0`: A is `0`, B is `1`, and so on.

## Saving the result

The basic example writes the default result directly to a file. To receive a
Node.js `Buffer` instead:

```js
const output = await template.build({ type: 'nodebuffer' });
fs.writeFileSync('./output.xlsx', output);
```

You can also use `uint8array`, `arraybuffer`, `blob`, or `base64`.

## Limitations

- The template must be saved as `.xlsx`. Formats `.xls`, `.xlsb`, and `.xlsm`
  are not supported.
- An array expands across cells in its own row, shifting subsequent cells
  right; an empty array removes the placeholder cell and shifts them left.
  A table expands downwards from the template row. An empty table clears its
  placeholders without removing the row. Only limited adjustments to other
  workbook structures are made:

  - Merged ranges and ordinary named cells/ranges starting to the right of the
    array on the same row are shifted right or left.
  - Merged ranges, Excel tables, and ordinary named cells/ranges below an
    expanded table are moved down. A merged range that starts on the template
    row is copied to the generated rows.
  - Hyperlinks anchored to moved cells are moved with them.
  - Ranges spanning several rows or crossing an insertion point need manual
    checking. In particular, merged cells across multiple rows may become
    invalid when those rows expand by different amounts. These updates are
    not a general-purpose rewrite of every reference in the workbook.
- Formula references are not rewritten when rows or columns are inserted. If a
  formula must refer to a table with an unknown number of rows or columns,
  either:

  - place the table at the end of a sheet and use a sufficiently large formula
    range, keeping other data and the formula cell outside that range; or
  - use an Excel table. Its range is expanded with the generated rows, so a
    structured reference such as `TableName[ColumnName]` includes all values.
    This means an actual named Excel table created in the template, not just
    cells styled to look like a table.
- Placeholder substitution is supported in ordinary cells and Excel tables,
  not inside pivot tables or other specialised Excel objects. The library does
  not refresh pivot tables.
