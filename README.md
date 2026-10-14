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

Create `template.xlsx` in Excel with placeholders such as `${extractDate}`,
`${dates}`, `${table:people.name}`, and `${table:people.age}` in its cells.
The [Tags](#tags) section explains how to lay them out.

Save the following code as `main.mjs` beside the template. After installing the
package, run `node main.mjs` from that directory. It loads the template, fills
one sheet, and writes `output.xlsx`:

```js
import { ExcelTemplate } from 'xlsx-template-next';
import fs from 'node:fs';

async function main() {
	const template = new ExcelTemplate();

	await template.load('./template.xlsx');

	await template.process(1, {
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

Use the shorter processing examples below in place of the `process(...)` block
inside `main()`, after `load()` and before `build()`. Where only a `data` object
is shown, pass it as the second argument to `process()`. Replace sample sheet
names such as `Invoice` and `Report` with names that exist in your workbook.
Constructor options are set when creating the `ExcelTemplate` instance, before
loading the file.

### Choosing a sheet

To fill a cell containing `${invoiceNumber}` on a sheet named `Invoice`:

```js
const data = { invoiceNumber: 'INV-001' };
await template.process('Invoice', data);
```

`process()` accepts a sheet name or number. A number first matches the sheet's
internal ID; if no ID matches, it is treated as a one-based position. After
sheets have been reordered, IDs and positions can differ, so using a name is
the clearest choice.

Alternatively, use the same `data` object to fill all sheets:

```js
await template.processAll(data);
```

The sheets can have different placeholders; include their values in `data`.
Missing values are replaced with empty text.

### Loading from a buffer

`load()` also accepts a `Buffer`, for example when the template was received
from an HTTP request or read from storage:

```js
import { ExcelTemplate } from 'xlsx-template-next';
import fs from 'node:fs';

async function main() {
	const template = new ExcelTemplate();
	const buffer = fs.readFileSync('./template.xlsx');

	await template.load(buffer);

	await template.process(1, {
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

## Tags

### Scalars

Use `${name}` for a single value. The placeholder may occupy a whole cell or be
part of its text. For example:

    | Extracted on: | ${extractDate} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2023 |

For a `Date` value, put the placeholder alone in the cell and apply a date
format in Excel. To include a date inside text, pass an already formatted
string instead.

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

Put `${total}` alone in cell B11 of the `Report` sheet, then supply a string
beginning with `=`:

```js
await template.process('Report', {
	total: '=SUM(B2:B10)'
});
```

The library writes the formula but does not calculate its result. Recalculate
the generated workbook in Excel to obtain the value. Inserting rows or columns
does not rewrite the references inside the formula (see [Limitations](#limitations)).

### Columns

An array in a placeholder creates one cell per item across columns. The
placeholder must be the only content of its cell. For example:

    | ${dates} |

might result in:

    | Jun-01-2023 | Jun-02-2023 | Jun-03-2023 |

### Tables

For a table, put `${table:items.field}` placeholders in adjacent cells on one
template row. Each placeholder must occupy its whole cell. The first item
fills that row; further items create additional rows:

    | Name                | Age                 |
    | ${table:people.name} | ${table:people.age} |

With an array of `{ name, age }` objects in `people`, the result is:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

An array field inside a table row is repeated across columns in the same way.
An empty table array clears its placeholder cells but keeps the template row.

### Images

Put the image placeholder alone in its cell:

    | My image: | ${image:imageName} |

Given data

```js
const data = {
	imageName: 'helloImage.jpg'
};
```

For images in a table, use:

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

Image data can be a `Buffer` containing image bytes, a raw base64 string
(without a `data:image/...;base64,` prefix), or a path to a local image file.
Image URLs are not downloaded automatically.

For example, with `${image:imageName}` in a cell on `Report`, read the image
into a `Buffer` using the `fs` import from the full example:

```js
await template.process('Report', {
	imageName: fs.readFileSync('./logo.png')
});
```

Relative image paths are resolved from the Node.js process's working
directory. Set `imageRootPath` to use a different directory and pass image
filenames relative to it:

```js
const option = {
	imageRootPath: '/path/to/your/image/dir'
};
const template = new ExcelTemplate(option);
```

An image in a merged cell is sized to fit that merged area.

Use `imageRatio` to scale image width and height in ordinary cells (percentage;
default `100`). This option does not affect images in merged cells:

```js
const option = {
	imageRatio: 75.4
};
const template = new ExcelTemplate(option);
```

## Sheet helpers

You can copy or remove sheets by name or numeric ID, and hide columns by their
zero-based indexes. Perform these operations after `load()` and before `build()`:

```js
const sheetName = 'Sheet1'; // replace with the name in your workbook
await template.copySheet(sheetName, 'Copy of Sheet1');
await template.hideCols(sheetName, [1, 3]); // hide columns B and D
await template.deleteSheet('Copy of Sheet1');
```

`hideCols` accepts integer indexes from 0 (A) to 16383 (XFD). It hides only
the requested columns, preserving the widths and styles of adjacent columns.
`copySheet` accepts a sheet name or numeric ID; the new name is optional.
Local defined names and comment threads are copied with the sheet.

## Saving the result

Call `build()` after processing the sheets and making any sheet changes.
Without options it returns a `Uint8Array` containing the `.xlsx` file, which
can be passed directly to `fs.writeFileSync()` as in the full examples.

For a Node.js `Buffer`, pass an options object:

```js
const output = await template.build({ type: 'nodebuffer' });
fs.writeFileSync('./output.xlsx', output);
```

Other output types include `uint8array`, `arraybuffer`, `blob`, and `base64`.
Whenever you pass an options object, include `type` explicitly.

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
