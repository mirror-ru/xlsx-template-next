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

## Limitations

The template must be an `.xlsx` file. Older `.xls`, binary `.xlsb`, and macro
enabled `.xlsm` files are not supported.

Most ordinary templates work as expected: a list expands across a row and a
table expands downwards. The library also moves the common things around that
content — merged cells, named ranges, Excel tables, and hyperlinks.

The tricky case is a complex layout around a growing table. If a table sits
inside merged cells that span several rows, or a range crosses the place where
new rows are inserted, open the generated file and check it in Excel. The
library cannot safely understand and rewrite every kind of reference in a
workbook.

Formulas are kept and Excel recalculates their values when the generated file
is opened. Their cell references are not changed automatically. For example,
if rows are inserted above `=SUM(B2:B10)`, the formula still refers to
`B2:B10`. For a table with an unknown number of rows, use an actual Excel table
in the template and a formula such as `TableName[ColumnName]`. Its range grows
with the generated rows. Another option is to put the table at the end of a
sheet and use a larger fixed range in the formula.

Tags work in ordinary cells and Excel tables. They do not work inside pivot
tables, and pivot tables are not refreshed by the library.
