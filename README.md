# XLSX Template Next

[![npm version](https://badge.fury.io/js/xlsx-template-next.svg)](https://badge.fury.io/js/xlsx-template-next) ![Downloads](https://img.shields.io/npm/dm/xlsx-template-next.svg)

This module provides a means of generating Excel reports (i.e. not CSV
files) in NodeJS applications.

## Installation

```sh
# with npm
$ npm i xlsx-template-next
```
## Basic template

From:

![Template xlsx](https://raw.githubusercontent.com/mirror-ru/xlsx-template-next/master/images/before.png)

To:

![Result process](https://raw.githubusercontent.com/mirror-ru/xlsx-template-next/master/images/after.png)

## Basic usage

Loading a document by filename:
```js
import { ExcelTemplate } from 'xlsx-template-next'
import fs from 'fs'

async function main() {
	const template = new ExcelTemplate();

	await template.load('./template.xlsx');

	await template.process(1, {
		extractDate: new Date(),
		dates: [ 
			new Date('2013-06-01'), 
			new Date('2013-06-02'), 
			new Date('2013-06-03')
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

Loading a document from buffer:
```js
import { ExcelTemplate } from 'xlsx-template-next'
import fs from 'fs'

async function main() {
	const template = new ExcelTemplate();

	const buffer = fs.readFileSync('./template.xlsx');

	await template.load(buffer);

	await template.process(1, {
		extractDate: new Date(),
		dates: [ 
			new Date('2013-06-01'), 
			new Date('2013-06-02'), 
			new Date('2013-06-03')
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

Simple placholders take the format `${name}`. Here, `name` is the name of a
key in the placeholders map. The value of this placholder here should be a
scalar, i.e. not an array or object. The placeholder may appear on its own in a
cell, or as part of a text string. For example:

    | Extracted on: | ${extractDate} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2023 |

Here, `extractDate` may be a date and the second cell may be formatted as a
number.

Inside scalars there possibility to use array indexers. 
For example: 

Given data

```js
let data = { 
	extractDates: [ 'Jun-01-2022', 'Jun-01-2023' ]
};
```

which will be applied to following template

    | Extracted on: | ${extractDates[0]} |

will results in the 

    | Extracted on: | Jun-01-2022 |

### Columns

You can use arrays as placeholder values to indicate that the placeholder cell
is to be replicated across columns. In this case, the placeholder cannot appear
inside a text string - it must be the only thing in its cell. For example,
if the placehodler value `dates` is an array of dates:

    | ${dates} |

might result in:

    | Jun-01-2023 | Jun-02-2023 | Jun-03-2023 |

### Tables

Finally, you can build tables made up of multiple rows. In this case, each
placeholder should be prefixed by `table:` and contain both the name of the
placeholder variable (a list of objects) and a key (in each object in the list).
For example:

    | Name                 | Age                 |
    | ${table:people.name} | ${table:people.age} |

If the replacement value under `people` is an array of objects, and each of
those objects have keys `name` and `age`, you may end up with something like:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

If a particular value is an array, then it will be repeated across columns as
above.

### Images

You can insert images with   

    | My image: | ${image:imageName} |

Given data

```js
let data = { 
	imageName: 'helloImage.jpg' 
};
```  

You can insert a list of images with   

    | My images | ${table:images.name:image} |

Given data

```js
let data = { 
    images: [ 
      { name: 'helloImage1.jpg' }, 
      { name: 'helloImage2.jpg' } 
    ]
};
``` 

Supported image format in given data: 
- Base64 string
- Base64 Buffer
- Absolute path file
- relative path file (absolute is prior to relative in test)

You can pass imageRootPath option for setting the root folder for your images.  

```js
let option = { 
	imageRootPath: '/path/to/your/image/dir' 
};
///...  
const template = new ExcelTemplate(option);
```

If the image Placeholders is in standard cell, image is insert normaly  
If the image Placeholders is in merge cell, image feet (at the best) the size of the merge cell.

You can pass imageRatio option for adjust the ratio image (in percent and for standard cell - not applied on merge cell)

```js
let option = { 
	imageRatio: 75.4 
};  
///...  
const template = new ExcelTemplate(option);
```

### Sheet helpers

You can copy or remove sheets and hide columns by their zero-based indexes:

```js
await template.copySheet('Sheet1', 'Copy of Sheet1');
await template.hideCols('Sheet1', [1, 3]); // hide columns B and D
await template.deleteSheet('Copy of Sheet1');
```

`hideCols` accepts integer indexes from 0 (A) to 16383 (XFD). It hides only
the requested columns, preserving the widths and styles of adjacent columns.
`copySheet` accepts a sheet name or numeric ID; the new name is optional.
Local defined names and comment threads are copied with the sheet.

At this stage, `data` is a string blob representing the compressed archive that
is the `.xlsx` file (that's right, a `.xlsx` file is a zip file of XML files,
if you didn't know). You can send this back to a client, store it to disk,
attach it to an email or do whatever you want with it.

You can pass options to `generate()` to set a different return type. use
`{ type: 'uint8array' }` to generate a `Uint8Array`, `arraybuffer`, `blob`,
`nodebuffer` to generate an `ArrayBuffer`, `Blob` or `nodebuffer`, or
`base64` to generate a base64-encoded string. Default: `{ type: 'uint8array' }`

## Caveats

* The spreadsheet must be saved in `.xlsx` format. `.xls`, `.xlsb` or `.xlsm`
  won't work.
* Column (array) and table (array-of-objects) insertions cause rows and cells to
  be inserted or removed. When this happens, only a limited number of
  adjustments are made:
    * Merged cells and named cells/ranges to the right of cells where insertions
      or deletions are made are moved right or left, appropriately. This may
      not work well if cells are merged across rows, unless all rows have the
      same number of insertions.
    * Merged cells, named tables or named cells/ranges below rows where further
      rows are inserted are moved down.
  Formulae are not adjusted.
* As a corollary to this, it is not always easy to build formulae that refer
  to cells in a table (e.g. summing all rows) where the exact number of rows
  or columns is not known in advance. There are two strategies for dealing
  with this:
    * Put the table as the last (or only) thing on a particular sheet, and
      use a formula that includes a large number of rows or columns in the
      hope that the actual table will be smaller than this number.
    * Use named tables. When a placeholder in a named table causes columns or
      rows to be added, the table definition (i.e. the cells included in the
      table) will be updated accordingly. You can then use things like
      `TableName[ColumnName]` in your formula to refer to all values in a given
      column in the table as a logical range.
* Placeholders only work in simple cells and tables, pivot tables or
  other such things.
