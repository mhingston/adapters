# adapters

A collection of adapters exposing the same API for reading and updating a data source.

## Installation

    npm install mhingston/adapters

## Adapters

### Spreadsheet: new Spreadsheet(options)

* `options` {Object} Configuration options.
  * `data` {any} See [`XLSX.read`](https://github.com/SheetJS/js-xlsx#parsing-functions).
  * `parsingOptions` {Object} See [Parsing Options](https://github.com/SheetJS/js-xlsx#parsing-options).
  * `writingOptions` {Object} See [Writing Options](https://github.com/SheetJS/js-xlsx#writing-options).
  * `workSheets` {Object[]} An array of worksheet objects to process.
    * `id` {Number|String} Worksheet ID, can be a string or number (zero indexed).
    * `columns` {Object[]} An array of column object to process
      * `id` {Number|String} Column ID, can be a string or number (zero indexed).

All adapters expose `read` and `update` methods.

`<Adapter>.read(options)` (async)

* `options` {Object} Configuration options.
  * `batchSize` {Number} How many rows to read.

Returns: {Object}
* `data` {Object[]} An array of objects.
* `meta`: {any} Meta data used by the adapter for data translation.

`<Adapter>.update(args)` (async)

* `args` {Object} Configuration options.
  * `data` {Object[]} The (modified) data array from `<Adapter>.read`.
  * `meta` {any} The (unmodified) meta data from `<Adapter>.read`.
  * `options` {Object} Adapter-specific configuration options.

# Examples

```javascript
const {Spreadsheet} = require('adapters');

const main = async () =>
{
    const spreadsheet = new Spreadsheet(
    {
        data: '/path/to/a/file',
        parsingOptions: // see https://docs.sheetjs.com/#parsing-options
        {
            type: 'file'
        },
        workSheets:
        [
            {
                id: 'Sheet1', // worksheet ID, can be a string or number (zero indexed)
                columns: // An array containing the columns you want to export from the worksheet
                [
                    {
                        id: 0 // column ID, can be a string or number (zero indexed)
                    }
                ]
            }
        ]
    });
    
    const {data, meta} = await spreadsheet.read(
    {
        batchSize: 20 // how many rows to read
    });

    // Do something with the data (in this case add a new column)
    data.forEach((row) => row['ExtraColumn'] = Math.random());

    await spreadsheet.update(
    {
        data,
        meta,
        options:
        {
            save: true // When this is true changes to the spreadsheet will be saved to disk. You probably only want to set this to true for the final update.
        }
    });
}

main();