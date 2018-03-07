const XLSX = require('xlsx');

const getWorksheet = (spreadsheet) =>
{
    const error = new Error();

    while(spreadsheet.worksheetIndex < spreadsheet.options.worksheets.length-1)
    {
        spreadsheet.worksheetIndex++;
        spreadsheet.rowIndex = 0;
        const worksheet = spreadsheet.options.worksheets[spreadsheet.worksheetIndex];
        
        if(typeof worksheet.id === 'number' && spreadsheet.workbook.SheetNames[worksheet.id])
        {
            return spreadsheet.workbook.SheetNames[worksheet.id];
        }

        else if(typeof worksheet.id === 'string' && spreadsheet.workbook.SheetNames.includes(worksheet.id))
        {
            return worksheet.id;
        }

        else if(worksheet.required)
        {
            error.required = true;
            break;
        }
    }

    error.message = `Failed to open worksheet ${spreadsheet.options.worksheets[spreadsheet.worksheetIndex]}.`;
    throw error;
}

class Spreadsheet
{
    constructor(options = {})
    {
        this.options = options;
        this.options.parsingOptions = options.parsingOptions || {};
        this.options.writingOptions = options.writingOptions || {};
        this.workbook = XLSX.read(options.data, Object.assign({}, options.parsingOptions));
        this.worksheetIndex = -1;
        this.worksheet = getWorksheet(this);
        this.worksheetData = XLSX.utils.sheet_to_json(this.workbook.Sheets[this.worksheet]);
    }

    async read(options = {})
    {
        let meta = {};
        let data = this.worksheetData.slice(this.rowIndex, this.rowIndex+options.batchSize)
        .map((rowData, index) =>
        {
            const keys = Object.keys(rowData);
            const row = {};
            this.options.worksheets[this.worksheetIndex].columns.forEach((column) =>
            {
                if(typeof column.id === 'number')
                {
                    if(keys[column.id] === undefined && column.required)
                    {
                        throw new Error(`Row ${index+2} on worksheet ${this.worksheet} is missing the column at index ${column.id+1}.`)
                    }

                    row[keys[column.id]] = rowData[keys[column.id]];
                }

                else if(typeof column.id === 'string')
                {
                    if(rowData[column.id] === undefined && column.required)
                    {
                        throw new Error(`Row ${index+2} on worksheet ${this.worksheet} is missing the column ${column.id}.`)
                    }

                    row[column.id] = rowData[column.id];
                }
            });
            return row;
        });
        meta.start = this.rowIndex;
        this.rowIndex += data.length;
        meta.end = this.rowIndex;
        meta.id = this.worksheet;
        const remaining = options.batchSize - data.length;

        if(data.length === 0)
        {
            return {data: null, meta: null};
        }

        else if(remaining > 0)
        {
            try
            {
                this.worksheet = getWorksheet(this);
                this.worksheetData = XLSX.utils.sheet_to_json(this.workbook.Sheets[this.worksheet]);
            }

            catch(error)
            {
                if(error.required)
                {
                    throw error;
                }
            }

            const additional = await this.read({batchSize: remaining});

            if(additional.data)
            {
                data = data.concat(additional.data);
                meta = Array.isArray(additional.meta) ? [meta, ...additional.meta] : [meta, additional.meta];
            }
        }

        return {data, meta: Array.isArray(meta) ? meta : [meta]};
    }

    async update({data, meta, options})
    {
        let index = 0;
        const keys = Object.keys(data[0]);
        meta.forEach((meta) =>
        {
            const worksheet = XLSX.utils.sheet_to_json(this.workbook.Sheets[meta.id]);

            for(let row=meta.start; row <= meta.end; row++)
            {
                for(const key of keys)
                {
                    worksheet[row][key] = data[index][key];
                }

                index++;
            }

            this.workbook.Sheets[meta.id] = XLSX.utils.json_to_sheet(worksheet);
        });

        if(options.save)
        {
            let path;
            
            if(this.options.writingOptions.path)
            {
                path = this.options.writingOptions.path;
            }

            else if(this.options.data && this.options.parsingOptions.type === 'file')
            {
                path = this.options.data;
            }

            else
            {
                throw new Error('Output filename not specified.');
            }

            XLSX.writeFile(this.workbook, path, this.options.writingOptions);
        }
    }
}

module.exports = Spreadsheet;
