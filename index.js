/* Dependencies */
var moment = require('moment');
var GoogleSpreadsheet = require("google-sheets-node-api");

var hasCreds = !!process.env.SHEETS_CREDENTIALS;
var parsedCreds = false;
var parsedCredsError;
var creds;
try
{
    creds = process.env.SHEETS_CREDENTIALS;
    creds = JSON.parse(creds.replace(/'/g, '"'));
    parsedCreds = true;
}
catch (e)
{
    parsedCredsError = new Error('Unable to parse Google sheets credentials. ' + e.message);
    parsedCredsError.code = 401;
    parsedCreds = false;
}

//Main export
module.exports = function (SHEETS_ID)
{
    var self = this;
    self.postError = postError;

    var starchSheet = new GoogleSpreadsheet(SHEETS_ID);
    return self;

    /* Public */
    function postError(data)
    {
        //If missing credentials, 404 error
        if (!hasCreds)
        {
            var missingCredsError = new Error('No Google sheets credentials found');
            missingCredsError.code = 404;
            return Promise.reject(missingCredsError);
        }
        //If cannot parse credentials, 401 error
        if (!parsedCreds)
        {
            if (parsedCredsError) return Promise.reject(parsedCredsError);
            else
            {
                var e = new Error('Unable to parse Google sheets credentials.');
                e.code = 401;
                return Promise.reject(e);
            }
        }
        if (!data || !data.type) return Promise.reject(new Error('Error data must have type specified'));

        return checkReady().then(function (sheet)
        {
            var worksheet = getWorksheetForError(data.type, sheet);

            //Notify caller that error type is unsupported by spreadsheet
            if (!worksheet)
            {
                var unsupported = new Error('Unable to find worksheet for error type: ' + data.type);
                unsupported.type = 'unsupported';
                unsupported.code = 404;
                throw unsupported;
            }
            return getRowsInWorksheet(worksheet);
        }).then(function (worksheet)
        {
            var matchingRow = findMatchingRow(data, worksheet.formattedRows);

            if (!matchingRow) return addRow(worksheet, data);
            else if (!matchingRow.date || !moment(matchingRow.date, 'MM/DD/YYYY').isSame(moment(), 'day'))
            {
                var rowToUpdate = worksheet.rowsById[matchingRow.id];
                if (!rowToUpdate) throw new Error('Could not update row with id: ' + row.id + ' for worksheet: ' + worksheet.title);
                return updateRow(rowToUpdate,
                {
                    date: moment().format('MM/DD/YYYY')
                });
            }
            else return;
        });
    }


    /* Private */
    function addRow(ws, data)
    {
        var newRow = {
            date: moment().format('MM/DD/YYYY')
        };
        for (var key in data)
        {
            if (!data.hasOwnProperty(key)) continue;
            newRow[key.toLowerCase()] = data[key];
        }
        return ws.addRow(newRow);
    }

    function updateRow(row, values)
    {
        for (key in values)
        {
            row[key] = values[key];
        }
        return row.save();
    }

    //Routes error to proper worksheet
    function getWorksheetForError(type, sheet)
    {
        if (!sheet.sheetInfo || !sheet.sheetInfo.worksheets || !sheet.sheetInfo.worksheets.length) return null;
        var ws = sheet.sheetInfo.worksheets.find(function (ws)
        {
            return ws.title.toLowerCase() === type.toLowerCase();
        });
        return ws;
    }

    function findMatchingRow(data, rows)
    {
        var updateDate;
        var matching;
        for (var i = 0; i < rows.length; i++)
        {
            var differentKeys = [];
            var row = rows[i];
            //Row is blank
            if (Object.keys(row).length === 1 && row.id) continue;
            for (var key in data)
            {
                //Compare values, ignoring date
                if (row[key.toLowerCase()] && data[key] && key.toLowerCase() !== 'date')
                {
                    //Only compare first 400 characters to see if error message is roughly the same
                    if (row[key.toLowerCase()].substring(0, 400) !== data[key].toString().substring(0, 400))
                    {
                        differentKeys.push(key);
                    }
                }
            }

            if (!differentKeys.length)
            {
                matching = row;
                break;
            }
        }
        return matching;
    }

    //Get the rows in current sheet
    function getRowsInWorksheet(ws)
    {
        var rows;
        return ws.getRows().then(function (theRows)
        {
            rows = theRows;

            //Determine the headers
            if (rows && rows[0] && rows[0]._xml) return getHeadersFromXML(ws, rows[0]._xml);
            else return getHeadersFromCells(ws);
        }).then(function (headers)
        {
            if (!headers || !headers.length) throw new Error('Unable to parse col headers for worksheet: ' + ws.title);

            var rowsById = {};

            //Formats the returned sheet data to json object, based on row column headers
            rows = rows.map(function (r)
            {
                rowsById[r.id] = r;
                var data = {
                    id: r.id,
                };

                headers.forEach(function (h)
                {
                    if (h in r) data[h] = r[h];
                });
                return data;
            });
            ws.rowsById = rowsById;
            ws.formattedRows = rows || [];
            return ws;
        });
    }

    //Grab headers from the first row's XML data
    function getHeadersFromXML(ws, xml)
    {
        //Format headers.  Sheets will lowercase headers
        var headers = xml.match(/<gsx:\w+>/g);
        if (!headers || !headers.length) throw new Error('Unable to parse col headers for worksheet: ' + ws.title);
        return headers.map(function (h)
        {
            return h.replace(/<gsx:|>/ig, '');
        });
    }

    //Grab headers from worksheet cell list (requires extra call, so xml preferred)
    function getHeadersFromCells(ws)
    {
        return ws.getCells().then(function (cells)
        {
            if (!cells || !cells.length) throw new Error('Unable to parse col headers for worksheet: ' + ws.title + '. Worksheet is empty');
            return cells.filter(function (c)
            {
                return c.row === 1;
            }).map(function (c)
            {
                return c.value;
            });
        });
    }


    //Check if requests have been authenticated.  If not, authenticate.
    function checkReady()
    {
        if (!self.ready || !starchSheet.sheetInfo) return authenticate(starchSheet);
        else return Promise.resolve(starchSheet);
    }

    function authenticate(sheet)
    {
        return sheet.useServiceAccountAuth(creds).then(function (res)
        {
            return sheet.getSpreadsheet();
        }).then(function (sheetInfo)
        {
            sheet.sheetInfo = sheetInfo;
            self.ready = true;
            return sheet;
        });
    }
};
