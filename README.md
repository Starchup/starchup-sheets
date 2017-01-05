# Starchup Sheets
Simple helper for posting Starchup errors to Google sheets.


#### Installation

npm install starchup-sheets

#### Required environment vars
export SHEETS\_ID="id\_of\_google\_spreadsheet"

export SHEETS\_CREDENTIALS="one\_line\_string\_of\_downloaded\_credentials\_file\_from\_google"


#### Initialization

```
var starchupSheets = require('starchup-sheets');
var sheet = new starchupSheets();
```


#### Example
```
....
}).catch(function(e) {
	sheet.postError(e);
});
```

Error object must have `type` property that corresponds to title of appropriate worksheet,
and should have properties for any applicable headers in the worksheet.
