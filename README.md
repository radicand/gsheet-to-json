gsheet-to-json
==============

Convert your Google Doc spreadsheets to JSON, in the browser.  This is not an API method.

It is for you to work with your data in Google Spreadsheets, then copy out the resulting JSON conversion to paste somewhere.

Row 1 must dictate the type of data you have in that column:
Supported Types:

 * `string` (or just nothing at all.  Values will be quoted and escaped)
 * `number` (data will attempt to be converted using parseFloat)
 * `dict` (data in cell as {"keyhere": "Value Here", ...})
 * `boolean` (true/false -- not specially handled, but just works)
 * `list` (data in cell as ["item 1", "item 2"...].  Double quotes mandatory)
 * `date` (sets date as timestamp in your local timezone)
 * `date-#` (sets date in GMT-# timestamp)

Row 2 must be the column headers.

Rows 3+ are your data.

Fork of code at http://code.google.com/p/pamelafox-samplecode/


Usage
=====

1. Format your data as per the above
2. Highlight your columns
3. Go to "Insert"->"Gadget" and choose Custom.  Add url `https://raw.github.com/radicand/gsheet-to-json/master/gadget.xml`
4. Be sure to re-highlight your columns when prompted, if needed
5. Enjoy!