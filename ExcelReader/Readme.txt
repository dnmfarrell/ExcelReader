ExcelReader Documentation

ExcelReader is a .Net class that is designed to make reading Excel files easy. It provides two objects for working with Excel: Book and Sheet.

Book Class
==========
Book is a class that represents an Excel workbook. It has the following methods:

New (string filepath to Excel file)
- Constructs a new Book object.

getSheetCount ()
- Returns a count of the number of worksheets in the Excel file.

getSheets ()
- Returns an array of Sheet objects.

getSheet (int index)
- Return a Sheet object at the given index location.


Sheet Class
===========
Sheet represents an Excel worksheet and provides the following methods:

getTitle ()
- Returns a string of the worksheet title

getIndex ()
- Returns an integer of the index location of the worksheet.

getRange (int startrow, int startcol, int endrow, int endcol)
- Returns a multidimensional string array representing the values in the requested range.

getColumn (int colIndex, int lastRow)
- Returns a string array of the values in the requested column.

getRow (int rowIndex, int lastColumn)
- Returns a string array of the values in the requested row.

