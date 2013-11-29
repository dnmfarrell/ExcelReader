using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class Sheet
    {
        private Worksheet xlWorksheet;
        private Application xlApp;

        internal Sheet(Application app, Worksheet sheet)
        {
            xlApp = app;
            xlWorksheet = sheet;
        }

        ~Sheet()
        {
            releaseObject(xlWorksheet);
        }

        public String getTitle()
        {
            return xlWorksheet.Name;
        }

        public String getIndex()
        {
            return Convert.ToString(xlWorksheet.Index);
        }

        public String[,] getRange(int startrow, int startcol, int endrow, int endcol)
        {
            String[,] data = new String[endrow - startrow + 1, endcol - startcol + 1];

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        data[i, j] = (((Range)xlWorksheet.Cells[i + startrow, j + startcol]).Value2 != null
                                    ? ((Range)xlWorksheet.Cells[i + startrow, j + startcol]).Value2.ToString()
                                    : ""); 
                    }
                }
            return data;
        }

        public string[] getColumn(int colIndex, int lastRow)
        {
            string[,] range = getRange(1, colIndex, lastRow, colIndex);
            string[] column = new string[range.GetLength(0)];
            for (int i = 0; i < column.Length; i++)
            {
                column[i] = range[i, 0];
            }
            return column;
        }

        public string[] getRow(int rowIndex, int lastColumn)
        {
            string[,] range = getRange(rowIndex, 1, rowIndex, lastColumn);
            string[] row = new string[range.GetLength(1)];
            for (int i = 0; i < row.Length; i++)
            {
                row[i] = range[0, i];
            }
            return row;
        }

        public void save(string filepath)
        {
            object misValue = System.Reflection.Missing.Value;
            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Worksheet sheet1 = xlWorkBook.Worksheets.get_Item(1);
            xlWorksheet.Copy(Type.Missing, sheet1);
            sheet1.Delete();

            //save
            xlWorkBook.SaveAs(filepath, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            releaseObject(sheet1);
            releaseObject(xlWorkBook);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
