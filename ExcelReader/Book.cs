using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class Book
    {
        private Workbook xlWorkbook;
        private Application xlApp;

        public Book(string filepath)
        {
            string fullpath = Path.GetFullPath(filepath);
            if (!File.Exists(fullpath))
            {
                Console.WriteLine("\nError loading file: " + fullpath);
                Environment.Exit(0);
            }
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Application();

            try
            {
                xlWorkbook = xlApp.Workbooks.Open(fullpath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nError opening file: " + fullpath + "\n" + ex);
                Environment.Exit(0);
            }
        }

        public int getSheetCount()
        {
            return xlWorkbook.Worksheets.Count;
        }

        public Sheet[] getSheets()
        {
            Sheet[] sheets = new Sheet[xlWorkbook.Worksheets.Count];
            for (int i = 0; i < sheets.Length; i++)
            {
                sheets[i] = getSheet(i + 1);
            }
            return sheets;
        }

        public Sheet getSheet(int i) 
        {
            return new Sheet(xlApp, xlWorkbook.Worksheets.get_Item(i));
        }

        ~Book()
        {
            object misValue = System.Reflection.Missing.Value;
            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkbook);
            releaseObject(xlApp);
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