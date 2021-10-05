using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Environment = System.Environment;
using File = System.IO.File;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReferenceKeyTesting_WPF
{
    public class ExcelInteraction
    {
        public static readonly string Path = Environment.CurrentDirectory + @"\ReferenceKeyData.xlsx";
        private static Excel.Workbook _objBook;
        private static Excel.Worksheet _objSheet;
        private int _counter;
        public ExcelInteraction()
        {
            CreateExcel(Path);
        }
        private struct ColumnNumbers
        {
            public static readonly int[] BodyCol = { 1, 2, 3 };
            public static readonly int[] FaceCol = { 5, 6, 7 };
            public static readonly int[] EdgeCol = { 9, 10, 11 };
            public static readonly int[] ContextKeyCol = {13, 14, 15};
        }
        public void AddDataToExcel(DataPoints data)
        {
            try
            {
                int[] columnNumber = { };
                switch (data.EntityType)
                {
                    case DataPoints.Entity.SurfaceBody:
                        columnNumber = ColumnNumbers.BodyCol;
                        break;
                    case DataPoints.Entity.Face:
                        columnNumber = ColumnNumbers.FaceCol;
                        break;
                    case DataPoints.Entity.Edge:
                        columnNumber = ColumnNumbers.EdgeCol;
                        break;
                    case DataPoints.Entity.ContextKey:
                        columnNumber = ColumnNumbers.ContextKeyCol;
                        break;
                }
                _objSheet.Cells[data.Counter + 1, columnNumber[0]] = data.Name;
                _objSheet.Cells[data.Counter + 1, columnNumber[1]] = data.KeyContext;
                _objSheet.Cells[data.Counter + 1, columnNumber[2]] = data.RefKey;
                _counter++;
                MainWindow.Main.UpdateLabel = _counter.ToString();
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
        }
        private void CreateExcel(string path)
        {
            try
            {
                Excel.Application objExcelApp;
                Excel.Workbooks objBooks;
                if (!File.Exists(path))
                {
                    objExcelApp = new Excel.Application();
                    objBooks = objExcelApp.Workbooks;
                    _objBook = objBooks.Add(Missing.Value);
                    Excel.Sheets objSheets = _objBook.Worksheets;
                    _objSheet = (Excel.Worksheet)objSheets.Item[1];
                    _objSheet.Name = "TopologyMap";
                    Excel.Range aRange = _objSheet.Range["A1", "G100"];
                    aRange.Columns.AutoFit();
                    _objBook.SaveAs(path);
                }
                else
                {
                    if (Extension.IsFileOpen(path))
                    {
                        objExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    else
                    {
                        objExcelApp = new Excel.Application();
                    }
                    objBooks = objExcelApp.Workbooks;
                    _objBook = objBooks.Open(path);
                }
                _objBook.Activate();
                objExcelApp.Visible = true;
                _objSheet = _objBook.Worksheets["TopologyMap"];
                Headers(_objSheet);
            }

            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
        }
        private void Headers(Excel.Worksheet objSheet)
        {
            objSheet.Cells[1, 1] = "Body Identifier";
            objSheet.Cells[1, 2] = "keyContext";
            objSheet.Cells[1, 3] = "key";
            objSheet.Cells[1, 5] = "Face Identifier";
            objSheet.Cells[1, 6] = "keyContext";
            objSheet.Cells[1, 7] = "key";
            objSheet.Cells[1, 9] = "Edge Identifier";
            objSheet.Cells[1, 10] = "keyContext";
            objSheet.Cells[1, 11] = "key";
        }

        public void CloseExcel()
        {
            try
            {
                _objBook.Save();
                Excel.Application app = _objBook.Parent;
                _objBook.Close();
                Marshal.ReleaseComObject(_objBook);
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
        }

    }
}
