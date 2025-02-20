using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace _4333Project
{
    public static class ExelReader
    {
        public static string [,] Read(string fileName)
        {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            string[,] list = new string[_rows, _columns];
            for(int j = 0; j < _columns; j++) for(int i = 0; i < _rows; i++) list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            return list;
        }
    }
}
