using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Win32;

namespace _4333Project
{
    /// <summary>
    /// Логика взаимодействия для _4333_Garifullin.xaml
    /// </summary>
    public partial class _4333_Garifullin : Window
    {
        public _4333_Garifullin()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (UserEntities usersEntities = new UserEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (list[i, 0] != "" && list[i, 0] != " ")
                    {
                        usersEntities.Rents.Add(new Rent()
                        {
                            Id = Convert.ToInt32(list[i, 0]),
                            Rent_Code = list[i, 1],
                            Date_of_creation = list[i, 2],
                            Creation_Time = list[i, 3],
                            Client_Code = Convert.ToInt32(list[i, 4]),
                            Servicess = list[i, 5],
                            Statuses = list[i, 6],
                            Date_of_closing = list[i, 7],
                            Rent_time = list[i, 8],
                        });
                    }
                }
                usersEntities.SaveChanges();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List<Rent> allRents;
            List<string> Names = new List<string>();
            using (UserEntities usersEntitiesEntities = new UserEntities())
            {
                allRents = usersEntitiesEntities.Rents.ToList().OrderBy(s => s.Statuses).ToList();
            }
            foreach (var rent in allRents)
            {
                Names.Add(rent.Statuses);
            }
            var StatusNames = Names.Distinct();
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = StatusNames.Count();
            Excel.Workbook wb = app.Workbooks.Add(Type.Missing);
            int g = 0;
            var StatusRents = allRents.GroupBy(s => s.Statuses).ToList();
            foreach ( string name in StatusNames)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[g + 1];
                worksheet.Name = name;
                worksheet.Cells[1][1] = "Id";
                worksheet.Cells[2][1] = "Код заказа";
                worksheet.Cells[3][1] = "Дата создания";
                worksheet.Cells[4][1] = "Код клиента";
                worksheet.Cells[5][1] = "Услуги";
                worksheet.Cells[1][1].Font.Bold = true;
                worksheet.Cells[2][1].Font.Bold = true;
                worksheet.Cells[3][1].Font.Bold = true;
                worksheet.Cells[4][1].Font.Bold = true;
                worksheet.Cells[5][1].Font.Bold = true;
                startRowIndex++;
                g++;
                foreach (var rents in StatusRents)
                {
                    if (rents.Key == name)
                    {
                        foreach (Rent rent in allRents)
                        {
                            if (rent.Statuses == rents.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = rent.Id;
                                worksheet.Cells[2][startRowIndex] = rent.Rent_Code;
                                worksheet.Cells[3][startRowIndex] = rent.Date_of_creation;
                                worksheet.Cells[4][startRowIndex] = rent.Client_Code;
                                worksheet.Cells[5][startRowIndex] = rent.Servicess;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
                    worksheet.Columns.AutoFit();
                }
            }
            app.Visible = true;
        }
    }
}
