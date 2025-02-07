using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace _4333Project
{
    /// <summary>
    /// Логика взаимодействия для _4333_Mustafina.xaml
    /// </summary>
    public partial class _4333_Mustafina : System.Windows.Window
    {
        public class UserInfo
        {
            public Users User { get; set; }
            public int Age
            {
                get
                {
                    DateTime dob;
                    if (DateTime.TryParse(User.DateOfBirth, out dob))
                    {
                        int age = DateTime.Now.Year - dob.Year;
                        if (dob > DateTime.Now.AddYears(-age)) age--;
                        return age;
                    }
                    return 0;
                }
            }
        }
        public _4333_Mustafina()
        {
            InitializeComponent();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
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
                {
                    for (int i = 0; i < _rows; i++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }
                int lastRow = 0;
                for (int i = 0; i < _rows; i++)
                {
                    if (list[i, 1] != string.Empty)
                    {
                        lastRow = i;
                    }
                }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
                using (MustafinaZR_4333Entities usersEntities = new MustafinaZR_4333Entities())
                {
                    for (int i = 1; i <= lastRow; i++)
                    {
                        var user = new Users()
                        {
                            Id = Convert.ToInt32(list[i, 1]),
                            FIO = list[i, 0],
                            DateOfBirth = list[i, 2],
                            IndexNum = Convert.ToInt32(list[i, 3]),
                            City = list[i, 4],
                            Street = list[i, 5],
                            House = Convert.ToInt32(list[i, 6]),
                            Apartment = Convert.ToInt32(list[i, 7]),
                            Email = list[i, 8]
                        };
                        usersEntities.Users.Add(user);
                    }
                    usersEntities.SaveChanges();
                }
                MessageBox.Show("Успешное импортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var allUsers = new List<UserInfo>();
                var groups = new[]
                {
                new { Id = 1, Name = "Категория 1  – от 20 до 29"},
                new { Id = 2, Name = "Категория 2  – от 30 до 39"},
                new { Id = 3, Name = "Категория 3  – от 40 "}
            };
                using (MustafinaZR_4333Entities usersEntities = new MustafinaZR_4333Entities())
                {
                    foreach (var user in usersEntities.Users)
                    {
                        allUsers.Add(new UserInfo { User = user });
                    }
                }

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = groups.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                var groupedUsers = allUsers.GroupBy(user =>
                {
                    int age = user.Age;
                    if (age >= 20 && age < 30) return 1;
                    else if (age >= 30 && age < 40) return 2;
                    else if (age >= 40) return 3;
                    else return 0;
                });

                for (int i = 0; i < groups.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(groups[i].Name);
                    worksheet.Cells[1][startRowIndex] = "Код клиента";
                    worksheet.Cells[2][startRowIndex] = "ФИО";
                    worksheet.Cells[3][startRowIndex] = "Email";
                    startRowIndex++;
                    foreach (var user in groupedUsers.FirstOrDefault(g => g.Key == i + 1))
                    {
                        worksheet.Cells[1][startRowIndex] = user.User.Id;
                        worksheet.Cells[2][startRowIndex] = user.User.FIO;
                        worksheet.Cells[3][startRowIndex] = user.User.Email;
                        startRowIndex++;
                    }
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
                MessageBox.Show("Успешное экспортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
