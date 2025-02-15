using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
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
using Word = Microsoft.Office.Interop.Word;

namespace _4333Project
{
    /// <summary>
    /// Логика взаимодействия для _4333_Mustafina.xaml
    /// </summary>
    public partial class _4333_Mustafina : System.Windows.Window
    {
        public class User
        {
            public User(string id, string fIO, string dateOfBirth, string indexNum, string city, string street, int house, int apartment, string email)
            {
                Id = id;
                FIO = fIO;
                DateOfBirth = dateOfBirth;
                IndexNum = indexNum;
                City = city;
                Street = street;
                House = house;
                Apartment = apartment;
                Email = email;
            }

            [JsonPropertyName("CodeClient")]
            public string Id { get; set; }

            [JsonPropertyName("FullName")]
            public string FIO { get; set; }

            [JsonPropertyName("BirthDate")]
            public string DateOfBirth { get; set; }

            [JsonPropertyName("Index")]
            public string IndexNum { get; set; }

            [JsonPropertyName("City")]
            public string City { get; set; }

            [JsonPropertyName("Street")]
            public string Street { get; set; }

            [JsonPropertyName("Home")]
            public int House { get; set; }

            [JsonPropertyName("Kvartira")]
            public int Apartment { get; set; }

            [JsonPropertyName("E_mail")]
            public string Email { get; set; }
        }

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

        private void JsonImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<User> persons = new List<User>();
                using (FileStream fs = new FileStream("C:/Users/user/Downloads/Telegram Desktop/3.json", FileMode.OpenOrCreate))
                {
                    persons = JsonSerializer.Deserialize<List<User>>(fs);
                }
                using (MustafinaZR_4333Entities usersEntities = new MustafinaZR_4333Entities())
                {
                    foreach (var person in persons)
                    {
                        var u = new Users()
                        {
                            Id = Convert.ToInt32(person.Id),
                            Apartment = Convert.ToInt32(person.Apartment),
                            City = person.City,
                            DateOfBirth = person.DateOfBirth,
                            Email = person.Email,
                            FIO = person.FIO,
                            House = person.House,
                            IndexNum = Convert.ToInt32(person.IndexNum),
                            Street = person.Street
                        };
                        usersEntities.Users.Add(u);
                    }
                    usersEntities.SaveChanges();
                }
                MessageBox.Show("Успешное импортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void JsonExportButton_Click(object sender, RoutedEventArgs e)
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

                var groupedUsers = allUsers.GroupBy(user =>
                {
                    int age = user.Age;
                    if (age >= 20 && age < 30) return 1;
                    else if (age >= 30 && age < 40) return 2;
                    else if (age >= 40) return 3;
                    else return 0;
                });
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                for (int i = 0; i < groups.Count(); i++)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(groups[i].Name);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    int count = groupedUsers.FirstOrDefault(g => g.Key == i + 1).Count();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table usersTable = document.Tables.Add(tableRange, count + 1, 3);
                    usersTable.Borders.InsideLineStyle = usersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    usersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = usersTable.Cell(1, 1).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = usersTable.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = usersTable.Cell(1, 3).Range;
                    cellRange.Text = "E-mail";
                    usersTable.Rows[1].Range.Bold = 1;
                    usersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int j = 1;
                    foreach (var user in groupedUsers.FirstOrDefault(g => g.Key == i + 1))
                    {
                        cellRange = usersTable.Cell(j + 1, 1).Range;
                        cellRange.Text = user.User.Id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = usersTable.Cell(j + 1, 2).Range;
                        cellRange.Text = user.User.FIO.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = usersTable.Cell(j + 1, 3).Range;
                        cellRange.Text = user.User.Email.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;
                document.SaveAs2("C:/Users/user/Desktop/outputFileWord.docx");
                document.SaveAs2("C:/Users/user/Desktop/outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
                MessageBox.Show("Успешное экспортирование данных", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
