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
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using Newtonsoft.Json;

namespace _4333Project
{
    /// <summary>
    /// Логика взаимодействия для _4333_Garifullin.xaml
    /// </summary>
    public partial class _4333_Garifullin : System.Windows.Window
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
            foreach (string name in StatusNames)
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

        private void WordExport(object sender, RoutedEventArgs e)
        {
            List<Rent> allRents;
            using (UserEntities usersEntitiesEntities = new UserEntities())
            {
                allRents = usersEntitiesEntities.Rents.ToList().OrderBy(s => s.Statuses).ToList();
                var RentGroups = allRents.GroupBy(s => s.Statuses).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                foreach (var group in RentGroups)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(allRents.Where(g => g.Statuses == group.Key).FirstOrDefault().Statuses);
                    paragraph.set_Style("Заголовок 2");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table RentsTable = document.Tables.Add(tableRange, group.Count() + 1, 5);
                    RentsTable.Borders.InsideLineStyle =
                    RentsTable.Borders.OutsideLineStyle =
                    Word.WdLineStyle.wdLineStyleSingle;
                    RentsTable.Range.Cells.VerticalAlignment =
                    Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = RentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Id";
                    cellRange = RentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = RentsTable.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = RentsTable.Cell(1, 4).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = RentsTable.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    RentsTable.Rows[1].Range.Bold = 1;
                    RentsTable.Rows[1].Range.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int i = 1;
                    foreach (var rent in group)
                    {
                        cellRange = RentsTable.Cell(i + 1, 1).Range;
                        cellRange.Text = rent.Id.ToString();
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = RentsTable.Cell(i + 1, 2).Range;
                        cellRange.Text = rent.Rent_Code;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = RentsTable.Cell(i + 1, 3).Range;
                        cellRange.Text = rent.Date_of_creation;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = RentsTable.Cell(i + 1, 4).Range;
                        cellRange.Text = rent.Client_Code.ToString();
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = RentsTable.Cell(i + 1, 5).Range;
                        cellRange.Text = rent.Servicess;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;
            }

        }

        private void JsonImport(object sender, RoutedEventArgs e)
        {
            List<Rent> allRents = new List<Rent>();
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            var json = File.ReadAllText(ofd.FileName);
                allRents = JsonConvert.DeserializeObject<List<Rent>>(json);
                using (UserEntities usersEntities = new UserEntities())
                {
                    foreach (var rent in allRents)
                    {
                            Rent u = new Rent()
                            {
                                Id = Convert.ToInt32(rent.Id),
                                Rent_Code = rent.Rent_Code,
                                Date_of_creation = rent.Date_of_creation,
                                Creation_Time = rent.Creation_Time,
                                Client_Code = Convert.ToInt32(rent.Client_Code),
                                Servicess = rent.Servicess,
                                Statuses = rent.Statuses,
                                Date_of_closing = rent.Date_of_closing,
                                Rent_time = rent.Rent_time,
                            };
                        usersEntities.Rents.Add(u);
                    }
                    usersEntities.SaveChanges();
                }
        }
    }
}
