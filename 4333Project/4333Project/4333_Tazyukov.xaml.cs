using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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

namespace _4333Project
{
    /// <summary>
    /// Interaction logic for _4333_Tazyukov.xaml
    /// </summary>
    public partial class _4333_Tazyukov : Window
    {
        public _4333_Tazyukov()
        {
            InitializeComponent();
        }

        private void ButtonImport_Click(object sender, RoutedEventArgs e)
        {
            ExelReader.Read(
                new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                    Title = "Выберите файл базы данных"
                }.FileName
            );

            const string connectionString = "connection string";

            using(SqlConnection connection = new SqlConnection(connectionString))
            {
            }

    }
}
