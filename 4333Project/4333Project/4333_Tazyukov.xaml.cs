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
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            openFileDialog.ShowDialog();
            var list = ExelReader.Read(openFileDialog.FileName);

            const string connectionString = "Server=DESKTOP-40D8MST\\Maksim;Database=test_DB;Trusted_Connection=True;";

            using(SqlConnection connection = new SqlConnection(connectionString))
            {
                //DataBaseInteractor.Copy(list, connection);

                using(SqlCommand command = new SqlCommand("INSERT INTO user VALUES (1, 2, 3, 4, 5)", connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
