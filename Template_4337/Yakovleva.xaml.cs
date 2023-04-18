using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
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
using OfficeOpenXml;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Yakovleva.xaml
    /// </summary>
    public partial class Yakovleva : Window
    {
        public Yakovleva()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Создание объекта для чтения xlsx-файла
            var package = new ExcelPackage(new FileInfo("C:\\Users\\tanus\\OneDrive\\Рабочий стол\\1.xlsx"));

            // Открытие соединения с базой данных
            using (t1Entities1 connection = new t1Entities1())
            {
                // Обход строк в xlsx-файле
                var worksheet = package.Workbook.Worksheets[1];
                for (int row = 2; row <= 13; row++)
                {
                    // Чтение данных из xlsx-файла
                    double id = Convert.ToDouble(worksheet.Cells[row, 1].Value);
                    string kodZakaza = (worksheet.Cells[row, 2].Value.ToString());
                    double dataSozdaniya = Convert.ToDouble(worksheet.Cells[row, 3].Value);
                    



                    myTable заказы = new myTable(id, kodZakaza, dataSozdaniya);
                    connection.myTables.Add(заказы);
                    connection.SaveChanges();
                    MessageBox.Show("Success! Data has been added to database!");
                }

            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           
            using (var connection = new SqlConnection("data source=LAPTOP-KH7CD52O;initial catalog=t1;integrated security=True;MultipleActiveResultSets=True"))
            {
                connection.Open();
                var command = new SqlCommand("SELECT [ID], [Наименование услуги], [Стоимость]\r\nFROM [t1].[dbo].[myTable]\r\nORDER BY [Стоимость]\r\n\r\n\r\n\r\n\r\n", connection);
                // Выборка данных из базы данных
                if ((bool)rd2.IsChecked)
                {
                     command = new SqlCommand("SELECT [ID], [Наименование услуги], [Стоимость]\r\nFROM [t1].[dbo].[myTable]\r\nORDER BY [Наименование услуги]\r\n\r\n\r\n\r\n\r\n", connection);
                }
                var dataReader = command.ExecuteReader();

                // Создание файла Excel и заполнение его данными
                var excelPackage = new ExcelPackage();
                var worksheet = excelPackage.Workbook.Worksheets.Add("Worksheet Name");
                worksheet.Cells.LoadFromDataReader(dataReader, true);

                // Сохранение файла Excel на диск
                var file = new FileInfo("C:\\Users\\tanus\\OneDrive\\Рабочий стол\\ExportedFile.xlsx");
                excelPackage.SaveAs(file);

                MessageBox.Show("Данные экспортированы успешно!");
            }
        }
    }
}
    