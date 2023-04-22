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
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using OfficeOpenXml;
using Microsoft.Office.Interop.Word;

namespace Template_4337
{
    public class Services
    {
        public int IdServices { get; set; }
        public string NameServices { get; set; }
        public string TypeOfService { get; set; }
        public string CodeService { get; set; }
        public int Cost { get; set; }
    }

    /// <summary>
    /// Логика взаимодействия для Yakovleva.xaml
    /// </summary>
    public partial class Yakovleva : System.Windows.Window
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
                    MessageBox.Show("Данные успешно импортированы!");
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

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            string json = "[{\"idservices\":1,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u043b\\u044b\\u0436\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"8hfjhg443\",\"cost\":1000},{\"idservices\":2,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u043b\\u044b\\u0436\\u043d\\u044b\\u0445 \\u043f\\u0430\\u043b\\u043e\\u043a\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"87fdjkhj\",\"cost\":100},{\"idservices\":3,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u0441\\u043d\\u043e\\u0443\\u0431\\u043e\\u0440\\u0434\\u0430\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"jur8r\",\"cost\":1200},{\"idservices\":4,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u043e\\u0431\\u0443\\u0432\\u0438 \\u0434\\u043b\\u044f \\u0441\\u043d\\u043e\\u0443\\u0431\\u043e\\u0440\\u0434\\u0430\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"jkfbj09\",\"cost\":400},{\"idservices\":5,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u0448\\u043b\\u0435\\u043c\\u0430\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"63748hf\",\"cost\":300},{\"idservices\":6,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u0437\\u0430\\u0449\\u0438\\u0442\\u043d\\u044b\\u0445 \\u043f\\u043e\\u0434\\u0443\\u0448\\u0435\\u043a \\u0434\\u043b\\u044f \\u0441\\u043d\\u043e\\u0443\\u0431\\u043e\\u0440\\u0434\\u0438\\u0441\\u0442\\u043e\\u0432\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"jfh7382\",\"cost\":300},{\"idservices\":7,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u043e\\u0447\\u043a\\u043e\\u0432 \\u0434\\u043b\\u044f \\u043b\\u044b\\u0436\\u043d\\u0438\\u043a\\u043e\\u0432\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"oijnb12\",\"cost\":150},{\"idservices\":8,\"nameservices\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442 \\u0432\\u0430\\u0442\\u0440\\u0443\\u0448\\u043a\\u0438\",\"typeofservice\":\"\\u041f\\u0440\\u043e\\u043a\\u0430\\u0442\",\"codeservice\":\"bsfbhv63\",\"cost\":200},{\"idservices\":9,\"nameservices\":\"\\u041e\\u0431\\u0443\\u0447\\u0435\\u043d\\u0438\\u0435 \\u043a\\u0430\\u0442\\u0430\\u043d\\u0438\\u044e \\u043d\\u0430 \\u0433\\u043e\\u0440\\u043d\\u044b\\u0445 \\u043b\\u044b\\u0436\\u0430\\u0445\",\"typeofservice\":\"\\u041e\\u0431\\u0443\\u0447\\u0435\\u043d\\u0438\\u0435\",\"codeservice\":\"hjbuje21j\",\"cost\":1000},{\"idservices\":10,\"nameservices\":\"\\u041e\\u0431\\u0443\\u0447\\u0435\\u043d\\u0438\\u0435 \\u043a\\u0430\\u0442\\u0430\\u043d\\u0438\\u044e \\u043d\\u0430 \\u0441\\u043d\\u043e\\u0443\\u0431\\u043e\\u0440\\u0434\\u0435\",\"typeofservice\":\"\\u041e\\u0431\\u0443\\u0447\\u0435\\u043d\\u0438\\u0435\",\"codeservice\":\"dhbgfy563\",\"cost\":1000},{\"idservices\":11,\"nameservices\":\"\\u041f\\u043e\\u0434\\u044a\\u0435\\u043c \\u043d\\u0430 1 \\u0443\\u0440\\u043e\\u0432\\u0435\\u043d\\u044c\",\"typeofservice\":\"\\u041f\\u043e\\u0434\\u044a\\u0435\\u043c\",\"codeservice\":\"jhvsjf6\",\"cost\":500},{\"idservices\":12,\"nameservices\":\"\\u041f\\u043e\\u0434\\u044a\\u0435\\u043c \\u043d\\u0430 2  \\u0443\\u0440\\u043e\\u0432\\u0435\\u043d\\u044c\",\"typeofservice\":\"\\u041f\\u043e\\u0434\\u044a\\u0435\\u043c\",\"codeservice\":\"djhgbs982\",\"cost\":750}]";
            List<Services> services = JsonConvert.DeserializeObject<List<Services>>(json);

            using (t1Entities1 db = new t1Entities1())
            {
                foreach (var service in services)
                {
                    myTable myTable = new myTable(service.IdServices, service.NameServices, service.Cost);
                    db.myTables.Add(myTable);
                }

                db.SaveChanges();
                MessageBox.Show("Данные успешно экспортированы!");

            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
           

  
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // создать новый документ Word
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add();

            using (t1Entities1 connection = new t1Entities1())
            {
                int i = 0;
                var services = (bool)RD1.IsChecked
                               ? connection.myTables.OrderBy(s => s.Стоимость).ToList()
                               : connection.myTables.OrderBy(s => s.Наименование_услуги).ToList();

                Microsoft.Office.Interop.Word.Table table = wordDoc.Tables.Add(wordDoc.Range(), services.Count, 6);

                foreach (var service in services)
                {
                    table.Cell(i, 1).Range.Text = "ID: " + service.ID;
                    table.Cell(i, 2).Range.Text = "Наименование: " + service.Наименование_услуги;
                    table.Cell(i, 3).Range.Text = "Стоимость: " + service.Стоимость;

                    i++;
                }

                wordDoc.SaveAs("C:\\Users\\tanus\\OneDrive\\Рабочий стол\\MyDocument.docx");

                // закрыть приложение Word
                wordApp.Quit();
                MessageBox.Show("Данные успешно импортированы!");
            }

        }
    }
}
    