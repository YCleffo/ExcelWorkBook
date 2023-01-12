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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorkBook
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void submitBtn_Click(object sender, RoutedEventArgs e)
        {
            /*создаем файл Excel*/

            Excel.Application aplication = new Excel.Application();
            aplication.Visible = true;


            /*количество листов*/

            aplication.SheetsInNewWorkbook = 1;

            /*добавляем рабочую книгу*/
            Excel.Workbook workbook = aplication.Workbooks.Add(Type.Missing);

            /*создаем лист*/
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.StandardWidth = 20;
            worksheet.Columns.ColumnWidth = 20;


            worksheet.Name = "Почта";

            /*заголовки вывод в Excel (в первую строку)*/
            worksheet.Cells[3][1] = "Почта";
            worksheet.Cells[3][1].Font.Size = 25;

            worksheet.Cells[2][3] = "Номер";
            worksheet.Cells[2][3].Font.Bold=true;
            worksheet.Cells[3][3] = "Наименование";
            worksheet.Cells[3][3].Font.Bold = true;
            worksheet.Cells[4][3] = "Дата отправки";
            worksheet.Cells[4][3].Font.Bold = true;

            worksheet.Cells[2][4] = "1290";
            worksheet.Cells[2][5] = "764";
            worksheet.Cells[2][6] = "6526";

            worksheet.Cells[3][4] = "Посылка";
            worksheet.Cells[3][5] = "Бандероль";
            worksheet.Cells[3][6] = "Письмо";

            worksheet.Cells[4][4] = "12.10.2015";
            worksheet.Cells[4][5] = "04.11.2012";
            worksheet.Cells[4][6] = "05.10.2012";

        }
    }
}
