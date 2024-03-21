using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

using System.Text.Json;
using System.Linq.Expressions;
namespace Template4333
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _4333_Kulikova win = new _4333_Kulikova();
            win.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
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
                    for (int i = 0; i < _rows; i++)
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
                using (ISRPOEntities4 usersEntities = new ISRPOEntities4())
                {

                    for (int i = 1; i < _rows; i++)
                    {
                        if (list[i, 1] != "" && list[i, 2] != "" && list[i, 3] != "" && list[i, 4] != "")
                        {

                            usersEntities.Ord.Add(new Ord()
                            {

                                Id_order = list[i, 1],
                                Date_of_creation = list[i, 2],
                                Creation_time = list[i, 3],
                                Id_client = list[i, 4],
                                Services = list[i, 5],
                                Status = list[i, 6],
                                Closing_date = list[i, 7],
                                Rental_time = list[i, 8]
                            });
                        }
                    }
                    usersEntities.SaveChanges();
                    MessageBox.Show("Данные успешно добавлены", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            { MessageBox.Show("Произошла ошибка при добавлении данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                List<Ord> allorders;
                List<string> status;
                using (ISRPOEntities4 usersEntities = new ISRPOEntities4())
                {
                    allorders = usersEntities.Ord.ToList().OrderBy(s => s.Status).ToList();
                    status = usersEntities.Ord.ToList().Select(Ord => Ord.Status.ToString()).Distinct().ToList();
                }
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = status.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                for (int i = 0; i < status.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(status[i]);
                    worksheet.Cells[1][startRowIndex] = "ID";
                    worksheet.Cells[2][startRowIndex] = "Код заказа";
                    worksheet.Cells[3][startRowIndex] = "Дата создания";
                    worksheet.Cells[4][startRowIndex] = "Код клиента";
                    worksheet.Cells[5][startRowIndex] = "Услуги";
                    startRowIndex++;
                    foreach (var order in allorders)
                    {
                        if (order.Status == status[i])
                        {
                            worksheet.Cells[1][startRowIndex] = order.Id.ToString();
                            worksheet.Cells[2][startRowIndex] = order.Id_order;
                            worksheet.Cells[3][startRowIndex] = order.Date_of_creation;
                            worksheet.Cells[4][startRowIndex] = order.Id_client;
                            worksheet.Cells[5][startRowIndex] = order.Services;
                            startRowIndex++;
                        }
                    }

                }
                app.Visible = true;
            }
            catch (Exception ex) { MessageBox.Show("Произошла ошибка при экспорте данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }

        }

        private async void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON files (*.json)|*.json";

                if (openFileDialog.ShowDialog() == true)
                {
                    string jsonFilePath = openFileDialog.FileName;

                    List<Ord> ordersData;

                    using (FileStream fs = new FileStream(jsonFilePath, FileMode.Open))
                    {
                        ordersData = await JsonSerializer.DeserializeAsync<List<Ord>>(fs);
                    }

                    using (ISRPOEntities4 usersEntities = new ISRPOEntities4())
                    {
                        foreach (var orderData in ordersData)
                        {
                            Ord newOrder = new Ord
                            {
                                Id_order = orderData.Id_order,
                                Date_of_creation = orderData.Date_of_creation,
                                Creation_time = orderData.Creation_time,
                                Id_client = orderData.Id_client,
                                Services = orderData.Services,
                                Status = orderData.Status,
                                Closing_date = orderData.Closing_date,
                                Rental_time = orderData.Rental_time
                            };
                            usersEntities.Ord.Add(newOrder);
                        }
                        usersEntities.SaveChanges();
                    }

                    MessageBox.Show("Данные успешно импортированы из JSON файла в таблицу БД.");
                }
            }
            catch (Exception ex) { MessageBox.Show("Произошла ошибка при добавлении данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
        }
    

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            try
            {
                List<Ord> allorders;

                using (ISRPOEntities4 usersEntities = new ISRPOEntities4())
                {
                    allorders = usersEntities.Ord.ToList().OrderBy(s => s.Status).ToList();

                }
                foreach (var group in allorders.GroupBy(o => o.Status))
                {
                    var app = new Word.Application();
                    Word.Document document = app.Documents.Add();

                    Word.Paragraph headerParagraph = document.Paragraphs.Add();
                    Word.Range headerRange = headerParagraph.Range;
                    headerRange.Text = $"Статус заказа: {group.Key}";
                    headerParagraph.set_Style("Заголовок 1");
                    headerRange.InsertParagraphAfter();

                    Word.Table ordersTable = document.Tables.Add(headerRange, group.Count() + 1, 5);
                    ordersTable.Borders.InsideLineStyle = ordersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    ordersTable.Rows[1].Range.Font.Bold = 1;
                    ordersTable.Rows[1].Range.Font.Italic = 1;
                    ordersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    ordersTable.Cell(1, 1).Range.Text = "Id";
                    ordersTable.Cell(1, 2).Range.Text = "Код заказа";
                    ordersTable.Cell(1, 3).Range.Text = "Дата создания";
                    ordersTable.Cell(1, 4).Range.Text = "Код клиента";
                    ordersTable.Cell(1, 5).Range.Text = "Услуги";

                    int i = 1;
                    foreach (var order in group)
                    {
                        i++;
                        ordersTable.Cell(i, 1).Range.Text = order.Id.ToString();
                        ordersTable.Cell(i, 2).Range.Text = order.Id_order;
                        ordersTable.Cell(i, 3).Range.Text = order.Date_of_creation.ToString();
                        ordersTable.Cell(i, 4).Range.Text = order.Id_client.ToString();
                        ordersTable.Cell(i, 5).Range.Text = order.Services;
                    }

                    string fileName = $"C:/Users/roni0/Desktop/ISRPO/outputFileWord_{group.Key}.docx";
                    document.SaveAs2(fileName);

                    app.Visible = true;
                }

            }
            catch (Exception ex) { MessageBox.Show("Произошла ошибка при экспорте данных: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }


        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {

        }
    }
}
