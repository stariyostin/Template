using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Starostin.xaml
    /// </summary>
    public partial class _4333_Starostin : Window
    {
        public _4333_Starostin()
        {
            InitializeComponent();
        }

        private void btnImport(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            //var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _rows = ObjWorkSheet.Cells[ObjWorkSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
            int _columns = ObjWorkSheet.Cells[1, ObjWorkSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            list = new string[_rows, _columns];

                for (int j = 0; j < _columns; j++)
                {
                        for (int i = 0; i < _rows; i++)
                        {
                            list[i, j] = ObjWorkSheet.Cells[i + 2, j + 1].Text;
                        }
                }
            ObjWorkBook.Close(false, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ISRPO2Entities isrpoEntities = new ISRPO2Entities())
            {
                for (int i = 0; i < _rows - 1; i++)
                {
                    DateTime dateOfBirth = DateTime.Parse(list[i, 2]);
                    int age = DateTime.Today.Year - dateOfBirth.Year;
                    if (dateOfBirth > DateTime.Today.AddYears(-age))
                        age--;
                    isrpoEntities.TableLaba2.Add(new TableLaba2()
                    {
                        ФИО = list[i, 0],
                        Код_клиента = Convert.ToInt32(list[i, 1]),
                        Дата_рождения = dateOfBirth,
                        Индекс = Convert.ToInt32(list[i, 3]),
                        Город = list[i, 4],
                        Улица = list[i, 5],
                        Дом = Convert.ToInt32(list[i, 6]),
                        Квартира = Convert.ToInt32(list[i, 7]),
                        E_mail = list[i, 8],
                        Возраст = age,
                    });
                }
                isrpoEntities.SaveChanges();
                MessageBox.Show("Успешный импорт");
            }
        }

        private void btnExport(object sender, RoutedEventArgs e)
        {
            List<TableLaba2> category_1;
            List<TableLaba2> category_2;
            List<TableLaba2> category_3;
            using (ISRPO2Entities isrpoEntities = new ISRPO2Entities())
            {
                category_1 = isrpoEntities.TableLaba2.Where(x => x.Возраст >= 20 && x.Возраст <= 29).ToList();
                category_2 = isrpoEntities.TableLaba2.Where(x => x.Возраст >= 30 && x.Возраст <= 39).ToList();
                category_3 = isrpoEntities.TableLaba2.Where(x => x.Возраст >= 40).ToList();
            }
            var allCategories = new List<List<TableLaba2>>()
            {
                category_1,
                category_2,
                category_3
            };
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < 3; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория {i + 1}";
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[1][startRowIndex].Font.Bold = true;
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[2][startRowIndex].Font.Bold = true;
                worksheet.Cells[3][startRowIndex] = "E-mail";
                worksheet.Cells[3][startRowIndex].Font.Bold = true;
                foreach (var person in allCategories[i])
                {
                    startRowIndex++;
                    worksheet.Cells[1][startRowIndex] = person.Код_клиента;
                    worksheet.Cells[2][startRowIndex] = person.ФИО;
                    worksheet.Cells[3][startRowIndex] = person.E_mail;
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
