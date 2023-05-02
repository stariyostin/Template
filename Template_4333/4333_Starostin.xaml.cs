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
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;

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

        class Person
        {
            public int Id { get; set; }
            public string FullName { get; set; }
            public string CodeClient { get; set; }
            public string BirthDate { get; set; }
            public string Index { get; set; }
            public string City { get; set; }
            public string Street { get; set; }
            public int Home { get; set; }
            public int Kvartira { get; set; }
            public string E_mail { get; set; }
        }

        private async void btnJson(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json |*.json",
                Title = "Выберите файл"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            List<Person> list;

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                list = await JsonSerializer.DeserializeAsync<List<Person>>(fs);
            }
            using (ISRPO2Entities db = new ISRPO2Entities())
            {
                foreach (Person person in list)
                {
                    DateTime dateOfBirth = DateTime.Parse(person.BirthDate.ToString());
                    int age = DateTime.Today.Year - dateOfBirth.Year;
                    if (dateOfBirth > DateTime.Today.AddYears(-age))
                        age--;

                    db.TableLaba2.Add(new TableLaba2()
                    {
                        ФИО = person.FullName,
                        Код_клиента = Convert.ToInt32(person.CodeClient),
                        Дата_рождения = dateOfBirth,
                        Индекс = Convert.ToInt32(person.Index),
                        Город = person.City,
                        Улица = person.Street,
                        Дом = Convert.ToInt32(person.Home),
                        Квартира = Convert.ToInt32(person.Kvartira),
                        E_mail = person.E_mail,
                        Возраст = age,
                    });

                }
                db.SaveChanges();
            }
        }

        private void btnExportW(object sender, RoutedEventArgs e)
        {
            List<TableLaba2> people = new List<TableLaba2>();
            using (ISRPO2Entities db = new ISRPO2Entities())
            {
                people = db.TableLaba2.ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

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
                int i = 1;
                foreach (var category in allCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = "Категория " + i;
                    i++;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table peopleTable = document.Tables.Add(tableRange, category.Count() + 1, 3);
                    peopleTable.Borders.InsideLineStyle = peopleTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    peopleTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = peopleTable.Cell(1, 1).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = peopleTable.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = peopleTable.Cell(1, 3).Range;
                    cellRange.Text = "E-mail";
                    peopleTable.Rows[1].Range.Bold = 1;
                    peopleTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int j = 1;
                    foreach (var person in category)
                    {
                        cellRange = peopleTable.Cell(j + 1, 1).Range;
                        cellRange.Text = person.Код_клиента.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = peopleTable.Cell(j + 1, 2).Range;
                        cellRange.Text = person.ФИО;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = peopleTable.Cell(j + 1, 3).Range;
                        cellRange.Text = person.E_mail;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }

                    Word.Paragraph countPeopleParagraph = document.Paragraphs.Add();
                    Word.Range countPeopleRange = countPeopleParagraph.Range;
                    countPeopleRange.Text = $"Количество людей в категории - {category.Count()}";
                    countPeopleRange.Font.Color = Word.WdColor.wdColorRed;
                    countPeopleRange.InsertParagraphAfter();

                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;
                document.SaveAs2(@"C:\Users\sasha\Desktop\ISRPO3\outputFileWord.docx");
                document.SaveAs2(@"C:\Users\sasha\Desktop\ISRPO3\outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
            }
        }
    }
}
