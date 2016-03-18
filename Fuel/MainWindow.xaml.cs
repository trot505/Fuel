using System.Windows;
using System.Windows.Forms;
using System.ComponentModel;
using System.Windows.Controls;
using System.Linq;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;


namespace Fuel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xamld
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Company> company = new ObservableCollection<Company>();
        private const string DIR_SEPARATOR = @"/";
        private string companyPatch = @"company.json";
        private string cellPatch = @"optionxls.json";
        cellExcel cell = new cellExcel();
        exelComp exc = new exelComp();
        List<Out> outArr = new List<Out>();
        Newtonsoft.Json.Linq.JObject arrOpt;



        public MainWindow()
        {
            InitializeComponent();
            if (File.Exists(companyPatch))
            {
                string fileC = File.ReadAllText(companyPatch, UTF8Encoding.UTF8);
                company = JsonConvert.DeserializeObject<ObservableCollection<Company>>(fileC);
                CompanyGrid.ItemsSource = company;
                CompanyGrid.UnselectAllCells();
            }
            else {
                System.Windows.MessageBox.Show("Файл со списком организайи не существует!", "ВНИМАНИЕ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            arrOpt = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(cellPatch, Encoding.UTF8));
            exc = JsonConvert.DeserializeObject<exelComp>(arrOpt["excelComp"].ToString());

        }


        //кнопка добавление взаимосявзи в списко компаний Наименование->Башнефть->Лукойл->
        private void add_Click(object sender, RoutedEventArgs e)
        {
            addcompany tr = new addcompany();
            if (tr.ShowDialog() == true)
            {
                company.Add(new Company(tr.name.Text, tr.fullName.Text, tr.nameBash.Text, tr.nameLuk.Text));
                tr.Close();
            }
            searchText.Text = "";
            clearSearch.Visibility = Visibility.Hidden;
        }

        // метод удаление взаимосвязи комании
        private void deleteCompany()
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
            }
        }

        //метод изменение взаимосвязи компании
        private void updateCompany(string n, string fn, string nb, string nl)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                Company newC = new Company(n, fn, nb, nl);
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
                company.Add(newC);
            }
        }

        //поиск по таблице взаимосвязей компаний
        private void searchText_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            var s = company.Where<Company>(r => (r.Name + r.FullName + r.NameBash + r.NameLuk).ToLower().Contains(searchText.Text.ToLower()));
            CompanyGrid.ItemsSource = (s.SequenceEqual(company)) ? company : s;
            clearSearch.Visibility = Visibility.Visible;
        }

        //кнопка измениения взаимосвязи компании
        private void updateC_Click(object sender, RoutedEventArgs e)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                addcompany tr = new addcompany();
                Company c = CompanyGrid.SelectedItem as Company;
                tr.name.Text = c.Name;
                tr.fullName.Text = c.FullName;
                tr.nameBash.Text = c.NameBash;
                tr.nameLuk.Text = c.NameLuk;
                if (tr.ShowDialog() == true)
                {
                    updateCompany(tr.name.Text, tr.fullName.Text, tr.nameBash.Text, tr.nameLuk.Text);
                }
                tr.Close();
            }
            else {
                System.Windows.MessageBox.Show("Необходимо выбрать фирму из списка.", "ВНИМЕНИЕ !!!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            CompanyGrid.UnselectAllCells();
        }

        //кнопка удаления взаимосвязи
        private void delC_Click(object sender, RoutedEventArgs e)
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                var res = System.Windows.MessageBox.Show("Вы собираетесьб удалить запись: \n"
                + (CompanyGrid.SelectedItem as Company).Name, "УДАЛЕНИЕ", MessageBoxButton.YesNo);
                if (res == MessageBoxResult.Yes)
                {
                    deleteCompany();
                    System.Windows.MessageBox.Show("Удаление записи прошло успешно", "СООБЩЕНИЕ", MessageBoxButton.OK);
                }
            }
            else {
                System.Windows.MessageBox.Show("Необходимо выбрать фирму из списка.", "ВНИМЕНИЕ !!!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            CompanyGrid.UnselectAllCells();
            searchText.Text = "";
        }

        //обработка при закрытии программы
        private void Window_Closed(object sender, EventArgs e)
        {
            // очищаем поле поиска компании
            searchText.Text = "";
            CompanyGrid.ItemsSource = company;
            if (!File.Exists(companyPatch))
            {
                File.Create(companyPatch);
            }
            //внесение изменений в файл json взаимосвязей компании
            using (StreamWriter file = new StreamWriter(companyPatch, false))
            {
                file.WriteLine(JsonConvert.SerializeObject(company));
            }
        }

        //внесение данных из массива номеров колонок (файла отчета) в поля формы для парсинга отчета
        private void CellText()
        {
            card.Text = cell.CellCard.ToString();
            holder.Text = cell.CellCompany.ToString();
            azs.Text = cell.CellAzs.ToString();
            adres.Text = cell.CellAdressAzs.ToString();
            date.Text = cell.CellDateFill.ToString();
            fuelT.Text = cell.CellFuelT.ToString();
            countF.Text = cell.CellCountF.ToString();
            operation.Text = cell.CellOperation.ToString();
            first.Text = cell.FirstRow.ToString();
            last.Text = cell.LastRow.ToString();
            folderPatch.Text = cell.FolderPatch.ToString();
            folderMonth.SelectedIndex = cell.FolderMonth;
            listExl.Text = cell.ListExl.ToString();
        }

        //клик по радиокнопке Башнефть формирование данных номеров колонок
        private void radioBash_Checked(object sender, RoutedEventArgs e)
        {
            //получение данных из файла optionxls.json касательно данных колонок Башнефть      
            cell = JsonConvert.DeserializeObject<cellExcel>(arrOpt["Bash"].ToString());
            CellText();
        }

        //клик по радиокнопке Лукойл формирование данных номеров колонок
        private void radioLuk_Checked(object sender, RoutedEventArgs e)
        {
            //получение данных из файла optionxls.json касательно данных колонок Лукойл            
            cell = JsonConvert.DeserializeObject<cellExcel>(arrOpt["Luk"].ToString());
            CellText();
        }

        //Удаляет лишние пробелы и переносы строк(заменяя на 1 пробел):	
        private string RemoveSpaces(string txt)
        {
            txt = txt.Trim();
            txt = txt.Replace("  ", " ");

            if (txt.Contains("  "))
                RemoveSpaces(txt);
            return txt;
        }

        //парсинг файла отчета для формирования List Out (всех транзакций) 
        private void parseExcelReport()
        {
            pgText.Text = "ФОРМИРОВАНИЕ МАССИВА ДАННЫХ \nИЗ ФАЙЛА ОБЩЕГО ОТЧЕТА";
            System.Windows.Forms.Application.DoEvents();

            //очишаем спарсенный массив excel данных из файла поставщика
            outArr.Clear();
            if (File.Exists(fileName.Text))
            {
                try
                {
                    //если новый офис
                    using (ExcelPackage execPac = new ExcelPackage(new FileInfo(fileName.Text)))
                    {
                        ExcelWorksheet execPage = execPac.Workbook.Worksheets[cell.ListExl];

                        for (int i = cell.FirstRow; i <= cell.LastRow; i++)
                        {
                            var c = (execPage.Cells[i, cell.CellCard].Value == null) ? "" : execPage.Cells[i, cell.CellCard].Value.ToString();
                            var s = (execPage.Cells[i, cell.CellAzs].Value == null) ? "" : execPage.Cells[i, cell.CellAzs].Value.ToString();
                            var a = (execPage.Cells[i, cell.CellAdressAzs].Value == null) ? "" : execPage.Cells[i, cell.CellAdressAzs].Value.ToString();
                            var d = (execPage.Cells[i, cell.CellDateFill].Value == null) ? "" : execPage.Cells[i, cell.CellDateFill].Value.ToString();
                            var o = (execPage.Cells[i, cell.CellOperation].Value == null) ? "" : execPage.Cells[i, cell.CellOperation].Value.ToString();
                            var t = (execPage.Cells[i, cell.CellFuelT].Value == null) ? "" : execPage.Cells[i, cell.CellFuelT].Value.ToString();
                            var co = (execPage.Cells[i, cell.CellCountF].Value == null) ? "" : execPage.Cells[i, cell.CellCountF].Value.ToString();
                            var n = (execPage.Cells[i, cell.CellCompany].Value == null) ? "" : execPage.Cells[i, cell.CellCompany].Value.ToString();

                            outArr.Add(new Out(c, s, a, d, o, t, co, n));

                            pgBar.Value = (i * 100) / cell.LastRow;
                            System.Windows.Forms.Application.DoEvents();
                        }
                        execPac.Dispose();
                    }
                }
                catch (Exception)
                {
                    //если старый офис

                    //Создаём приложение.
                    Excel.Application ObjExcel = new Excel.Application();
                    //Открываем книгу.                                                                                                                                                       
                    Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName.Text, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Выбираем таблицу(лист).
                    Excel.Worksheet exePage;
                    exePage = (Excel.Worksheet)ObjWorkBook.Sheets[cell.ListExl];
                    
                    for (int i = cell.FirstRow; i <= cell.LastRow; i++)
                    {
                        //Выбираем область таблицы. (в нашем случае просто ячейку)
                        var c = (exePage.Cells[i, cell.CellCard].Value == null) ? "" : exePage.Cells[i, cell.CellCard].Value.ToString();
                        var s = (exePage.Cells[i, cell.CellAzs].Value == null) ? "" : exePage.Cells[i, cell.CellAzs].Value.ToString();
                        var a = (exePage.Cells[i, cell.CellAdressAzs].Value == null) ? "" : exePage.Cells[i, cell.CellAdressAzs].Value.ToString();
                        var d = (exePage.Cells[i, cell.CellDateFill].Value == null) ? "" : exePage.Cells[i, cell.CellDateFill].Value.ToString();
                        var o = (exePage.Cells[i, cell.CellOperation].Value == null) ? "" : exePage.Cells[i, cell.CellOperation].Value.ToString();
                        var t = (exePage.Cells[i, cell.CellFuelT].Value == null) ? "" : exePage.Cells[i, cell.CellFuelT].Value.ToString();
                        var co = (exePage.Cells[i, cell.CellCountF].Value == null) ? "" : exePage.Cells[i, cell.CellCountF].Value.ToString();
                        var n = (exePage.Cells[i, cell.CellCompany].Value == null) ? "" : exePage.Cells[i, cell.CellCompany].Value.ToString();

                        outArr.Add(new Out(c, s, a, d, o, t, co, n));

                        pgBar.Value = (i * 100) / cell.LastRow;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                    ObjExcel.Quit();
                }
            }
            else {
                System.Windows.MessageBox.Show("Не выбран файл для парсинга данных!", "ВНИАМЕНИЕ !", MessageBoxButton.OK);
                return;
            }            
            pgBar.Value = 0;
            System.Windows.Forms.Application.DoEvents();
        }



        private void parseBtn_Click(object sender, RoutedEventArgs e)
        {
            pg.Visibility = Visibility.Visible;
            System.Windows.Forms.Application.DoEvents();

            //внесение изменений в массив опций для парсинга файлов
            cellLukBash();

            //запускаем парсинг файла отчета
            parseExcelReport();

            // запуск формирования отчетов
            creationRepeatAll();

            System.Windows.MessageBox.Show("Формирование файлов отчета зваершено!", "ИНФОРМАЦИЯ", MessageBoxButton.OK);
            pg.Visibility = Visibility.Hidden;
            pgText.Text = string.Empty;
            pgTName.Text = string.Empty;
            pgBar.Value = 0;
            System.Windows.Forms.Application.DoEvents();
        }


        //сохранение данных формы по колонкам excel если были изменения и внесение их в массив Башнефть
        // для последующего сохрание в файл и работы с данными по формированию отчета
        private void cellLukBash()
        {
            string prov = (radioBash.IsChecked.Value) ? "Bash" : "Luk";

            cell.CellCard = Convert.ToInt32(card.Text);
            cell.CellAzs = Convert.ToInt32(azs.Text);
            cell.CellCompany = Convert.ToInt32(holder.Text);
            cell.CellAdressAzs = Convert.ToInt32(adres.Text);
            cell.CellDateFill = Convert.ToInt32(date.Text);
            cell.CellOperation = Convert.ToInt32(operation.Text);
            cell.CellFuelT = Convert.ToInt32(fuelT.Text);
            cell.CellCountF = Convert.ToInt32(countF.Text);
            cell.FirstRow = Convert.ToInt32(first.Text);
            cell.LastRow = Convert.ToInt32(last.Text);
            cell.FolderPatch = folderPatch.Text;
            cell.FolderMonth = folderMonth.SelectedIndex;
            cell.ListExl = Convert.ToInt32(listExl.Text);

            //изменение массива для последующего внесения изменений в файл optionxls.json
            arrOpt[prov] = JsonConvert.SerializeObject(cell, Formatting.Indented);
            //сохранение данных в файл
            saveOption();

        }

        //внесение изменений в файл настроек json 
        private void saveOption()
        {
            using (StreamWriter file = new StreamWriter(cellPatch, false))
            {
                file.WriteLine(JsonConvert.SerializeObject(arrOpt));
            }
        }


        // очистка поля поиска компании и вывод всех компаний в Грид
        private void clearSearch_Click(object sender, RoutedEventArgs e)
        {
            searchText.Text = "";
            CompanyGrid.ItemsSource = company;
            clearSearch.Visibility = Visibility.Hidden;
        }

        // парсинго компаний
        private void parseComp_Click(object sender, RoutedEventArgs e)
        {
            pg.Visibility = Visibility.Visible;
            pgText.Text = "ОБРАБОТКА ФАЙЛА КОМПАНИЙ";
            System.Windows.Forms.Application.DoEvents();
            exC pf = new exC();
            if (exc != null)
            {
                pf.briefName.Text = exc.BriefName.ToString();
                pf.fullName.Text = exc.FullName.ToString();
                pf.bashName.Text = exc.BashName.ToString();
                pf.lukName.Text = exc.LukName.ToString();
                pf.rangeFirst.Text = exc.RangeFirst.ToString();
                pf.rangeLast.Text = exc.RangeLast.ToString();
                pf.listPage.Text = exc.ListPage.ToString();
            }
            if (pf.ShowDialog() == true)
            {
                exc.BriefName = Convert.ToInt32(pf.briefName.Text);
                exc.FullName = Convert.ToInt32(pf.fullName.Text);
                exc.BashName = Convert.ToInt32(pf.bashName.Text);
                exc.LukName = Convert.ToInt32(pf.lukName.Text);
                exc.RangeFirst = Convert.ToInt32(pf.rangeFirst.Text);
                exc.RangeLast = Convert.ToInt32(pf.rangeLast.Text);
                exc.ListPage = Convert.ToInt32(pf.listPage.Text);

                try
                {
                    using (ExcelPackage exlPac = new ExcelPackage(new FileInfo(pf.fileC.Text)))
                    {
                        ExcelWorksheet exlPage = exlPac.Workbook.Worksheets[exc.ListPage];
                        
                        company.Clear();
                        for (int i = exc.RangeFirst; i <= exc.RangeLast; i++)
                        {
                            var n = (exlPage.Cells[i, exc.BriefName].Value == null) ? "" : exlPage.Cells[i, exc.BriefName].Value.ToString();
                            var fn = (exlPage.Cells[i, exc.FullName].Value == null) ? "" : exlPage.Cells[i, exc.FullName].Value.ToString();
                            var bn = (exlPage.Cells[i, exc.BashName].Value == null) ? "" : exlPage.Cells[i, exc.BashName].Value.ToString();
                            var ln = (exlPage.Cells[i, exc.LukName].Value == null) ? "" : exlPage.Cells[i, exc.LukName].Value.ToString();
                            company.Add(new Company(n, fn, bn, ln));
                            pgBar.Value = (i * 100) / exc.RangeLast;
                            System.Windows.Forms.Application.DoEvents();
                        }
                        exlPac.Dispose();
                    }
                }
                catch (Exception)
                {
                    //если старый офис

                    //Создаём приложение.
                    Excel.Application ObjExcel = new Excel.Application();
                    //Открываем книгу.                                                                                                                                                       
                    Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pf.fileC.Text, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Выбираем таблицу(лист).
                    Excel.Worksheet exlcPage;
                    exlcPage = (Excel.Worksheet)ObjWorkBook.Sheets[exc.ListPage];
                    
                    company.Clear();
                    for (int i = exc.RangeFirst; i <= exc.RangeLast; i++)
                    {
                        var n = (exlcPage.Cells[i, exc.BriefName].Value == null) ? "" : exlcPage.Cells[i, exc.BriefName].Value.ToString();
                        var fn = (exlcPage.Cells[i, exc.FullName].Value == null) ? "" : exlcPage.Cells[i, exc.FullName].Value.ToString();
                        var bn = (exlcPage.Cells[i, exc.BashName].Value == null) ? "" : exlcPage.Cells[i, exc.BashName].Value.ToString();
                        var ln = (exlcPage.Cells[i, exc.LukName].Value == null) ? "" : exlcPage.Cells[i, exc.LukName].Value.ToString();
                        company.Add(new Company(n, fn, bn, ln));
                        pgBar.Value = (i * 100) / exc.RangeLast;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                    ObjExcel.Quit();
                }
                System.Windows.MessageBox.Show("Добавление данных по организациям \nзавершено успешно!", "ИНФОРМАЦИЯ.", MessageBoxButton.OK);
                pg.Visibility = Visibility.Hidden;
                pgText.Text = "";
                pgBar.Value = 0;
                System.Windows.Forms.Application.DoEvents();
            }
            else
            {
                return;
            }
        }

        //выбор каталога сохранения отчетов
        private void folderBtn_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPatch.Text = folder.SelectedPath;
            }
        }

        //выбор файла с исходным отчетом
        private void reportBtn_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Файлы excel|*.xls;*.xlsx;*.xlsm";
            OPF.Title = "Выберите файл excel";
            if (OPF.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName.Text = OPF.FileName;
            }
        }

        // формирование данных в файлы по каждой компании и формирование общего отчета
        private void creationRepeatAll()
        {
            
            pgText.Text = "ФОРМИРОВАНИЕ ФАЙЛОВ ОТЧЕТА \nПО КАЖДОЙ КОМПАНИИ";            
            System.Windows.Forms.Application.DoEvents();

            OfficeOpenXml.Drawing.ExcelPicture img = null;
            string nameCompFile = "";
            string provider = (radioBash.IsChecked.Value) ? "Башнефть" : "Лукойл";
            string outD = folderPatch.Text + DIR_SEPARATOR + folderMonth.Text;
            string outDir = outD + DIR_SEPARATOR + provider;
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            // создание файла Сводной таблицы по всем компаниям
            ExcelPackage tw = new ExcelPackage(new FileInfo(outD + DIR_SEPARATOR + "Общий отчет " + provider + ".xlsx"));
            ExcelWorksheet tws = tw.Workbook.Worksheets.Add("Сводная таблица");


            tws.Workbook.Properties.Title = "Отчет за " + cell.FolderMonth + " " + provider;
            tws.Workbook.Properties.Author = "Директор";
            tws.Workbook.Properties.Company = "ООО Регионсбыт";

            tws.PrinterSettings.Orientation = eOrientation.Portrait;
            tws.PrinterSettings.PaperSize = ePaperSize.A4;

            tws.PrinterSettings.LeftMargin = 0.6m;
            tws.PrinterSettings.RightMargin = tws.PrinterSettings.TopMargin = tws.PrinterSettings.BottomMargin = 0.4m;


            tws.DefaultColWidth = 8;
            tws.Column(1).Width = 15;
            tws.Column(2).Width = 19;

            tws.Cells[1, 1].Value = @"СВОДНЫЙ ОТЧЕТ ПО ОРГАНИЗАЦИЯМ";
            tws.Cells[1, 1, 1, 9].Merge = true;
            tws.Cells[2, 1].Value = @"за " + folderMonth.Text + " " + DateTime.Now.Year.ToString() + " г";
            tws.Cells[2, 1, 2, 9].Merge = true;

            tws.Cells[4, 1].Value = @"ГРУППА";
            tws.Cells[4, 2].Value = @"НАИМЕНОВАНИЕ ОРГАНИЗАЦИИ";
            tws.Cells[4, 3].Value = @"АИ-80";
            tws.Cells[4, 4].Value = @"АИ-92";
            tws.Cells[4, 5].Value = @"АИ-95";
            tws.Cells[4, 6].Value = @"ДТ";
            tws.Cells[4, 7].Value = @"ГАЗ";
            tws.Cells[4, 8].Value = @"ПРОЧЕЕ";
            tws.Cells[4, 9].Value = @"ИТОГО";
            using (var tp = tws.Cells[1, 1, 4, 9])
            {
                tp.Style.Font.Bold = true;
                tp.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                tp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                tp.Style.WrapText = true;
            }


            int aLine = 5;

            var oneC = outArr.GroupBy(f => f.NameCompany).Distinct();
            oneC = oneC.Where(s => s.Key.Trim().Length > 0).OrderBy(nf => nf.Key).ToList();

            var docPDF = new Document();
            docPDF.Info.Title = "Отчет за " + cell.FolderMonth + " " + provider;
            docPDF.Info.Subject = "Директор";
            docPDF.Info.Author = "ООО Регионсбыт";
            var style = docPDF.Styles["Normal"];
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            //style.Font.Bold = true;
            style.Font.Color = Colors.Black;
            //Получает или задает значение, указывающее, является ли разрыв страницы вставляется перед абзацем.
            style.ParagraphFormat.PageBreakBefore = true;
            //Возвращает или задает пространство, включить после абзаца.
            style.ParagraphFormat.SpaceAfter = 3;
            // Set KeepWithNext for all headings to prevent headings from appearing all alone
            // at the bottom of a page. The other headings inherit this from Heading1.
            style.ParagraphFormat.KeepWithNext = true;


            var section = docPDF.AddSection();
            //section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            section.PageSetup.StartingNumber = 1;
            

            //var paragraph = document.LastSection.AddParagraph("Table Overview", "Heading1");
            //paragraph.AddBookmark("Tables");  


            int i = 0;
            foreach (var row in oneC)
            {
                              
                double cai80 = 0, cai92 = 0, cai95 = 0, cdt = 0, cgaz = 0, cdef = 0;
                double ai80 = 0, ai92 = 0, ai95 = 0, dt = 0, gaz = 0, def = 0;
                var s = (radioBash.IsChecked.Value) ? company.Where(c => RemoveSpaces(c.NameBash.ToLower()) == RemoveSpaces(row.Key.ToLower())).Select(k => k)
                    : (radioLuk.IsChecked.Value) ? company.Where(c => RemoveSpaces(c.NameLuk.ToLower()) == RemoveSpaces(row.Key.ToLower())).Select(k => k)
                    : company.Where(c => RemoveSpaces(c.Name.ToLower()) == RemoveSpaces(row.Key.ToLower())).Select(k => k);

                if (s.Count() == 1)
                {
                    nameCompFile = s.ElementAt(0).Name.ToString();
                }
                else if (s.Count() > 1)
                {
                    var res = System.Windows.MessageBox.Show("В списке оргинизаций имеются \nдублированные записи.\nПерейти к редактированию", "", MessageBoxButton.OK, MessageBoxImage.Information);
                    if (res == MessageBoxResult.OK)
                    {
                        CompanyTab.IsSelected = true;
                        searchText.Text = row.Key;
                        //Directory.Delete(outDir);
                        return;
                    }
                }
                else {
                    var res = System.Windows.MessageBox.Show("В списке оргинизаций отсутствует взаимосвязь \n с \"" + row.Key + "\".\nПерейти к добавлению организации", "", MessageBoxButton.OK, MessageBoxImage.Information);
                    if (res == MessageBoxResult.OK)
                    {
                        addcompany addC = new addcompany();
                        if (radioBash.IsChecked.Value)
                        {
                            addC.nameBash.Text = row.Key;
                        }
                        if (radioLuk.IsChecked.Value)
                        {
                            addC.nameLuk.Text = row.Key;
                        }
                        if (addC.ShowDialog() == true)
                        {
                            company.Add(new Company(addC.name.Text, addC.fullName.Text, addC.nameBash.Text, addC.nameLuk.Text));
                            nameCompFile = addC.name.Text;
                            addC.Close();
                        }
                    }
                }
                char[] charInvalidFileChars = Path.GetInvalidFileNameChars();

                foreach (char charInvalid in charInvalidFileChars)
                {
                    nameCompFile = nameCompFile.Replace(charInvalid, ' ');
                }
                pgTName.Text = nameCompFile;
                System.Windows.Forms.Application.DoEvents();


                string fOutOne = outDir + DIR_SEPARATOR + row.Key + " (" + nameCompFile.Replace(".","");
                ExcelPackage cp = new ExcelPackage(new FileInfo(fOutOne + ").xlsx"));
                var str = outArr.Where(r => r.NameCompany == row.Key);
                ExcelWorksheet compPage = cp.Workbook.Worksheets.Add("Отчет по картам");


                compPage.Workbook.Properties.Title = "Отчет за " + cell.FolderMonth + " " + provider;
                compPage.Workbook.Properties.Author = "директор";
                compPage.Workbook.Properties.Company = "ООО Регионсбыт";
                                                   
                compPage.PrinterSettings.Orientation = eOrientation.Portrait;
                compPage.PrinterSettings.PaperSize = ePaperSize.A4;
                compPage.PrinterSettings.LeftMargin = 0.4m;                
                compPage.PrinterSettings.RightMargin = tws.PrinterSettings.TopMargin = tws.PrinterSettings.BottomMargin = 0.2m;

                compPage.Column(1).Width = 11;
                compPage.Column(2).Width = 22;
                compPage.Column(3).Width = 15;
                compPage.Column(4).Width = 17;
                compPage.Column(5).Width = 10;
                compPage.Column(6).Width = 12;
                compPage.Column(7).Width = 10;



                compPage.Cells[1, 2].Value = @"ПОСТАВЩИК/ПРОДАВЕЦ:";
                compPage.Cells[1, 2, 1, 3].Merge = true;
                compPage.Cells[1, 4].Value = @"ЗАКАЗЧИК/ПОКУПАТЕЛЬ:";
                compPage.Cells[1, 4, 1, 7].Merge = true;

                img = compPage.Drawings.AddPicture("log", System.Drawing.Image.FromFile(@"log.png"));
                img.SetPosition(1, 2, 0, 2);

                compPage.Cells[2, 2].Value = "ООО \"Регионсбыт\"";
                compPage.Cells[2, 2, 2, 3].Merge = true;


                compPage.Cells[2, 4].Value = s.ElementAt(0).FullName.ToString();//полное наименование компании                
                compPage.Cells[2, 4, 2, 7].Merge = true;

                compPage.Cells[4, 1].Value = @"ОТЧЕТ ПО ТОПЛИВНЫМ КАРТАМ";
                compPage.Cells[4, 1, 4, 7].Merge = true;
                compPage.Cells[5, 1].Value = @"за " + folderMonth.Text + " " + DateTime.Now.Year.ToString() + " г";
                compPage.Cells[5, 1, 5, 7].Merge = true;


                compPage.Cells[7, 1].Value = @"№ КАРТЫ";
                compPage.Cells[7, 2].Value = @"АДРЕС АЗС";
                compPage.Cells[7, 3].Value = @"№ АЗС";
                compPage.Cells[7, 4].Value = @"ДАТА";
                compPage.Cells[7, 5].Value = @"ТИП ТОПЛИВА";
                compPage.Cells[7, 6].Value = @"ВИД ОПЕРАЦИИ";
                compPage.Cells[7, 7].Value = @"ЗАПР. ЛИТРОВ";

                using (var nf = compPage.Cells[1, 1, 7, 7])
                {
                    nf.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    nf.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    nf.Style.WrapText = true;
                    nf.Style.Font.Bold = true;
                    nf.Style.WrapText = true;
                }

                compPage.Row(2).Height = 60;
                compPage.Cells[4, 1, 5, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                using (var hb = compPage.Cells[7, 1, 7, 7])
                {
                    hb.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    hb.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    var border = hb.Style.Border;
                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                }

                compPage.Cells[6, 1].Value = @"Держатель";
                compPage.Cells[6, 2].Value = row.Key;
                compPage.Cells[6, 5].Value = @"АЗС";
                compPage.Cells[6, 6].Value = provider;
                compPage.Cells[6, 1, 6, 7].Style.Font.Size = 9;
                compPage.Cells[6, 1, 6, 7].Style.Font.Bold = false;

                var cardR = str.GroupBy(c => c.Card).Distinct();
                int compLine = 8;
                //начало по каждой карте компании
                foreach (var cr in cardR)
                {
                    var crd = str.Where(ca => ca.Card == cr.Key);

                    if (provider == "Лукойл")
                    {
                        compPage.Cells[compLine, 1].Value = cr.Key;
                        compPage.Cells[compLine, 1, compLine, 7].Merge = true;
                        using (var k = compPage.Cells[compLine, 1, compLine, 7])
                        {
                            k.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            k.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            k.Style.Font.Bold = true;
                            k.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            var border = k.Style.Border;
                            border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                        }
                        compLine++;
                    }
                    else
                    {
                        compPage.Cells[compLine, 1].Value = cr.Key;
                        compPage.Cells[compLine, 1].Style.Font.Bold = true;
                    }


                    //переменный для каждой карты количество топлива
                    ai80 = 0; ai92 = 0; ai95 = 0; dt = 0; gaz = 0; def = 0;
                    // собираем и формируем отчет по каждой отдельной карте
                    foreach (var r in crd)
                    {
                        Regex r80 = new Regex(@"80", RegexOptions.IgnoreCase);
                        Match mr80 = r80.Match(r.TypeFuel);
                        Regex r92 = new Regex(@"92", RegexOptions.IgnoreCase);
                        Match mr92 = r92.Match(r.TypeFuel);
                        Regex r95 = new Regex(@"95", RegexOptions.IgnoreCase);
                        Match mr95 = r95.Match(r.TypeFuel);
                        Regex rdt = new Regex(@"дт|диз.+ое", RegexOptions.IgnoreCase);
                        Match mrdt = rdt.Match(r.TypeFuel);
                        Regex rgaz = new Regex(@"газ", RegexOptions.IgnoreCase);
                        Match mrgaz = rgaz.Match(r.TypeFuel);
                        Regex raz = new Regex(@"\[.+\]", RegexOptions.IgnoreCase);
                        Regex ic = new Regex(@"\d", RegexOptions.IgnoreCase);

                        compPage.Cells[compLine, 2].Value = r.AdressAzs;
                        compPage.Cells[compLine, 2].Style.WrapText = true;
                        compPage.Cells[compLine, 2].Style.Font.Size = 8;
                        compPage.Cells[compLine, 3].Value = r.Azs = (provider == "Башнефть") ? raz.Replace(r.Azs, "") : r.Azs;
                        compPage.Cells[compLine, 3].Style.Font.Size = 10;
                        compPage.Cells[compLine, 4].Value = r.DateFill;
                        compPage.Cells[compLine, 4].Style.Font.Size = 10;

                        double total = (provider == "Башнефть") ? -Convert.ToDouble(r.CountFuel) : (ic.IsMatch(r.CountFuel) ? Convert.ToDouble(r.CountFuel) : 0);
                        r.CountFuel = total.ToString();
                        if (mr80.Success)
                        {
                            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-80";                            
                            ai80 += total;
                        }
                        if (mr92.Success)
                        {
                            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-92";
                            ai92 += total;
                        }
                        if (mr95.Success)
                        {
                            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-95";
                            ai95 += total;
                        }
                        if (mrdt.Success)
                        {
                            compPage.Cells[compLine, 5].Value = r.TypeFuel = "ДТ";
                            dt += total;
                        }
                        if (mrgaz.Success)
                        {
                            compPage.Cells[compLine, 5].Value = r.TypeFuel = "ГАЗ";
                            gaz += total;
                        }

                        compPage.Cells[compLine, 6].Value = r.Operation;
                        compPage.Cells[compLine, 6].Style.Font.Size = 8;
                        compPage.Cells[compLine, 7].Value = total;
                        using (var allR = compPage.Cells[compLine, 1, compLine, 7])
                        {
                            allR.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            allR.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            allR.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            var border = allR.Style.Border;
                            border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                        }
                        compPage.Cells[compLine, 2].Style.HorizontalAlignment =
                        compPage.Cells[compLine, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        if (compPage.Cells[compLine, 5].Value == null)
                        {
                            compLine++;
                            compPage.Cells[compLine, 1].Value = "Расшифровка : " + r.TypeFuel;
                            compPage.Cells[compLine, 1, compLine, 7].Merge = true;
                            using (var d = compPage.Cells[compLine, 1, compLine, 7])
                            {
                                d.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                d.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                d.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                var border = d.Style.Border;
                                border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                            }
                            compPage.Cells[compLine, 1, compLine, 7].Style.WrapText = true;
                            def += total;
                        }

                        Tables.DefineTables(docPDF, r);
                        
                        compLine++;
                    }

                   

                    //конец по каждой отдельной карте

                    //формирование раздела итогов по каждой отдельной карте
                    compPage.Cells[compLine, 1].Value = @"ИТОГО по карте (" + cr.Key + ") :";
                    compPage.Cells[compLine, 1, compLine, 5].Merge = true;
                    compPage.Cells[compLine, 6].Value = ai80 + ai92 + ai95 + dt + gaz + def;
                    compPage.Cells[compLine, 6, compLine, 7].Merge = true;

                    using (var cel = compPage.Cells[compLine, 1, compLine, 7])
                    {
                        cel.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        cel.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cel.Style.Font.Bold = true;
                        cel.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        var border = cel.Style.Border;
                        border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    compLine++;

                    compPage.Cells[compLine, 1].Value = @"в т.ч :";
                    compPage.Cells[compLine, 2].Value = @"АИ80 :  " + ai80;
                    compPage.Cells[compLine, 3].Value = @"АИ92 :  " + ai92;
                    compPage.Cells[compLine, 4].Value = @"АИ95 :  " + ai95;
                    compPage.Cells[compLine, 5].Value = @"ДТ :  " + dt;
                    compPage.Cells[compLine, 6].Value = @"ГАЗ :  " + gaz;
                    compPage.Cells[compLine, 7].Value = @"ПРОЧ :  " + def;
                    using (var re = compPage.Cells[compLine, 1, compLine, 7])
                    {
                        re.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        re.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        re.Style.Font.Bold = true;
                        re.Style.Font.Size = 9;
                        re.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        var border = re.Style.Border;
                        border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    // конец вывода итогов по карте
                    compLine++;
                    //присвоение данных для сводного отчета (количество заправленного топлива)
                    cai80 += ai80; cai92 += ai92; cai95 += ai95; cdt += dt; cgaz += gaz; cdef += def;
                }
                // конец по всем картам компании
                // итоги по компании
                compLine = compLine + 2;
                compPage.Cells["A" + compLine].Value = @"Итого по типам топлива";
                using (var ac = compPage.Cells["A" + compLine + ":G" + compLine])
                {
                    ac.Merge = true;
                    ac.Style.Font.Bold = true;
                    ac.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ac.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                compLine++;
                compPage.Cells["A" + compLine].Value = @"АИ-80";
                compPage.Cells["B" + compLine].Value = @"АИ-92";
                compPage.Cells["C" + compLine].Value = @"АИ-95";
                compPage.Cells["D" + compLine].Value = @"ДТ";
                compPage.Cells["E" + compLine].Value = @"ГАЗ";
                compPage.Cells["F" + compLine].Value = @"ПРОЧЕЕ";
                compPage.Cells["G" + compLine].Value = @"ИТОГО";
                compLine++;
                compPage.Cells["A" + compLine].Value = cai80;
                compPage.Cells["B" + compLine].Value = cai92;
                compPage.Cells["C" + compLine].Value = cai95;
                compPage.Cells["D" + compLine].Value = cdt;
                compPage.Cells["E" + compLine].Value = cgaz;
                compPage.Cells["F" + compLine].Value = cdef;
                compPage.Cells["G" + compLine].Formula = string.Format("SUM({0}:{1})", "A" + (compLine), "F" + (compLine));

                compPage.Cells["A" + (compLine - 1) + ":G" + (compLine - 1)].Style.Font.Bold = true;
                using (var res = compPage.Cells["A" + (compLine - 1) + ":G" + compLine])
                {
                    res.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    res.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    res.Style.WrapText = true;
                    res.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    var border = res.Style.Border;
                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                }

                compLine = compLine + 2;

                img = compPage.Drawings.AddPicture("stamp", System.Drawing.Image.FromFile(@"stamp.png"));
                img.SetPosition(compLine, 2, 1, 2);

                img = compPage.Drawings.AddPicture("sign", System.Drawing.Image.FromFile(@"sign.png"));
                img.SetPosition(compLine++, 2, 2, 2);

                compLine = compLine + 2;
                compPage.Cells[compLine, 1].Value = @"Директор ООО Регионсбыт";
                compPage.Cells[compLine, 1, compLine, 2].Merge = true;
                compPage.Cells[compLine, 4].Value = @"М.А. Хомченко";
                compPage.Cells[compLine, 4, compLine, 5].Merge = true;
                compPage.Cells[compLine, 1, compLine, 5].Style.Font.Bold = true;
                // конец итогов по компании
                
                //сохранение файла по компании                    
                cp.Save();
                cp.Dispose(); //закрытие файла компании

                //(fOutOne + ").xlsx", fOutOne + ").pdf");
                
                using (ExcelPackage cpr = new ExcelPackage(new FileInfo(fOutOne + ").xlsx")))
                {

                    ExcelWorksheet cmpr = cpr.Workbook.Worksheets[1];
                    cmpr.Drawings.Remove("stamp");
                    cmpr.Drawings.Remove("sign");
                    //cpr.SaveAs(new FileInfo(fOutOne + ").jpg"));
                    cpr.Save();
                }

                // формирование данных в сводном отчете
                tws.Cells[aLine, 1].Value = row.Key;
                tws.Cells[aLine, 1].Style.Font.Size = 10;
                tws.Cells[aLine, 2].Value = nameCompFile;
                tws.Cells[aLine, 2].Style.Font.Size = 9;
                tws.Cells[aLine, 3].Value = cai80;
                tws.Cells[aLine, 4].Value = cai92;
                tws.Cells[aLine, 5].Value = cai95;
                tws.Cells[aLine, 6].Value = cdt;
                tws.Cells[aLine, 7].Value = cgaz;
                tws.Cells[aLine, 8].Value = cdef;
                tws.Cells[aLine, 9].Formula = string.Format("SUM({0}:{1})", "B" + aLine, "G" + aLine);

                using (var ares = tws.Cells[aLine, 1, aLine, 9])
                {
                    ares.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ares.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ares.Style.WrapText = true;
                    ares.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    var border = ares.Style.Border;
                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                }
                tws.Cells[aLine, 1, aLine, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                i++;
                pgBar.Value = (i * 100) / row.Count();
                System.Windows.Forms.Application.DoEvents();
                aLine++;

                MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(docPDF, "MigraDoc.mdddl");

                var renderer = new PdfDocumentRenderer(true);
                renderer.Document = docPDF;

                renderer.RenderDocument();
                renderer.PdfDocument.Save(fOutOne + ").pdf");
                
                renderer.PdfDocument.Close();
                GC.Collect();

            }

            

            tws.Cells[aLine, 2].Value = @"ОБЩИЕ ИТОГИ :";
            tws.Cells[aLine, 3].Formula = string.Format("SUM({0}:{1})", "C5", "C" + (aLine - 1));
            tws.Cells[aLine, 4].Formula = string.Format("SUM({0}:{1})", "D5", "D" + (aLine - 1));
            tws.Cells[aLine, 5].Formula = string.Format("SUM({0}:{1})", "E5", "E" + (aLine - 1));
            tws.Cells[aLine, 6].Formula = string.Format("SUM({0}:{1})", "F5", "F" + (aLine - 1));
            tws.Cells[aLine, 7].Formula = string.Format("SUM({0}:{1})", "G5", "G" + (aLine - 1));
            tws.Cells[aLine, 8].Formula = string.Format("SUM({0}:{1})", "H5", "H" + (aLine - 1));
            tws.Cells[aLine, 9].Formula = string.Format("SUM({0}:{1})", "I5", "I" + (aLine - 1));
            using (var at = tws.Cells[aLine, 1, aLine, 9])
            {
                at.Style.Font.Bold = true;
                at.Style.Font.Size = 10;
                at.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                at.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                at.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                var border = at.Style.Border;
                border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
            }


            //сохранение и закрытие файла со сводным отчетом
            tw.Save();
            tw.Dispose();      

        }
        
     
     
    }
}