
using System.Windows;
using System.Windows.Forms;
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
using System.ComponentModel;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;


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
        Newtonsoft.Json.Linq.JObject arr;

        public MainWindow()
        {
            InitializeComponent();            
            if (File.Exists(companyPatch))
            {
                string fileC = File.ReadAllText(companyPatch, UTF8Encoding.UTF8);
                company = JsonConvert.DeserializeObject<ObservableCollection<Company>>(fileC);
                CompanyGrid.ItemsSource = company;
                CompanyGrid.UnselectAllCells();
            } else {
                System.Windows.MessageBox.Show("Файл со списком организайи не существует!", "ВНИМАНИЕ", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            arr = Newtonsoft.Json.Linq.JObject.Parse(File.ReadAllText(cellPatch, Encoding.UTF8));
            exc = JsonConvert.DeserializeObject<exelComp>(arr["excelComp"].ToString());

        }
        
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

        private void deleteCompany()
        {
            if (CompanyGrid.SelectedIndex >= 0)
            {
                Company c = CompanyGrid.SelectedItem as Company;
                company.Remove(c);
            }
        }

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

        private void searchText_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            var s = company.Where<Company>(r => (r.Name + r.FullName + r.NameBash + r.NameLuk).ToLower().Contains(searchText.Text.ToLower()));
            CompanyGrid.ItemsSource = (s.SequenceEqual(company)) ? company:s;
            clearSearch.Visibility = Visibility.Visible;
        }

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

        private void Window_Closed(object sender, EventArgs e)
        {
            searchText.Text = "";
            if (!File.Exists(companyPatch))
            {
                File.Create(companyPatch);
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(companyPatch, false))
            {
                file.WriteLine(JsonConvert.SerializeObject(company));
            }
        }
        
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

        private void radioBash_Checked(object sender, RoutedEventArgs e)
        {           
            
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Bash"].ToString());
            CellText();            
        }

        private void radioLuk_Checked(object sender, RoutedEventArgs e)
        {            
            cell = JsonConvert.DeserializeObject<cellExcel>(arr["Luk"].ToString());
            adres.Text = cell.CellAdressAzs.ToString();
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

        private void parseBtn_Click(object sender, RoutedEventArgs e)
        {
            //сщхранение данных формы по колонкам excel
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

            //очишаем спарсенный массив excel данных из файла поставщика
            outArr.Clear();

        //this.Visibility = Visibility.Hidden;
        //pg.Visibility = Visibility.Visible;
        //pgText.Text = "СКАНИРОВАНИЕ ФАЙЛА ДЛЯ ФОРМИРОВАНИЯ ДАННЫХ \nВЫПОЛНЕНО НА:";
        //System.Threading.Thread.Sleep(5000);// пауза
            string nameCompFile = "";
            if (File.Exists(fileName.Text))
            {
                try
                {
                    //если новый офис
                    using (ExcelPackage execPac = new ExcelPackage(new FileInfo(fileName.Text)))
                    {
                        ExcelWorksheet execPage = execPac.Workbook.Worksheets[cell.ListExl];
                        pgBar.Maximum = cell.LastRow;
                        for (int i = cell.FirstRow; i <= cell.LastRow; i++)
                        {
                            pgBar.Value = i;
                            var c = (execPage.Cells[i, cell.CellCard].Value == null) ? "" : execPage.Cells[i, cell.CellCard].Value.ToString();
                            var s = (execPage.Cells[i, cell.CellAzs].Value == null) ? "" : execPage.Cells[i, cell.CellAzs].Value.ToString();
                            var a = (execPage.Cells[i, cell.CellAdressAzs].Value == null) ? "" : execPage.Cells[i, cell.CellAdressAzs].Value.ToString();
                            var d = (execPage.Cells[i, cell.CellDateFill].Value == null) ? "" : execPage.Cells[i, cell.CellDateFill].Value.ToString();
                            var o = (execPage.Cells[i, cell.CellOperation].Value == null) ? "" : execPage.Cells[i, cell.CellOperation].Value.ToString();
                            var t = (execPage.Cells[i, cell.CellFuelT].Value == null) ? "" : execPage.Cells[i, cell.CellFuelT].Value.ToString();
                            var co = (execPage.Cells[i, cell.CellCountF].Value == null) ? "" : execPage.Cells[i, cell.CellCountF].Value.ToString();
                            var n = (execPage.Cells[i, cell.CellCompany].Value == null) ? "" : execPage.Cells[i, cell.CellCompany].Value.ToString();

                            outArr.Add(new Out(c, s, a, d, o, t, co, n));
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

                    //pgBar.Maximum = cell.LastRow;
                    for (int i = cell.FirstRow; i <= cell.LastRow; i++)
                    {
                        //pgBar.Value = i;
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
                    }
                    //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                    ObjExcel.Quit();
                }
            }
            else {
                System.Windows.MessageBox.Show("Не выбран файл для парсинга данных!", "ВНИАМЕНИЕ !", MessageBoxButton.OK);
                return;
            }

            //pgText.Text = "ФОРМИРОВАНИЕ ФАЙЛОВ ОТЧЕТА \nВЫПОЛНЕНО НА:";
            //pgBar.Value = 0;

            string provider = (radioBash.IsChecked.Value) ? "Башнефть" : "Лукойл";
            string outD = folderPatch.Text + DIR_SEPARATOR + folderMonth.Text;
            string outDir = outD + DIR_SEPARATOR + provider;
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            // создание файла Сводной таблицы по всем компаниям
            ExcelPackage totalPack = new ExcelPackage(new FileInfo(outD + DIR_SEPARATOR + "Общий отчет " + provider + ".xlsx"));
            ExcelWorksheet totalPage = totalPack.Workbook.Worksheets.Add("Сводная таблица");
            
            totalPage.Workbook.Properties.Title = "Отчет за " + cell.FolderMonth + " " + provider;
            totalPage.Workbook.Properties.Author = "директор";
            totalPage.Workbook.Properties.Company = "ООО Регионсбыт";

            totalPage.PrinterSettings.Orientation = eOrientation.Landscape;
            totalPage.PrinterSettings.PaperSize = ePaperSize.A4;
            totalPage.PrinterSettings.LeftMargin = 0.7m;
            totalPage.PrinterSettings.RightMargin = totalPage.PrinterSettings.TopMargin = totalPage.PrinterSettings.BottomMargin = 0.5m;
            totalPage.DefaultColWidth = 15;
            totalPage.Column(1).Width = 20;
            totalPage.Column(2).Width = 25;

            totalPage.Cells[1, 1].Value = @"СВОДНЫЙ ОТЧЕТ ПО ОРГАНИЗАЦИЯМ";
            totalPage.Cells[1, 1, 1, 8].Merge = true;
            totalPage.Cells[2, 1].Value = @"за " + folderMonth.Text + " " + DateTime.Now.Year.ToString() + " г";
            totalPage.Cells[2, 1, 2, 8].Merge = true;

            totalPage.Cells[4, 1].Value = @"ГРУППА";
            totalPage.Cells[4, 2].Value = @"НАИМЕНОВАНИЕ ОРГАНИЗАЦИИ";
            totalPage.Cells[4, 3].Value = @"АИ-80";
            totalPage.Cells[4, 4].Value = @"АИ-92";
            totalPage.Cells[4, 5].Value = @"АИ-95";
            totalPage.Cells[4, 6].Value = @"ДТ";
            totalPage.Cells[4, 7].Value = @"ГАЗ";
            totalPage.Cells[4, 8].Value = @"ИТОГО";
            using (var tp = totalPage.Cells[1, 1, 4, 8])
            {
                tp.Style.Font.Bold = true;
                tp.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                tp.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                tp.Style.WrapText = true;
            }

            int aLine = 5;
            
            var oneC = outArr.GroupBy(f => f.NameCompany).Distinct();
            //pgBar.Maximum = oneC.Count();
            //int pgC = 0;
            oneC = oneC.Where(s => s.Key.Trim().Length > 0).OrderBy(nf => nf.Key).ToList();
            
            foreach (var row in oneC)
            {
                double cai80 = 0, cai92 = 0, cai95 = 0, cdt = 0, cgaz = 0;
                double ai80 = 0, ai92 = 0, ai95 = 0, dt = 0, gaz = 0;
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
                ExcelPackage compPack = new ExcelPackage(new FileInfo(outDir + DIR_SEPARATOR + row.Key + " (" + nameCompFile + ").xlsx"));
                var str = outArr.Where(r => r.NameCompany == row.Key);

                ExcelWorksheet compPage = compPack.Workbook.Worksheets.Add("Отчет по картам");

                compPage.Workbook.Properties.Title = "Отчет за " + cell.FolderMonth + " " + provider;
                compPage.Workbook.Properties.Author = "директор";
                compPage.Workbook.Properties.Company = "ООО Регионсбыт";
                
                compPage.PrinterSettings.Orientation = eOrientation.Portrait;
                compPage.PrinterSettings.PaperSize = ePaperSize.A4;
                compPage.PrinterSettings.LeftMargin = 0.5m;
                compPage.PrinterSettings.RightMargin = totalPage.PrinterSettings.TopMargin = totalPage.PrinterSettings.BottomMargin = 0.2m;
                compPage.Column(1).Width = 10;
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

                compPage.Cells[2, 2].Value = @"ООО Регионсбыт";
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
                compPage.Row(4).Style.HorizontalAlignment = compPage.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
                using (var hb = compPage.Cells[7, 1, 7, 7])
                {
                    hb.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
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

                    compPage.Cells[compLine, 1].Value = cr.Key;
                    compPage.Cells[compLine, 1].Style.WrapText = true;
                    //переменный для каждой карты количество топлива
                    ai80 = 0; ai92 = 0; ai95 = 0; dt = 0; gaz = 0;
                    // собираем и формируем отчет по каждой отдельной карте
                    foreach (var r in crd)
                    {

                        Regex r80 = new Regex(@"80", RegexOptions.IgnoreCase);
                        Match mr80 = r80.Match(r.TypeFuel);
                        Regex r92 = new Regex(@"92", RegexOptions.IgnoreCase);
                        Match mr92 = r92.Match(r.TypeFuel);
                        Regex r95 = new Regex(@"95", RegexOptions.IgnoreCase);
                        Match mr95 = r95.Match(r.TypeFuel);
                        Regex rdt = new Regex(@"дт|диз+", RegexOptions.IgnoreCase);
                        Match mrdt = rdt.Match(r.TypeFuel);
                        Regex rgaz = new Regex(@"газ", RegexOptions.IgnoreCase);
                        Match mrgaz = rgaz.Match(r.TypeFuel);
                        Regex raz = new Regex(@"\[.+\]", RegexOptions.IgnoreCase);
                        Regex ic = new Regex(@"\d", RegexOptions.IgnoreCase);

                        compPage.Cells[compLine, 2].Value = r.AdressAzs;
                        compPage.Cells[compLine, 2].Style.WrapText = true;
                        compPage.Cells[compLine, 2].Style.Font.Size = 8;
                        compPage.Cells[compLine, 3].Value = (provider == "Башнефть") ?raz.Replace(r.Azs,""):r.Azs;
                        compPage.Cells[compLine, 3].Style.Font.Size = 10;
                        compPage.Cells[compLine, 4].Value = r.DateFill;
                        compPage.Cells[compLine, 4].Style.Font.Size = 10;
                        
                        double total = (provider == "Башнефть") ? -Convert.ToDouble(r.CountFuel) : (ic.IsMatch(r.CountFuel) ? Convert.ToDouble(r.CountFuel) : 0);
                        if (mr80.Success)
                        {
                            compPage.Cells[compLine, 5].Value = "АИ-80";
                            ai80 += total;
                        }
                        if (mr92.Success)
                        {
                            compPage.Cells[compLine, 5].Value = "АИ-92";
                            ai92 += total;
                        }
                        if (mr95.Success)
                        {
                            compPage.Cells[compLine, 5].Value = "АИ-95";
                            ai95 += total;
                        }
                        if (mrdt.Success)
                        {
                            compPage.Cells[compLine, 5].Value = "ДТ";
                            dt += total;
                        }
                        if (mrgaz.Success)
                        {
                            compPage.Cells[compLine, 5].Value = "ГАЗ";
                            gaz += total;
                        }
                        
                        compPage.Cells[compLine, 6].Value = r.Operation;
                        compPage.Cells[compLine, 6].Style.Font.Size = 9;                        
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
                        compPage.Cells[compLine, 1].Style.Font.Bold = true;

                        compLine++;
                    }
                    
                    //конец по каждой отдельной карте

                    //формирование раздела итогов по каждой отдельной карте
                    compPage.Cells[compLine, 1].Value = @"ИТОГО по карте (" + cr.Key + ") :";
                    compPage.Cells[compLine, 1, compLine, 5].Merge = true;
                    compPage.Cells[compLine, 6].Value = ai80 + ai92 + ai95 + dt + gaz;
                    compPage.Cells[compLine, 6, compLine, 7].Merge = true;

                    using (var cel = compPage.Cells[compLine, 1, compLine, 7])
                    {
                        cel.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        cel.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cel.Style.Font.Bold = true;
                        //cel.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //cel.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        var border = cel.Style.Border;
                        border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    compLine++;

                    compPage.Cells[compLine, 2].Value = @"в том числе:";
                    compPage.Cells[compLine, 3].Value = @"АИ80 :  " + ai80 + "л.";
                    compPage.Cells[compLine, 4].Value = @"АИ92 :  " + ai92 + "л.";
                    compPage.Cells[compLine, 5].Value = @"АИ95 :  " + ai95 + "л.";
                    compPage.Cells[compLine, 6].Value = @"ДТ :  " + dt + "л.";
                    compPage.Cells[compLine, 7].Value = @"ГАЗ :  " + gaz + "л.";

                    using (var re = compPage.Cells[compLine, 1, compLine, 7])
                    {
                        re.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        re.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        re.Style.Font.Bold = true;
                        re.Style.Font.Size = 9;                        
                        //re.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //re.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        var border = re.Style.Border;
                        border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    // конец вывода итогов по карте
                    compLine++;
                    //присвоение данных для сводного отчета (количество заправленного топлива)
                    cai80 += ai80; cai92 += ai92; cai95 += ai95; cdt += dt; cgaz += gaz;
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
                compPage.Cells["F" + compLine].Value = @"ИТОГО";
                compLine++;
                compPage.Cells["A" + compLine].Value = cai80;
                compPage.Cells["B" + compLine].Value = cai92;
                compPage.Cells["C" + compLine].Value = cai95;
                compPage.Cells["D" + compLine].Value = cdt;
                compPage.Cells["E" + compLine].Value = cgaz;
                compPage.Cells["F" + compLine].Formula = string.Format("SUM({0}:{1})", "A" + (compLine), "E" + (compLine));

                compPage.Cells["A" + (compLine - 1) + ":F" + (compLine - 1)].Style.Font.Bold = true;
                using (var res = compPage.Cells["A" + (compLine - 1) + ":F" + compLine])
                {
                    res.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    res.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    res.Style.WrapText = true;
                    res.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    //res.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //res.Style.Fill.BackgroundColor.SetColor(Color.IndianRed);
                    var border = res.Style.Border;
                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                }

                compLine = compLine + 4;

                compPage.Cells[compLine, 1].Value = @"Директор ООО Регионсбыт";
                compPage.Cells[compLine, 1, compLine, 2].Merge = true;
                compPage.Cells[compLine, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                compPage.Cells[compLine, 4].Value = @"М.А. Хомченко";
                compPage.Cells[compLine, 4, compLine, 5].Merge = true;
                compPage.Cells[compLine, 1, compLine, 5].Style.Font.Bold = true;

                //compPage.Cells[1, 1, compLine, 7].Style.Font.Name = "Times New Roman";
                // конец итогов по компании
                    
                compPack.Save(); //сохранение файла по компании                
                compPack.Dispose(); //закрытие файла компании

                // формирование данных в сводном отчете
                totalPage.Cells[aLine, 1].Value = row.Key;
                totalPage.Cells[aLine, 2].Value = nameCompFile;
                totalPage.Cells[aLine, 3].Value = cai80;
                totalPage.Cells[aLine, 4].Value = cai92;
                totalPage.Cells[aLine, 5].Value = cai95;
                totalPage.Cells[aLine, 6].Value = cdt;
                totalPage.Cells[aLine, 7].Value = cgaz;
                totalPage.Cells[aLine, 8].Formula = string.Format("SUM({0}:{1})", "B" + aLine, "F" + aLine);

                using (var ares = totalPage.Cells[aLine, 1, aLine, 8])
                {
                    ares.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ares.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ares.Style.WrapText = true;
                    ares.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    var border = ares.Style.Border;
                    border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
                }

                aLine++;
                //pgC++;
                //pgBar.Value = pgC;
            }

            totalPage.Cells[aLine, 2].Value = @"ОБЩИЕ ИТОГИ :";
            totalPage.Cells[aLine, 3].Formula = string.Format("SUM({0}:{1})", "C5", "C" + (aLine - 1));
            totalPage.Cells[aLine, 4].Formula = string.Format("SUM({0}:{1})", "D5", "D" + (aLine - 1));
            totalPage.Cells[aLine, 5].Formula = string.Format("SUM({0}:{1})", "E5", "E" + (aLine - 1));
            totalPage.Cells[aLine, 6].Formula = string.Format("SUM({0}:{1})", "F5", "F" + (aLine - 1));
            totalPage.Cells[aLine, 7].Formula = string.Format("SUM({0}:{1})", "G5", "G" + (aLine - 1));
            totalPage.Cells[aLine, 8].Formula = string.Format("SUM({0}:{1})", "H5", "H" + (aLine - 1));
            using (var at = totalPage.Cells[aLine, 1, aLine, 8])
            {
                at.Style.Font.Bold = true;
                at.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                at.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //at.Style.Fill.PatternType = ExcelFillStyle.Solid;
                //at.Style.Fill.BackgroundColor.SetColor(Color.IndianRed);
                var border = at.Style.Border;
                border.Top.Style = border.Left.Style = border.Bottom.Style = border.Right.Style = ExcelBorderStyle.Thin;
            }


            //сохранение и закрытие файла со сводным отчетом
            totalPack.Save();
            totalPack.Dispose();
            System.Windows.MessageBox.Show("Формирование файлов отчета зваершено!", "ИНФОРМАЦИЯ", MessageBoxButton.OK);
            //pg.Visibility = Visibility.Hidden;
            //pgBar.Value = 0;
            //this.Visibility = Visibility.Visible;
        }
        
        private void clearSearch_Click_1(object sender, RoutedEventArgs e)
        {
            searchText.Text = "";
            CompanyGrid.ItemsSource = company;
            clearSearch.Visibility = Visibility.Hidden;
        }

        // парсинго компаний
        private void parseComp_Click(object sender, RoutedEventArgs e)
        {
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
            if(pf.ShowDialog() == true)
            {
                //pg.Visibility = Visibility.Visible;
                //pgText.Text = @"ЗАГРУЗКА ДАННЫХ ПО КОМПАНИЯМ";
                //pgBar.Value = 0;

                exc.BriefName = Convert.ToInt32(pf.briefName.Text);
                exc.FullName = Convert.ToInt32(pf.fullName.Text);
                exc.BashName = Convert.ToInt32(pf.bashName.Text);
                exc.LukName = Convert.ToInt32(pf.lukName.Text);
                exc.RangeFirst = Convert.ToInt32(pf.rangeFirst.Text);
                exc.RangeLast = Convert.ToInt32(pf.rangeLast.Text);
                exc.ListPage = Convert.ToInt32(pf.listPage.Text);
                
                using (ExcelPackage exlPac = new ExcelPackage(new FileInfo(pf.fileC.Text)))
                {
                    ExcelWorksheet exlPage = exlPac.Workbook.Worksheets[exc.ListPage];
                    //pgBar.Maximum = exc.RangeLast;
                    company.Clear();
                    for (int i = exc.RangeFirst; i <= exc.RangeLast; i++)
                    {
                        //pgBar.Value = i;
                        var n = (exlPage.Cells[i, exc.BriefName].Value == null)?"":exlPage.Cells[i, exc.BriefName].Value.ToString();
                        var fn = (exlPage.Cells[i, exc.FullName].Value == null) ? "" : exlPage.Cells[i, exc.FullName].Value.ToString();
                        var bn = (exlPage.Cells[i, exc.BashName].Value == null)? "": exlPage.Cells[i, exc.BashName].Value.ToString();
                        var ln = (exlPage.Cells[i, exc.LukName].Value == null) ? "" : exlPage.Cells[i, exc.LukName].Value.ToString();                        company.Add(new Company(n,fn,bn,ln));
                    }
                    exlPac.Dispose();
                }
                System.Windows.MessageBox.Show("Добавление данных по организациям \nзавершено успешно!", "ИНФОРМАЦИЯ.", MessageBoxButton.OK);
               // pg.Visibility = Visibility.Hidden;
               // pgText.Text = "";
               // pgBar.Value = 0;

            }
            else
            {
                return;
            }
        }

        private void folderBtn_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPatch.Text = folder.SelectedPath;
            }
        }

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
    }
}