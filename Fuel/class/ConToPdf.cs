using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Fuel
{
    public class ConToPdf
    {

        public static void HeadDoc(Document document,string FullNameComp)
        {
            var table = new Table();
            table.Borders.Visible = false;
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Format.Font.Bold = true;
            table.TopPadding = 20;
           
            table.AddColumn(Unit.FromCentimeter(2.2));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(3.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            var row = table.AddRow();
            row.Cells[0].AddImage(@"log.png");
            row.Cells[0].MergeDown = 1;
            row.Cells[1].AddParagraph(@"ПОСТАВЩИК/ПРОДАВЕЦ:");
            row.Cells[1].MergeRight = 2;
            row.Cells[4].AddParagraph(@"ЗАКАЗЧИК/ПОКУПАТЕЛЬ:");
            row.Cells[4].MergeRight = 2;

            row = table.AddRow();            
            row.Cells[1].AddParagraph("ООО \"Регионсбыт\"");
            row.Cells[1].MergeRight = 2;
            row.Cells[4].AddParagraph(FullNameComp);//полное наименование компании                
            row.Cells[4].MergeRight = 2;
            document.LastSection.Add(table);
        }

        public static void ThLineCard(Document document, string mon, string nComp, string nPost)
        {
            var par = new Table();
            par.Borders.Visible = false;
            par.Format.Alignment = ParagraphAlignment.Center;
            par.Format.Font.Bold = true;            
            par.AddColumn(Unit.FromCentimeter(20));            
            var r = par.AddRow();            
            r.Cells[0].AddParagraph(@"ОТЧЕТ ПО ТОПЛИВНЫМ КАРТАМ");
            r.Cells[0].Format.SpaceBefore = Unit.FromCentimeter(0.5);
            r = par.AddRow();
            r.Cells[0].AddParagraph(@"за " + mon + " " + DateTime.Now.Year.ToString() + " г");            
            document.LastSection.Add(par);

            var table = new Table();
            table.Borders.Width = 0.75;
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Format.Font.Bold = true;                      
            table.AddColumn(Unit.FromCentimeter(2.2));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(3.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            var row = table.AddRow();
            row.Borders.Visible = false;
            row.Format.Font.Size = 9;
            row.Cells[0].AddParagraph(@"Держатель");
            row.Cells[1].AddParagraph(nComp);
            row.Cells[5].AddParagraph(@"АЗС");
            row.Cells[6].AddParagraph(nPost);
            row = table.AddRow();
            row.Cells[0].AddParagraph(@"№ КАРТЫ");
            row.Cells[1].AddParagraph(@"АДРЕС АЗС");
            row.Cells[2].AddParagraph(@"№ АЗС");
            row.Cells[3].AddParagraph(@"ДАТА");
            row.Cells[4].AddParagraph(@"ТИП ТОПЛИВА");
            row.Cells[5].AddParagraph(@"ВИД ОПЕРАЦИИ");
            row.Cells[6].AddParagraph(@"ЗАПР. ЛИТРОВ");
            document.LastSection.Add(table);
        }

        public static void LineCard(Document document, Out r)// List<Out> str, string provider)
        {
            #region
            var table = new Table();
            table.Borders.Width = 0.75;

            table.AddColumn(Unit.FromCentimeter(2.2));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(3.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));

            var row = table.AddRow();
            row.Format.Alignment = ParagraphAlignment.Center;
            var cell = row.Cells[0];
            cell.AddParagraph(r.Card);

            cell = row.Cells[1];
            cell.Format.Font.Size = 9;
            cell.Format.Alignment = ParagraphAlignment.Left;
            cell.AddParagraph(r.AdressAzs);

            cell = row.Cells[2];
            cell.AddParagraph(r.Azs);

            cell = row.Cells[3];
            cell.AddParagraph(r.DateFill);

            cell = row.Cells[5];
            cell.Format.Font.Size = 9;
            cell.AddParagraph(r.Operation);

            cell = row.Cells[6];
            cell.AddParagraph(Convert.ToDouble(r.CountFuel).ToString());

            if (r.TypeFuel.Length > 5)
            {
                cell = row.Cells[4];
                cell.AddParagraph(@"Прочее*");
                row = table.AddRow();
                cell = row.Cells[0];
                cell.AddParagraph(@"* Расшифровка: " + r.TypeFuel);
                cell.MergeRight = 6;
            }
            else {
                cell = row.Cells[4];
                cell.AddParagraph(r.TypeFuel);
            }
            document.LastSection.Add(table);
            #endregion

            //var table = new Table();
            //table.Borders.Width = 0.75;

            //table.AddColumn(Unit.FromCentimeter(2.2));
            //table.AddColumn(Unit.FromCentimeter(3.5));
            //table.AddColumn(Unit.FromCentimeter(2.3));
            //table.AddColumn(Unit.FromCentimeter(3.3));
            //table.AddColumn(Unit.FromCentimeter(2.3));
            //table.AddColumn(Unit.FromCentimeter(2.3));
            //table.AddColumn(Unit.FromCentimeter(2.3));

            ////начало по каждой карте компании
            //var cardR = str.GroupBy(c => c.Card).Distinct();
            //foreach (var cr in cardR)
            //{
            //    var crd = str.Where(ca => ca.Card == cr.Key);

            //    var row = table.AddRow();

            //    if (provider == "Лукойл")
            //    {
            //        row.Cells[0].AddParagraph(cr.Key);
            //        row.Cells[0].MergeRight = 6;
            //        row = table.AddRow();
            //    }
            //    else
            //    {
            //        compPage.Cells[compLine, 1].Value = cr.Key;
            //        compPage.Cells[compLine, 1].Style.Font.Bold = true;
            //    }


            //    //переменный для каждой карты количество топлива
            //    ai80 = 0; ai92 = 0; ai95 = 0; dt = 0; gaz = 0; def = 0;
            //    // собираем и формируем отчет по каждой отдельной карте
            //    foreach (var r in crd)
            //    {
            //        Regex r80 = new Regex(@"80", RegexOptions.IgnoreCase);
            //        Match mr80 = r80.Match(r.TypeFuel);
            //        Regex r92 = new Regex(@"92", RegexOptions.IgnoreCase);
            //        Match mr92 = r92.Match(r.TypeFuel);
            //        Regex r95 = new Regex(@"95", RegexOptions.IgnoreCase);
            //        Match mr95 = r95.Match(r.TypeFuel);
            //        Regex rdt = new Regex(@"дт|диз.+ое", RegexOptions.IgnoreCase);
            //        Match mrdt = rdt.Match(r.TypeFuel);
            //        Regex rgaz = new Regex(@"газ", RegexOptions.IgnoreCase);
            //        Match mrgaz = rgaz.Match(r.TypeFuel);
            //        Regex raz = new Regex(@"\[.+\]", RegexOptions.IgnoreCase);
            //        Regex ic = new Regex(@"\d", RegexOptions.IgnoreCase);

            //        compPage.Cells[compLine, 2].Value = r.AdressAzs;
            //        compPage.Cells[compLine, 2].Style.WrapText = true;
            //        compPage.Cells[compLine, 2].Style.Font.Size = 8;
            //        compPage.Cells[compLine, 3].Value = r.Azs = (provider == "Башнефть") ? raz.Replace(r.Azs, "") : r.Azs;
            //        compPage.Cells[compLine, 3].Style.Font.Size = 10;
            //        compPage.Cells[compLine, 4].Value = r.DateFill;
            //        compPage.Cells[compLine, 4].Style.Font.Size = 10;

            //        double total = (provider == "Башнефть") ? -Convert.ToDouble(r.CountFuel) : (ic.IsMatch(r.CountFuel) ? Convert.ToDouble(r.CountFuel) : 0);
            //        r.CountFuel = total.ToString();
            //        if (mr80.Success)
            //        {
            //            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-80";
            //            ai80 += total;
            //        }
            //        if (mr92.Success)
            //        {
            //            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-92";
            //            ai92 += total;
            //        }
            //        if (mr95.Success)
            //        {
            //            compPage.Cells[compLine, 5].Value = r.TypeFuel = "АИ-95";
            //            ai95 += total;
            //        }
            //        if (mrdt.Success)
            //        {
            //            compPage.Cells[compLine, 5].Value = r.TypeFuel = "ДТ";
            //            dt += total;
            //        }
            //        if (mrgaz.Success)
            //        {
            //            compPage.Cells[compLine, 5].Value = r.TypeFuel = "ГАЗ";
            //            gaz += total;
            //        }

            //        compPage.Cells[compLine, 6].Value = r.Operation;
            //        compPage.Cells[compLine, 6].Style.Font.Size = 8;
            //        compPage.Cells[compLine, 7].Value = total;

            //        compPage.Cells[compLine, 2].Style.HorizontalAlignment =
            //        compPage.Cells[compLine, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            //        if (compPage.Cells[compLine, 5].Value == null)
            //        {
            //            compLine++;
            //            compPage.Cells[compLine, 1].Value = "Расшифровка : " + r.TypeFuel;
            //            compPage.Cells[compLine, 1, compLine, 7].Merge = true;

            //            compPage.Cells[compLine, 1, compLine, 7].Style.WrapText = true;
            //            def += total;
            //        }



            //        compLine++;
            //    }



            //    //конец по каждой отдельной карте

            //    //формирование раздела итогов по каждой отдельной карте
            //    compPage.Cells[compLine, 1].Value = @"ИТОГО по карте (" + cr.Key + ") :";
            //    compPage.Cells[compLine, 1, compLine, 5].Merge = true;
            //    compPage.Cells[compLine, 6].Value = ai80 + ai92 + ai95 + dt + gaz + def;
            //    compPage.Cells[compLine, 6, compLine, 7].Merge = true;



            //    compLine++;

            //    compPage.Cells[compLine, 1].Value = @"в т.ч :";
            //    compPage.Cells[compLine, 2].Value = @"АИ80 :  " + ai80;
            //    compPage.Cells[compLine, 3].Value = @"АИ92 :  " + ai92;
            //    compPage.Cells[compLine, 4].Value = @"АИ95 :  " + ai95;
            //    compPage.Cells[compLine, 5].Value = @"ДТ :  " + dt;
            //    compPage.Cells[compLine, 6].Value = @"ГАЗ :  " + gaz;
            //    compPage.Cells[compLine, 7].Value = @"ПРОЧ :  " + def;

            //    // конец вывода итогов по карте
            //    compLine++;
            //    //присвоение данных для сводного отчета (количество заправленного топлива)
            //    cai80 += ai80; cai92 += ai92; cai95 += ai95; cdt += dt; cgaz += gaz; cdef += def;
            //}

        }
    }
}
