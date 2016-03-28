using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Fuel
{
    public class ConToPdf
    {
        
        //добавление шапки документа по отчету и 
        #region HeadDoc
        public static void HeadDoc(Document document, string FullNameComp)
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
        #endregion

        // наименований столбцов
        #region ThLineCard
        public static void ThLineCard(Document document, string[] dateP, string nComp, string nPost)
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
            r.Cells[0].AddParagraph(@"за период c " + dateP[0] + " по " + dateP[1]);
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
        #endregion

        // добавление данных по картам
        #region LineCard 
        public static void LineCard(Document document, IEnumerable<Out> str, string provider)
        {
            var table = new Table();
            table.Borders.Width = 0.75;
            table.Format.Alignment = ParagraphAlignment.Center;

            table.AddColumn(Unit.FromCentimeter(2.2));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(3.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));

            //начало по каждой карте компании
            var cardR = str.GroupBy(c => c.Card).Distinct();
            foreach (var cr in cardR)
            {
                var crd = str.Where(ca => ca.Card == cr.Key);

                var row = table.AddRow();

                if (provider == "Лукойл")
                {
                    row.Cells[0].AddParagraph(cr.Key);
                    row.Cells[0].Format.Font.Bold = true;
                    row.Cells[0].MergeRight = 6;
                    row = table.AddRow();
                }
                else
                {
                    row.Cells[0].AddParagraph(cr.Key);
                    row.Cells[0].Format.Font.Bold = true;
                }

                //переменный для каждой карты количество топлива
                double ai80 = 0; double ai92 = 0; double ai95 = 0; double dt = 0; double gaz = 0; double def = 0;

                // собираем и формируем отчет по каждой отдельной карте
                foreach (var r in crd)
                {
                    
                    row.Cells[1].AddParagraph(r.AdressAzs);
                    row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].Format.Font.Size = 9;
                    row.Cells[2].AddParagraph(r.Azs);
                    row.Cells[2].Format.Font.Size = 10;
                    row.Cells[3].AddParagraph(r.DateFill);
                    row.Cells[3].Format.Font.Size = 10;
                    row.Cells[5].AddParagraph(r.Operation);
                    row.Cells[5].Format.Font.Size = 9;
                    row.Cells[6].AddParagraph(r.CountFuel);

                    if (r.TypeFuel.Length > 5)
                    {
                        row.Cells[4].AddParagraph(@"Прочее*");
                        row.Cells[4].Format.Font.Size = 9;
                        row = table.AddRow();
                        row.Cells[0].AddParagraph("*Расшифровка: " + r.TypeFuel);
                        row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
                        row.Cells[0].MergeRight = 6;
                    }
                    else
                    {
                        row.Cells[4].AddParagraph(r.TypeFuel);
                    }
                    def += Convert.ToDouble(r.CountFuel);

                    row = table.AddRow();
                }

                //конец по каждой отдельной карте
                //формирование раздела итогов по каждой отдельной карте
                row.Cells[0].AddParagraph(@"ИТОГО по карте (" + cr.Key + ") :");
                row.Cells[0].Format.Font.Bold = true;
                row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
                row.Cells[0].MergeRight = 4;

                row.Cells[5].AddParagraph((ai80 + ai92 + ai95 + dt + gaz + def).ToString());
                row.Cells[5].Format.Font.Bold = true;
                row.Cells[5].MergeRight = 1;

                row = table.AddRow();
                row.Format.Font.Bold = true;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.Cells[0].AddParagraph(@"в т.ч :");
                row.Cells[1].AddParagraph(@"АИ80 :  " + ai80);
                row.Cells[2].AddParagraph(@"АИ92 :  " + ai92);
                row.Cells[3].AddParagraph(@"АИ95 :  " + ai95);
                row.Cells[4].AddParagraph(@"ДТ :  " + dt);
                row.Cells[5].AddParagraph(@"ГАЗ :  " + gaz);
                row.Cells[6].AddParagraph(@"ПРОЧ :  " + def);
                // конец вывода итогов по карте               
            }
            document.LastSection.Add(table);
        }
        #endregion

        // формирование общих итогов по копании и подписей с печатью
        #region FooterCard
        public static void FooterCard(Document document,double[] genTotal)
        {
            var table = new Table();
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Format.Font.Bold = true;
            table.Borders.Width = 0.75;

            table.AddColumn(Unit.FromCentimeter(2.2));
            table.AddColumn(Unit.FromCentimeter(3.5));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(3.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));
            table.AddColumn(Unit.FromCentimeter(2.3));

            var row = table.AddRow();
            row.Borders.Visible = false;
            row.Format.SpaceBefore = Unit.FromCentimeter(1);
            row.Format.SpaceAfter = Unit.FromCentimeter(0.5);
            row.Cells[0].AddParagraph((@"Итого по типам топлива").ToUpper());
            row.Cells[0].MergeRight = 6;

            row = table.AddRow();
            row.Cells[0].AddParagraph(@"АИ-80");
            row.Cells[1].AddParagraph(@"АИ-92");
            row.Cells[2].AddParagraph(@"АИ-95");
            row.Cells[3].AddParagraph(@"ДТ");
            row.Cells[4].AddParagraph(@"ГАЗ");
            row.Cells[5].AddParagraph("ПРОЧЕЕ");
            row.Cells[6].AddParagraph(@"ИТОГО");

            row = table.AddRow();
            row.Cells[0].AddParagraph(genTotal[0].ToString());
            row.Cells[1].AddParagraph(genTotal[1].ToString());
            row.Cells[2].AddParagraph(genTotal[2].ToString());
            row.Cells[3].AddParagraph(genTotal[3].ToString());
            row.Cells[4].AddParagraph(genTotal[4].ToString());
            row.Cells[5].AddParagraph(genTotal[5].ToString());
            row.Cells[6].AddParagraph(genTotal[6].ToString());

            row = table.AddRow();
            row.Borders.Visible = false;
            row.Format.SpaceBefore = Unit.FromCentimeter(2);
            row.TopPadding = Unit.FromCentimeter(0.5);
            
            row.Cells[0].AddParagraph(@"Директор ООО Регионсбыт");
            row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 1;
            row.Cells[2].AddImage(@"stamp.png");
            row.Cells[3].AddParagraph().AddImage(@"sign.png");
            row.Cells[3].Format.SpaceBefore = Unit.FromCentimeter(1);
            row.Cells[4].AddParagraph(@"М.А. Хомченко");
            row.Cells[4].MergeRight = 1;

            document.LastSection.Add(table);
        }
        #endregion
    }
}
