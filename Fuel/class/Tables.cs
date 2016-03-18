using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace Fuel
{
    public class Tables
    {
        public static void DefineTables(Document document, Out r)
        {
               

            var table = document.LastSection.AddTable();
            

            table.Borders.Visible = true;
            table.TopPadding = 3;
            table.BottomPadding = 3;
            
            var column = table.AddColumn(100);
            column = table.AddColumn(120);            
            column = table.AddColumn(50);            
            column = table.AddColumn(90);            
            column = table.AddColumn(50);            
            column = table.AddColumn(70);            
            column = table.AddColumn(50);
            //column.Format.Alignment = ParagraphAlignment.Center;
            //column.Format.Font.Size = 9;


            //table.Rows.HeightRule = RowHeightRule.Auto;

            var row = table.AddRow();

            Table c0 = new Table();
            c0 = row.Cells[0].AddTextFrame().AddTable();
            c0.AddColumn();
            c0.AddRow();
            c0[0,0].AddParagraph(r.Card);

            Table c1 = new Table();
            c1 = row.Cells[1].AddTextFrame().AddTable();
            c1.AddColumn();
            c1.AddRow();
            c1[0,0].AddParagraph(r.AdressAzs);

            Table c2 = new Table();
            c2 = row.Cells[2].AddTextFrame().AddTable();
            c2.AddColumn();
            c2.AddRow();
            c2[0,0].AddParagraph(r.Azs);

            Table c3 = new Table();
            c3 = row.Cells[3].AddTextFrame().AddTable();
            c3.AddColumn();
            c3.AddRow();
            c3[0,0].AddParagraph(r.DateFill);

            Table c5 = new Table();
            c5 = row.Cells[5].AddTextFrame().AddTable();
            c5.AddColumn();
            c5.AddRow();
            c5[0,0].AddParagraph(r.Operation);

            Table c6 = new Table();
            c6 = row.Cells[6].AddTextFrame().AddTable();
            c6.AddColumn();
            c6.AddRow();
            c6[0,0].AddParagraph(r.CountFuel);

           

            if (r.TypeFuel.Length > 5)
            {
                row = table.AddRow();                
                row.Cells[0].MergeRight = 6;
                Table c4 = new Table();
                c4 = row.Cells[0].AddTextFrame().AddTable();
                c4.AddColumn();
                c4.AddRow();
                c4[0,0].AddParagraph(r.TypeFuel);
            }
            else {
                Table c4 = new Table();
                c4 = row.Cells[4].AddTextFrame().AddTable();
                c4.AddColumn();
                c4.AddRow();
                c4[0,0].AddParagraph(r.TypeFuel);
            }



            // DemonstrateCellMerge(document);
        }

       

        public static void DemonstrateCellMerge(Document document)
        {
            document.LastSection.AddParagraph("Cell Merge", "Heading3");

            var table = document.LastSection.AddTable();
            table.Borders.Visible = true;
            table.TopPadding = 5;
            table.BottomPadding = 5;

            var column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Right;

            table.Rows.Height = 35;

            var row = table.AddRow();
            row.Cells[0].AddParagraph("Merge Right");
            row.Cells[0].MergeRight = 1;

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Merge Down");


            row.Cells[1].AddImage(@"sign.png");
            row.Cells[1].Format.WidowControl = false;
            row.Cells[1].AddParagraph("gjlgbcm");
            row.Cells[0].AddImage(@"stamp.png");

            table.AddRow();
        }
    }
}
