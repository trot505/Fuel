using MigraDoc.DocumentObjectModel;

namespace Fuel
{
    public class Styles
    {
        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        public static void DefineStyles(Document document)
        {
            // Получить стандартный стиль обычный.
            var style = document.Styles["Normal"];
            //Потому что все стили являются производными от нормальных, 
            //следующая строка изменяет размер шрифта всего документа. 
            //Или, точнее, она меняется шрифт всех стилей и абзацев, не переопределить шрифт.
            style.Font.Name = "Segoe UI";


            //Heading1 к Heading9 предопределенных стилей с уровнем структуры.
            //Уровень структуры, чем другие OutlineLevel.BodyText автоматически
            //создает контур (или закладок) в pdf.

            style = document.Styles["Heading1"];
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

            style = document.Styles["Heading2"];
            style.Font.Size = 14;
            //style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading3"];
            style.Font.Size = 12;
            //style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;

            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called TextBox based on style Normal.
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            //TODO: Colors
            style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

            // Create a new style called TOC based on style Normal.
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Colors.Blue;
        }
    }
}
