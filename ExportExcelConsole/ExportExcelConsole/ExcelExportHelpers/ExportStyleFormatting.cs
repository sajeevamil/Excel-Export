using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportStyleFormatting
    {
        public static void ApplyFontFormatting(HtmlNode node, ExcelWorksheet worksheet, int currentRow, int currentColumn = 1, bool bold = false, bool underline = false)
        {

            //ExcelRichTextCollection richText = worksheet.Cells[currentRow, currentColumn].RichText;
            //string text = $" {node.InnerText}";
            //ExcelRichText textPart = richText.Add(text);
            //if (bold)
            //    textPart.Bold = true;
            //if (underline)
            //    textPart.UnderLine = true;
            ExcelRange cell = worksheet.Cells[currentRow, currentColumn];
            if (bold)
                cell.Style.Font.Bold = true;
            if (underline)
                cell.Style.Font.UnderLine = true;
        }
    }
}
