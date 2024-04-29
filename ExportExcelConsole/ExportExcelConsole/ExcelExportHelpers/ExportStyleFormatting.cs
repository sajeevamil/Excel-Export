using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportStyleFormatting
    {
        public static void ApplyFontFormatting(HtmlNode node, ExcelWorksheet worksheet, int currentRow, bool bold = false, bool underline = false)
        {
            ExcelRange cell = worksheet.Cells[currentRow, 1];
            if (bold)
                cell.Style.Font.Bold = true;
            if (underline)
                cell.Style.Font.UnderLine = true;
        }
    }
}
