using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class ExportStyleFormatting
    {
        public ExportStyleFormatting() { }
        public void ApplyFontFormattingToCell(HtmlNode node, ExcelWorksheet worksheet, int currentRow, int currentColumn = 1, bool bold = false, bool underline = false)
        {
            ExcelRange cell = worksheet.Cells[currentRow, currentColumn];
            if (bold)
                cell.Style.Font.Bold = true;
            if (underline)
                cell.Style.Font.UnderLine = true;
        }

        public void ApplyJustifyToTheContent(ExcelWorksheet worksheet, int currentRow, int currentColumn = 1)
        {
            ExcelRange cell = worksheet.Cells[currentRow, currentColumn];
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
        }
    }
}
