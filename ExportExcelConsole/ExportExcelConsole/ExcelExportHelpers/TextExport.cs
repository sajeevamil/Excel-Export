using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class TextExport
    {
        public static void ExportTextToExcel(HtmlNode node, ExcelWorksheet worksheet, ref int currentRow)
        {            
            string text = node.InnerText.Trim();
            ExportHelper.MergeCellsAndApplyWrapText(worksheet, currentRow, text);
            ExportHelper.SetRowHeight(worksheet, text, currentRow);
            ExportStyleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);
            // currentRow++;
        }
    }
}
