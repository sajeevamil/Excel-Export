using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class TextExport
    {
        private ExportHelper _exportHelper;
        private ExportStyleFormatting _styleFormatting;

        public TextExport()
        {
            _exportHelper = new ExportHelper();
            _styleFormatting = new ExportStyleFormatting();
        }
        public void ExportTextToExcel(HtmlNode node, ExcelWorksheet worksheet, ref int currentRow)
        {
            string text = node.InnerText.Trim();
            _exportHelper.MergeCellsAndApplyWrapText(worksheet, currentRow, text);
            _exportHelper.SetRowHeight(worksheet, text, currentRow);
            _styleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);
        }
    }
}
