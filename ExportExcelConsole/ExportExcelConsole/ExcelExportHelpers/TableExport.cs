using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class TableExport
    {
        public static void ExportTableToExcel(HtmlNode tableNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            // int row = 1;
            foreach (HtmlNode rowNode in tableNode.SelectNodes(".//tr"))
            {
                int column = 1;
                foreach (HtmlNode cellNode in rowNode.SelectNodes(".//td"))
                {
                    worksheet.Cells[currentRow, column].Value = cellNode.InnerText.Trim();
                    column++;
                }
                currentRow++;
            }
        }
    }
}
