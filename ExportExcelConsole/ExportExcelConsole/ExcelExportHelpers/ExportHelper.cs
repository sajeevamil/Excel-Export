using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportHelper
    {
        public static int numberOfColumnsForExcel = 1;
        public static void MergeCellsAndApplyWrapText(ExcelWorksheet worksheet, int currentRow, string cellContent)
        {
            // Merge cells horizontally and add the cell text content
            ExcelRange paragraphCell = worksheet.Cells[currentRow + 1, 1, currentRow + 1, numberOfColumnsForExcel];
            paragraphCell.Merge = true;
            paragraphCell.Value = cellContent;

            // Enable text wrapping for the merged cell
            paragraphCell.Style.WrapText = true;
        }

        public static void SetRowHeight(ExcelWorksheet worksheet, string cellContent, int currentRow)
        {
            // Calculate the height required for the content based on max characters per line
            int maxCharactersPerLine = 80;
            double rowHeight = cellContent.Length / maxCharactersPerLine * 15; // Assuming font size 15

            if (rowHeight > 0)
            {
                worksheet.Row(currentRow + 1).Height = rowHeight;
            }
        }

        public static int GetNumberOfColumns(HtmlDocument doc)
        {
            // Select the first table element
            HtmlNode tableNode = doc.DocumentNode.SelectSingleNode("//table");
            if (tableNode == null)
                return 1;

            // Select the first row element within the table
            HtmlNode rowNode = tableNode.SelectSingleNode(".//tr");
            if (rowNode == null)
                return 1;

            // Select all cell elements within the row
            HtmlNodeCollection cellNodes = rowNode.SelectNodes(".//td");
            return cellNodes.Count;
        }
        public static void AdjustColumnWidth(ExcelWorksheet worksheet, int numberOfColumns, int columnWidth)
        {
            for (int i = 1; i <= numberOfColumns; i++)
            {
                worksheet.Column(i).Width = columnWidth;
            }
        }
    }
}
