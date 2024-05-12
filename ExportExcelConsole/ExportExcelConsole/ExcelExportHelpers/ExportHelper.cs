using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportHelper
    {
        public static int numberOfColumnsForExcel = 1;
        public static void MergeCellsAndApplyWrapText(ExcelWorksheet worksheet, int currentRow, string cellContent)
        {
            // Merge cells horizontally and add the cell text content
            ExcelRange paragraphCell = worksheet.Cells[currentRow, 1, currentRow, numberOfColumnsForExcel];
            paragraphCell.Merge = true;
            var existingContent = "";
            if (paragraphCell.Value is not null)
            {
                if (paragraphCell.Value is string stringValue)
                {
                    existingContent = stringValue;
                }
                else if (paragraphCell.Value is object[,] richTextArray)
                {
                    for (int i = 0; i < richTextArray.GetLength(0); i++)
                    {
                        for (int j = 0; j < richTextArray.GetLength(1); j++)
                        {
                            object value = richTextArray[i, j];

                            if (value is not null)
                                existingContent = $"{existingContent} {value}";
                        }
                    }
                }
            }
            paragraphCell.Value = !existingContent.Equals("") ? $"{existingContent} {cellContent}" : cellContent;

            // Enable text wrapping for the merged cell
            paragraphCell.Style.WrapText = true;
        }

        public static void SetRowHeight(ExcelWorksheet worksheet, string cellContent, int currentRow)
        {
            // Calculate the height required for the content based on max characters per line
            int maxCharactersPerLine = 90;
            double rowHeight = cellContent.Length / maxCharactersPerLine * 15; // Assuming font size 15

            if (rowHeight > 0)
            {
                worksheet.Row(currentRow).Height = rowHeight;
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
