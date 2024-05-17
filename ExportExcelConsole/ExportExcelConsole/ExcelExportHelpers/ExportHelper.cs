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

        public static void SetRowHeight(ExcelWorksheet worksheet, string cellContent, int currentRow, int expectedFontSize = 16, int maxCharactersPerLine= 114)
        {
            double height = Math.Floor(((double)cellContent.Length / maxCharactersPerLine) * expectedFontSize);
            double rowHeight = Math.Ceiling(height / expectedFontSize) * expectedFontSize;

            if (rowHeight > 0)
            {
                worksheet.Row(currentRow).Height = rowHeight;
            }
        }

        public static int GetNumberOfColumns(HtmlDocument doc)
        {
            int maxColumns = 0;

            // Select all table elements
            HtmlNodeCollection tableNodes = doc.DocumentNode.SelectNodes("//table");

            if (tableNodes == null)
                return 1;

            foreach (var tableNode in tableNodes)
            {
                // Select the first row element within the table
                HtmlNode rowNode = tableNode.SelectSingleNode(".//tr");
                if (rowNode == null)
                    continue;

                // Select all cell elements within the row
                HtmlNodeCollection cellNodes = rowNode.SelectNodes(".//td");
                int numberOfColumns = cellNodes?.Count ?? 0;

                // Update maxColumns if this table has more columns
                if (numberOfColumns > maxColumns)
                {
                    maxColumns = numberOfColumns;
                }
            }

            // If no tables were found, default to 1 column
            if (maxColumns == 0)
            {
                maxColumns = 1;
            }

            return maxColumns;
        }
        public static void AdjustColumnWidth(ExcelWorksheet worksheet, int numberOfColumns, int columnWidth)
        {
            for (int i = 1; i <= numberOfColumns; i++)
            {
                worksheet.Column(i).Width = columnWidth;
            }
        }

        public static bool HasBoldText(ExcelRichTextCollection richTextCollection)
        {
            foreach (var richText in richTextCollection)
            {
                if (richText.Bold)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
