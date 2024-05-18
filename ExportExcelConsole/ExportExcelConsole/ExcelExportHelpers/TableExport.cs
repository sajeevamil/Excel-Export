using HtmlAgilityPack;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class TableExport
    {
        private ExportHelper _exportHelper;
        public TableExport()
        {
            _exportHelper = new ExportHelper();
        }
        public void ExportTableToExcel(HtmlNode tableNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            var definedNumberOfColumns = ExportConstants.numberOfColumnsForExcel;
            var currentTableNumberOfColumns = GetNumberOfColumnsOfTable(tableNode);

            foreach (HtmlNode rowNode in tableNode.SelectNodes(".//tr"))
            {
                int column = 1;
                HtmlNodeCollection cellNodes = rowNode.SelectNodes(".//td");
                if (cellNodes is null)
                    continue;

                for (int i = 0; i < cellNodes.Count; i++)
                {
                    int cellIndex = i;
                    HtmlNode cellNode = cellNodes[i];
                    foreach (HtmlNode childNode in cellNode.ChildNodes)
                    {
                        var currentCell = worksheet.Cells[currentRow, column];
                        _exportHelper.ApplyFormatting(childNode, currentCell);
                    }

                    if (definedNumberOfColumns == currentTableNumberOfColumns)
                        column++;
                    else
                    {
                        var numOfColumnsToMerge = definedNumberOfColumns / currentTableNumberOfColumns;
                        var prevMergeEndColumn = column;
                        column += numOfColumnsToMerge;
                        ExcelRange currentCell = worksheet.Cells[currentRow, prevMergeEndColumn, currentRow, column-1];
                        currentCell.Merge = true;
                    }
                }
                currentRow++;
            }
        }

        public static int GetNumberOfColumnsOfTable(HtmlNode tableNode)
        {
            HtmlNode rowNode = tableNode.SelectSingleNode(".//tr");
            if (rowNode == null)
                return 0;

            // Select all cell elements within the row
            HtmlNodeCollection cellNodes = rowNode.SelectNodes(".//td");
            int numberOfColumns = cellNodes?.Count ?? 0;
            return numberOfColumns;
        }
    }
}
