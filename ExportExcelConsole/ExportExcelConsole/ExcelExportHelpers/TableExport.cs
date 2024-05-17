using HtmlAgilityPack;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class TableExport
    {
        public static void ExportTableToExcel(HtmlNode tableNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            var definedNumberOfColumns = ExportHelper.numberOfColumnsForExcel;
            var currentTableNumberOfColumns = GetNumeberOfColumnsOfTable(tableNode);

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
                        ExportHelper.ApplyFormatting(childNode, currentCell);
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

        //private static void ApplyFormatting(HtmlNode node, ExcelRange cell, bool bold = false, bool underline = false)
        //{
        //    foreach (var childNode in node.ChildNodes)
        //    {
        //        if (childNode.NodeType == HtmlNodeType.Text)
        //        {
        //            var richText = cell.RichText.Add(childNode.InnerText);
        //            richText.Bold = bold;
        //            richText.UnderLine = underline;
        //        }
        //        else if (childNode.NodeType == HtmlNodeType.Element)
        //        {
        //            if (childNode.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
        //            {
        //                ApplyFormatting(childNode, cell, bold: true, underline: underline);
        //            }
        //            else if (childNode.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
        //            {
        //                ApplyFormatting(childNode, cell, bold, underline: true);
        //            }
        //            else
        //            {
        //                ApplyFormatting(childNode, cell, bold, underline);
        //            }
        //        }
        //    }
        //}

        public static int GetNumeberOfColumnsOfTable(HtmlNode tableNode)
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
