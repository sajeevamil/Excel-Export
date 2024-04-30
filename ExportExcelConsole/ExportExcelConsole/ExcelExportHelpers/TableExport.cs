using HtmlAgilityPack;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class TableExport
    {
        public static void ExportTableToExcel(HtmlNode tableNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            foreach (HtmlNode rowNode in tableNode.SelectNodes(".//tr"))
            {
                int column = 1;
                if (rowNode.SelectNodes(".//td") is not null)
                {
                    foreach (HtmlNode cellNode in rowNode.SelectNodes(".//td"))
                    {
                        var children = cellNode.ChildNodes;
                        foreach (HtmlNode childNode in cellNode.ChildNodes)
                        {
                            ParseNodeToExcel(childNode, worksheet, column, ref currentRow);
                        }
                        column++;
                    }
                    currentRow++;
                }
            }
        }

        static void ParseNodeToExcel(HtmlNode node, ExcelWorksheet worksheet, int columnNum, ref int currentRow)
        {
            if (node.NodeType == HtmlNodeType.Element)
            {
                foreach (HtmlNode childNode in node.ChildNodes)
                {
                    ParseNodeToExcel(childNode, worksheet, columnNum, ref currentRow);
                    if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, columnNum, true);
                    }
                    else if (node.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, columnNum, false, true);
                    }
                }
            }
            else
            {
                worksheet.Cells[currentRow, columnNum].Value = node.InnerText.Trim();
            }
        }
    }
}
