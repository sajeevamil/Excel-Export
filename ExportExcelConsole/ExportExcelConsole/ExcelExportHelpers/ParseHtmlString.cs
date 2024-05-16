using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ParseHtmlString
    {
        public static void ParseHtmlStringToExcel(HtmlDocument doc, ExcelWorksheet worksheet)
        {
            int currentRow = 0;
            foreach (HtmlNode node in doc.DocumentNode.ChildNodes)
            {
                ParseNodeToExcel(node, worksheet, ref currentRow);
            }
        }

        static void ParseNodeToExcel(HtmlNode node, ExcelWorksheet worksheet, ref int currentRow)
        {
            if (node.NodeType == HtmlNodeType.Element)
            {
                if (node.Name.Equals("table", StringComparison.OrdinalIgnoreCase))
                {
                    currentRow++; // this is added, faced an issue like table get overrid the already rendered paragraph
                    TableExport.ExportTableToExcel(node, worksheet, ref currentRow);
                }
                else if (node.Name.Equals("ul", StringComparison.OrdinalIgnoreCase) ||
                             node.Name.Equals("ol", StringComparison.OrdinalIgnoreCase))
                {
                    currentRow++;
                    ListExport.ExportListToExcel(node, worksheet, ref currentRow);
                }
                //else if (node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
                //{
                //    currentRow++;
                //}
                else if (node.Name.Equals("p", StringComparison.OrdinalIgnoreCase))
                {
                    if(currentRow == 0)
                    {
                        currentRow++;
                    }
                    ExportParagraph.ExportParagraphToExcel(node, worksheet, ref currentRow);
                }
                else 
                {
                    if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormattingToCell(node, worksheet, currentRow, 1, true);
                        foreach (HtmlNode childNode in node.ChildNodes)
                        {
                            ParseNodeToExcel(childNode, worksheet, ref currentRow);
                        }
                    }
                    else if (node.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormattingToCell(node, worksheet, currentRow, 1, false, true);
                        foreach (HtmlNode childNode in node.ChildNodes)
                        {
                            ParseNodeToExcel(childNode, worksheet, ref currentRow);
                        }
                    }
                }
            }
            else if (node.NodeType == HtmlNodeType.Text)
            {
                TextExport.ExportTextToExcel(node, worksheet, ref currentRow);
            }
        }
    }
}
