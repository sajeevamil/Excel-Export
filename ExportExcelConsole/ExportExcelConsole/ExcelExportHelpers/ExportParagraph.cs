using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportParagraph
    {
        public static void ExportParagraphToExcel(HtmlNode paragraphNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            ExcelRange paragraphCell = worksheet.Cells[currentRow, 1, currentRow, ExportHelper.numberOfColumnsForExcel];
            paragraphCell.Merge = true;

            // Apply formatting recursively
            ApplyFormatting(paragraphNode, paragraphCell);

            paragraphCell.Style.WrapText = true;
            ExportStyleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);

            ExcelRichTextCollection richTextCollection = paragraphCell.RichText;
            bool hasBold = ExportHelper.HasBoldText(richTextCollection);
            if(hasBold)
            {
                ExportHelper.SetRowHeight(worksheet, paragraphCell.Text, currentRow, 18); // 18 added for a workaround, for bold text increase expected font size
            }
            else
            {
                ExportHelper.SetRowHeight(worksheet, paragraphCell.Text, currentRow);

            }
            currentRow++;
        }

        private static void ApplyFormatting(HtmlNode node, ExcelRange cell, bool bold = false, bool underline = false)
        {
            foreach (var childNode in node.ChildNodes)
            {
                if (childNode.NodeType == HtmlNodeType.Text)
                {
                    var richText = cell.RichText.Add(childNode.InnerText);
                    richText.Bold = bold;
                    richText.UnderLine = underline;
                }
                else if (childNode.NodeType == HtmlNodeType.Element)
                {
                    if (childNode.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        ApplyFormatting(childNode, cell, bold: true, underline: underline);
                    }
                    else if (childNode.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        ApplyFormatting(childNode, cell, bold, underline: true);
                    }
                    else
                    {
                        ApplyFormatting(childNode, cell, bold, underline);
                    }
                }
            }
        }
    }
}
