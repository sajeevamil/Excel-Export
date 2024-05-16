using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public static class ExportParagraph
    {
        public static void ExportParagraphToExcel(HtmlNode paragraphNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            ExcelRange paragraphCell = worksheet.Cells[currentRow, 1, currentRow, ExportHelper.numberOfColumnsForExcel];
            paragraphCell.Merge = true;

            foreach (var childNode in paragraphNode.ChildNodes)
            {
                if (childNode.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                {
                    paragraphCell.RichText.Add(childNode.InnerText).Bold = true;
                }
                else if (childNode.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                {
                    paragraphCell.RichText.Add(childNode.InnerText).UnderLine = true;
                }
                else
                {
                    var richText = paragraphCell.RichText.Add(childNode.InnerText);
                    richText.Bold = false;
                    richText.UnderLine = false;
                }

            }
            paragraphCell.Style.WrapText = true;
            ExportHelper.SetRowHeight(worksheet, paragraphCell.Text, currentRow);
            ExportStyleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);
            currentRow++;
        }
    }
}
