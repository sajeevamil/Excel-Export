using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class ExportParagraph
    {
        private ExportHelper _exportHelper;
        private ExportStyleFormatting _styleFormatting;
        public ExportParagraph()
        {
            _exportHelper = new ExportHelper();
            _styleFormatting = new ExportStyleFormatting();
        }

        public void ExportParagraphToExcel(HtmlNode paragraphNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            ExcelRange paragraphCell = worksheet.Cells[currentRow, 1, currentRow, ExportConstants.numberOfColumnsForExcel];
            paragraphCell.Merge = true;

            // Apply formatting recursively
            _exportHelper.ApplyFormatting(paragraphNode, paragraphCell);

            paragraphCell.Style.WrapText = true;
            _styleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);

            ExcelRichTextCollection richTextCollection = paragraphCell.RichText;
            bool hasBold = _exportHelper.HasBoldText(richTextCollection);
            if (hasBold)
            {
                _exportHelper.SetRowHeight(worksheet, paragraphCell.Text, currentRow, 18); // 18 added for a workaround, for bold text increase expected font size
            }
            else
            {
                _exportHelper.SetRowHeight(worksheet, paragraphCell.Text, currentRow);

            }
            currentRow++;
        }
    }
}
