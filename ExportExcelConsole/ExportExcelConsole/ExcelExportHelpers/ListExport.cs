using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class ListExport
    {
        private ExportHelper _exportHelper;
        private ExportStyleFormatting _styleFormatting;
        public ListExport()
        {
            _exportHelper = new ExportHelper();
            _styleFormatting = new ExportStyleFormatting();
        }
        public void ExportListToExcel(HtmlNode listNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            bool isOrderedList = string.Equals(listNode.Name, "ol", StringComparison.OrdinalIgnoreCase);

            int listItemNumber = 1; // Counter for ordered lists

            foreach (HtmlNode listItem in listNode.SelectNodes(".//li"))
            {
                string listItemText = listItem.InnerText.Trim();

                if (isOrderedList)
                {
                    // Add numbers to ordered list items
                    listItemText = $"{listItemNumber}. {listItemText}";
                    listItemNumber++;
                }
                else
                {
                    // Add bullet points to unordered list items
                    listItemText = $"• {listItemText}";
                }

                _exportHelper.MergeCellsAndApplyWrapText(worksheet, currentRow, listItemText);
                _exportHelper.SetRowHeight(worksheet, listItemText, currentRow);
                _styleFormatting.ApplyJustifyToTheContent(worksheet, currentRow);
                currentRow++;
            }
        }
    }
}
