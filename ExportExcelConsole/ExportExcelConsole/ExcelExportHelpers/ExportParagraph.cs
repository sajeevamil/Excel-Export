﻿using HtmlAgilityPack;
using OfficeOpenXml;

namespace ExportExcelConsole.ExcelExportHelpers
{
    // TODO: Remove later, nowhere using this
    public static class ExportParagraph
    {
        public static void ExportParagraphToExcel(HtmlNode paragraphNode, ExcelWorksheet worksheet, ref int currentRow)
        {
            string paragraphText = paragraphNode.InnerText.Trim();
            ExportHelper.MergeCellsAndApplyWrapText(worksheet, currentRow, paragraphText);
            ExportHelper.SetRowHeight(worksheet, paragraphText, currentRow);
            currentRow++;
        }
    }
}
