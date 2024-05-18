using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportExcelConsole.ExcelExportHelpers
{
    public class ParseHtmlString
    {
        private ExportParagraph _exportParagraph;
        private ListExport _listExport;
        private ExportStyleFormatting _styleFormatting;
        private TableExport _tableExport;
        private TextExport _textExport;
        public ParseHtmlString()
        {
            _exportParagraph = new ExportParagraph();
            _listExport = new ListExport();
            _styleFormatting = new ExportStyleFormatting();
            _tableExport = new TableExport();
            _textExport = new TextExport();
        }

        public void ParseHtmlStringToExcel(HtmlDocument doc, ExcelWorksheet worksheet)
        {
            int currentRow = 0;
            foreach (HtmlNode node in doc.DocumentNode.ChildNodes)
            {
                ParseNodeToExcel(node, worksheet, ref currentRow);
            }
        }

        void ParseNodeToExcel(HtmlNode node, ExcelWorksheet worksheet, ref int currentRow)
        {
            if (node.NodeType == HtmlNodeType.Element)
            {
                if (node.Name.Equals("table", StringComparison.OrdinalIgnoreCase))
                {
                    currentRow++; // this is added, faced an issue like table get overrid the already rendered paragraph
                    _tableExport.ExportTableToExcel(node, worksheet, ref currentRow);
                }
                else if (node.Name.Equals("ul", StringComparison.OrdinalIgnoreCase) ||
                             node.Name.Equals("ol", StringComparison.OrdinalIgnoreCase))
                {
                    currentRow++;
                    _listExport.ExportListToExcel(node, worksheet, ref currentRow);
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
                    _exportParagraph.ExportParagraphToExcel(node, worksheet, ref currentRow);
                }
                else 
                {
                    if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        _styleFormatting.ApplyFontFormattingToCell(node, worksheet, currentRow, 1, true);
                        foreach (HtmlNode childNode in node.ChildNodes)
                        {
                            ParseNodeToExcel(childNode, worksheet, ref currentRow);
                        }
                    }
                    else if (node.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        _styleFormatting.ApplyFontFormattingToCell(node, worksheet, currentRow, 1, false, true);
                        foreach (HtmlNode childNode in node.ChildNodes)
                        {
                            ParseNodeToExcel(childNode, worksheet, ref currentRow);
                        }
                    }
                }
            }
            else if (node.NodeType == HtmlNodeType.Text)
            {
                _textExport.ExportTextToExcel(node, worksheet, ref currentRow);
            }
        }
    }
}
