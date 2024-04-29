using ExportExcelConsole.ExcelExportHelpers;
using HtmlAgilityPack;
using OfficeOpenXml;

class Program
{
    static int columnWidthForA4 = 100;
    static void Main(string[] args)
    {
        // Set EPPlus license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Sample HTML content
        // string htmlContent = @"<table><tbody><tr><td>Product Id</td><td>Product Name</td><td>Price</td></tr><tr><td>1</td><td>Chai</td><td>18</td></tr></tbody></table><ul><li>Text formatting &amp; alignment</li><li>Bulleted and numbered lists</li><li>Hyperlink and image dialogs</li></ul><p>Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.</p>";
        // string htmlContent = @"<p>Unsecured loan from Directors represents interest free loan from Mr. Director 1 and Mrs. Director 3 </p><p>50% each. This amount was payable to Shibili Drug House (a partnership firm of Mr. Director 1 and </p>";
        string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p>";
        ExportToExcel(htmlContent);
    }

    static void ExportToExcel(string htmlContent)
    {
        // Load HTML content into HtmlDocument
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;

            int numberOfColumnsForExcel = ExportHelper.GetNumberOfColumns(doc);
            int columnWidth = columnWidthForA4 / numberOfColumnsForExcel;
            ExportHelper.AdjustColumnWidth(worksheet, numberOfColumnsForExcel, columnWidth);

            ParseHtmlStringToExcel(doc, worksheet);

            // Save the Excel package to a file
            string filePath = $"D:\\Work\\Nithesh's POC\\ExportedExcels\\exported_data_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            File.WriteAllBytes(filePath, package.GetAsByteArray());

            Console.WriteLine("Excel file exported successfully!");
            Console.WriteLine($"File Path: {Path.GetFullPath(filePath)}");
        }
    }

    static void ParseHtmlStringToExcel(HtmlDocument doc, ExcelWorksheet worksheet)
    {
        int currentRow = 1;
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
                TableExport.ExportTableToExcel(node, worksheet, ref currentRow);
            }
            else if (node.Name.Equals("ul", StringComparison.OrdinalIgnoreCase) ||
                         node.Name.Equals("ol", StringComparison.OrdinalIgnoreCase))
            {
                ListExport.ExportListToExcel(node, worksheet, ref currentRow);
            }
            else
            {
                foreach (HtmlNode childNode in node.ChildNodes)
                {
                    ParseNodeToExcel(childNode, worksheet, ref currentRow);

                    if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, true);
                    }
                    else if (node.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, false, true);
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