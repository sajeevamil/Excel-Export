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
        // string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p>";
        // string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p><p>The financial statements of the company have been prepared under the historical cost </p><p>convention on an accrual basis of accounting in accordance with the Generally Accepted </p><p>Accounting Principles in India to comply with the Accounting Standards notified under Section </p><p>133 of Companies Act, 2013 read with Rule 7 of the Companies (Accounts) Rules, 2014 and </p><p>relevant provisions of the Companies Act, 2013 (“the 2013 Act”). The accounting policies </p><p>applied are consistent with the previous period.</p>";
        // string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p><p>The financial statements of the company have been prepared under the historical cost </p><p>convention on an accrual basis of accounting in accordance with the Generally Accepted </p><p>Accounting Principles in India to comply with the Accounting Standards notified under Section </p><p>133 of Companies Act, 2013 read with Rule 7 of the Companies (Accounts) Rules, 2014 and </p><p>relevant provisions of the Companies Act, 2013 (“the 2013 Act”). The accounting policies </p><p>applied are consistent with the previous period.</p><p><br></p><p><u><strong>Use of estimates and judgements</strong></u></p>";
        // string htmlContent = @"<p>applied are consistent with the previous period.</p><p><br></p><p><u><strong>Use of estimates and judgements</strong></u></p>";
        // string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p><p>The financial statements of the company have been prepared under the historical cost </p><p>convention on an accrual basis of accounting in accordance with the Generally Accepted </p><p>Accounting Principles in India to comply with the Accounting Standards notified under Section </p><p>133 of Companies Act, 2013 read with Rule 7 of the Companies (Accounts) Rules, 2014 and </p><p>relevant provisions of the Companies Act, 2013 (“the 2013 Act”). The accounting policies </p><p>applied are consistent with the previous period.</p><p><br></p><p><u><strong>Use of estimates and judgements</strong></u></p><p>The preparation of financial statements requires management to make judgments, estimates </p><p>and assumptions, that affect the application of accounting policies and the reported amounts </p><p>of assets, liabilities, income, expenses and disclosures of contingent liabilities at the date of </p><p>these financial statements. Actual results may differ from these estimates.</p><p><br></p><p>Estimates and underlying assumptions are reviewed at each balance sheet date. Revisions to </p><p>accounting estimates are recognised in the period in which the estimate is revised and future </p><p>periods affected.</p><p><br></p><p><u><strong>Property, plant and equipment</strong></u></p><p>Property, plant and equipment are disclosed at historical cost less depreciation. All direct </p><p>expenses incurred on construction/installation of assets are capitalized.</p><p>Property, plant and equipment are capitalized at acquisition cost including the GST when input </p><p>credit is ineligible under Section.17(5) of the CGST Act 2017. In other cases, assets are </p><p>recognized at the cost excluding the GST and the GST paid is claimed as input credit to pay off </p><p>the GST liability. </p><p>Depreciation is provided on the Straight-Line Method (SLM) over the estimated useful lives of </p><p>the assets considering the nature, estimated usage, operating conditions, past history of </p><p>replacement, anticipated technological changes, manufacturers warranties and maintenance </p><p>support.</p><p><br></p>";
        // string htmlContent = @"<table class=""se-table-layout-fixed""><tbody><tr><td><div><u><strong>Type of asset</strong></u></div></td><td><div><br></div></td><td><div><u><strong>Estimated useful life</strong></u></div></td></tr></tbody></table>";
        string htmlContent = @"<table class=""se-table-layout-fixed""><tbody><tr><td><div><u><strong>Typeofasset</strong></u></div></td><td><div><br></div></td><td><div><u><strong>Estimatedusefullife</strong></u></div></td></tr><tr></tr><tr><td><div>FurnitureandFittings</div></td><td><div><br></div></td><td class=""se-table-selected-cell""><div>10years</div></td></tr><tr><td><div>MotorVehicles</div></td><td><div><br></div></td><td><div>8-10years</div></td></tr><tr><td><div>PlantandMachinery</div></td><td><div><br></div></td><td><div>15years</div></td></tr><tr><td><div>ElectricalInstallationsandFittings</div></td><td><div><br></div></td><td><div>10years</div></td></tr><tr><td><div>OfficeEquipments</div></td><td><div><br></div></td><td><div>5years</div></td></tr><tr><td><div>ComputerandDataProcessingUnits</div></td><td><div><br></div></td><td><div>3years</div></td></tr><tr><td><div>IntangibleAssets</div></td><td><div><br></div></td><td><div>5years</div></td></tr></tbody></table>";
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
            else if (node.Name.Equals("br", StringComparison.OrdinalIgnoreCase))
            {
                currentRow++;
            }
            else // expecting it as paragraph
            {
                foreach (HtmlNode childNode in node.ChildNodes)
                {
                    ParseNodeToExcel(childNode, worksheet, ref currentRow);

                    if (node.Name.Equals("strong", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, 1, true);
                    }
                    else if (node.Name.Equals("u", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportStyleFormatting.ApplyFontFormatting(node, worksheet, currentRow, 1, false, true);
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