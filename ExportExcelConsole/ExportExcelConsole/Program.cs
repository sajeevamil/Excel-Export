﻿using ExportExcelConsole.ExcelExportHelpers;
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
        // string htmlContent = @"<table class=""se-table-layout-fixed""><tbody><tr><td><div><u><strong>Typeofasset</strong></u></div></td><td><div><br></div></td><td><div><u><strong>Estimatedusefullife</strong></u></div></td></tr><tr></tr><tr><td><div>FurnitureandFittings</div></td><td><div><br></div></td><td class=""se-table-selected-cell""><div>10years</div></td></tr><tr><td><div>MotorVehicles</div></td><td><div><br></div></td><td><div>8-10years</div></td></tr><tr><td><div>PlantandMachinery</div></td><td><div><br></div></td><td><div>15years</div></td></tr><tr><td><div>ElectricalInstallationsandFittings</div></td><td><div><br></div></td><td><div>10years</div></td></tr><tr><td><div>OfficeEquipments</div></td><td><div><br></div></td><td><div>5years</div></td></tr><tr><td><div>ComputerandDataProcessingUnits</div></td><td><div><br></div></td><td><div>3years</div></td></tr><tr><td><div>IntangibleAssets</div></td><td><div><br></div></td><td><div>5years</div></td></tr></tbody></table>";
        string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p><p>The financial statements of the company have been prepared under the historical cost </p><p>convention on an accrual basis of accounting in accordance with the Generally Accepted </p><p>Accounting Principles in India to comply with the Accounting Standards notified under Section </p><p>133 of Companies Act, 2013 read with Rule 7 of the Companies (Accounts) Rules, 2014 and </p><p>relevant provisions of the Companies Act, 2013 (“the 2013 Act”). The accounting policies </p><p>applied are consistent with the previous period.</p><p><br></p><p><u><strong>Use of estimates and judgements</strong></u></p><p>The preparation of financial statements requires management to make judgments, estimates </p><p>and assumptions, that affect the application of accounting policies and the reported amounts </p><p>of assets, liabilities, income, expenses and disclosures of contingent liabilities at the date of </p><p>these financial statements. Actual results may differ from these estimates.</p><p><br></p><p>Estimates and underlying assumptions are reviewed at each balance sheet date. Revisions to </p><p>accounting estimates are recognised in the period in which the estimate is revised and future </p><p>periods affected.</p><p><br></p><p><u><strong>Property, plant and equipment</strong></u></p><p>Property, plant and equipment are disclosed at historical cost less depreciation. All direct </p><p>expenses incurred on construction/installation of assets are capitalized.</p><p>Property, plant and equipment are capitalized at acquisition cost including the GST when input </p><p>credit is ineligible under Section.17(5) of the CGST Act 2017. In other cases, assets are </p><p>recognized at the cost excluding the GST and the GST paid is claimed as input credit to pay off </p><p>the GST liability. </p><p>Depreciation is provided on the Straight-Line Method (SLM) over the estimated useful lives of </p><p>the assets considering the nature, estimated usage, operating conditions, past history of </p><p>replacement, anticipated technological changes, manufacturers warranties and maintenance </p><p>support.</p><p><br></p><p>Estimated useful lives of assets are as follows:</p><table class=""se-table-layout-fixed""><tbody><tr><td><div><u><strong>Type of asset</strong></u></div></td><td><div><br></div></td><td><div><u><strong>Estimated useful life</strong></u></div></td></tr><tr></tr><tr><td><div>Furniture and Fittings</div></td><td><div><br></div></td><td class=""se-table-selected-cell""><div>10 years</div></td></tr><tr><td><div>Motor Vehicles</div></td><td><div><br></div></td><td><div>8-10 years</div></td></tr><tr><td><div>Plant and Machinery</div></td><td><div><br></div></td><td><div>15 years</div></td></tr><tr><td><div>Electrical Installations and Fittings</div></td><td><div><br></div></td><td><div>10 years</div></td></tr><tr><td><div>Office Equipments</div></td><td><div><br></div></td><td><div>5 years</div></td></tr><tr><td><div>Computer and Data Processing Units</div></td><td><div><br></div></td><td><div>3 years</div></td></tr><tr><td><div>Intangible Assets</div></td><td><div><br></div></td><td><div>5 years</div></td></tr></tbody></table><p><br></p><p>Impairment tests are performed on property, plant and equipment when there is an indicator</p><p>that they may be impaired. When the carrying amount of an item of property, plant and </p><p>equipment is assessed to be higher than the estimated recoverable amount, an impairment </p><p>loss is recognised immediately in profit or loss to bring the carrying amount in line with the </p><p>recoverable amount.</p><p><br></p><p><u><strong>Inventories</strong></u></p><p>Inventories are valued and certified by the management at cost or net realizable value </p><p>whichever is lower on FIFO basis.</p><p><br></p><p>Net realisable value is the estimated selling price in the ordinary course of business </p><p>less the estimated costs of completion and the estimated costs necessary</p><p><br></p><p>The cost of inventories comprises of all costs of purchase, costs of conversion and </p><p>other costs incurred in bringing the inventories to their present location and condition.</p><p><br></p><p><u><strong>Cash and cash equivalents</strong></u></p><p>Cash and cash equivalents comprise cash on hand and bank deposits free of </p><p>encumbrance with a maturity date of three months or less from the date of deposit, net of </p><p>temporary overdraft.</p><p><br></p><p><u><strong>Revenue Recognition</strong></u></p><p><br></p><p>Revenue is recognized to the extent it is probable that the economic benefits will flow </p><p>to the entity and the revenue can be reliably measured. Revenue is reduced for </p><p>estimated customer returns, rebates and other similar allowances.</p><p><br></p><p>The following specific recognition criteria must also be met before revenue is recognized:</p><p><br></p><p>Sale of goods</p><p><br></p><p>Revenue is recognized when significant risks and rewards of ownership of goods are </p><p>transferred to the buyer, usually on delivery of goods.</p><p><br></p><p><u><strong>Cost recognition</strong></u></p><p>Costs and expenses are recognized when incurred and are classified according to their nature.</p><p><br></p><p>Expenditure capitalized represents various expenses incurred for construction and </p><p>installation of fixed assets including product development undertaken by the firm.</p><p><br></p><p><u><strong>Employee benefits</strong></u></p><p><br></p><p>The cost of short-term employee benefits, (those payable within 12 months after the </p><p>service is rendered), are recognised in the period in which the service is rendered.</p><p><br></p><p>The company provides a lumpsum payment to its vested employees on retirement or </p><p>termination of employment based on the employees’ last drawn salary and years of </p><p>employment with the company. Such expense is booked in the statement of profit and loss </p><p>in the period it is paid or becomes payable.</p><p><br></p><p><u><strong>Borrowing Cost</strong></u></p><p><br></p><p>Borrowing costs that are directly attributable to the acquisition, construction or </p><p>production of a qualifying asset are capitalised as part of the cost of that asset </p><p>are capitalised as part of the cost of that asset until such time as the asset is ready </p><p>for its intended use.</p><p><br></p><p>All other borrowing costs are recognised as an expense in the period in which they are </p><p>incurred.</p><p><br></p><p><u><strong>Taxation</strong></u></p><p><br></p><p>Current tax is determined as the amount payable in respect of taxable income for </p><p>the year.</p><p><br></p><p>Provision for taxation for the year is ascertained on the basis of assessable profits </p><p>computed in accordance with the provisions of the Income Tax Act, 1961 and is recognized </p><p>in the profit and loss statement for the year.</p><p><br></p><p>Deferred Tax Liability/ Asset is accounted for at current tax rate on timing </p><p>differences on profit as per books and that as per tax provision, to the extent to </p><p>which the timing differences are expected to reverse in the future years.</p><p><br></p><p><u><strong>Impairment of Assets</strong></u></p><p><br></p><p>At each Balance Sheet date, the firm assesses whether there is any indication that </p><p>the fixed assets with finite lives may be impaired. If any such indication exists, </p><p>the recoverable amount of the asset is estimated in order to determine the extent </p><p>of the impairment, if any. Where it is not possible to estimate the recoverable </p><p>amount of individual asset, the firm estimates the recoverable amount of the cash-</p><p>generating unit to which the asset belongs.</p><p><br></p><p>The recoverable amount of an asset or a cash-generating unit is the higher of its </p><p>fair value less costs to sell and its value in use. If the recoverable amount of an </p><p>asset is less than its carrying amount, the carrying amount of the asset is reduced to </p><p>its recoverable amount and the impairment loss is recognised in the profit and loss </p><p>statement.</p><p><br></p><p><u><strong>Earnings per share</strong></u></p><p><br></p><p>Basic earnings per share has been computed by dividing profit/loss for the year by the </p><p>weighted average number of shares outstanding during the year. Partly paid up shares are </p><p>included as fully paid equivalents according to the fraction paid up.</p><p><br></p><p>Earnings considered in ascertaining the Company's earnings per share is the net profit </p><p>for the period after deducting preference dividends and any attributable tax thereto for </p><p>the period.</p><p><br></p><p>The weighted average number of equity shares outstanding during the period is adjusted </p><p>for events such as bonus issue, bonus element in a right issue, share split and reverse </p><p>share split (consolidation of shares) that have changed the number of equity shares </p><p>outstanding without a corresponding change in resources.</p><p><br></p><p>Diluted earnings per share has been computed using the weighted average number of </p><p>shares and dilutive potential shares, except where the result would be anti-dilutive.</p><p><br></p><p><u><strong>Provisions and contingencies</strong></u></p><p><br></p><p>Provisions are recognized when the company has a present legal or constructive </p><p>obligation as a result of past events; it is probable that an outflow of resources </p><p>embodying economic benefits will be required to settle the obligation; and a reliable </p><p>estimate can be made of the amount of the obligation. The expense relating to any </p><p>provision is recognized in the statement of profit or loss, net of any reimbursement.</p><p><br></p><p>For potential losses that are considered possible, but not probable, the company </p><p>provides disclosure in the financial statements but does not record a liability in its </p><p>accounts unless the loss becomes probable.</p>";
        // string htmlContent = @"<p><u><strong>Basis of preparation</strong></u></p><p>The financial statements of the company have been prepared under the historical cost </p><table class=""se-table-layout-fixed""><tbody><tr><td><div><u><strong>Typeofasset</strong></u></div></td><td><div><br></div></td><td><div><u><strong>Estimatedusefullife</strong></u></div></td></tr></tbody></table>";
        // string htmlContent = @"<p>Basis of preparation</p><table class=""se-table-layout-fixed""><tbody><tr><td>Typeofasset</td><td></td><td>Estimatedusefullife</td></tr></tbody></table>";

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
            ExportHelper.numberOfColumnsForExcel = numberOfColumnsForExcel;
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