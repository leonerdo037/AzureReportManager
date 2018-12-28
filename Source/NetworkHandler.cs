using Microsoft.Azure.Management.Fluent;
using Microsoft.Azure.Management.Network.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public class NetworkHandler: AzureHandler
{
    public void GetNSG()
    {
        // Creating Sheets
        var package = new ExcelPackage();
        ExcelWorksheet worksheet = null;
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > NSG Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            worksheet = package.Workbook.Worksheets.Add(sub.DisplayName);
            // Adding Column Headers
            worksheet.Cells[1, 1].Value = "NSG Name";
            worksheet.Cells[1, 2].Value = "Priority";
            worksheet.Cells[1, 3].Value = "Rule Name";
            worksheet.Cells[1, 4].Value = "Protocol";
            worksheet.Cells[1, 5].Value = "Source Address Prefix";
            worksheet.Cells[1, 6].Value = "Source Port Range";
            worksheet.Cells[1, 7].Value = "Direction";
            worksheet.Cells[1, 8].Value = "Destination Address Prefix";
            worksheet.Cells[1, 9].Value = "Destination Port Range";
            worksheet.Cells[1, 10].Value = "Action";
            int row = 2;
            foreach (INetworkSecurityGroup nsg in AzHandler.NetworkSecurityGroups.List())
            {
                worksheet.Cells[row, 1].Value = nsg.Name;
                foreach (string ruleName in nsg.SecurityRules.Keys)
                {
                    worksheet.Cells[row, 2].Value = nsg.SecurityRules[ruleName].Priority;
                    worksheet.Cells[row, 3].Value = nsg.SecurityRules[ruleName].Name;
                    worksheet.Cells[row, 4].Value = nsg.SecurityRules[ruleName].Protocol;
                    // Source Address
                    if (nsg.SecurityRules[ruleName].SourceAddressPrefix == null)
                    {
                        worksheet.Cells[row, 5].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourceAddressPrefixes);
                    }
                    else
                    {
                        worksheet.Cells[row, 5].Value = nsg.SecurityRules[ruleName].SourceAddressPrefix;
                    }
                    // Source Port
                    if (nsg.SecurityRules[ruleName].SourcePortRange == null)
                    {
                        worksheet.Cells[row, 6].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourcePortRanges);
                    }
                    else
                    {
                        worksheet.Cells[row, 6].Value = nsg.SecurityRules[ruleName].SourcePortRange;
                    }
                    worksheet.Cells[row, 7].Value = nsg.SecurityRules[ruleName].Direction;
                    // Destination Address
                    if (nsg.SecurityRules[ruleName].DestinationAddressPrefix == null)
                    {
                        worksheet.Cells[row, 8].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationAddressPrefixes);
                    }
                    else
                    {
                        worksheet.Cells[row, 8].Value = nsg.SecurityRules[ruleName].DestinationAddressPrefix;
                    }
                    // Destination Port
                    if (nsg.SecurityRules[ruleName].DestinationPortRange == null)
                    {
                        worksheet.Cells[row, 9].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationPortRanges);
                    }
                    else
                    {
                        worksheet.Cells[row, 9].Value = nsg.SecurityRules[ruleName].DestinationPortRange;
                    }
                    worksheet.Cells[row, 10].Value = nsg.SecurityRules[ruleName].Access;
                    row++;
                }
                // Style
                using (var range = worksheet.Cells[1, 1, 1, 10])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }
            }
        }
        // Saving
        var xlFile = new FileInfo(string.Format("{0}_{1}-{2}-{3}", TextHandler.CurrentPath + @"\NSG",
                                 DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
        package.SaveAs(xlFile);
    }
}
