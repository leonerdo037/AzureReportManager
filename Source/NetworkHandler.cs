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
            worksheet.Cells[1, 2].Value = "Resource Group";
            worksheet.Cells[1, 3].Value = "Region";
            worksheet.Cells[1, 4].Value = "Priority";
            worksheet.Cells[1, 5].Value = "Rule Name";
            worksheet.Cells[1, 6].Value = "Protocol";
            worksheet.Cells[1, 7].Value = "Source Address Prefix";
            worksheet.Cells[1, 8].Value = "Source Port Range";
            worksheet.Cells[1, 9].Value = "Direction";
            worksheet.Cells[1, 10].Value = "Destination Address Prefix";
            worksheet.Cells[1, 11].Value = "Destination Port Range";
            worksheet.Cells[1, 12].Value = "Action";
            foreach (INetworkSecurityGroup nsg in AzHandler.NetworkSecurityGroups.List())
            {
                int row = 2;
                foreach (string ruleName in nsg.SecurityRules.Keys)
                {
                    
                    worksheet.Cells[row, 1].Value = nsg.Name;
                    worksheet.Cells[row, 2].Value = nsg.ResourceGroupName;
                    worksheet.Cells[row, 3].Value = nsg.RegionName;
                    worksheet.Cells[row, 4].Value = nsg.SecurityRules[ruleName].Priority;
                    worksheet.Cells[row, 5].Value = nsg.SecurityRules[ruleName].Name;
                    worksheet.Cells[row, 6].Value = nsg.SecurityRules[ruleName].Protocol;
                    // Source Address
                    if (nsg.SecurityRules[ruleName].SourceAddressPrefix == null)
                    {
                        worksheet.Cells[row, 7].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourceAddressPrefixes);
                    }
                    else
                    {
                        worksheet.Cells[row, 7].Value = nsg.SecurityRules[ruleName].SourceAddressPrefix;
                    }
                    // Source Port
                    if (nsg.SecurityRules[ruleName].SourcePortRange == null)
                    {
                        worksheet.Cells[row, 8].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourcePortRanges);
                    }
                    else
                    {
                        worksheet.Cells[row, 8].Value = nsg.SecurityRules[ruleName].SourcePortRange;
                    }
                    worksheet.Cells[row, 9].Value = nsg.SecurityRules[ruleName].Direction;
                    // Destination Address
                    if (nsg.SecurityRules[ruleName].DestinationAddressPrefix == null)
                    {
                        worksheet.Cells[row, 10].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationAddressPrefixes);
                    }
                    else
                    {
                        worksheet.Cells[row, 10].Value = nsg.SecurityRules[ruleName].DestinationAddressPrefix;
                    }
                    // Destination Port
                    if (nsg.SecurityRules[ruleName].DestinationPortRange == null)
                    {
                        worksheet.Cells[row, 11].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationPortRanges);
                    }
                    else
                    {
                        worksheet.Cells[row, 11].Value = nsg.SecurityRules[ruleName].DestinationPortRange;
                    }
                    worksheet.Cells[row, 12].Value = nsg.SecurityRules[ruleName].Access;
                    row++;
                }
                // Style
                using (var range = worksheet.Cells[1, 1, 1, 12])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }
            }
        }
        // Saving
        var xlFile = new FileInfo(string.Format("{0}_{1}-{2}-{3}.{4}", TextHandler.CurrentPath + @"\NSG",
                                 DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
        package.SaveAs(xlFile);
    }

    public void GetUDR()
    {
        // Creating Sheets
        var package = new ExcelPackage();
        ExcelWorksheet worksheet = null;
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > UDR Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            worksheet = package.Workbook.Worksheets.Add(sub.DisplayName);
            // Adding Column Headers
            worksheet.Cells[1, 1].Value = "UDR Name";
            worksheet.Cells[1, 2].Value = "Resource Group";
            worksheet.Cells[1, 3].Value = "Region";
            worksheet.Cells[1, 4].Value = "Next Hop IP";
            worksheet.Cells[1, 5].Value = "Next Hop Type";
            worksheet.Cells[1, 6].Value = "Destination Address Prefix";
            foreach (INetwork network in AzHandler.Networks.List())
            {
                foreach(ISubnet subnet in network.Subnets.Values)
                {
                    int row = 2;
                    IRouteTable routes = subnet.GetRouteTable();
                    if(routes == null)
                    {
                        continue;
                    }
                    foreach (IRoute route in routes.Routes.Values)
                    {
                        worksheet.Cells[row, 1].Value = route.Name;
                        worksheet.Cells[row, 2].Value = routes.ResourceGroupName;
                        worksheet.Cells[row, 3].Value = routes.RegionName;
                        worksheet.Cells[row, 4].Value = route.NextHopIPAddress;
                        worksheet.Cells[row, 5].Value = route.NextHopType;
                        worksheet.Cells[row, 6].Value = route.DestinationAddressPrefix;
                        row++;
                    }
                    // Style
                    using (var range = worksheet.Cells[1, 1, 1, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                        range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    }
                }
            }
        }
        // Saving
        var xlFile = new FileInfo(string.Format("{0}_{1}-{2}-{3}.{4}", TextHandler.CurrentPath + @"\UDR",
                                 DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
        package.SaveAs(xlFile);
    }

    public void GetVNET()
    {
        // Creating Sheets
        var package = new ExcelPackage();
        ExcelWorksheet worksheet = null;
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > VNET Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            worksheet = package.Workbook.Worksheets.Add(sub.DisplayName);
            // Adding Column Headers
            worksheet.Cells[1, 1].Value = "VNET Name";
            worksheet.Cells[1, 2].Value = "Resource Group";
            worksheet.Cells[1, 3].Value = "Region";
            worksheet.Cells[1, 4].Value = "Address Space";
            worksheet.Cells[1, 5].Value = "Subnet Name";
            worksheet.Cells[1, 6].Value = "Address Prefix";
            worksheet.Cells[1, 7].Value = "Attached NSG";
            worksheet.Cells[1, 8].Value = "Attached UDR";
            worksheet.Cells[1, 9].Value = "Available IPs";
            int row = 2;
            foreach (INetwork network in AzHandler.Networks.List())
            {
                foreach (ISubnet subnet in network.Subnets.Values)
                {
                    worksheet.Cells[row, 1].Value = network.Name;
                    worksheet.Cells[row, 2].Value = network.ResourceGroupName;
                    worksheet.Cells[row, 3].Value = network.RegionName;
                    worksheet.Cells[row, 4].Value = string.Join(",", network.AddressSpaces);
                    worksheet.Cells[row, 5].Value = subnet.Name;
                    worksheet.Cells[row, 6].Value = subnet.AddressPrefix;
                    try
                    {
                        worksheet.Cells[row, 7].Value = subnet.GetNetworkSecurityGroup().Name;
                    }
                    catch (Exception)
                    {
                        worksheet.Cells[row, 7].Value = "NA";
                    }
                    try
                    {
                        worksheet.Cells[row, 8].Value = subnet.GetRouteTable().Name;
                    }
                    catch (Exception)
                    {
                        worksheet.Cells[row, 8].Value = "NA";
                    }
                    worksheet.Cells[row, 9].Value = string.Join(",", subnet.ListAvailablePrivateIPAddresses());
                    row++;
                }
            }
            // Style
            using (var range = worksheet.Cells[1, 1, 1, 9])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                range.Style.Font.Color.SetColor(System.Drawing.Color.White);
            }
        }
        // Saving
        var xlFile = new FileInfo(string.Format("{0}_{1}-{2}-{3}.{4}", TextHandler.CurrentPath + @"\VNET",
                                 DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
        package.SaveAs(xlFile);
    }
}
