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
    private ExcelHandler ExHandler;

    public void GetNSG()
    {
        ExHandler = new ExcelHandler();
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > NSG Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            ExHandler.CreateSheet(sub.DisplayName);
            // Adding Column Headers
            ExHandler.worksheet.Cells[1, 1].Value = "NSG Name";
            ExHandler.worksheet.Cells[1, 2].Value = "Resource Group";
            ExHandler.worksheet.Cells[1, 3].Value = "Region";
            ExHandler.worksheet.Cells[1, 4].Value = "Priority";
            ExHandler.worksheet.Cells[1, 5].Value = "Rule Name";
            ExHandler.worksheet.Cells[1, 6].Value = "Protocol";
            ExHandler.worksheet.Cells[1, 7].Value = "Source Address Prefix";
            ExHandler.worksheet.Cells[1, 8].Value = "Source Port Range";
            ExHandler.worksheet.Cells[1, 9].Value = "Direction";
            ExHandler.worksheet.Cells[1, 10].Value = "Destination Address Prefix";
            ExHandler.worksheet.Cells[1, 11].Value = "Destination Port Range";
            ExHandler.worksheet.Cells[1, 12].Value = "Action";
            int row = 2;
            foreach (INetworkSecurityGroup nsg in AzHandler.NetworkSecurityGroups.List())
            {
                foreach (string ruleName in nsg.SecurityRules.Keys)
                {
                    
                    ExHandler.worksheet.Cells[row, 1].Value = nsg.Name;
                    ExHandler.worksheet.Cells[row, 2].Value = nsg.ResourceGroupName;
                    ExHandler.worksheet.Cells[row, 3].Value = nsg.RegionName;
                    ExHandler.worksheet.Cells[row, 4].Value = nsg.SecurityRules[ruleName].Priority;
                    ExHandler.worksheet.Cells[row, 5].Value = nsg.SecurityRules[ruleName].Name;
                    ExHandler.worksheet.Cells[row, 6].Value = nsg.SecurityRules[ruleName].Protocol;
                    // Source Address
                    if (nsg.SecurityRules[ruleName].SourceAddressPrefix == null)
                    {
                        ExHandler.worksheet.Cells[row, 7].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourceAddressPrefixes);
                    }
                    else
                    {
                        ExHandler.worksheet.Cells[row, 7].Value = nsg.SecurityRules[ruleName].SourceAddressPrefix;
                    }
                    // Source Port
                    if (nsg.SecurityRules[ruleName].SourcePortRange == null)
                    {
                        ExHandler.worksheet.Cells[row, 8].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourcePortRanges);
                    }
                    else
                    {
                        ExHandler.worksheet.Cells[row, 8].Value = nsg.SecurityRules[ruleName].SourcePortRange;
                    }
                    ExHandler.worksheet.Cells[row, 9].Value = nsg.SecurityRules[ruleName].Direction;
                    // Destination Address
                    if (nsg.SecurityRules[ruleName].DestinationAddressPrefix == null)
                    {
                        ExHandler.worksheet.Cells[row, 10].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationAddressPrefixes);
                    }
                    else
                    {
                        ExHandler.worksheet.Cells[row, 10].Value = nsg.SecurityRules[ruleName].DestinationAddressPrefix;
                    }
                    // Destination Port
                    if (nsg.SecurityRules[ruleName].DestinationPortRange == null)
                    {
                        ExHandler.worksheet.Cells[row, 11].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationPortRanges);
                    }
                    else
                    {
                        ExHandler.worksheet.Cells[row, 11].Value = nsg.SecurityRules[ruleName].DestinationPortRange;
                    }
                    ExHandler.worksheet.Cells[row, 12].Value = nsg.SecurityRules[ruleName].Access;
                    row++;
                }
            }
            // Style
            ExHandler.SetStyle(ExcelHandler.StyleNames.Header, 1, 1, 1, 12);
            ExHandler.SetStyle(ExcelHandler.StyleNames.BorderBox, 1, 1, row-1, 12);
        }
        // Saving
        ExHandler.SaveSheet(TextHandler.CurrentPath + @"\NSG");
    }

    public void GetUDR()
    {
        ExHandler = new ExcelHandler();
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > UDR Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            ExHandler.CreateSheet(sub.DisplayName);
            // Adding Column Headers
            ExHandler.worksheet.Cells[1, 1].Value = "UDR Name";
            ExHandler.worksheet.Cells[1, 2].Value = "Resource Group";
            ExHandler.worksheet.Cells[1, 3].Value = "Region";
            ExHandler.worksheet.Cells[1, 4].Value = "Next Hop IP";
            ExHandler.worksheet.Cells[1, 5].Value = "Next Hop Type";
            ExHandler.worksheet.Cells[1, 6].Value = "Destination Address Prefix";
            int row = 2;
            foreach (INetwork network in AzHandler.Networks.List())
            {
                foreach(ISubnet subnet in network.Subnets.Values)
                {
                    IRouteTable routes = subnet.GetRouteTable();
                    if(routes == null)
                    {
                        continue;
                    }
                    foreach (IRoute route in routes.Routes.Values)
                    {
                        ExHandler.worksheet.Cells[row, 1].Value = route.Name;
                        ExHandler.worksheet.Cells[row, 2].Value = routes.ResourceGroupName;
                        ExHandler.worksheet.Cells[row, 3].Value = routes.RegionName;
                        ExHandler.worksheet.Cells[row, 4].Value = route.NextHopIPAddress;
                        ExHandler.worksheet.Cells[row, 5].Value = route.NextHopType;
                        ExHandler.worksheet.Cells[row, 6].Value = route.DestinationAddressPrefix;
                        row++;
                    }
                }
            }
            // Style
            ExHandler.SetStyle(ExcelHandler.StyleNames.Header, 1, 1, 1, 6);
            ExHandler.SetStyle(ExcelHandler.StyleNames.BorderBox, 1, 1, row-1, 6);
        }
        // Saving
        ExHandler.SaveSheet(TextHandler.CurrentPath + @"\UDR");
    }

    public void GetVNET()
    {
        ExHandler = new ExcelHandler();
        // Iterating over subscriptions
        foreach (ISubscription sub in AzHandler.Subscriptions.List())
        {
            TextHandler.Banner("Home > Network > VNET Report");
            TextHandler.ShowMsg("Fetching data on subscription: " + sub.DisplayName, currentState: TextHandler.MessageState.Information);
            AzHandler = Azure.Configure().Authenticate(AzCredentials).WithSubscription(sub.SubscriptionId);
            ExHandler.CreateSheet(sub.DisplayName);
            // Adding Column Headers
            ExHandler.worksheet.Cells[1, 1].Value = "VNET Name";
            ExHandler.worksheet.Cells[1, 2].Value = "Resource Group";
            ExHandler.worksheet.Cells[1, 3].Value = "Region";
            ExHandler.worksheet.Cells[1, 4].Value = "Address Space";
            ExHandler.worksheet.Cells[1, 5].Value = "Subnet Name";
            ExHandler.worksheet.Cells[1, 6].Value = "Address Prefix";
            ExHandler.worksheet.Cells[1, 7].Value = "Attached NSG";
            ExHandler.worksheet.Cells[1, 8].Value = "Attached UDR";
            ExHandler.worksheet.Cells[1, 9].Value = "Available IPs";
            int row = 2;
            foreach (INetwork network in AzHandler.Networks.List())
            {
                foreach (ISubnet subnet in network.Subnets.Values)
                {
                    ExHandler.worksheet.Cells[row, 1].Value = network.Name;
                    ExHandler.worksheet.Cells[row, 2].Value = network.ResourceGroupName;
                    ExHandler.worksheet.Cells[row, 3].Value = network.RegionName;
                    ExHandler.worksheet.Cells[row, 4].Value = string.Join(",", network.AddressSpaces);
                    ExHandler.worksheet.Cells[row, 5].Value = subnet.Name;
                    ExHandler.worksheet.Cells[row, 6].Value = subnet.AddressPrefix;
                    try
                    {
                        ExHandler.worksheet.Cells[row, 7].Value = subnet.GetNetworkSecurityGroup().Name;
                    }
                    catch (Exception)
                    {
                        ExHandler.worksheet.Cells[row, 7].Value = "NA";
                    }
                    try
                    {
                        ExHandler.worksheet.Cells[row, 8].Value = subnet.GetRouteTable().Name;
                    }
                    catch (Exception)
                    {
                        ExHandler.worksheet.Cells[row, 8].Value = "NA";
                    }
                    ExHandler.worksheet.Cells[row, 9].Value = string.Join(",", subnet.ListAvailablePrivateIPAddresses());
                    row++;
                }
            }
            // Style
            ExHandler.SetStyle(ExcelHandler.StyleNames.Header, 1, 1, 1, 9);
            ExHandler.SetStyle(ExcelHandler.StyleNames.BorderBox, 1, 1, row-1, 9);
        }
        // Saving
        ExHandler.SaveSheet(TextHandler.CurrentPath + @"\VNET");
    }
}

