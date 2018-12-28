using Microsoft.Azure.Management.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent.Authentication;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public class AzureHandler
{
    private AzureCredentials AzCredentials;
    private IAzure AzHandler;
    protected string TenantID;
    protected string ClientID;
    protected string Secret;

    public AzureHandler()
    {
        // Reading Config File
        string RawConfig = File.OpenText(TextHandler.configFile).ReadToEnd();
        JToken ConfigData = JsonConvert.DeserializeObject<JToken>(RawConfig);
        TenantID = ConfigData["Tenant ID"].ToString();
        ClientID = ConfigData["Client ID"].ToString();
        Secret = ConfigData["Client Secret"].ToString();
        // Authenticating Azure
        AzCredentials = SdkContext.AzureCredentialsFactory.FromServicePrincipal(ClientID, Secret, TenantID, AzureEnvironment.AzureGlobalCloud);
        AzHandler = Azure.Configure().Authenticate(AzCredentials).WithDefaultSubscription();
    }

    public void NSG()
    {
        // Creating Directory
        if (!System.IO.Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Jarvis"))
        {
            System.IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Jarvis");
        }
        // Saving NSG data to Excel
        // Creating Sheets
        var package = new ExcelPackage();
        ExcelWorksheet worksheet = null;
        azure = Azure.Configure().Authenticate(credentials).WithSubscription((CB_Sub.SelectedItem as ISubscription).SubscriptionId);
        int sheet = 1;
        foreach (INetworkSecurityGroup nsg in azure.NetworkSecurityGroups.List())
        {
            worksheet = package.Workbook.Worksheets.Add(sheet + " " + nsg.Name);
            // Iterating over Data
            int col = 1;
            foreach (DataColumn dc in DT.Columns)
            {
                // Add the headers
                worksheet.Cells[1, col].Value = dc.ColumnName;
                col++;
            }
            int row = 2;
            foreach (string ruleName in nsg.SecurityRules.Keys)
            {
                worksheet.Cells[row, 1].Value = nsg.SecurityRules[ruleName].Priority;
                worksheet.Cells[row, 2].Value = nsg.SecurityRules[ruleName].Name;
                worksheet.Cells[row, 3].Value = nsg.SecurityRules[ruleName].Protocol;
                // Source Address
                if (nsg.SecurityRules[ruleName].SourceAddressPrefix == null)
                {
                    worksheet.Cells[row, 4].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourceAddressPrefixes);
                }
                else
                {
                    worksheet.Cells[row, 4].Value = nsg.SecurityRules[ruleName].SourceAddressPrefix;
                }
                // Source Port
                if (nsg.SecurityRules[ruleName].SourcePortRange == null)
                {
                    worksheet.Cells[row, 5].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.SourcePortRanges);
                }
                else
                {
                    worksheet.Cells[row, 5].Value = nsg.SecurityRules[ruleName].SourcePortRange;
                }
                worksheet.Cells[row, 6].Value = nsg.SecurityRules[ruleName].Direction;
                // Destination Address
                if (nsg.SecurityRules[ruleName].DestinationAddressPrefix == null)
                {
                    worksheet.Cells[row, 7].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationAddressPrefixes);
                }
                else
                {
                    worksheet.Cells[row, 7].Value = nsg.SecurityRules[ruleName].DestinationAddressPrefix;
                }
                // Destination Port
                if (nsg.SecurityRules[ruleName].DestinationPortRange == null)
                {
                    worksheet.Cells[row, 8].Value = string.Join(",", nsg.SecurityRules[ruleName].Inner.DestinationPortRanges);
                }
                else
                {
                    worksheet.Cells[row, 8].Value = nsg.SecurityRules[ruleName].DestinationPortRange;
                }
                worksheet.Cells[row, 9].Value = nsg.SecurityRules[ruleName].Access;
                row++;
            }
            // Style
            using (var range = worksheet.Cells[1, 1, 1, col - 1])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                range.Style.Font.Color.SetColor(System.Drawing.Color.White);
            }
            sheet++;
        }
        // Saving
        var xlFile = new System.IO.FileInfo(string.Format("{0}{1}{2}_{3}_{4}_{5}_{6}_{7}_{8}_{9}", Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            @"\Jarvis\", (CB_Sub.SelectedValue as ISubscription).DisplayName, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, ".xlsx"));
        package.SaveAs(xlFile);
        MessageBox.Show("Report saved in Jarvis folder which is located in your desktop !", "Jarvis - Network Manager", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}