using Microsoft.Azure.Management.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent.Authentication;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;

public class AzureHandler
{
    protected AzureCredentials AzCredentials;
    protected IAzure AzHandler;
    protected string TenantID;
    protected string ClientID;
    protected string Secret;

    public AzureHandler()
    {
        // Creating Reports Directory
        if (!Directory.Exists(TextHandler.ReportPath))
        {
            Directory.CreateDirectory(TextHandler.ReportPath);
        }
        // Creating the Day's Directory
        if (!Directory.Exists(TextHandler.CurrentPath))
        {
            Directory.CreateDirectory(TextHandler.CurrentPath);
        }
        // Reading Config File
        string RawConfig = File.OpenText(TextHandler.ConfigFile).ReadToEnd();
        JToken ConfigData = JsonConvert.DeserializeObject<JToken>(RawConfig);
        TenantID = ConfigData["Tenant ID"].ToString();
        ClientID = ConfigData["Client ID"].ToString();
        Secret = ConfigData["Client Secret"].ToString();
        // Authenticating Azure
        AzCredentials = SdkContext.AzureCredentialsFactory.FromServicePrincipal(ClientID, Secret, TenantID, AzureEnvironment.AzureGlobalCloud);
        AzHandler = Azure.Configure().Authenticate(AzCredentials).WithDefaultSubscription();
    }
}