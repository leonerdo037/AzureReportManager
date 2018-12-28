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
}