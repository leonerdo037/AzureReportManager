using Microsoft.Azure.Management.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent.Authentication;
using System;
using System.Collections.Generic;
using System.Text;

public class AzureHandler
{
    private AzureCredentials credentials;
    private IAzure azure;

    public AzureHandler()
    {
    }
}