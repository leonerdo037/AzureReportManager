using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ARM
{
    class Program
    {
        static void Main(string[] args)
        {
            TextHandler.SetColor();
            if (!File.Exists(TextHandler.ConfigFile))
            {
                Setup();
            }
            else
            {
                Home();
            }
        }

        static void Home()
        {
            TextHandler.Banner("Home");
            TextHandler.ShowMsg("1. About");
            TextHandler.ShowMsg("2. Edit Configuration");
            TextHandler.ShowMsg("3. Network Reports");
            TextHandler.ShowMsg("4. Compute Reports");
            TextHandler.ShowMsg("0. Exit", tailBreak:true);
            switch (TextHandler.ReadInput("Enter your choice:"))
            {
                case "0":
                    Environment.Exit(0);
                    break;
                case "1":
                    Home();
                    break;
                case "2":
                    Setup();
                    break;
                case "3":
                    Network();
                    break;
                default:
                    Home();
                    break;
            }
        }

        static void Setup()
        {
            TextHandler.Banner("Settings");
            // Creating Configuration File
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            StreamWriter stw;
            TextHandler.ShowMsg("Creating configuration file . . .", headBreak:true, 
                        tailBreak:true, currentState: TextHandler.MessageState.Information);
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartObject();
                writer.WritePropertyName("Tenant ID");
                writer.WriteValue(TextHandler.ReadInput("Enter Azure Tenant ID: "));
                writer.WritePropertyName("Client ID");
                writer.WriteValue(TextHandler.ReadInput("Enter Azure Client ID: "));
                writer.WritePropertyName("Client Secret");
                writer.WriteValue(TextHandler.ReadInput("Enter Azure Client Secret: "));
                writer.WriteEndObject();
            }
            // Saving the configuration file
            stw = new StreamWriter(TextHandler.ConfigFile);
            try
            {
                stw.Write(sb);
                TextHandler.ShowMsg("Configuration file saved . . .", headBreak:true
                                    , currentState: TextHandler.MessageState.Success);
                TextHandler.Pause();
            }
            catch (Exception ex)
            {
                TextHandler.ShowMsg("Error: " + ex.Message, headBreak: true,
                                    currentState: TextHandler.MessageState.Error);
                TextHandler.Pause();
            }
            finally
            {
                stw.Flush();
                stw.Close();
                Home();
            }
        }

        static void Network()
        {
            NetworkHandler NwHandler = new NetworkHandler();
            TextHandler.Banner("Home > Network");
            TextHandler.ShowMsg("1. NSG Report");
            TextHandler.ShowMsg("2. UDR Report");
            TextHandler.ShowMsg("3. VNET Report");
            TextHandler.ShowMsg("4. Load Balancer Report");
            TextHandler.ShowMsg("9. Go Back");
            TextHandler.ShowMsg("0. Exit", tailBreak: true);
            switch (TextHandler.ReadInput("Enter your choice:"))
            {
                case "0":
                    Environment.Exit(0);
                    break;
                case "1":
                    try
                    {
                        NwHandler.GetNSG();
                        TextHandler.ShowMsg(string.Format("Report saved in the path: {0}", TextHandler.CurrentPath),
                        headBreak: true, currentState: TextHandler.MessageState.Success);
                    }
                    catch (Exception ex)
                    {
                        TextHandler.ShowMsg("Error: " + ex.Message, headBreak: true,
                                    currentState: TextHandler.MessageState.Error);
                    }
                    TextHandler.Pause();
                    Network();
                    break;
                case "2":
                    try
                    {
                        NwHandler.GetUDR();
                        TextHandler.ShowMsg(string.Format("Report saved in the path: {0}", TextHandler.CurrentPath),
                        headBreak: true, currentState: TextHandler.MessageState.Success);
                    }
                    catch (Exception ex)
                    {
                        TextHandler.ShowMsg("Error: " + ex.Message, headBreak: true,
                                    currentState: TextHandler.MessageState.Error);
                    }
                    TextHandler.Pause();
                    Network();
                    break;
                case "3":
                    NwHandler.GetVNET();
                    try
                    {
                        
                        TextHandler.ShowMsg(string.Format("Report saved in the path: {0}", TextHandler.CurrentPath),
                        headBreak: true, currentState: TextHandler.MessageState.Success);
                    }
                    catch (Exception ex)
                    {
                        TextHandler.ShowMsg("Error: " + ex.Message, headBreak: true,
                                    currentState: TextHandler.MessageState.Error);
                    }
                    TextHandler.Pause();
                    Network();
                    break;
                case "4":
                    Network();
                    break;
                case "9":
                    Home();
                    break;
                default:
                    Network();
                    break;
            }
        }
    }
}