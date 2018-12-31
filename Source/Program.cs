using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
            // Reading Args
            if (args.Length == 0)
            {
                Interactive();
            }
            else if(args.Length == 1 && (args[0].ToLower() == "-h" || args[0].ToLower() == "-help"))
            {
                Help();
            }
            else if(args.Length == 4)
            {
                CLI(args);
            }
            else
            {
                Help();
            }
        }

        static void Help()
        {
            Console.WriteLine("Azure Report Manager CLI v1.0");
            Console.WriteLine("Interactive Usage: dotnet ARM.dll");
            Console.WriteLine("Command-line Usage: dotnet ARM.dll [arguments]" + TextHandler.NewLine);
            Console.WriteLine("Arguments(Not case sensitive):");
            Console.WriteLine("  -r|-resource  Flag to specify the Azure resource");
            Console.WriteLine("     eg: dotnet ARM.dll -r Network|Compute" + TextHandler.NewLine);
            Console.WriteLine("  -t|-type  Flag to specify the report type");
            Console.WriteLine("     eg: dotnet ARM.dll -r Network -t NSG|UDR|VNET|LB");
            Console.WriteLine("         dotnet ARM.dll -r Compute -t VM|RG|DISK" + TextHandler.NewLine);
        }

        static void CLI(string[] args)
        {
            if((args[0].ToLower() == "-r" || args[0].ToLower() == "-resource") &&
               (args[2].ToLower() == "-t" || args[2].ToLower() == "-type"))
            {
                // Generating Report
                switch (args[1].ToLower())
                {
                    case "network":
                        NetworkHandler NwHandler = new NetworkHandler();
                        switch (args[3].ToLower())
                        {
                            case "nsg":
                                NwHandler.GetNSG();
                                break;
                            case "udr":
                                NwHandler.GetUDR();
                                break;
                            case "vnet":
                                NwHandler.GetVNET();
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        Help();
                        break;
                }
            }
            else
            {
                Help();
            }
            switch (args[0].ToLower())
            {
                case "-r":
                case "-resource":
                    // Resource Type
                    switch (args[2].ToLower())
                    {
                        case "-t":
                        case "-type":
                            // Fetching Report
                            switch (args[1].ToLower())
                            {
                                case "network":
                                    break;
                                default:
                                    break;
                            }
                            break;
                        default:
                            Help();
                            break;
                    }
                    break;
                default:
                    Help();
                    break;
            }
        }

        static void Interactive()
        {
            // Interactive Mode
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
                    try
                    {
                        NwHandler.GetVNET();
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