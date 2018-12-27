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
        private static string curPath = Environment.CurrentDirectory;
        private static string configFile = curPath + @"\config.json";

        static void Main(string[] args)
        {
            TextHandler.SetColor();
            if (!File.Exists(configFile))
            {
                Setup();
            }
            else
            {
                Home();
            }
        }

        static void Banner(string header)
        {
            Console.Clear();
            TextHandler.ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+", currentState: TextHandler.MessageState.Banner);
            TextHandler.ShowMsg("|A|Z|U|R|E| |R|E|P|O|R|T| |M|A|N|A|G|E|R|", currentState: TextHandler.MessageState.Banner);
            TextHandler.ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+", currentState: TextHandler.MessageState.Banner);
            TextHandler.ShowMsg("ARM > " + header, currentState: TextHandler.MessageState.Banner);
            string line = "++++++";
            foreach (char letter in header)
            {
                line = line + "+";
            }
            TextHandler.ShowMsg(line, tailBreak:true, currentState: TextHandler.MessageState.Banner);
        }

        static void Home()
        {
            Banner("Home");
            TextHandler.ShowMsg("1. View Reports");
            TextHandler.ShowMsg("2. Create Reports");
            TextHandler.ShowMsg("3. Delete Reports");
            TextHandler.ShowMsg("4. Edit Configuration");
            TextHandler.ShowMsg("5. About");
            TextHandler.ShowMsg("0. Exit", tailBreak:true);
            switch (TextHandler.ReadInput("Enter your choice:"))
            {
                case "0":
                    Environment.Exit(0);
                    break;
                case "1":
                    break;
                case "2":
                    break;
                case "3":
                    break;
                case "4":
                    Setup();
                    break;
                case "5":
                    break;
                default:
                    Home();
                    break;
            }
        }

        static void Setup()
        {
            Banner("Settings");
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
            stw = new StreamWriter(configFile);
            try
            {
                stw.Write(sb);
                TextHandler.ShowMsg("Configuration file created . . .", headBreak:true
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
    }
}