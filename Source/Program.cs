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
        private static string NewLine = Environment.NewLine;

        static void Main(string[] args)
        {
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
            ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+");
            ShowMsg("|A|Z|U|R|E| |R|E|P|O|R|T| |M|A|N|A|G|E|R|");
            ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+");
            ShowMsg("ARM > " + header);
            string line = "++++++";
            foreach (char letter in header)
            {
                line = line + "+";
            }
            ShowMsg(line, tailBreak:true);
        }

        static void Pause()
        {
            ShowMsg("Press any key to continue . . .", headBreak:true, singleLine:true);
            Console.ReadKey();
        }

        static void ShowMsg(string msg, bool headBreak=false, bool tailBreak=false, bool singleLine=false)
        {
            // Handling Line Breaks
            if (headBreak && tailBreak)
            {
                msg = NewLine + msg + NewLine;
            }
            else if (!headBreak && tailBreak)
            {
                msg = msg + NewLine;
            }
            else if (headBreak && !tailBreak)
            {
                msg = NewLine + msg;
            }
            // Printing Message
            if (singleLine)
            {
                Console.Write(msg);
            }
            else
            {
                Console.WriteLine(msg);
            }
        }

        static string ReadInput(string msg)
        {
            ShowMsg(msg, headBreak:false, tailBreak:false, singleLine:true);
            return Console.ReadLine();
        }

        static void Home()
        {
            Banner("Home");

        }

        static void Setup()
        {
            Banner("Settings");
            // Creating Configuration File
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            StreamWriter stw;
            ShowMsg("Creating configuration file . . .", headBreak:true, tailBreak:true);
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartObject();
                writer.WritePropertyName("Tenant ID");
                writer.WriteValue(ReadInput("Enter Azure Tenant ID: "));
                writer.WritePropertyName("Client ID");
                writer.WriteValue(ReadInput("Enter Azure Client ID: "));
                writer.WritePropertyName("Client Secret");
                writer.WriteValue(ReadInput("Enter Azure Client Secret: "));
                writer.WriteEndObject();
            }
            // Saving the configuration file
            stw = new StreamWriter(configFile);
            try
            {
                stw.Write(sb);
                ShowMsg("Configuration file created . . .", headBreak:true);
                Pause();
            }
            catch (Exception ex)
            {
                ShowMsg("Error: " + ex.Message, headBreak: true);
                Pause();
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