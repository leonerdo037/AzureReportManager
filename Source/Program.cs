using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace ARM
{
    class Program
    {
        static string curPath = Environment.CurrentDirectory + @"\etc\network\";
        static string configFile = curPath + @"\ARM.config";

        static void Main(string[] args)
        {
            Banner();
            if (!File.Exists(configFile))
            {
                Setup();
            }
        }

        static void Banner()
        {
            Console.Clear();
            Console.WriteLine("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+");
            Console.WriteLine("|A|Z|U|R|E| |R|E|P|O|R|T| |M|A|N|A|G|E|R|");
            Console.WriteLine("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+");
            Console.WriteLine("Version 0.1" + Environment.NewLine);
        }

        static void Setup()
        {
            Banner();

        }
    }
}