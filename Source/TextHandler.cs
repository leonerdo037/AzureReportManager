﻿using System;
using System.Collections.Generic;
using System.Text;

public static class TextHandler
{
    public static string AppPath = Environment.CurrentDirectory;
    public static string ConfigFile = AppPath + @"\config.json";
    public static string ReportPath = AppPath + @"\Reports";
    public static string CurrentPath = string.Format("{0}\\{1}-{2:MMMM}-{3}", ReportPath, 
                                                    DateTime.Now.Day, DateTime.Now, 
                                                    DateTime.Now.Year);
    public static string NewLine = Environment.NewLine;
    public enum MessageState
    {
        Default,
        Banner,
        Input,
        Information,
        Pause,
        Warning,
        Success,
        Error
    }

    public static void Banner(string header)
    {
        Console.Clear();
        ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+", currentState: MessageState.Banner);
        ShowMsg("|A|Z|U|R|E| |R|E|P|O|R|T| |M|A|N|A|G|E|R|", currentState: MessageState.Banner);
        ShowMsg("+-+-+-+-+-+ +-+-+-+-+-+-+ +-+-+-+-+-+-+-+", currentState: MessageState.Banner);
        ShowMsg("ARM > " + header, currentState: MessageState.Banner);
        string line = "++++++";
        foreach (char letter in header)
        {
            line = line + "+";
        }
        ShowMsg(line, tailBreak: true, currentState: MessageState.Banner);
    }

    public static void Pause()
    {
        ShowMsg("Press any key to continue . . .", headBreak: true, singleLine: true
                , currentState: MessageState.Pause);
        Console.ReadKey();
    }

    public static void ShowMsg(string msg, bool headBreak = false, 
                               bool tailBreak = false, bool singleLine = false, 
                               MessageState currentState=MessageState.Default)
    {
        // Setting Color
        switch (currentState)
        {
            case MessageState.Information:
                SetColor(foreground: ConsoleColor.Blue, background: ConsoleColor.White);
                break;
            case MessageState.Warning:
                SetColor(foreground: ConsoleColor.DarkYellow, background: ConsoleColor.White);
                break;
            case MessageState.Success:
                SetColor(foreground: ConsoleColor.DarkGreen, background: ConsoleColor.White);
                break;
            case MessageState.Error:
                SetColor(foreground: ConsoleColor.Red, background: ConsoleColor.White);
                break;
            case MessageState.Banner:
                SetColor(foreground: ConsoleColor.DarkMagenta, background: ConsoleColor.White);
                break;
            case MessageState.Input:
                SetColor(foreground: ConsoleColor.White, background: ConsoleColor.Black);
                break;
            case MessageState.Pause:
                SetColor(foreground: ConsoleColor.White, background: ConsoleColor.DarkGray);
                break;
            default:
                SetColor();
                break;
        }
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
        // Resetting Color
        SetColor();
    }

    public static string ReadInput(string msg)
    {
        ShowMsg(msg, headBreak: false, tailBreak: false, singleLine: true, currentState: MessageState.Input);
        Console.Write(" ");
        return Console.ReadLine();
    }

    public static void SetColor(ConsoleColor foreground = ConsoleColor.Black, ConsoleColor background = ConsoleColor.White)
    {
        Console.BackgroundColor = background;
        Console.ForegroundColor = foreground;
    }
}
