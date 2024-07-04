using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Runtime.CompilerServices;
using System.Security;

public class DefaultHostUserInterface : PSHostUserInterface, IHostUISupportsMultipleChoiceSelection
{
    private PSHostRawUserInterface _psRawUserInterface = new DefaultHostRawUserInterface();
    private InformationDelegate _onInformation;

    public delegate void InformationDelegate(string information);

    public event InformationDelegate OnInformation
    {
        [MethodImpl(MethodImplOptions.Synchronized)]
        add
        {
            _onInformation = (InformationDelegate)Delegate.Combine(_onInformation, value);
            GC.KeepAlive(this);
        }
        [MethodImpl(MethodImplOptions.Synchronized)]
        remove
        {
            _onInformation = (InformationDelegate)Delegate.Remove(_onInformation, value);
            GC.KeepAlive(this);
        }
    }

    protected void RaiseOnInformation(string value)
    {
        _onInformation?.Invoke(value);
    }

    private static string[] GetHotkeyAndLabel(string input)
    {
        string[] result = new string[2];
        string[] parts = input.Split('&');
        if (parts.Length == 2)
        {
            if (parts[1].Length > 0)
            {
                result[0] = parts[1][0].ToString().ToUpper(CultureInfo.CurrentCulture);
            }
            result[1] = (parts[0] + parts[1]).Trim();
        }
        else
        {
            result[1] = input;
        }
        return result;
    }

    private static string[,] BuildHotkeysAndPlainLabels(Collection<ChoiceDescription> choices)
    {
        string[,] result = new string[2, choices.Count];
        for (int i = 0; i < choices.Count; i++)
        {
            string[] hotkeyAndLabel = GetHotkeyAndLabel(choices[i].Label);
            result[0, i] = hotkeyAndLabel[0];
            result[1, i] = hotkeyAndLabel[1];
        }
        return result;
    }

    public override PSHostRawUserInterface RawUI => _psRawUserInterface;

    public override string ReadLine()
    {
        throw new NotImplementedException();
    }

    public override SecureString ReadLineAsSecureString()
    {
        throw new NotImplementedException();
    }

    public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
    {
        RaiseOnInformation(value);
    }

    public override void Write(string value)
    {
        RaiseOnInformation(value);
    }

    public override void WriteLine(string value)
    {
        RaiseOnInformation(value);
    }

    public override void WriteErrorLine(string value)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(value);
        Console.ResetColor();
    }

    public override void WriteDebugLine(string message)
    {
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    public override void WriteProgress(long sourceId, ProgressRecord record)
    {
        Console.WriteLine($"{record.PercentComplete}% {record.Activity} {record.StatusDescription}");
    }

    public override void WriteVerboseLine(string message)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    public override void WriteWarningLine(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions)
    {
        throw new NotImplementedException("Prompt method is not implemented.");
    }

    public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
    {
        throw new NotImplementedException("PromptForCredential method with options is not implemented.");
    }

    public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName)
    {
        throw new NotImplementedException("PromptForCredential method is not implemented.");
    }

    public virtual Collection<int> PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, IEnumerable<int> defaultChoices)
    {
        throw new NotImplementedException("PromptForChoice method is not implemented.");
    }

    public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice)
    {
        throw new NotImplementedException("PromptForChoice method is not implemented.");
    }
}