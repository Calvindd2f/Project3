#pragma managed

#include <msclr/marshal_cppstd.h>
#include <vcclr.h>
#include <string>
#include "../../../../../Program Files (x86)/Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8.1/System.dll"

using namespace System;
using namespace std;
using namespace System::Management::Automation;
using namespace System::Management::Automation::Runspaces;
using namespace msclr::interop;
using namespace System::IO;
using namespace System::Diagnostics;
using namespace System::Reflection;
using namespace System::Management::Automation::Host;
using namespace System::Runtime::InteropServices;

typedef void (*LogCallbackFunction)(const char*, int);

struct Callbacks
{
    LogCallbackFunction LogCallback;
};

public
enum class LogOutputType
{
    Error = 1,
    Verbose = 2,
    Information = 3,
    Warning = 4,
    Debug = 5
};

public
ref class PowerShellExecutionResult
{
public:
    property bool Success;
    property String^ Output;
    property String^ ErrorMessage;
};

public
ref class PowerShellExecutor
{
public:
    PowerShellExecutor(IntPtr callbacks);
    PowerShellExecutionResult^ ExecutePowerShell(String^ script, bool isInlinePowershell);

private:
    IntPtr callbacksPtr;
    int activityLogCounter;
    int verboseLinesProcessed;
    int warningLinesProcessed;
    int errorLinesProcessed;
    int informationLinesProcessed;
    static const int ActivityLogThreshold = 1000;

    void SendLog(String^ logOutput, LogOutputType logType);
    void BindEvents(PowerShell^ ps, PSHost^ host);
    void Debug_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void Progress_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void Error_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void Verbose_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void Warning_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void Information_DataAdded(Object^ sender, DataAddedEventArgs^ e);
    void OnOutputDataReceived(Object^ sender, DataReceivedEventArgs^ e);
    void Host_OnInformation(String^ information);
    void HandleInformation(String^ logOutput);
};