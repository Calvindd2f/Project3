#include "Header.h"


PowerShellExecutor::PowerShellExecutor(IntPtr callbacks)
    : callbacksPtr(callbacks), activityLogCounter(0), verboseLinesProcessed(0),
    warningLinesProcessed(0), errorLinesProcessed(0), informationLinesProcessed(0)
{
}

void PowerShellExecutor::SendLog(String^ logOutput, LogOutputType logType)
{
    if (activityLogCounter > ActivityLogThreshold)
    {
        throw gcnew Exception("Activity Log threshold exceeded.");
    }

    auto callbacksNative = static_cast<Callbacks*>(callbacksPtr.ToPointer());
    std::string utf8LogOutput = marshal_as<std::string>(logOutput);
    callbacksNative->LogCallback(utf8LogOutput.c_str(), static_cast<int>(logType));

    activityLogCounter++;
}

void PowerShellExecutor::Debug_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto debugRecords = safe_cast<PSDataCollection<DebugRecord^>^>(sender);
    String^ text = String::Join(Environment::NewLine, debugRecords);
    SendLog(text, LogOutputType::Debug);
}

void PowerShellExecutor::Progress_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto progressRecords = safe_cast<PSDataCollection<ProgressRecord^>^>(sender);
    for each (ProgressRecord ^ record in progressRecords)
    {
        String^ message = String::Format("Activity: {0}, Status: {1}, Progress: {2}%",
            record->Activity, record->StatusDescription, record->PercentComplete);
        SendLog(message, LogOutputType::Information);
    }
}

void PowerShellExecutor::Error_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto errorRecords = safe_cast<PSDataCollection<ErrorRecord^>^>(sender);
    for each (ErrorRecord ^ errorRecord in errorRecords)
    {
        SendLog(errorRecord->ToString(), LogOutputType::Error);

        if (errorRecord->Exception != nullptr)
        {
            SendLog(errorRecord->Exception->ToString(), LogOutputType::Error);

            auto webEx = dynamic_cast<System::Net::WebException^>(errorRecord->Exception->InnerException);
            if (webEx != nullptr)
            {
                try
                {
                    String^ responseText = gcnew System::IO::StreamReader(webEx->Response->GetResponseStream())->ReadToEnd();
                    SendLog("Web Exception Details: " + responseText, LogOutputType::Error);
                }
                catch (Exception^ ex)
                {
                    SendLog("Error reading web exception details: " + ex->ToString(), LogOutputType::Error);
                }
            }
        }
    }
}

void PowerShellExecutor::Verbose_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto verboseRecords = safe_cast<PSDataCollection<VerboseRecord^>^>(sender);
    String^ text = String::Join(Environment::NewLine, verboseRecords);
    SendLog(text, LogOutputType::Verbose);
    verboseLinesProcessed += verboseRecords->Count;
}

void PowerShellExecutor::Warning_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto warningRecords = safe_cast<PSDataCollection<WarningRecord^>^>(sender);
    String^ text = String::Join(Environment::NewLine, warningRecords);
    SendLog(text, LogOutputType::Warning);
    warningLinesProcessed += warningRecords->Count;
}

void PowerShellExecutor::Information_DataAdded(Object^ sender, DataAddedEventArgs^ e)
{
    auto informationRecords = safe_cast<PSDataCollection<InformationRecord^>^>(sender);
    String^ text = String::Join(Environment::NewLine, informationRecords);
    SendLog(text, LogOutputType::Information);
    informationLinesProcessed += informationRecords->Count;
}

void PowerShellExecutor::OnOutputDataReceived(Object^ sender, DataReceivedEventArgs^ e)
{
    if (!String::IsNullOrEmpty(e->Data))
    {
        SendLog(e->Data, LogOutputType::Information);
    }
}

void PowerShellExecutor::Host_OnInformation(String^ information)
{
    String^ text = information->Trim();
    if (!String::IsNullOrEmpty(text))
    {
        SendLog(text, LogOutputType::Information);
    }
}

void PowerShellExecutor::HandleInformation(String^ logOutput)
{
    if (!String::IsNullOrEmpty(logOutput))
    {
        SendLog(logOutput, LogOutputType::Information);
    }
}

void PowerShellExecutor::BindEvents(PowerShell^ ps, PSHost^ host)
{
    ps->Streams->Debug->DataAdded += gcnew EventHandler<DataAddedEventArgs^>(this, &PowerShellExecutor::Debug_DataAdded);
    ps->Streams->Error->DataAdded += gcnew EventHandler<DataAddedEventArgs^>(this, &PowerShellExecutor::Error_DataAdded);
    ps->Streams->Progress->DataAdded += gcnew EventHandler<DataAddedEventArgs^>(this, &PowerShellExecutor::Progress_DataAdded);
    ps->Streams->Verbose->DataAdded += gcnew EventHandler<DataAddedEventArgs^>(this, &PowerShellExecutor::Verbose_DataAdded);
    ps->Streams->Warning->DataAdded += gcnew EventHandler<DataAddedEventArgs^>(this, &PowerShellExecutor::Warning_DataAdded);

    auto informationProperty = ps->GetType()->GetProperty("Information");
    if (informationProperty != nullptr)
    {
        Object^ informationObject = informationProperty->GetValue(ps->Streams, nullptr);
        EventInfo^ dataAddedEvent = informationObject->GetType()->GetEvent("DataAdded");
        MethodInfo^ informationMethod = this->GetType()->GetMethod("Information_DataAdded", System::Reflection::BindingFlags::Instance | System::Reflection::BindingFlags::NonPublic);
        Delegate^ handler = Delegate::CreateDelegate(dataAddedEvent->EventHandlerType, this, informationMethod);
        dataAddedEvent->AddEventHandler(informationObject, handler);
    }
    else
    {
        auto defaultHost = dynamic_cast<DefaultHost^>(host);
        if (defaultHost != nullptr)
        {
            defaultHost->OnInformation += gcnew Action<String^>(this, &PowerShellExecutor::Host_OnInformation);
        }
    }
}

PowerShellExecutionResult^ ExecutePowerShell(String^ script, bool isInlinePowershell)
{
    auto result = gcnew PowerShellExecutionResult();
    try
    {
        auto runspace = RunspaceFactory::CreateRunspace(gcnew DefaultHost(CultureInfo::CurrentCulture, CultureInfo::CurrentUICulture));
        runspace->Open();

        auto ps = PowerShell::Create();
        ps->Runspace = runspace;

        String^ initScript = R"(null)";

        ps->AddScript(initScript);
        ps->Invoke();

		//PowerShellExecutor::BindEvents(ps, ps->Host);
        //String ^ deserializedScript = DeserializeScriptVariables(script);

		if (!isInlinePowershell)
		{
			ps->AddScript(script);
		}
		else
		{
			ps->AddCommand(script);
            ps->AddCommand("ExecuteActivity");
		}

        auto psInvocationSettings = gcnew PSInvocationSettings();
        psInvocationSettings->Host = runspace->SessionStateProxy->Host;

        auto outputCollection = gcnew PSDataCollection<PSObject^>();
        ps->Invoke(nullptr, outputCollection, psInvocationSettings);

        if (outputCollection->Count == 0)
        {
            throw gcnew Exception("Activity did not return a result and/or failed while executing");
        }
        if (outputCollection->Count > 1)
        {
            throw gcnew Exception("Activity returned more than one result. See below for details");
        }

        result->Output = outputCollection[0]->ToString();
        result->Success = true;
    }
    catch (Exception^ ex)
	{
		result->Output = ex->Message;
		result->Success = false;
	}
	return result;
}