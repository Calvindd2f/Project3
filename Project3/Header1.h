#pragma managed

#include "ntaqre3j.h"
#include <msclr/marshal_cppstd.h>
#include <vcclr.h>
#include <string>

using namespace System;
using namespace std;
using namespace System::Management::Automation;
using namespace System::Management::Automation::Runspaces;
using namespace msclr::interop;

extern "C"
{
	__declspec(dllexport) void __cdecl ntaqre3j();
}

typedef void (*LogCallbackFunction)(const char*, int);