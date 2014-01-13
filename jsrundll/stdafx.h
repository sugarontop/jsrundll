#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN   

#include <windows.h>

#define DLLEXPORT     extern "C" __declspec( dllexport )

#define __Script_Debug__				// speed is 1/3

#pragma warning(disable: 4482 )
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS 

#include <atlbase.h>
#include <atlstr.h>
#include <atlcom.h>
#include <atlenc.h>

#include <ActivScp.h>
#include <string>
#include <map>
#include <vector>
#include <wincrypt.h>
#include <memory>

#ifdef __Script_Debug__
#include <ActivDbg.h>
#endif

