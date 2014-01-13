// stdafx.h : 標準のシステム インクルード ファイルのインクルード ファイル、または
// 参照回数が多く、かつあまり変更されない、プロジェクト専用のインクルード ファイル
// を記述します。
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Windows ヘッダーから使用されていない部分を除外します。
// Windows ヘッダー ファイル:
#include <windows.h>

#define DLLEXPORT     extern "C" __declspec( dllexport )

#define __Script_Debug__				// speedが1/3になる

#pragma warning(disable: 4482 )
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // 一部の CString コンストラクターは明示的です。

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

