// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#if !defined(AFX_STDAFX_H__300839D4_82CE_4A7D_B6AE_56E24A376F33__INCLUDED_)
#define AFX_STDAFX_H__300839D4_82CE_4A7D_B6AE_56E24A376F33__INCLUDED_

#pragma warning(disable : 4786) // VC6 complaining about truncated debug names.

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#if _MSC_VER < 1900
#define noexcept //-V1059
#define override
#endif

#ifndef _WIN32_WINNT_WIN7
#define _WIN32_WINNT_WIN7                   0x0601 
#endif

#ifndef VC6_MSC_VERSION
#define VC6_MSC_VERSION 1200
#endif

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#endif
#define _ATL_APARTMENT_THREADED

#include <string>
#include <vector>
#include <map>
#include <afxwin.h>
#include <afxdisp.h>
#include <atlbase.h>
// You may derive a class from CComModule and use it if you want to override
// something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atlctl.h>
#include <ocidl.h> // Added by ClassView
#include "Stopwatch.h"

#ifndef TRACE
#define TRACE ATLTRACE
#endif

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before
// the previous line.

#endif // !defined(AFX_STDAFX_H__300839D4_82CE_4A7D_B6AE_56E24A376F33__INCLUDED)
