#pragma once

class CListItem;
#include "StdAfx.h"
#include "alphanum.hpp"
#include <shlwapi.h>

#if _MSC_VER <= VC6_MSC_VERSION
	LWSTDAPI_(int) StrCmpLogicalW(LPCWSTR psz1, LPCWSTR psz2);
#endif

#pragma comment (lib, "Shlwapi.lib") // StrCmpLogicalW

namespace my {
	struct CmpLessCiNatural {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			LPCTSTR a_beg = (LPCTSTR)lia->m_listItemInfo.text;
			LPCTSTR b_beg = (LPCTSTR)lib->m_listItemInfo.text;
#ifndef UNICODE
			return doj::alphanum_impl(a_beg, b_beg) < 0;
#else
			return StrCmpLogicalW(a_beg, b_beg) < 0;
#endif
			
		}
	};

	struct CmpGreaterCiNatural {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			LPCTSTR a_beg = (LPCTSTR)lia->m_listItemInfo.text;
			LPCTSTR b_beg = (LPCTSTR)lib->m_listItemInfo.text;
#ifndef UNICODE
			return doj::alphanum_impl(a_beg, b_beg) > 0;
#else
			return StrCmpLogicalW(a_beg, b_beg) > 0;
			//return TRUE;
#endif
		}
	};

	struct CmpLessNoCase {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			return lia->m_listItemInfo.text.CompareNoCase(lib->m_listItemInfo.text) < 0;
		}

	};
	struct CmpGreaterNoCase {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			return lia->m_listItemInfo.text.CompareNoCase(lib->m_listItemInfo.text) > 0;
		}
	};

	struct CmpLess {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			return lia->m_listItemInfo.text.Compare(lib->m_listItemInfo.text) < 0;
		}

	};

	struct CmpGreater {
		__inline bool operator()(const IDispatch* a, const IDispatch* b) const
		{
			CListItem* lia = (CListItem*)a;
			CListItem* lib = (CListItem*)b;
			return lia->m_listItemInfo.text.Compare(lib->m_listItemInfo.text) > 0;
		}

	};
}