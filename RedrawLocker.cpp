// RedrawLocker.cpp : Implementation of CRedrawLocker
#include "stdafx.h"
#include "ATLListView.h"
#include "RedrawLocker.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CRedrawLocker

STDMETHODIMP CRedrawLocker::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IRedrawLocker
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (my::InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
