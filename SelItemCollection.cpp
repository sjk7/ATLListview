// SelItemCollection.cpp : Implementation of CSelItemCollection
#include "stdafx.h"
#include "ATLListView.h"
#include "SelItemCollection.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CSelItemCollection

STDMETHODIMP CSelItemCollection::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ISelItemCollection
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (my::InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
