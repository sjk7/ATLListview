// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com
// InterfaceCollection.cpp : Implementation of CInterfaceCollection
#include "stdafx.h"
#include "ATLListView.h"
#include "InterfaceCollection.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CInterfaceCollection

STDMETHODIMP CInterfaceCollection::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IInterfaceCollection};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}


