// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#include "stdafx.h"
#include "ATLListView.h"
#include "SelItemCollection.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CSelItemCollection

STDMETHODIMP CSelItemCollection::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_ISelItemCollection};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}
