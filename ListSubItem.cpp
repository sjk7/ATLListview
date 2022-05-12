// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com
// ListSubItem.cpp : Implementation of CListSubItem
#include "stdafx.h"
#include "ATLListView.h"
#include "ListSubItem.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CListSubItem

STDMETHODIMP CListSubItem::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IListSubItem};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}
