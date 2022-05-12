// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// RedrawLock.cpp : Implementation of CRedrawLock
#include "stdAfx.h"
#include "ATLListView.h"
#include "RedrawLock.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CRedrawLock

STDMETHODIMP CRedrawLock::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IRedrawLock};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP CRedrawLock::get_ListControlObject(IListControl** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    if (m_pLvw) {
        *pVal = m_pLvw;
        (*pVal)->AddRef();
    }

    return S_OK;
}

STDMETHODIMP CRedrawLock::putref_ListControlObject(IListControl* newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    if (m_pLvw) {
        m_pLvw->SetRedraw(my::BoolToVB(m_prevRedraw));
        m_prevRedraw = TRUE;
        m_pLvw = 0;
    }

    if (newVal) {
        VARIANT_BOOL vb = VARIANT_TRUE;
        CComPtr<IListControl> tmp;
        HRESULT hr = newVal->QueryInterface(&tmp);
        ASSERT(SUCCEEDED(hr));
        hr = tmp->get_RedrawEnabled(&vb);
        ASSERT(SUCCEEDED(hr));
        m_prevRedraw = (vb == VARIANT_TRUE ? TRUE : FALSE);
        m_pLvw = tmp; //-V820
    } else {
        m_pLvw = 0;
    }

    if (m_pLvw) {
        m_pLvw->SetRedraw(VARIANT_FALSE);
    }
    return S_OK;
}
