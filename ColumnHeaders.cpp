// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ColumnHeaders.cpp : Implementation of CColumnHeaders
#include "stdafx.h"
#include "ATLListView.h"
#include "ColumnHeaders.h"
#include "ListControl.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CColumnHeaders

CColumnHeaders::CColumnHeaders()
    : m_plv(0)
    , m_col_interface(0)
    , m_scaleUnits(my::win32::pixelUnits)
    , m_hWnd(0)
    , m_heightInPixels(-1)
    , m_hWndLv(0) {

    HRESULT hr
        = CComObject<CInterfaceCollection>::CreateInstance(&m_col_interface);
    m_col_interface->AddRef();
    ASSERT(SUCCEEDED(hr));
    m_col_interface->setVectorData(m_cols.vec_data());
}
STDMETHODIMP CColumnHeaders::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IColumnHeaders};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP CColumnHeaders::get_Count(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = (LONG)m_cols.size();

    return S_OK;
}

// HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, BOOL& bHandled
BOOL CColumnHeaders::apiSetHeight(
    HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {

    (void)bHandled;
    (void)wParam;
    LPHDLAYOUT pHL = reinterpret_cast<LPHDLAYOUT>(lParam);
    RECT* pRect = pHL->prc;
    WINDOWPOS* pWPos = pHL->pwpos;
    BOOL r = m_plv->hdrCallDefWndProc(hWnd, uMsg, 0, lParam);
    if (r != FALSE) {
        int h = pWPos->cy;
        if (m_heightInPixels > 0) {
            h = m_heightInPixels;
        }
        pWPos->cy = h;
        pRect->top = h;
    }
    return r;
}

STDMETHODIMP CColumnHeaders::get_Item(LONG index, IColumnHeader** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    HRESULT hr = S_OK;
    const long sz = (long)m_cols.size();
    index -= 1;

    if (index < 0 || index >= sz) {
        hr = DISP_E_BADINDEX;
    } else {
        hr = m_cols.at(index)->QueryInterface(IID_IDispatch, (void**)pVal);
    }

    return hr;
}

int CColumnHeaders::addColumn(const CString& txt, int width, VARIANT* key) {
    CComObject<ColumnHeader> ch;
    CComObject<ColumnHeader>* pch;
    HRESULT hr = ch.CreateInstance(&pch);
    if (FAILED(hr)) return -1;
    m_cols.push_back(pch); // AddRef() is here
    size_t idx = m_cols.size() - 1;
    ColumnHeader* ptr = (ColumnHeader*)m_cols[idx];
    ptr->setKey(key);
    ptr->setName(txt);
    ptr->setVBIndex(idx + 1);
    ASSERT(m_plv && m_plv->lvhWnd());
    ptr->setup(
        m_plv, m_plv->lvhWnd(), this, ListView_GetHeader(m_plv->lvhWnd()));
    ptr->setListview(m_plv);
    if (width >= 0) ptr->m_width = width;

    ptr->applyWidth();
    ptr->m_VBsubItemIndex = (short)idx; // the first column is NOT a subitem, so
                                        // gets 0 here (rem: VBIndex,right?)

    return my::lvInsertTextColumn(m_plv->lvhWnd(), txt);
}

STDMETHODIMP CColumnHeaders::Add(VARIANT* Index, VARIANT* Key, VARIANT* Text,
    VARIANT* Width, VARIANT* Alignment, VARIANT* Icon, IColumnHeader** retVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    //(void)Icon;
    //(void)Alignment;
    //(void)Width;

	this->Fire_ColumnClick(NULL);
    int vbindex = 0;
    HRESULT hr = my::VariantToInt(Index, vbindex, true);
    if (FAILED(hr)) return hr;
    int index = vbindex - 1;
    if (index < 0 || index >= m_cols.isize()) {
        index = m_cols.isize(); // append
    }
    int width = 100;

    hr = my::VariantToInt(Width, width, true);
    if (FAILED(hr)) return hr;

    CString txt;
    hr = my::VariantToCString(Text, txt, true);
    if (FAILED(hr)) {
        return reportError(L"Bad type in text argument to ColumnHeaders::Add()",
            DISP_E_TYPEMISMATCH);
    }

    int added = addColumn(txt, width, Key);
    ASSERT(added == index);
    if (static_cast<int>(m_cols.size()) != index + 1) {
        return reportError(
            L"Unable to add column: probably out of memory", E_OUTOFMEMORY);
    }

    hr = m_cols[index]->QueryInterface(IID_IColumnHeader, (void**)retVal);
    return hr;
}

STDMETHODIMP CColumnHeaders::get__NewEnum(IUnknown** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_col_interface->setVectorData(m_cols.vec_data());
    HRESULT hr = m_col_interface->get__NewEnum(pVal);

    return hr;
}

STDMETHODIMP CColumnHeaders::get_HeightInPixels(long* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    *pVal = heightInPixels();

    return S_OK;
}

STDMETHODIMP CColumnHeaders::put_HeightInPixels(long newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    return this->heightInPixelsSet(newVal);
}

STDMETHODIMP CColumnHeaders::Clear() {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    const int cnt = my::lvGetColumnCount(m_hWndLv);
    int mycnt = cnt - 1;
    while (mycnt >= 0) {
        BOOL del = ListView_DeleteColumn(m_hWndLv, mycnt);
        ASSERT(del == TRUE);
        IDispatch* p = m_cols.at(mycnt);
        IColumnHeader* pcol = 0;

        HRESULT hr = p->QueryInterface(IID_IColumnHeader, (void**)&pcol);
        if (SUCCEEDED(hr)) m_plv->Fire_ColumnRemoved((ColumnHeader*)pcol);
        if (pcol) pcol->Release();

        mycnt--;
    }

    m_cols.clear();

    return S_OK;
}

STDMETHODIMP CColumnHeaders::Remove(LONG indexToRemove) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    const int apiIndex = indexToRemove - 1;
    if (apiIndex < 0) return DISP_E_BADINDEX;
    int del = ListView_DeleteColumn(m_hWndLv, apiIndex);
    ASSERT(del);
    CComPtr<IColumnHeader> pcol = 0;
    HRESULT hr
        = m_cols.at(apiIndex)->QueryInterface(IID_IColumnHeader, (void**)&pcol);
    if (FAILED(hr)) return hr;
    m_plv->Fire_ColumnRemoved((ColumnHeader*)pcol.p);
    m_cols.delete_at(apiIndex);
    return S_OK;
}
