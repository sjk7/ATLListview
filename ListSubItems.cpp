// ListSubItems.cpp : Implementation of CListSubItems
#include "stdafx.h"
#include "ATLListView.h"
#include "ListSubItems.h"
#include "ListControl.h"
#include "ColumnHeader.h"
#include "ColumnHeaders.h"

/////////////////////////////////////////////////////////////////////////////
// CListSubItems

HRESULT CListSubItems::mySubItemInit(
    CListControl* pControl, CColumnHeaders* pColHeaders) {
    ASSERT(pControl);
    ASSERT(pColHeaders);
    m_pControl = pControl;
    m_pColumnHeaders = pColHeaders;
    return S_OK;
}

//#ifdef _DEBUG
#define FILL_SUBITEM_TEXT_AUTO
//#endif

HRESULT CListSubItems::myResize(
    CListItem* plitem, CListControl* pControl, CColumnHeaders* pColHeaders) {
    if (pControl) { // If you send one, please send both
        ASSERT(pColHeaders);
        mySubItemInit(pControl, pColHeaders);
    }
    this->m_pParentItem = plitem;
    int how_many_subitems = pColHeaders->m_cols.isize() - 1;
    if (how_many_subitems < 0) how_many_subitems = 0;
    ASSERT(m_pControl && m_pColumnHeaders); // who forgot to call myInit()?
    // this->m_subItems->m_coll.resize(how_many_subitems);

    HRESULT hr = S_OK;
    for (int i = 0; i < how_many_subitems; i++) {
        CComPtr<IDispatch> p;
        if (TRUE) { // if we already have one, leave it alone
            CComObject<CListSubItem>* ptr = 0;
            hr = CComObject<CListSubItem>::CreateInstance(&ptr);
            if (FAILED(hr)) return hr;
            hr = ptr->QueryInterface(IID_IDispatch, (void**)&p);
            if (FAILED(hr)) return hr;
            LPTSTR ptxt = 0;
#ifdef FILL_SUBITEM_TEXT_AUTO
            CString cs = my::to_string(i, _T("ListSubItem: "));
            ptxt = cs.GetBuffer(cs.GetLength());
#endif
            hr = ptr->SubItemInit(
                this, plitem, m_pControl, m_pColumnHeaders, i, ptxt);
            if (FAILED(hr)) return hr;

            ASSERT(m_subItems->m_coll[i]);
            CListSubItem* ptrback = (CListSubItem*)m_subItems->m_coll[i];
            p = 0;
            ASSERT(ptrback->m_dwRef == 1); // from QI above
            ASSERT(ptrback->m_info.apiIndex == i);
        }
    }
    return S_OK;
}

HRESULT CListSubItems::addSubItem(CListSubItem* p) {

    p->m_info.apiIndex = m_subItems->isize();
    if (!p->m_info.key.IsEmpty()) {
        std::pair<CString, CListSubItem*> pair_add;
        pair_add.first = p->m_info.key;
        pair_add.second = p;
        std::pair<subitem_map_t::iterator, bool> pair_ret
            = m_map.insert(pair_add);
        if (!pair_ret.second) {
            std::wstring wserr(L"Subitem, with key: ");
            CComBSTR bs(p->m_info.key);
            wserr += bs.m_str;
            wserr += L" is already in the subitem collection";
            return my::atlReportError(
                CLSID_ListSubItems, IID_IListSubItems, wserr, DISP_E_BADINDEX);
        }
    }
    m_subItems->addItem(p);

    return S_OK;
}

STDMETHODIMP CListSubItems::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IListSubItems};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) {
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP CListSubItems::get_ItemByKey(BSTR Key, IListItem** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    CString s(Key);
    my::com_map<CString, CListSubItem*>::map_it it = m_map.find(s);
    if (it == m_map.end()) {
        *pVal = 0;
        return DISP_E_BADINDEX;
    }

    HRESULT hr = it->second->QueryInterface(IID_IListSubItem, (void**)pVal);
    if (FAILED(hr)) return hr;
    (*pVal)->AddRef();

    return S_OK;
}

HRESULT CListSubItems::myAddRange(LONG, std::vector<CString>) {
    return E_NOTIMPL;
}

HRESULT CListSubItems::myAddListSubItem(LONG, const CString&, const CString&,
    const CString&, VARIANT*, IListSubItem**) {

    return E_NOTIMPL;
}

STDMETHODIMP CListSubItems::Add(VARIANT* Index, VARIANT* Key, VARIANT* Text,
    VARIANT* ReportIcon, VARIANT* ToolTipText, IListSubItem** out) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    int index = 0;
    HRESULT hr = my::VariantToInt(Index, index, true);
    if (FAILED(hr)) return hr;
    if (hr == S_FALSE) {
        index = (int)m_subItems->m_coll.size(); // append by default
    }
    CString text, key, toolTipText;
    hr = my::VariantToCString(Text, text, true);
    if (FAILED(hr)) return hr;
    hr = my::VariantToCString(Key, key, true);
    if (FAILED(hr)) return hr;

    hr = my::VariantToCString(ToolTipText, toolTipText, true);
    if (FAILED(hr)) return hr;

    return myAddListSubItem(index, text, key, toolTipText, ReportIcon, out);
}
