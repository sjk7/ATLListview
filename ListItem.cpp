// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListItem.cpp : Implementation of CListItem
#include "stdafx.h"
#include "ATLListView.h"
#include "ListItem.h"
#include "my.h"
#include "ListControl.h"
#include "ListSubItems.h"
/////////////////////////////////////////////////////////////////////////////
// CListItem

CListItem::CListItem() {

    CComObject<CListSubItems>* p;
    HRESULT hr = CComObject<CListSubItems>::CreateInstance(&p);
    ASSERT(SUCCEEDED(hr));
    hr = p->QueryInterface(&m_IsubItems);

    ASSERT(SUCCEEDED(hr));
    m_subItems = p;
}

STDMETHODIMP CListItem::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IListItem};

    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

void CListItem::notifySelChanged(BOOL selected) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    ASSERT(m_listItemInfo.m_pLv);
    if (this->m_listItemInfo.m_pLv) {
        m_listItemInfo.m_pLv->Fire_ItemSelectionChanged(
            this, my::BoolToVB(selected));
    }
}

HRESULT CListItem::resizeSubItems(
    CListItem* pli, CListControl* pControl, CColumnHeaders* pcolHeaders) {

    HRESULT hr = m_subItems->myResize(pli, pControl, pcolHeaders);
    return hr;
}

LPTSTR CListItem::getSubItemText(int isubItem) const {
    CListSubItems* plsis = m_subItems;
    ASSERT(plsis->m_pColumnHeaders
        && plsis->m_pControl); // init() was not called on the subitems.
    const std::vector<IDispatch*>& v = plsis->subItemCollection();
    ASSERT(isubItem >= 0 && isubItem < (int)v.size());

    CListSubItem* plsi = (CListSubItem*)v[isubItem];
    ASSERT(plsi);
    if (!plsi) return NULL;

    CString& text = plsi->m_info.text;
    const int len = text.GetLength();
    return text.GetBuffer(len);
}

STDMETHODIMP CListItem::get_Text(BSTR* pVal) {

    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    *pVal = m_listItemInfo.text.AllocSysString();
    return S_OK;
}

STDMETHODIMP CListItem::put_Text(BSTR newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    this->m_listItemInfo.text = newVal;

    return S_OK;
}
STDMETHODIMP_(HRESULT __stdcall)
CListItem::get_ListSubItems(IListSubItems** ppVal) {
    HRESULT hr
        = this->m_subItems->QueryInterface(IID_IListSubItems, (void**)ppVal);
    return hr;
}
void CListItem::setSelected(BOOL selected) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    m_listItemInfo.selected = selected;
    /*/
    if (selected) {
        TRACE(_T("ListItem %d is SELECTED\n"), m_listItemInfo.apiIndex);

    } else {
        TRACE(_T("ListItem %d is *NOT* SELECTED\n"), m_listItemInfo.apiIndex);
    }
    /*/
    if (selected) {
        m_listItemInfo.m_pLv->setLastSelItemIndex(m_listItemInfo.apiIndex);
    }
    notifySelChanged(selected);
    // this->m_listItemInfo.m_pLv->selItemChanged(this);
}
HRESULT CListItem::setListItemInfo(const ListItemInfo& info) noexcept {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    m_listItemInfo = info;
    ASSERT(m_subItems);
    ASSERT(info.m_pLv && info.pcolHeaders);
    HRESULT hr = m_subItems->myResize(this, info.m_pLv, info.pcolHeaders);
    return hr;
}
STDMETHODIMP CListItem::get_Index(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = static_cast<LONG>(m_listItemInfo.VBIndex());

    return S_OK;
}
