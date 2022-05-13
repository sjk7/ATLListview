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
/////////////////////////////////////////////////////////////////////////////
// CListItem

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
        m_listItemInfo.m_pLv->selItemChange(this);
        m_listItemInfo.m_pLv->Fire_ItemSelectionChanged(
            this, my::BoolToVB(selected));
    }
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
void CListItem::setSelected(BOOL selected) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    m_listItemInfo.selected = selected;
    if (selected) {
        TRACE(_T("ListItem %d is SELECTED\n"), m_listItemInfo.apiIndex);

    } else {
        TRACE(_T("ListItem %d is *NOT* SELECTED\n"), m_listItemInfo.apiIndex);
    }

    if (selected) {
        m_listItemInfo.m_pLv->setLastSelItemIndex(m_listItemInfo.apiIndex);
    }
    notifySelChanged(selected);
    // this->m_listItemInfo.m_pLv->selItemChanged(this);
}
void CListItem::setListItemInfo(const ListItemInfo& info) noexcept {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    m_listItemInfo = info;
}
STDMETHODIMP CListItem::get_Index(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = static_cast<LONG>(m_listItemInfo.VBIndex());

    return S_OK;
}
