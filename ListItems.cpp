// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListItems.cpp : Implementation of CListItems
#include "stdafx.h"
#include "ATLListView.h"
#include "ListItems.h"

#include "ListControl.h"
#include "my.h"

/////////////////////////////////////////////////////////////////////////////
// CListItems

static inline HRESULT getVariantIndex(
    VARIANT* Index, const listitem_collection& items, int& oneBasedIndex) {
    int myindex = 0;
    if (Index) {
        if (Index->vt == VT_I4 || Index->vt == VT_I2) {
            myindex = Index->lVal;
        } else {
            if (Index->vt == VT_ERROR) {
                // wasn't filled in by caller, ok
                myindex = items.isize() + 2; // one-based, assure append
            } else {
                return DISP_E_BADINDEX;
            }
        }
    }
    if (myindex <= 0) {
        return DISP_E_BADINDEX;
    }
    if (myindex > items.isize() + 2) myindex = items.isize() + 2;
    oneBasedIndex = myindex;
    return S_OK;
}

STDMETHODIMP CListItems::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IListItems};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP CListItems::Add(VARIANT* Index, VARIANT* Key, VARIANT* Text,
    VARIANT* Icon, VARIANT* SmallIcon, IListItem** out) {

    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    int myOneBasedIndex = -1;
    HRESULT hr = getVariantIndex(Index, m_items, myOneBasedIndex);
    if (FAILED(hr)) {
        return hr;
    }

    CString txt;
    hr = my::CStringFromVariant(txt, Text);
    if (FAILED(hr)) {
        return reportError(L"Bad type for listitem text. Should be a string",
            DISP_E_TYPEMISMATCH);
    }

    CString key;
    if (Key && Key->vt != VT_ERROR) {
        hr = my::CStringFromVariant(key, Key);
        if (FAILED(hr)) {
            return reportError(L"Bad type for listitem key. Should be a string",
                DISP_E_TYPEMISMATCH);
        }
    }
    int ret = 0;
    const int myindex = myOneBasedIndex - 1;
    const add_info info(1);
    ret = addItem(info, txt, myindex, key, out);

    if (ret < 0) {
        return DISP_E_EXCEPTION;
    }
    return S_OK;
}

STDMETHODIMP CListItems::get__NewEnum(IUnknown** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_col_interface->setVectorData(m_items.vec_data());
    const HRESULT ret = this->m_col_interface->get__NewEnum(pVal);

    return ret;
}

HRESULT CListItems::createListItem(
    CListControl* lv, const CString& text, IListItem** out) {

    CComObject<CListItem>* pli;
    HRESULT hr = m_factory.CreateInstance(&pli);
    if (FAILED(hr)) {
        ASSERT("CreateInstance for listview item failed" == 0);
        return reportError(L"CreateInstance for listview item failed");
    }

    m_items.add(pli); // AddRef() is here
    const size_t idx = m_items.size() - 1;
    ListItemInfo info(text, idx, lv->m_hdr, lv);
    pli->setListItemInfo(info);
    if (out) {
        hr = pli->QueryInterface(IID_IListItem, reinterpret_cast<void**>(out));
    }
    return hr;
}

HRESULT CListItems::resizeMe(
    const int curSize, const add_info& addInfo, const int newCount) {
    if (addInfo.n == 0) {
        static const size_t vec_contingency = 20000;
        if (m_items.icapacity() < newCount) {
            try {
                m_items.reserve(newCount + vec_contingency);
            } catch (...) {
                try {
                    m_items.reserve(newCount);
                } catch (...) {
                    return reportError(L"Unable to reserve memory");
                }
            }
        }

        if (m_lv->viewItemCountGet() < addInfo.howMany + curSize) {
            BOOL ok = m_lv->viewItemCountSet(newCount + vec_contingency);
            if (!ok) {
                ok = m_lv->viewItemCountSet(newCount);
            }
            if (!ok) {
                return reportError(L"Unable to set the view count");
            }
        }
    }
    return S_OK;
}

// NOTE: index here is ZERO-BASED. You must subtract one from index BEFORE
// CALLING ME if the caller is VB.
HRESULT CListItems::addItem(const add_info& addInfo, const CString& text,
    const int index, const CString& key, IListItem** out) {

    if (!key.IsEmpty()) {
        IDispatch* p = m_items.find(key);
        if (p) {
            return reportError(L"Key already in collection", DISP_E_BADINDEX);
        }
    }
    const HRESULT hr = createListItem(m_lv, text, out);
    if (FAILED(hr)) return hr;
    HWND hWnd = m_lv->lvhWnd();

    const int curSize = m_items.isize();
    const int newCount = curSize + addInfo.howMany;
    my::lv_insert_info info(index, m_lv->m_virtualMode, curSize);
    resizeMe(curSize, addInfo, newCount);
    info.viewItemCount = newCount;

    int ret = 0;
    if (!info.isVirtual) {
        ret = my::lvInsertItem(hWnd, text, info);
        ASSERT(ret >= 0);
    }

    if (m_items.size() == 1 && ret >= 0) {
        // first one: make sure there's an indication we tabbed to the control.
        my::lvSetSelectedItem(hWnd, 0, 1);
    }

    return hr;
}

HRESULT CListItems::addItems(SAFEARRAY* sa) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    BSTR* raw = 0;
    HRESULT hr = SafeArrayAccessData(sa, (void**)&raw);
    if (FAILED(hr)) return hr;

    long lowerBound, upperBound; // get array bounds
    hr = SafeArrayGetLBound(sa, 1, &lowerBound);
    if (FAILED(hr)) return hr;
    hr = SafeArrayGetUBound(sa, 1, &upperBound);
    if (FAILED(hr)) return hr;

    const long cnt_elements = upperBound - lowerBound + 1;

    VARIANT_BOOL was_redraw = VARIANT_TRUE;
    hr = m_lv->get_RedrawEnabled(&was_redraw);
    if (FAILED(hr)) return hr;

    m_lv->SetRedraw(VARIANT_FALSE);
    int idx = m_items.size();
    add_info ainfo;
    ainfo.howMany = cnt_elements;
    ainfo.n = 0;
    int& i = ainfo.n;
    while (ainfo.n < cnt_elements) {
        const wchar_t* wc = raw[i];
        CString cs(wc);
        // TRACE("%s\n", cs);
        hr = addItem(ainfo, cs, idx++, "", NULL);
        if (FAILED(hr)) return hr;
        ++i;
    }

    m_lv->SetRedraw(was_redraw);
    return SafeArrayUnaccessData(sa);
}

STDMETHODIMP CListItems::AddRangeEx(SAFEARRAY** arItemTexts) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    addItems(*arItemTexts);

    return S_OK;
}

STDMETHODIMP CListItems::get_Count(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = static_cast<LONG>(m_items.size());

    return S_OK;
}

STDMETHODIMP CListItems::get_Item(LONG index, IListItem** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    index -= 1;
    if (index < 0 || index >= static_cast<LONG>(m_items.size())) {
        return DISP_E_BADINDEX;
    }

    // *pVal = m_items[index];
    const HRESULT hr
        = m_items[index]->QueryInterface(IID_IListItem, (void**)pVal);
    return hr;
}

STDMETHODIMP CListItems::Clear() {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    m_items.clear();

    return S_OK;
}

BOOL CListItems::reserve(const size_t howManyExtra) {
    m_items.reserve(m_items.size() + howManyExtra);
    const BOOL ok = m_lv->viewItemCountSet(m_items.size() + howManyExtra);
    ASSERT(ok);
    return ok;
}

STDMETHODIMP CListItems::AddRange(
    LONG howMany, VARIANT* Index, IInterfaceCollection** out) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    HRESULT hr = S_OK;
    if (howMany <= 0) return hr;
    int myindex = 0;
    hr = getVariantIndex(Index, m_items, myindex);
    if (FAILED(hr)) return hr;
    myindex -= 1;

    if (out && *out) {
        (*out)->Release();
    }
    CComObject<CInterfaceCollection>* pcol;
    hr = CComObject<CInterfaceCollection>::CreateInstance(&pcol);
    if (FAILED(hr)) return E_OUTOFMEMORY;

    const BOOL ok = reserve(howMany);
    ASSERT(ok);
    pcol->reserve(howMany);
    myindex = this->m_items.isize(); // FIXME for true inserts

    if (!ok) {
        return E_OUTOFMEMORY;
    }
    CComPtr<IListItem> pli;
    this->m_lv->SetRedraw(VARIANT_FALSE);
    add_info info;
    info.howMany = howMany;
    int& i = info.n;
    info.n = 0;

    while (i < howMany) {
        addItem(info, "", myindex++, "", &pli);
        if (!pli) {
            this->m_lv->SetRedraw(VARIANT_TRUE);
            return E_OUTOFMEMORY;
        }
        pcol->addItem(pli);

        pli = 0;
        ++i;
    } // while i < howMany

    if (out) {
        hr = pcol->QueryInterface(
            IID_IInterfaceCollection, reinterpret_cast<void**>(out));
        if (FAILED(hr)) {
            return hr;
        }
    }
    this->m_lv->SetRedraw(VARIANT_TRUE);
    return S_OK;
}

STDMETHODIMP CListItems::Remove(LONG Index) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    const int apiIndex = Index - 1;
    if (apiIndex < 0) return DISP_E_BADINDEX;
    BOOL del = ListView_DeleteItem(m_lv->lvhWnd(), apiIndex);
    if (del == FALSE) {
        return reportError(L"Failed remove item from the grid");
    }
    IListItem* p = 0;
    HRESULT hr = m_items.at(apiIndex)->QueryInterface(IID_IListItem, (void**)p);
    if (FAILED(hr)) return reportError(L"Failed to get item to remove", hr);
    // m_plv->Fire_ColumnRemoved((ColumnHeader*)pcol);
    CListItem* pli = (CListItem*)p;
    if (!pli->m_sKey.IsEmpty()) {
        m_items.m_map.remove(pli->m_sKey);
    }
    m_items.delete_at(pli->apiIndex());
    p->Release();

    return S_OK;
}

__inline bool listItemTextComparer(const IDispatch* a, const IDispatch* b) {
    CListItem* lhs = (CListItem*)a;
    CListItem* rhs = (CListItem*)b;
    return lhs->m_listItemInfo.text < rhs->m_listItemInfo.text;
}
