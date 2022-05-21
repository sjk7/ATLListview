// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListSubItem.h : Declaration of the CListSubItem

#ifndef __LISTSUBITEM_H_
#define __LISTSUBITEM_H_

//#include "resource.h" // main symbols
//#include "ATLListViewCP.h"
#include "ListSubItems.h"
class CListControl;
class CColumnHeaders;

struct SubItemInfo_t {
    SubItemInfo_t(int api_idx, CListControl* plc, CColumnHeaders* pch,
        CListSubItems* plis, CListItem* pli, LPTSTR txt = 0, LPTSTR Key = 0)
        : apiIndex(api_idx)
        , pControl(plc)
        , pColHeaders(pch)
        , pListSubItems(plis)
        , pListItem(pli)
        , text(txt)
        , key(Key) {}

    SubItemInfo_t()
        : apiIndex(-1)
        , pControl(0)
        , pColHeaders(0)
        , pListSubItems(0)
        , pListItem(0) {}
    int apiIndex;
    CListControl* pControl;
    CColumnHeaders* pColHeaders;
    CListSubItems* pListSubItems;
    CListItem* pListItem;
    CString text;
    CString key;
};
/////////////////////////////////////////////////////////////////////////////
// CListSubItem
class ATL_NO_VTABLE CListSubItem
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CListSubItem, &CLSID_ListSubItem>,
      public ISupportErrorInfo,
      public IDispatchImpl<IListSubItem, &IID_IListSubItem,
          &LIBID_ATLLISTVIEWLib>,
      public IConnectionPointContainerImpl<CListSubItem> {
    public:
    CListSubItem(){};
    SubItemInfo_t m_info;

    HRESULT SubItemInit(CListSubItems* psubItems, CListItem* pItem,
        CListControl* pControl, CColumnHeaders* pColumnHeaders, int apiIndex,
        LPTSTR text = 0, LPTSTR key = 0) {
        SubItemInfo_t inf(
            apiIndex, pControl, pColumnHeaders, psubItems, pItem, text, key);
        m_info = inf;
        return mySubItemInit(m_info);
    }

    private:
    HRESULT mySubItemInit(const SubItemInfo_t& inf);

    public:
    DECLARE_REGISTRY_RESOURCEID(IDR_LISTSUBITEM)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CListSubItem)
    COM_INTERFACE_ENTRY(IListSubItem)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IListSubItem
    public:
    STDMETHOD(get_Index)(/*[out, retval]*/ LONG* pVal) {
        *pVal = m_info.apiIndex + 1;
        return S_OK;
    }
    STDMETHOD(get_Text)(/*[out, retval]*/ BSTR* pVal) {
        *pVal = m_info.text.AllocSysString();
        return S_OK;
    }
    STDMETHOD(put_Text)(/*[in]*/ BSTR newVal) {
        m_info.text = newVal;
        return S_OK;
    }

    public:
    BEGIN_CONNECTION_POINT_MAP(CListSubItem)
    // CONNECTION_POINT_ENTRY(DIID__IListControlEvents)
    END_CONNECTION_POINT_MAP()
};

#endif //__LISTSUBITEM_H_
