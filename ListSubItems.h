// ListSubItems.h : Declaration of the CListSubItems

#ifndef __LISTSUBITEMS_H_
#define __LISTSUBITEMS_H_

#include "resource.h" // main symbols
#include "my.h"
#include "InterfaceCollection.h"
#include "ListSubItem.h"
#include <map>

class CListControl;
class CColumnHeaders;
class CListSubItem;

typedef std::map<CString, CListSubItem*> subitem_map_t;
/////////////////////////////////////////////////////////////////////////////
// CListSubItems
class ATL_NO_VTABLE CListSubItems
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CListSubItems, &CLSID_ListSubItems>,
      public ISupportErrorInfo,
      public IDispatchImpl<IListSubItems, &IID_IListSubItems,
          &LIBID_ATLLISTVIEWLib> {
    public:
    CListSubItems() : m_pControl(0), m_pColumnHeaders(0), m_pParentItem(0) {
        CComObject<CInterfaceCollection>* p;
        HRESULT hr = CComObject<CInterfaceCollection>::CreateInstance(&p);

        ASSERT(SUCCEEDED(hr));
        hr = p->QueryInterface(&m_subItems_interface);
        ASSERT(SUCCEEDED(hr));
        m_subItems = p;
    }

    ContainerType& subItemCollection() const { return m_subItems->m_coll; }

    int isize() const { return static_cast<int>(m_subItems->m_coll.size()); }
    HRESULT CListSubItems::myResize(CListItem* pli, CListControl* pControl = 0,
        CColumnHeaders* pcolHeaders = 0);

    HRESULT addSubItem(CListSubItem* p);

    HRESULT mySubItemInit(CListControl* pControl, CColumnHeaders* pColHeaders);
    CColumnHeaders* m_pColumnHeaders;
    CListItem* m_pParentItem;

    private:
    // CInterfaceCollection has only a vector
    // so we do the map thing.
    subitem_map_t m_map;
    CComPtr<IInterfaceCollection> m_subItems_interface;
    CComObject<CInterfaceCollection>* m_subItems;

    public:
    CListControl* m_pControl;

    DECLARE_REGISTRY_RESOURCEID(IDR_LISTSUBITEMS)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CListSubItems)
    COM_INTERFACE_ENTRY(IListSubItems)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IListSubItems
    public:
    HRESULT myAddListSubItem(LONG index, const CString& text = _T(""),
        const CString& key = _T(""), const CString& toolTipText = _T(""),
        VARIANT* reportIcon = NULL, IListSubItem** out = NULL);
    HRESULT myAddRange(LONG index, std::vector<CString> texts);
    STDMETHOD(Add)
    (/*[in, optional]*/ VARIANT* Index, /*[in, optional]*/ VARIANT* Key,
        /*[in, optional]*/ VARIANT* Text,
        /*[in, optional]*/ VARIANT* ReportIcon,
        /*[in, optional]*/ VARIANT* ToolTipText,
        /*[out, retval]*/ IListSubItem** out);

    STDMETHOD(get_ItemByKey)
    (/*[in]*/ BSTR Key, /*[out, retval]*/ IListItem** pVal);
    STDMETHOD(Clear)() {
        m_subItems->clearvec(m_subItems->m_coll);
        m_map.clear();
        return S_OK;
    }
    STDMETHOD(get_Item)
    (/*[in]*/ LONG index, /*[out, retval]*/ IListSubItem** pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        ASSERT(pVal);
        if (pVal && *pVal) {
            (*pVal)->Release();
        }
        index--;
        CComPtr<IDispatch> pi = m_subItems->subitemAt(index);
        HRESULT hr = S_OK;
        if (pi) {
            hr = pi->QueryInterface(__uuidof(IListSubItem), (void**)pVal);
            ASSERT(SUCCEEDED(hr));
        } else {
            *pVal = 0;
            return S_OK; // its ok if we have null items int the subitems
                         // vector, just don't try to dereference them!
        }
        return hr;
    }

    STDMETHOD(get_Count)(/*[out, retval]*/ LONG* pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        return m_subItems->get_Count(pVal);
    }

    STDMETHOD(get__NewEnum)(/*[out, retval]*/ IUnknown** pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        return m_subItems->get__NewEnum(pVal);
    }
};

#endif //__LISTSUBITEMS_H_
