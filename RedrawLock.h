// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __REDRAWLOCK_H_
#define __REDRAWLOCK_H_

#include "resource.h" // main symbols
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CRedrawLock
class ATL_NO_VTABLE CRedrawLock
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CRedrawLock, &CLSID_RedrawLock>,
      public ISupportErrorInfo,
      public IDispatchImpl<IRedrawLock, &IID_IRedrawLock,
          &LIBID_ATLLISTVIEWLib> {
    public:
    CRedrawLock() : m_prevRedraw(TRUE) {}

    BOOL m_prevRedraw;

    virtual ~CRedrawLock() {
        VARIANT_BOOL vb = my::BoolToVB(m_prevRedraw);
        if (m_pLvw) m_pLvw->SetRedraw(vb);
    }

    CComPtr<IListControl> m_pLvw;

    DECLARE_REGISTRY_RESOURCEID(IDR_REDRAWLOCK)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CRedrawLock)
    COM_INTERFACE_ENTRY(IRedrawLock)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IRedrawLock
    public:
    STDMETHOD(get_ListControlObject)(/*[out, retval]*/ IListControl** pVal);
    STDMETHOD(putref_ListControlObject)(/*[in]*/ IListControl* newVal);
};

#endif //__REDRAWLOCK_H_
