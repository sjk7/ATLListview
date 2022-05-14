// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __INTERFACECOLLECTION_H_
#define __INTERFACECOLLECTION_H_

#include <mmsystem.h>
#pragma comment(lib, "winmm") // timeGetTime
#include "resource.h" // main symbols
#include "VCUE_Copy.h"

typedef std::vector<IDispatch*> ContainerType;

typedef IDispatch* CollectionExposedType;
typedef IInterfaceCollection CollectionInterface;

// Use IEnumVARIANT as the enumerator for VB compatibility
typedef VARIANT EnumeratorExposedType;
typedef IEnumVARIANT EnumeratorInterface;

typedef VCUE::GenericCopy<EnumeratorExposedType, ContainerType::value_type>
    EnumeratorCopyType;
typedef VCUE::GenericCopy<CollectionExposedType, ContainerType::value_type>
    CollectionCopyType;

typedef CComEnumOnSTL<EnumeratorInterface, &__uuidof(EnumeratorInterface),
    EnumeratorExposedType, EnumeratorCopyType, ContainerType>
    EnumeratorType;
typedef ICollectionOnSTLImpl<CollectionInterface, ContainerType,
    CollectionExposedType, CollectionCopyType, EnumeratorType>
    CollectionType;

typedef VCUE::GenericCopy<EnumeratorExposedType, ContainerType::value_type>
    EnumeratorCopyType;
typedef VCUE::GenericCopy<CollectionExposedType, ContainerType::value_type>
    CollectionCopyType;

/////////////////////////////////////////////////////////////////////////////
// CInterfaceCollection
class ATL_NO_VTABLE CInterfaceCollection
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CInterfaceCollection, &CLSID_InterfaceCollection>,
      public ISupportErrorInfo,
      public IDispatchImpl<CollectionType, &IID_IInterfaceCollection> {
    public:
    CInterfaceCollection() {}

    virtual ~CInterfaceCollection() { clearvec(m_coll); }

    void clearvec(ContainerType& v) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        typedef ContainerType::iterator it_t;
        for (it_t it = v.begin(); it < v.end(); ++it) {
            (*it)->Release();
        }
    }

    void setVectorData(const ContainerType& v) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        if (!m_coll.empty()) {
            clearvec(m_coll);
        }
        m_coll = ContainerType(v);
        typedef ContainerType::iterator it_t;
        for (it_t it = m_coll.begin(); it < m_coll.end(); ++it) {
            (*it)->AddRef();
        }
    }

    int addItem(IDispatch* p) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_coll.push_back(p);
        p->AddRef();
        return static_cast<int>(m_coll.size()) - 1;
    }

    void reserve(size_t how_many) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_coll.reserve(how_many);
    }

    DECLARE_REGISTRY_RESOURCEID(IDR_INTERFACECOLLECTION)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CInterfaceCollection)
    COM_INTERFACE_ENTRY(IInterfaceCollection)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    // COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IInterfaceCollection
    public:
    // STDMETHOD(get__NewEnum)(/*[out, retval]*/ IUnknown** pVal);
    // STDMETHOD(get_Count)(LONG* pRetval);
    // STDMETHOD(get_Item)(LONG Index, IUnknown** pVal);
    public:
    BEGIN_CONNECTION_POINT_MAP(CInterfaceCollection)
    // CONNECTION_POINT_ENTRY(DIID__IListControlEvents)
    END_CONNECTION_POINT_MAP()
};

#endif //__INTERFACECOLLECTION_H_
