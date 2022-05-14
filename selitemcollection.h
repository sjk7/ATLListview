// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __SELITEMCOLLECTION_H_
#define __SELITEMCOLLECTION_H_

#include "resource.h" // main symbols
#include "my.h"

// my_map_collection.hpp
#include "VCUE_Collection.h"
#include "VCUE_Copy.h"
#include "ListItem.h"
#include "Stopwatch.h"

typedef std::vector<IDispatch*> myContainerType;

typedef IDispatch* myCollectionExposedType;
typedef ISelItemCollection myCollectionInterface;

// Use IEnumVARIANT as the enumerator for VB compatibility
typedef VARIANT myEnumeratorExposedType;
typedef IEnumVARIANT myEnumeratorInterface;

typedef VCUE::GenericCopy<myEnumeratorExposedType, myContainerType::value_type>
    myEnumeratorCopyType;
typedef VCUE::GenericCopy<myCollectionExposedType, myContainerType::value_type>
    myCollectionCopyType;

typedef CComEnumOnSTL<myEnumeratorInterface, &__uuidof(myEnumeratorInterface),
    myEnumeratorExposedType, myEnumeratorCopyType, myContainerType>
    myEnumeratorType;
typedef ICollectionOnSTLImpl<myCollectionInterface, myContainerType,
    myCollectionExposedType, myCollectionCopyType, myEnumeratorType>
    myCollectionType;

typedef VCUE::GenericCopy<myEnumeratorExposedType, myContainerType::value_type>
    myEnumeratorCopyType;
typedef VCUE::GenericCopy<myCollectionExposedType, myContainerType::value_type>
    myCollectionCopyType;

// InterfaceCollection only implements vector,
// and we want to be sure selected items are unique and quickly removed
typedef std::map<unsigned long, CListItem*> map_type;
/////////////////////////////////////////////////////////////////////////////
// CSelItemCollection
class ATL_NO_VTABLE CSelItemCollection
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CSelItemCollection, &CLSID_SelItemCollection>,
      public ISupportErrorInfo,
      public IDispatchImpl<myCollectionType, &IID_ISelItemCollection,
          &LIBID_ATLLISTVIEWLib>
// public IDispatchImpl<CollectionType, &IID_IInterfaceCollection>
{
    public:
    CSelItemCollection() {}
    virtual ~CSelItemCollection() {}
    std::map<int, IListItem*> m_map;

    DECLARE_REGISTRY_RESOURCEID(IDR_SELITEMCOLLECTION)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CSelItemCollection)
    COM_INTERFACE_ENTRY(ISelItemCollection)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);
    void setData(const std::vector<IDispatch*>& vecItems) {

        my::Stopwatch sw("setData for populating selected items");
        m_map.clear();
        m_coll.clear();
        m_coll.reserve(vecItems.size());
        typedef std::vector<IDispatch*>::const_iterator ci_t;
        typedef std::map<int, IListItem*>::iterator m_it;

        for (ci_t it = vecItems.begin(); it < vecItems.end(); ++it) {
            CListItem* pi = (CListItem*)*it;

            if (pi->m_listItemInfo.selected) {
                std::pair<int, CListItem*> pr
                    = std::make_pair(pi->m_listItemInfo.apiIndex, pi);

                std::pair<m_it, bool> insert_pair = m_map.insert(pr);
                if (insert_pair.second) {
                    m_coll.push_back(*it);
                }
            }
        }
    }
    // ISelItemCollection
    public:
    STDMETHOD(get__NewEnum)(IUnknown** pVal) {
        HRESULT hr = myCollectionType::get__NewEnum(pVal);
        return hr;
    }

    STDMETHOD(get_Item)
    (LONG index, IListItem** pVal) {

        index -= 1;
        if (index < 0 || index >= long(m_coll.size())) return DISP_E_BADINDEX;
        HRESULT hr = m_coll[index]->QueryInterface(IID_IListItem, (void**)pVal);
        return hr;
    }

    STDMETHOD(get_Count)(LONG* pVal) {
        *pVal = this->m_coll.size();
        return S_OK;
    }
};

#endif //__SELITEMCOLLECTION_H_
