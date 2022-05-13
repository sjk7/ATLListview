// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListItems.h : Declaration of the CListItems

#ifndef __LISTITEMS_H_
#define __LISTITEMS_H_

#include "resource.h" // main symbols
#include "my.h"
#include "InterfaceCollection.h"
#include "ListItem.h"
#include "my_cmp.h"
#include "SelItemCollection.h"


class CListControl;

struct SortInfo {
	
	SortInfo() : key(0), order((ListSortOrderFlags)(lvwNoCase | lvwNatural)), sorted(FALSE) {}
	// which column to sort by. zero indexed. zero means sort by the 'text' (this is the first text column)
	int key;
	ListSortOrderFlags order;
	BOOL sorted;
	inline bool operator == (const SortInfo& rhs) {
		if (rhs.key == key && rhs.order == order && sorted == rhs.sorted) return true;
		return false;
	}
};

typedef my::idispatch_collection listitem_collection;

struct add_info {
    add_info(const int how_many = -1) : howMany(how_many), n(0) {}
    int howMany;
    int n;
};

/////////////////////////////////////////////////////////////////////////////
// CListItems
class ATL_NO_VTABLE
    CListItems // NOLINT(cppcoreguidelines-special-member-functions)
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CListItems, &CLSID_ListItems>,
      public ISupportErrorInfo,
      public IDispatchImpl<IListItems, &IID_IListItems,
          &LIBID_ATLLISTVIEWLib> 
{
    public:

		void sortByText(const SortInfo& si) {

			my::Stopwatch sw("Sorting");
			if (si.order & lvwNatural) {
				if (si.order & lvwAscending) {
					std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpLessCiNatural());
				}
				else {
					std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpGreaterCiNatural());
				}
			}
			else {
				if (si.order & lvwNoCase) {
					if (si.order & lvwAscending) {
						std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpLessNoCase());
					}
					else {
						std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpGreaterNoCase());
					}

				}
				else {
					if (si.order == lvwAscending) {
						std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpLess());
					}
					else {
						std::stable_sort(m_items.getVector().begin(), m_items.getVector().end(), my::CmpGreater());
					}
				}
			}
			m_sortInfo = si;
			indexListItemCollection(m_items.getVector());
		}

		void sort(const SortInfo& si) {
			if (si.order == lvwNone || si.key < 0) return;
			if (m_sortInfo == si) return;
			if (si.key == 0) {
				sortByText(si);
			}
		}

    CListItems() : m_col_interface(0), m_lv(0) {
        m_items.reserve(50000);

        HRESULT hr = CComObject<CInterfaceCollection>::CreateInstance(
            &m_col_interface);
        m_col_interface->AddRef();
        ASSERT(SUCCEEDED(hr));
        m_col_interface->setVectorData(m_items.vec_data());
    }

    virtual ~CListItems() {
        if (m_col_interface) m_col_interface->Release();
    }

    HRESULT reportError(const std::wstring& what, HRESULT hr = E_FAIL) const {
        return my::atlReportError(CLSID_ListItems, IID_IListItems, what, hr);
    }

    CComObject<CInterfaceCollection>* m_col_interface;

    DECLARE_REGISTRY_RESOURCEID(IDR_LISTITEMS1)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CListItems)
    COM_INTERFACE_ENTRY(IListItems)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);
    HRESULT addItems(SAFEARRAY* sa);

    // IListItems
    public:

    STDMETHOD(AddRange)
    (/*[in]*/ LONG howMany, /*[in, optional]*/ VARIANT* Index,
        /*[out, retval]*/ IInterfaceCollection** out);
    STDMETHOD(AddRangeEx)(SAFEARRAY** arItemTexts);
    STDMETHOD(Clear)();
    STDMETHOD(get_Item)
    (/*[in]*/ LONG index, /*[out, retval]*/ IListItem** pVal);

    STDMETHOD(get_Count)(/*[out, retval]*/ LONG* pVal);
    listitem_collection m_items;
    CListControl* m_lv;
    STDMETHOD(Add)
    (/*[in, optional]*/ VARIANT* Index, /*[in, optional]*/ VARIANT* Key,
        /*[in, optional]*/ VARIANT* Text, /*[in, optional]*/ VARIANT* Icon,
        /*[in,optional]*/ VARIANT* SmallIcon,
        /*[out, retval]*/ IListItem** out);

    STDMETHOD(get__NewEnum)(/*[out, retval]*/ IUnknown** pVal) override;

    BOOL reserve(size_t howManyExtra);
	__inline void reIndexItems() {
		indexListItemCollection(m_items.getVector());
	}
	SortInfo m_sortInfo;
    private:
    HRESULT addItem(const add_info& info, const CString& text, int index = -1, const CString& key = "",
        IListItem** out = 0);
    static CComObject<CListItem> m_factory;
	// This is where the item is actually added to the collection
    HRESULT createListItem(
        CListControl* lv, const CString& text, IListItem** out = 0);
    HRESULT resizeMe(
        const int curSize, const add_info& addInfo, const int newCount);
public:
	STDMETHOD(Remove)(LONG Index);
	size_t size() const {
		return m_items.size();
	}
};



#endif //__LISTITEMS_H_
