// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com


#ifndef __LISTITEM_H_
#define __LISTITEM_H_

#include "resource.h" // main symbols

class CColumnHeaders;
class CListControl;

struct ListItemInfo {
	ListItemInfo(const CString& txt, int APIIndex,
		CColumnHeaders* pColumnHeaders, CListControl* listControl) noexcept
		: text(txt)
		, apiIndex(APIIndex)
		, pcolHeaders(pColumnHeaders)
		, m_pLv(listControl)
		, selected(FALSE)
	{}
	ListItemInfo() noexcept : apiIndex(-1), pcolHeaders(0), m_pLv(0) {}
	__inline int VBIndex() const noexcept { return apiIndex + 1; }
	mutable CString text;
	int apiIndex;
	CColumnHeaders* pcolHeaders;
	CListControl* m_pLv;
	BOOL selected;

	__inline LPTSTR toString() const noexcept {
		const int len = text.GetLength();
		return text.GetBuffer(len);
	}
};
/////////////////////////////////////////////////////////////////////////////
// CListItem
class ATL_NO_VTABLE CListItem
	: public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CListItem, &CLSID_ListItem>,
	public ISupportErrorInfo,
	public IDispatchImpl<IListItem, &IID_IListItem, &LIBID_ATLLISTVIEWLib> {
public:
	CListItem() {}
	virtual ~CListItem() {}

	DECLARE_REGISTRY_RESOURCEID(IDR_LISTITEM)

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	BEGIN_COM_MAP(CListItem)
		COM_INTERFACE_ENTRY(IListItem)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
	END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	// IListItem
public:
	STDMETHOD(get_Index)(/*[out, retval]*/ LONG* pVal);
	STDMETHOD(get_Text)(/*[out, retval]*/ BSTR* pVal);
	STDMETHOD(put_Text)(/*[in]*/ BSTR newVal);
	std::vector<CString> m_subitems;
	CString m_sKey;

	void notifySelChanged(BOOL selected);

	// int addItem(const CString& txt);
	ListItemInfo m_listItemInfo;
	void setSelected(BOOL selected);
	void setListItemInfo(const ListItemInfo& info) noexcept;

	int VBIndex() const noexcept { return m_listItemInfo.apiIndex + 1; }
	int apiIndex() const noexcept { return m_listItemInfo.apiIndex; }
};

static __inline void indexListItemCollection(std::vector<IDispatch*>& v) {
	
	typedef std::vector<IDispatch*>::iterator my_it;
	int i = 0;
	for (my_it it = v.begin(); it != v.end(); ++it) {
		//HRESULT hr = (*it)->QueryInterface(IID_IListItem, (void**)&p);
		//ASSERT(SUCCEEDED(hr));
		CListItem* ptr = (CListItem*)(*it);
		ptr->m_listItemInfo.apiIndex = i++;
	}
}

#endif //__LISTITEM_H_
