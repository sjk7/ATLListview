// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListSubItem.h : Declaration of the CListSubItem

#ifndef __LISTSUBITEM_H_
#define __LISTSUBITEM_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CListSubItem
class ATL_NO_VTABLE CListSubItem : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CListSubItem, &CLSID_ListSubItem>,
	public ISupportErrorInfo,
	public IDispatchImpl<IListSubItem, &IID_IListSubItem, &LIBID_ATLLISTVIEWLib>
{
public:
	CListSubItem()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_LISTSUBITEM)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CListSubItem)
	COM_INTERFACE_ENTRY(IListSubItem)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IListSubItem
public:
};

#endif //__LISTSUBITEM_H_
