// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __SELITEMCOLLECTION_H_
#define __SELITEMCOLLECTION_H_

#include "resource.h"       // main symbols
#include "my.h"
#include "InterfaceCollection.h"

// InterfaceCollection only implements vector,
// and we want to be sure selected items are unique and quickly removed
typedef std::map<unsigned long, CListItem*> map_type;
/////////////////////////////////////////////////////////////////////////////
// CSelItemCollection
class ATL_NO_VTABLE CSelItemCollection : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSelItemCollection, &CLSID_SelItemCollection>,
	public ISupportErrorInfo,
	public IDispatchImpl<ISelItemCollection, &IID_ISelItemCollection, &LIBID_ATLLISTVIEWLib>
{
public:
	CSelItemCollection()
	{
	}
	
	CComObject<CInterfaceCollection>* m_col_interface;
	
DECLARE_REGISTRY_RESOURCEID(IDR_SELITEMCOLLECTION)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CSelItemCollection)
	COM_INTERFACE_ENTRY(ISelItemCollection)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ISelItemCollection
public:
	STDMETHOD(get__NewEnum)(/*[out, retval]*/ IUnknown** pVal){
		HRESULT hr = m_col_interface->get__NewEnum(pVal);
		return hr;
	}

	 STDMETHOD(get_Item)
		 (/*[in]*/ LONG index, /*[out, retval]*/ IListItem** pVal){
		 if (index < 0 || index > (int)m_col_interface->m_coll.size())
			 return DISP_E_BADINDEX;

		 HRESULT hr = m_col_interface->QueryInterface(IID_IInterfaceCollection, (void**)pVal);
		 return hr;

	 }

	 STDMETHOD(get_Count)(/*[out, retval]*/ LONG* pVal){
		*pVal = m_col_interface->m_coll.size();
		return S_OK;
	 }

};

#endif //__SELITEMCOLLECTION_H_
