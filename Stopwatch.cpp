// Stopwatch.cpp : Implementation of CStopwatch
#include "stdafx.h"
#include "ATLListView.h"
#include "Stopwatch.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// CStopwatch

STDMETHODIMP CStopwatch::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IStopwatch
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (my::InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CStopwatch::Start(BSTR info)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	
		m_sInfo = info;
	m_start = timeGetTime();

	return S_OK;
}

STDMETHODIMP CStopwatch::get_Disarmed(VARIANT_BOOL *pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	*pVal = m_VBDisarmed;

	return S_OK;
}

STDMETHODIMP CStopwatch::put_Disarmed(VARIANT_BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	m_VBDisarmed = newVal;

	return S_OK;
}

STDMETHODIMP CStopwatch::Stop(VARIANT_BOOL PrintInfo)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

		if (m_VBDisarmed){
			m_stopped = timeGetTime();
			return S_OK;
		}

		// output to VB's immediate window?



	return S_OK;
}
