// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

// ListControl.cpp : Implementation of CListControl

#include "stdafx.h"
#include "ATLListView.h"
#include "ListControl.h"
#define _CRTDBG_MAP_ALLOC
#include <stdlib.h>
#include <crtdbg.h>
#include <comdef.h> // _bstr_t

// Added fake code begins here
#if 0
class CAppModule : 
    public CComModule
{
};
	CAppModule _Module;

// Added fake code ends here, below is regular ATL project stuff
#endif



/////////////////////////////////////////////////////////////////////////////
// CListControl
CListControl::CListControl()
    : m_lvw(WC_LISTVIEW, this, 1)
    , m_litems(0)
    , m_selItems(0)
    , m_hdr(0)
    , m_clrBackColor(0x80000005)
    , m_clrBorderColor(0)
    , m_bEnabled(1)
    , m_clrForeColor(0)
    , m_nMousePointer(0)
    , m_bTabStop(0)
    , m_nBorderStyle(ccNone)
    , m_nAppearance(cc3D)
    , m_sid("Listview Control")
    , m_bMouseActivate(FALSE)
    , m_bRedrawEnabled(TRUE)
    , m_virtualMode(FALSE)
    , m_hWndVBHostSite(0)
    , m_viewItemCount(0)
    , m_phdrSubclass(0)
    , m_hdrWnd(0)
    , m_scaleUnitsEnum(my::win32::pixelUnits)
    , m_bMultiSelect(FALSE)
    , m_lastSelItemIndex(-1)

{
    // ::_CrtSetBreakAlloc(413);
    static bool have_init_cc = false;
    if (!have_init_cc) {
        have_init_cc = true;
        InitCommonControls();
        ::timeBeginPeriod(1); // make timings more reliable
    }

    FONTDESC fd = {sizeof(FONTDESC), OLESTR("Arial"), FONTSIZE(9), FW_NORMAL,
        ANSI_CHARSET, TRUE, FALSE, FALSE};
    HRESULT hr = OleCreateFontIndirect(&fd, IID_IFont, (void**)&m_pFont);

    ASSERT(SUCCEEDED(hr));

    m_bWindowOnly = TRUE;
    CComObject<CColumnHeaders>::CreateInstance(&m_hdr);
    ASSERT(m_hdr);
    if (m_hdr) {
        m_hdr->AddRef();
        m_hdr->m_plv = this; // NOT addrefed to avoid circularity#
        m_hdr->m_scaleUnits = m_scaleUnitsEnum;
    }
}

// called when user changes property in HOST (not us)
STDMETHODIMP CListControl::OnAmbientPropertyChange(DISPID dispid) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    HRESULT hr = E_FAIL;
    if (dispid == DISPID_AMBIENT_UIDEAD) return S_OK;

    if (dispid == DISPID_AMBIENT_SCALEUNITS) {
        hr = setScaleUnits();
    }

    if (dispid == DISPID_AMBIENT_FONT) {
        IFont* pfont = 0;
        hr = GetAmbientFont(&pfont);
        if (SUCCEEDED(hr)) {
            m_pFont = 0;
            hr = pfont->QueryInterface(IID_IFontDisp, (void**)&m_pFont);
            ASSERT(SUCCEEDED(hr));
        }
    } else if (dispid == DISPID_AMBIENT_BACKCOLOR) {
        if (!InDesignMode()) {
            hr = GetAmbientBackColor(m_clrBackColor);
            ASSERT(hr == S_OK);
        }
        hr = S_OK;
    } else if (dispid == DISPID_AMBIENT_FORECOLOR) {
        if (!InDesignMode()) {
            hr = GetAmbientForeColor(m_clrForeColor);
            ASSERT(hr == S_OK);
        }
        hr = S_OK;
    } else if (dispid == DISPID_AMBIENT_DISPLAYNAME) {
        hr = S_OK;
    }

    ASSERT(SUCCEEDED(hr)); // did you handle it?
    applyProperties();
    return S_OK;
}

STDMETHODIMP CListControl::get_hWnd(OLE_HANDLE* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = (OLE_HANDLE)m_lvw.m_hWnd;

    return S_OK;
}

STDMETHODIMP CListControl::get_BorderStyle(BorderStyleConstants* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    *pVal = (BorderStyleConstants)this->m_nBorderStyle;
    return S_OK;
}

STDMETHODIMP CListControl::put_BorderStyle(BorderStyleConstants newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    this->m_nBorderStyle = newVal;
    this->SetDirty(TRUE);
    applyProperties();

    return S_OK;
}

STDMETHODIMP CListControl::get_Appearance(AppearanceConstants* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    *pVal = (AppearanceConstants)this->m_nAppearance;
    return S_OK;
}

STDMETHODIMP CListControl::put_Appearance(AppearanceConstants newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    this->m_nAppearance = newVal;
    this->SetDirty(TRUE);
    applyProperties();

    return S_OK;
}

STDMETHODIMP CListControl::get_ColumnHeaders(IColumnHeaders** pVal) {

    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    HRESULT hr = m_hdr->QueryInterface(IID_IColumnHeaders, (void**)pVal);
    return hr;
}

STDMETHODIMP CListControl::get_ListItems(IListItems** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    *pVal = m_litems;
    (*pVal)->AddRef();

    return S_OK;
}

STDMETHODIMP CListControl::get_Font(IFontDisp** pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    // HRESULT hr = OleCreateFontIndirect(&_fontdesc, IID_IFont, (void **)pVal);

    HRESULT hr = m_pFont->QueryInterface(IID_IFont, (void**)pVal);

    return hr;
}

STDMETHODIMP CListControl::put_Font(IFontDisp* newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_pFont = newVal;
    SetDirty(TRUE);
    return S_OK;
}

STDMETHODIMP CListControl::putref_Font(IFontDisp* newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    SetDirty(TRUE);
    m_pFont = newVal;
    return S_OK;
}

STDMETHODIMP CListControl::SetRedraw(VARIANT_BOOL ShouldRedraw) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_bRedrawEnabled = (ShouldRedraw == VARIANT_TRUE ? TRUE : FALSE);
    m_lvw.SetRedraw(m_bRedrawEnabled);

    if (m_bRedrawEnabled) {
        FireViewChange();
        if (this->m_virtualMode && m_litems->m_items.size() > 0) {
            my::lvSetItemCount(lvhWnd(), m_litems->m_items.size());
        }

        sizeToFit();
    }

    return S_OK;
}

STDMETHODIMP CListControl::get_AxHWnd(OLE_HANDLE* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = (OLE_HANDLE)m_hWnd;

    return S_OK;
}

STDMETHODIMP CListControl::get_ContainerhWnd(OLE_HANDLE* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())
    CWindow wnd = GetParent();

    *pVal = (OLE_HANDLE)wnd.m_hWnd;

    return S_OK;
}

STDMETHODIMP CListControl::get_RedrawEnabled(VARIANT_BOOL* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    if (m_bRedrawEnabled)
        *pVal = VARIANT_TRUE;
    else
        *pVal = VARIANT_FALSE;

    return S_OK;
}

STDMETHODIMP CListControl::put_RedrawEnabled(VARIANT_BOOL newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    SetRedraw(newVal);

    return S_OK;
}

STDMETHODIMP CListControl::get_LayoutSuspended(VARIANT_BOOL* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    if (m_bRedrawEnabled)
        *pVal = VARIANT_FALSE;
    else
        *pVal = VARIANT_TRUE;

    return S_OK;
}

STDMETHODIMP CListControl::SuspendLayout() {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    SetRedraw(VARIANT_FALSE);

    return S_OK;
}

STDMETHODIMP CListControl::ResumeLayout() {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    SetRedraw(VARIANT_TRUE);

    return S_OK;
}

STDMETHODIMP CListControl::get_VirtualMode(VARIANT_BOOL* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = my::BoolToVB(m_virtualMode);

    return S_OK;
}

STDMETHODIMP CListControl::put_VirtualMode(VARIANT_BOOL newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_virtualMode = my::VBToBool(newVal);
    // Cannot simply use m_lvw.ModifyStyle here since virtual listviews
    // need the LVS_OWNERDATA flag set at CREATION time only :-(
    RECT r;
    this->GetWindowRect(&r);
    this->CreateControlWindow(GetParent(), r);

    this->Refresh();

    return S_OK;
}
