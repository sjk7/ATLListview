// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com
//

// ColumnHeader.cpp : Implementation of ColumnHeader
#include "stdafx.h"
#include "ATLListView.h"
#include "ColumnHeader.h"
#include "ListControl.h"
#include "my.h"
/////////////////////////////////////////////////////////////////////////////
// ColumnHeader

int ColumnHeader::apiColumnWidth(int index) const noexcept {
    return my::lvGetColumnWidth(m_plv->lvhWnd(), index);
}

STDMETHODIMP ColumnHeader::InterfaceSupportsErrorInfo(REFIID riid) {
    static const IID* arr[] = {&IID_IColumnHeader};
    for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
        if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
    }
    return S_FALSE;
}

STDMETHODIMP ColumnHeader::get_Text(BSTR* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_name.AllocSysString();
    return S_OK;
}

BOOL ColumnHeader::apiSetText() noexcept {
    // ListView_SetColumn()
    LVCOLUMN col = {0};
    col.pszText = m_name.GetBuffer(256);
    col.cchTextMax = m_name.GetLength();
    col.fmt = this->m_align;
    DWORD mask = LVCF_TEXT | LVCF_FMT;
    col.mask = mask;

    BOOL b = ListView_SetColumn(m_plv->lvhWnd(), m_VBIndex - 1, &col);
    ASSERT(b);
    return b;
}

long ColumnHeader::heightInPixels() const noexcept {
    return m_hdrs->heightInPixels();
}

STDMETHODIMP ColumnHeader::put_Text(BSTR* newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    CString n(*newVal);
    m_name = n;
    if (apiSetText()) {
        return S_OK;
    } else {
        return reportError(L"Failed to set the text of the columnHeader");
    }
}

STDMETHODIMP ColumnHeader::get_Index(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_VBIndex;

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_Width(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_width;

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_Width(LONG newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    ASSERT(m_hdrs);
    // negative numbers: eg LVSCW_AUTOSIZE_USEHEADER
    if (m_hdrs->m_scaleUnits == my::win32::twipsUnits && newVal > 0) {
        newVal /= my::win32::twipsPerPixel();
    }

    m_width = newVal;
    if (m_plv) {
        HWND hWnd = m_plv->lvhWnd();
        if (::IsWindow(hWnd)) {
            int ncols = my::lvGetColumnCount(hWnd);
            int apiIndex = m_VBIndex - 1;
            if (apiIndex >= 0 && apiIndex < ncols) {

                BOOL ret = ::SendMessage(
                    hWnd, LVM_SETCOLUMNWIDTH, apiIndex, m_width);
                ASSERT(ret);
                if (!ret) {
                    return reportError(
                        L"Setting columnHeader width: index out of bounds",
                        DISP_E_BADINDEX);
                }
            }
        }
    }

    m_widthInPixels = this->apiColumnWidth(m_VBIndex - 1);
    return S_OK;
}

STDMETHODIMP ColumnHeader::get_ResizeMode(ListColumnResizeMode* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = resizeMode();

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_ResizeMode(ListColumnResizeMode newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    resizeModeSet(newVal);

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_ColumnContentKind(ColumnContentType* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = contentType();

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_ColumnContentKind(ColumnContentType newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    contentTypeSet(newVal);

    return S_OK;
}

void ColumnHeader::setListview(CListControl* p) {
    m_plv = p;
    m_lvhWnd = p->lvhWnd();
}

long ColumnHeader::left_in_pixels() {

    long otherWidths = 0;
    if (m_plv && m_plv->m_hdr) {

        for (int i = 0; i < m_plv->m_hdr->m_cols.isize(); ++i) {

            IDispatch* raw = m_plv->m_hdr->m_cols.at(i);
            ColumnHeader* ch = static_cast<ColumnHeader*>(raw);
            if (ch->m_VBIndex == m_VBIndex) {
                break;
            }
            otherWidths = otherWidths + ch->m_width;
        }
    }

    return otherWidths;
}

STDMETHODIMP ColumnHeader::EnsureVisible() {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    int colLeft = 0;
    if (this->m_plv) {
        SCROLLINFO si = m_plv->hScrollInfo();
        // long ScrollBarPosNow = si.nPos;
        RECT rc;
        ::GetClientRect(m_plv->lvhWnd(), &rc);
        long ViewUpTo = 0;
        colLeft = m_leftInPixels;
        if (colLeft > ViewUpTo) {
            BOOL ok = ::SendMessage(m_plv->lvhWnd(), LVM_SCROLL, colLeft, 0);
            if (!ok) {
                return reportError(L"Column EnsureVisible failed: ");
            }
        }

    } else {
        return reportError(L"Cannot call EnsureVisible() if a column does not "
                           L"belong to a listview!");
    }

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_LeftInPixels(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = left_in_pixels();

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_Alignment(ListColumnAlignmentConstants* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_align;

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_Alignment(ListColumnAlignmentConstants newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    if (!m_plv) {
        return reportError(L"Cannot set a value on a ColumnHeader that is not "
                           L"associated with a listview!");
    }
    m_align = newVal;
    LVCOLUMN col = {0};

    col.fmt = m_align;
    col.mask = LVCF_FMT;

    BOOL ok = ListView_SetColumn(m_plv->lvhWnd(), m_VBIndex - 1, &col);
    if (!ok) {
        return reportError(L"Setting alignment on column failed");
    }

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_FixedWidth(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_fixedWidth;

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_FixedWidth(LONG newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_fixedWidth = newVal;

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_Key(BSTR* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_sKey.AllocSysString();

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_SubItemIndex(short* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_VBsubItemIndex;

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_Tag(VARIANT* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    *pVal = m_tag;

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_Tag(VARIANT newVal) //-V813
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState())

    m_tag = newVal;

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_WidthInPixels(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    *pVal = m_widthInPixels;

    return S_OK;
}

STDMETHODIMP ColumnHeader::get_HeightInPixels(LONG* pVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    *pVal = m_hdrs->heightInPixels();

    return S_OK;
}

STDMETHODIMP ColumnHeader::put_HeightInPixels(LONG newVal) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    HRESULT hr = m_hdrs->heightInPixelsSet(newVal);

    return hr;
}
