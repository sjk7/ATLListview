// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __COLUMNHEADERS_H_
#define __COLUMNHEADERS_H_

#include "resource.h" // main symbols

#include "ColumnHeader.h"
#include "my.h"
#include "InterfaceCollection.h"
#include "ATLListView.h"
//#include "_IColumnHeaderEvents_CP.H"
#include "ATLListViewCP.h"

class CListControl;
struct columnheaders : my::com_vector<IDispatch*> {};

class CMyContainer : public CMessageMap {
    public:
    CContainedWindow m_hdr;

    CMyContainer() : m_hdr(this, 0) {}
};

/////////////////////////////////////////////////////////////////////////////
// CColumnHeaders
class ATL_NO_VTABLE CColumnHeaders
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CColumnHeaders, &CLSID_ColumnHeaders>,
      public ISupportErrorInfo,
      public IConnectionPointContainerImpl<CColumnHeaders>,
      public IDispatchImpl<IColumnHeaders, &IID_IColumnHeaders,
          &LIBID_ATLLISTVIEWLib>,
      public CProxy_IColumnHeaderEvents<CColumnHeaders>,
      public CMyContainer

{
    public:
    BEGIN_MSG_MAP(CColumnHeaders)
    MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
    MESSAGE_HANDLER(WM_LBUTTONDOWN, OnMouseDown)
    MESSAGE_HANDLER(WM_LBUTTONUP, OnMouseUp)
    MESSAGE_HANDLER(WM_RBUTTONDOWN, OnMouseDown)
    MESSAGE_HANDLER(WM_RBUTTONUP, OnMouseUp)
    MESSAGE_HANDLER(HDM_LAYOUT, OnHeaderLayout)
    END_MSG_MAP()

    LRESULT OnHeaderLayout(UINT, WPARAM wp, LPARAM lp, BOOL& bHandled) {
        apiSetHeight(m_hWnd, HDM_LAYOUT, wp, lp, bHandled);
        UpdateWindow(m_hWnd);
        return 1;
    }
    /*/
                case HDM_LAYOUT:
                apiSetHeight(hWnd, msg, wParam, lParam, bHandled);
                UpdateWindow(m_hWnd);
                return 1;
                break;
    /*/

    CColumnHeaders();
    columnheaders m_cols;
    virtual ~CColumnHeaders() {
        if (m_col_interface) {
            m_col_interface->Release();
        }
    }

    HWND m_hWndLv;

    void setup(
        CListControl* p, const HWND lvw, const my::win32::ScaleUnits scale) {
        this->m_scaleUnits = scale;
        if (m_hWnd) {
            CMyContainer::m_hdr.Detach();
            m_hWnd = 0;
        }
        ASSERT(m_hWnd == 0); // who has set this behind our back?
        ASSERT(::IsWindow(lvw));
        ASSERT(my::isListView(lvw));

        m_hWndLv = lvw;
        m_plv = p;
        m_hWnd = ListView_GetHeader(lvw);
        ASSERT(::IsWindow(m_hWnd));
        ASSERT(my::isListviewHeader(m_hWnd));
        CMyContainer::m_hdr.SwitchMessageMap(0);
        BOOL b = CMyContainer::m_hdr.SubclassWindow(m_hWnd);
        ASSERT(b);
    }

    LRESULT OnMouseUp(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnits, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseUp(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);
        bHandled = FALSE;

        return 0;
    }

    LRESULT OnMouseDown(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnits, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseDown(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);

        bHandled = FALSE;
        return 0;
    }

    LRESULT OnMouseMove(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnits, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseMove(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);
        bHandled = FALSE;
        return 0;
    }

    private:
    CListControl* m_plv;
    HWND m_hWnd;
    int m_heightInPixels;
    CComObject<CInterfaceCollection>* m_col_interface;

    public:
    int addColumn(const CString& txt, int wid = 100, VARIANT* key = 0);

    /*/
    inline ColumnHeader* findColumnHeader(VARIANT* with_key) {
    if (with_key == 0) return 0;

    for (int i = 0; i < m_cols.isize(); ++i) {
    ColumnHeader* p = (ColumnHeader*)m_cols.at(i);
    if (p->m_key == *with_key) {
    return p;
    }
    }
    return 0;
    }
    /*/

    void applyWidths() {
        for (int i = 0; i < m_cols.isize(); ++i) {
            ColumnHeader* p = (ColumnHeader*)m_cols.at(i);
            p->applyWidth();
        }
    }

    inline HRESULT reportError(
        const std::wstring& what, HRESULT hr = E_FAIL) const {
        return my::atlReportError(
            CLSID_ColumnHeaders, IID_IColumnHeaders, what, hr);
    }

    BOOL apiSetHeight(
        HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    void applyColHeights() {
        // this does the trick for making the new column height apply right
        // away:
        LVCOLUMN c = {0};
        c.mask = LVCF_WIDTH;
        c.cx = 0;
        long was = my::lvGetColumnCount(m_hWndLv);
        long inserted = ListView_InsertColumn(
            m_hWndLv, static_cast<int>(m_cols.size() + 1), &c);
        ListView_DeleteColumn(m_hWndLv, inserted);
        long now = my::lvGetColumnCount(m_hWndLv);
        ASSERT(now == was);
    }

    HRESULT heightInPixelsSet(long newHeight) {
        HD_LAYOUT hdl = {0};
        RECT r;
        ::GetClientRect(m_hWnd, &r);

        hdl.prc = &r;
        WINDOWPOS pos = {0};
        hdl.pwpos = &pos;
        const long oldH = r.bottom - r.top;
        m_heightInPixels = newHeight;
        BOOL ret = ::SendMessage(m_hWnd, HDM_LAYOUT, 0, (LRESULT)&hdl);
        ASSERT(ret);
        if (ret == FALSE) {
            m_heightInPixels = oldH;
            return reportError(L"Unable to set the header height: ");
        }
        m_heightInPixels = newHeight;
        applyColHeights();
        ::GetClientRect(m_hWnd, &r);
        int hip = r.bottom - r.top;
        m_heightInPixels = hip;
        ASSERT(hip == newHeight);
        return S_OK;
    }

    my::win32::ScaleUnits m_scaleUnits;

    long heightInPixels() const noexcept {
        // LVCOLUMN col = {0};
        ASSERT(m_hWnd);
        RECT rc = {0};
        ::GetClientRect(m_hWnd, &rc);
        long ret = rc.bottom - rc.top;

        return ret;
    }

    public:
    BEGIN_CONNECTION_POINT_MAP(CColumnHeaders)
    CONNECTION_POINT_ENTRY(DIID__IColumnHeaderEvents)
    // CONNECTION_POINT_ENTRY(DIID__IListControlEvents)
    END_CONNECTION_POINT_MAP()

    DECLARE_REGISTRY_RESOURCEID(IDR_COLUMNHEADERS)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CColumnHeaders)
    COM_INTERFACE_ENTRY(IColumnHeaders)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
    COM_INTERFACE_ENTRY(IConnectionPointContainer)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IColumnHeaders
    public:
    STDMETHOD(HitTest)
    (FLOAT x, FLOAT y, /*[out, retval]*/ IColumnHeader** columnHeader);
    // STDMETHOD(get_Index)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(get__NewEnum)(/*[out, retval]*/ IUnknown** pVal);
    STDMETHOD(Add)
    (/*[in, optional]*/ VARIANT* Index, /*[in, optional]*/ VARIANT* Key,
        /*[in, optional]*/ VARIANT* Text,
        /*[in, optional]*/ VARIANT* Width,
        /*[in, optional]*/ VARIANT* Alignment, /*[in, optional]*/ VARIANT* Icon,
        IColumnHeader** pRetVal);
    STDMETHOD(get_Item)(LONG index, /*[out, retval]*/ IColumnHeader** pVal);
    STDMETHOD(get_Count)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(get_HeightInPixels)(long* pVal);
    STDMETHOD(put_HeightInPixels)(long newVal);
    STDMETHOD(Clear)();
    STDMETHOD(Remove)(LONG indexToRemove);
};

#endif //__COLUMNHEADERS_H_
