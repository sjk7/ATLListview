// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __LISTCONTROL_H_
#define __LISTCONTROL_H_

#pragma warning(disable : 4995) // deprecated property nonsense
#if _MSC_VER > 1200
#pragma warning(disable : 26812) // enums
#endif
#include "resource.h" // main symbols
#include <atlctl.h>
#include <afxcmn.h>
#include "ColumnHeaders.h"
#include "ListItems.h"
#include "my_disp_ids.h"
#include "ATLListView.h"
#include "my.h"
#include "myIListControl.h"
#include "ATLListViewCP.h"
//#include "_IListControlEvents_CP.h"

typedef enum tagKEYMODIFIERS {
    KEYMOD_NONE = 0X00000000,
    KEYMOD_SHIFT = 0x00000001,
    KEYMOD_CONTROL = 0x00000002,
    KEYMOD_ALT = 0x00000004
} KEYMODIFIERS;

/////////////////////////////////////////////////////////////////////////////
// CListControl
class ATL_NO_VTABLE CListControl
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CStockPropImpl<CListControl, IListControl, &IID_IListControl,
          &LIBID_ATLLISTVIEWLib>,
      public CComControl<CListControl>,
      public IPersistStreamInitImpl<CListControl>,
      public IOleControlImpl<CListControl>,
      public IOleObjectImpl<CListControl>,
      public IOleInPlaceActiveObjectImpl<CListControl>,
      public IViewObjectExImpl<CListControl>,
      public IOleInPlaceObjectWindowlessImpl<CListControl>,
      public ISupportErrorInfo,
      public IConnectionPointContainerImpl<CListControl>,
      public IPersistStorageImpl<CListControl>,
      public ISpecifyPropertyPagesImpl<CListControl>,
      public IQuickActivateImpl<CListControl>,
      public IDataObjectImpl<CListControl>,
      public IProvideClassInfo2Impl<&CLSID_ListControl,
          &DIID__IListControlEvents, &LIBID_ATLLISTVIEWLib>,
      public IPropertyNotifySinkCP<CListControl>,
      public CComCoClass<CListControl, &CLSID_ListControl>,
      public CProxy_IListControlEvents<CListControl> {

    public:
    DECLARE_REGISTRY_RESOURCEID(IDR_LISTCONTROL)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CListControl)
    COM_INTERFACE_ENTRY(IListControl)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(IViewObjectEx)
    COM_INTERFACE_ENTRY(IViewObject2)
    COM_INTERFACE_ENTRY(IViewObject)
    COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceObject)
    COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
    COM_INTERFACE_ENTRY(IOleControl)
    COM_INTERFACE_ENTRY(IOleObject)
    COM_INTERFACE_ENTRY(IPersistStreamInit)
    COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    COM_INTERFACE_ENTRY(IConnectionPointContainer)
    COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
    COM_INTERFACE_ENTRY(IQuickActivate)
    COM_INTERFACE_ENTRY(IPersistStorage)
    COM_INTERFACE_ENTRY(IDataObject)
    COM_INTERFACE_ENTRY(IProvideClassInfo)
    COM_INTERFACE_ENTRY(IProvideClassInfo2)
    COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
    END_COM_MAP()

    BEGIN_PROP_MAP(CListControl)
    /*/
    PROP_ENTRY_TYPE(szDesc,  dispid, clsid, vt)
    /*/
    PROP_DATA_ENTRY("BackColor", m_clrBackColor, VT_I4)
    PROP_ENTRY("Font", MY_FONT_DISP_ID, CLSID_StockFontPage)
    PROP_ENTRY("Enabled", DISPID_ENABLED, CLSID_NULL)
    PROP_DATA_ENTRY("BorderStyle", m_nBorderStyle, VT_I4)
    PROP_DATA_ENTRY("Appearance", m_nAppearance, VT_I2)
    PROP_ENTRY("ForeColor", DISPID_FORECOLOR, CLSID_StockColorPage)
    PROP_ENTRY("TabStop", DISPID_TABSTOP, CLSID_NULL)
    PROP_DATA_ENTRY("LabelEdit", m_labelEdit, VT_I4)
    // PROP_PAGE(CLSID_StockColorPage)
    END_PROP_MAP()

    BEGIN_CONNECTION_POINT_MAP(CListControl)
    // CONNECTION_POINT_ENTRY(DIID__IListControlEvents)
    CONNECTION_POINT_ENTRY(__uuidof(_IListControlEvents))
    CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
    END_CONNECTION_POINT_MAP()

    BEGIN_MSG_MAP(CListControl)

    //  note that we call OnSetCtlFocus instead of OnSetFocus
    //  when the focus is set to the control
    MESSAGE_HANDLER(WM_SETFOCUS, OnSetCtlFocus) //-V1048
    // because we disable the message chain, we should
    // put our own message handler for WM_KILLFOCUS
    MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
    // comment out the following line to disable the message chain
    // CHAIN_MSG_MAP(CComControl<CMyCtl>)
    MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
    MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnDblClick);
    MESSAGE_HANDLER(LVN_BEGINLABELEDIT, OnLabelEdit);

    ALT_MSG_MAP(1)
    // Replace this with message map entries for superclassed Edit
    // if the focus is set to the Edit control we call OnSetFocus
    MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus) //-V1048
    // This OnMouseActivate handler is used to identify if the control
    // is activated by the mouse click
    MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)

    MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
    MESSAGE_HANDLER(WM_LBUTTONDOWN, OnMouseDown)
    MESSAGE_HANDLER(WM_LBUTTONUP, OnMouseUp)
    MESSAGE_HANDLER(WM_RBUTTONDOWN, OnMouseDown)
    MESSAGE_HANDLER(WM_RBUTTONUP, OnMouseUp)
    // MESSAGE_HANDLER(WM_CLOSE, OnLvwClose)
    // MESSAGE_HANDLER(WM_NOTIFY, OnNotifyLvw)
    MESSAGE_HANDLER(WM_DESTROY, OnLvClose)
    MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnDblClick);
    MESSAGE_HANDLER(WM_KEYUP, OnKeyUp);

    END_MSG_MAP()

    BOOL m_bRedrawEnabled;
    OLE_COLOR m_clrBackColor;
    OLE_COLOR m_clrBorderColor;
    CComBSTR m_bstrCaption;
    BOOL m_bEnabled;
    OLE_COLOR m_clrForeColor;
    LONG m_nMousePointer;
    BOOL m_bTabStop;
    CComPtr<IFontDisp> m_pFont;
    BorderStyleConstants m_nBorderStyle;
    AppearanceConstants m_nAppearance;
    typedef AppearanceConstants AppearanceType;
    CComBSTR m_name;
    CString m_scaleUnits;
    my::win32::ScaleUnits m_scaleUnitsEnum;

    typedef CComControl<CListControl> ControlBaseType;
    CContainedWindow m_lvw;
    ListLabelEditConstants m_labelEdit;
    CContainedWindow* listview() { return &m_lvw; }

    HWND lvhWnd() const { return m_lvw.m_hWnd; }

    CComObject<CColumnHeaders>* m_hdr;

    CComObject<CListItems>* m_litems;
    CComObject<CSelItemCollection>* m_selItems;
    BOOL m_bMouseActivate;
    BOOL m_virtualMode;
    CListControl();
    virtual ~CListControl() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        if (m_hdr) m_hdr->Release();

        if (m_litems) m_litems->Release();
        if (m_selItems) m_selItems->Release();
    }

    SCROLLINFO hScrollInfo() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        SCROLLINFO tsi = {0};
        tsi.cbSize = sizeof(tsi);
        tsi.fMask = SIF_TRACKPOS | SIF_RANGE | SIF_PAGE;
        ::GetScrollInfo(lvhWnd(), SB_HORZ, &tsi);
        return tsi;
    }

    HRESULT setScaleUnits() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        BSTR bstr = 0;
        HRESULT hr = this->GetAmbientScaleUnits(bstr);
        ASSERT(SUCCEEDED(hr));
        if (SUCCEEDED(hr)) {
            m_scaleUnits = bstr;
            TRACE(_T("%s %ld\n"), m_scaleUnits.GetBuffer(255),
                m_scaleUnits.GetLength());
            ::SysFreeString(bstr);
            m_scaleUnitsEnum = my::win32::ScaleUnitsFromString(m_scaleUnits);
        }
        if (m_hdr) {
            m_hdr->m_scaleUnits = m_scaleUnitsEnum;
        }
        return hr;
    }

    // Provided because SetObjectRects gets the listview size slightly
    // incorrect.
    void sizeToFit() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        RECT rc;
        CComControl<CListControl>::GetClientRect(&rc);
        int subtract = m_nAppearance == cc3D ? 0 : 0;
        m_lvw.SetWindowPos(NULL, 0, 0, rc.right - rc.left - subtract,
            rc.bottom - rc.top - subtract, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    // called when control dropped onto a form in GUI Design Mode. Like
    // InitProperties in VB6 user control. Note: hWnds are null here :-(
    // This is where you would set the defaults for our properties. But
    // I did it in the constructor already.
    STDMETHOD(InitNew)() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        return IPersistStreamInitImpl<CListControl>::InitNew();
    }

    HRESULT hitTest(
        FLOAT x, FLOAT y, IListItem** retVal, LONG* SubItemIndex = 0) {
        POINT pt = my::win32::VBToPt(x, y, m_scaleUnitsEnum);
        LVHITTESTINFO hti = my::lvHitTest(lvhWnd(), pt);
        const int index = hti.iItem;
        if (retVal && *retVal) {
            (*retVal)->Release();
        }
        if (index >= 0 && index < m_litems->m_items.isize()) {
            IDispatch* pi = m_litems->m_items[index];
            if (pi) {
                HRESULT hr = pi->QueryInterface(IID_IListItem, (void**)retVal);
                if (hr == S_OK) {
                    if (SubItemIndex) {
                        *SubItemIndex = hti.iSubItem;
                    }
                }
                return hr;
            }
        } else {
            *retVal = 0;
            return S_FALSE;
        }

        return S_OK;
    }

    STDMETHOD(HitTest)(FLOAT x, FLOAT y, IListItem** retVal) {
        return hitTest(x, y, retVal);
    }

    STDMETHOD(HitTestEx)
    (FLOAT x, FLOAT y, LONG* SubItemIndex, IListItem** retVal) {
        ASSERT(SubItemIndex);
        HRESULT hr = hitTest(x, y, retVal, SubItemIndex);
        if (hr == S_OK) {
            *SubItemIndex = (*SubItemIndex) + 1;
        }
        return hr;
    }

    int m_viewItemCount;

    inline BOOL viewItemCountSet(int newCount, BOOL ForceToExactSize = FALSE) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        // view's item counts always need to be exact in virtual mode:
        // if (m_virtualMode) ForceToExactSize = TRUE;
        if (m_viewItemCount > newCount && !ForceToExactSize) {
            return TRUE;
        }
        const BOOL ret = my::lvSetItemCount(lvhWnd(), newCount);
        if (ret) {
            m_viewItemCount = newCount;
        }
        return ret;
    }

    LRESULT OnLabelEdit(UINT, WPARAM, LPARAM, BOOL&) {
        TRACE(_T("LabelEdit\n"));
        return 0;
    }

    LRESULT OnKeyUp(UINT, WPARAM wParam, LPARAM, BOOL& bHandled) {
        vbShiftConstants shift = my::win32::getVBKeyStates();
        Fire_KeyUp((short)wParam, (short)shift);

        bHandled = FALSE;
        return 0;
    }

    LRESULT OnDblClick(UINT, WPARAM, LPARAM, BOOL& bHandled) {
        bHandled = FALSE;
        Fire_DblClick();
        return 0;
    }

    LRESULT OnMouseUp(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnitsEnum, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseUp(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);
        bHandled = FALSE;
        return 0;
    }

    LRESULT OnMouseDown(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnitsEnum, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseDown(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);
        bHandled = FALSE;
        return 0;
    }

    LRESULT OnMouseMove(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        my::win32::InputInfo ii(
            m_scaleUnitsEnum, &lParam, 0, wParam == 0 ? NULL : &wParam);
        Fire_MouseMove(static_cast<SHORT>(ii.button),
            static_cast<SHORT>(ii.shift), ii.point.x, ii.point.y);
        bHandled = FALSE;
        return 0;
    }

    // cached, because it is slow to ask the actual listview about anything,
    // much!
    inline int viewItemCountGet() const { return m_viewItemCount; }

    LRESULT OnLvClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        (void)bHandled;

        // if (m_phdrSubclass) delete m_phdrSubclass; //-V809
        // m_phdrSubclass = 0;
        return m_lvw.DefWindowProc(uMsg, wParam, lParam);
    }

    inline void showColumnsCtxMnu(HWND hWnd, const POINT& pointScreen) {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        HMENU hPopupMenu = ::CreatePopupMenu();
        ASSERT(m_hdr);
        if (!m_hdr) return;
        const int sz = m_hdr->m_cols.isize();
        for (int i = 0; i < sz; ++i) {
            ColumnHeader* ph = (ColumnHeader*)m_hdr->m_cols.at(i);
            if (ph->m_bVisible) {
                ::InsertMenu(hPopupMenu, (UINT)-1,
                    MF_CHECKED | MF_BYCOMMAND | MF_STRING, i, ph->m_name);
            } else {
                ::InsertMenu(hPopupMenu, (UINT)-1,
                    MF_UNCHECKED | MF_BYCOMMAND | MF_STRING, i, ph->m_name);
            }
        }

        ::InsertMenu(hPopupMenu, (UINT)-1, MF_STRING, sz, _T("Cancel"));
        int which = TrackPopupMenu(hPopupMenu,
            TPM_TOPALIGN | TPM_LEFTALIGN | TPM_RETURNCMD, pointScreen.x,
            pointScreen.y, 0, hWnd, NULL);
        TRACE(_T("Clicked on menu: %d\n"), which);
        ::DestroyMenu(hPopupMenu);
        if (which >= sz) {
            TRACE(_T("Cancel clicked on Listview header popupMenu\n"));
            return;
        }
        if (which >= 0) {
            ColumnHeader* ph = (ColumnHeader*)m_hdr->m_cols.at(which);
            ph->toggleVisible();
        }
    }

    inline HRESULT getColumnHeader(int index, IColumnHeader** p) const {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        HRESULT hr = DISP_E_BADINDEX;
        if (m_hdr && index >= 0 && index < m_hdr->m_cols.isize()) {
            IDispatch* pd = m_hdr->m_cols.at(index);
            if (*p) (*p)->Release();
            hr = pd->QueryInterface(
                IID_IColumnHeader, reinterpret_cast<void**>(p));
            ASSERT(SUCCEEDED(hr));
            return hr;
        }
        return DISP_E_BADINDEX;
    }

    HRESULT doFireRightColClick(VARIANT_BOOL& bDoDefault, short ivbshift,
        my::win32::VBPOINTF vbPoint, IColumnHeader* pcolHeader) {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        // There is no need for all this passing by pointer:
        // It's to make it API compat with Listview_API
        this->AddRef();
        HRESULT hr = this->Fire_ColumnRightClick(
            &bDoDefault, ivbshift, vbPoint.x, vbPoint.y, pcolHeader);
        if (pcolHeader) pcolHeader->Release();
        this->Release();
        ASSERT(SUCCEEDED(hr));
        return hr;
    }

    inline LRESULT rightClickColumnHandler(const HWND hWndLvw) {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        HWND hChildWnd = 0;
        const POINT pointScreen = my::win32::ScreenPoint();
        POINT PointLVClient = pointScreen;
        POINT pointHdr = pointScreen; //-V821

        ::ScreenToClient(hWndLvw, &PointLVClient);
        // Because the header turns out to be a child of the
        // listview control, we obtain its handle here.
        hChildWnd = ::ChildWindowFromPoint(hWndLvw, PointLVClient);

        // NULL hChildWnd means R-CLICKED outside the listview.
        // hChildWnd == hWndLvw means listview got clicked: NOT the
        // header.
        if (hChildWnd == 0) {
            ASSERT("Right Clicked on listview, not a column" == 0);
        } else {
            if ((hChildWnd) && (hChildWnd != hWndLvw)) { //-V560
                ASSERT(hChildWnd == this->m_hdrWnd);
                // Transform to client coordinates
                // relative to HEADER control, NOT the listview!
                // Otherwise, incorrect column number is returned.
                HD_HITTESTINFO hdhti;
                my::lvHeaderHitTest(hChildWnd, pointHdr, hdhti);
                short ivbshift = 0;
                my::win32::getVBKeyStates(&ivbshift);
                IColumnHeader* pcolHeader = 0;
                if (hdhti.iItem >= 0) {
                    HRESULT hr = getColumnHeader(hdhti.iItem, &pcolHeader);
                    if (FAILED(hr)) return 0;
                } else {
                    // -1 if clicked on a header with no columns
                    pcolHeader = 0;
                }
                my::win32::VBPOINTF vbPoint
                    = my::win32::ptToVB(pointHdr, m_scaleUnitsEnum);
                ///////////////////////////////////////////////////////////////////
                VARIANT_BOOL bDoDefault = VARIANT_TRUE;
                doFireRightColClick(bDoDefault, ivbshift, vbPoint, pcolHeader);
                ///////////////////////////////////////////////////////////////////
                TRACE(_T("Header %d got right clicked\n"), hdhti.iItem);
                if (bDoDefault == VARIANT_TRUE) {
                    showColumnsCtxMnu(hWndLvw, pointScreen);
                }
            }
        }

        return 0L;
    }

    HRESULT leftClickColumnHandler(NMHEADER* phdr) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        ASSERT(phdr);
        if (!phdr) return E_INVALIDARG;
        IColumnHeader* pcolHeader = 0;
        HRESULT hr = m_hdr->get_Item(phdr->iItem + 1, &pcolHeader);
        ASSERT(SUCCEEDED(hr));

        this->AddRef();
        if (SUCCEEDED(hr)) hr = Fire_ColumnClick(pcolHeader);
        pcolHeader->Release();
        this->Release();
        ASSERT(SUCCEEDED(hr)); // check the Fire_ ... is returning the result of
        // the Invoke call, and not the variant scode
        // See the included text file: Fire_byRef_HOWTO.txt: VC6 wizard
        // gets it all wrong!
        return hr;
    }

    /*/
    wParam: The identifier of the common control sending the message. This
    identifier is not guaranteed to be unique. An application should use the
    hwndFrom or idFrom member of the NMHDR structure (passed as the lParam
    parameter) to identify the control.

      lParam:	A pointer to an NMHDR structure that contains the notification
      code and additional information. For some notification messages, this
      parameter points to a larger structure that has the NMHDR structure as its
      first member.
    /*/
    LRESULT OnNotify(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {

        bHandled = TRUE;
        // int id = (int)wParam;
        NMHDR* pnmhdr = (NMHDR*)lParam;
        HWND hWnd = pnmhdr->hwndFrom;
        (void)wParam;
        (void)uMsg;

        if (hWnd == m_lvw.m_hWnd) {
            LRESULT retval = my::lvHandleNotify(wParam, lParam, this,
                m_virtualMode, m_litems->m_items, pnmhdr, bHandled);
            return retval;

        } else if (pnmhdr->hwndFrom == m_hdrWnd) {
            if (pnmhdr->code == HDN_ITEMCLICK) {
                NMHEADER* phdr = (NMHEADER*)lParam;
                leftClickColumnHandler(phdr);
            } else if (pnmhdr->code == NM_RCLICK) {
                rightClickColumnHandler(lvhWnd());
            }
        }
        bHandled = FALSE;
        return 0;
    }
    int m_lastSelItemIndex;
    // helper for buggy virtual listview state notifications
    virtual int getLastSelItemIndex() { return m_lastSelItemIndex; }
    virtual void setLastSelItemIndex(int idx) { m_lastSelItemIndex = idx; }

    LRESULT OnMouseActivate(
        UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        // Set the flag to be true so that we know the control is
        // activated with mouse click
        m_bMouseActivate = TRUE;
        LRESULT lRes = CComControl<CListControl>::OnMouseActivate(
            uMsg, wParam, lParam, bHandled);
        return lRes;
    }

    LRESULT OnSetCtlFocus(
        UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        // We only call the OnSetFocus if it's not activated by
        // mouse click
        if (!m_bMouseActivate) {
            OnSetFocus(uMsg, wParam, lParam, bHandled);
        }
        m_bMouseActivate = FALSE;
        // Fire_GotFocus();
        return 0;
    }
#if _MSC_VER > 1200
#pragma warning(disable : 6387)
#endif

    LRESULT OnSetFocus(

        UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        LRESULT lRes = CComControl<CListControl>::OnSetFocus(
            uMsg, wParam, lParam, bHandled);
        if (m_bInPlaceActive) {
            DoVerbUIActivate(&m_rcPos, 0);
            if (!IsChild(::GetFocus())) {
                m_lvw.SetFocus();
            }
        }
        return lRes;
    }
#if _MSC_VER > 1200
#pragma warning(default : 6387)
#endif

    // This is better than many claimed 'fixes' for tab issues,
    // as it works even if the control is sited in a VB Usercontrol.
    BOOL PreTranslateAccelerator(LPMSG pMsg, HRESULT& hRet) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        if (pMsg->message == WM_KEYDOWN
            && (pMsg->wParam == VK_LEFT || pMsg->wParam == VK_RIGHT
                || pMsg->wParam == VK_UP || pMsg->wParam == VK_DOWN)) {
            hRet = S_OK;
            m_lvw.SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
            return TRUE;
        }

        my::lvHandleKeyPressFromPreTranslate(this, pMsg);

        return FALSE;
    }

    inline BOOL InDesignMode() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        BOOL usermode = FALSE;
        HRESULT hr = this->GetAmbientUserMode(usermode);
        if (FAILED(hr)) return FALSE;
        if (usermode) return FALSE;
        return TRUE;
    }

    static inline LONG TranslateColor(OLE_COLOR ole_color) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        COLORREF retval = 0;
        HRESULT hr = ::OleTranslateColor(ole_color, 0, &retval);
        if (!SUCCEEDED(hr)) {
            ASSERT(0);
            TRACE(_T("Translate Color did not succeed\n"));
        }

        return retval;
    }

    STDMETHODIMP Refresh() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        if (IsWindow() && m_lvw.IsWindow()) {
            this->Invalidate();
            m_lvw.Invalidate();
            // this->UpdateWindow();
            // m_lvw.Invalidate();
            FireViewChange();
            sizeToFit();
        }
        return S_OK;
    }

    // Called when a property is set on the HOST control in the client.
    // Allows us to 'inherit' changes made on the container form.
    // Note, from Steve: to get this event you need to copy n paste
    // the code in IMPLEMENT_STOCKPROP in atlctl.h from the VB6 one, to the one
    // we are using.
    STDMETHOD(OnAmbientPropertyChange)(DISPID dispid);

    HRESULT FireOnChanged(DISPID dispID) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        // saved (stock) property changed either in runtime or in design mode.
        if (dispID == DISPID_APPEARANCE || dispID == DISPID_BORDERSTYLE
            || dispID == DISPID_BACKCOLOR) {
        }
        if (InDesignMode()) {
            if (dispID == DISPID_FONT) {
                TRACE(_T("Setting font\n"));
            }
            if (IsWindow()) {
                applyProperties();
            }
        } else {
            // always apply all things in running mode

            if (IsWindow()) applyProperties();
        }
        return ControlBaseType::FireOnChanged(dispID);
    }

    void initProperties() {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_pFont = 0;
        HRESULT hr = GetAmbientFontDisp(&m_pFont);
        ASSERT(SUCCEEDED(hr));

        applyProperties();
    }

    void applyProperties() {
        // if (!InDesignMode()) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        ListView_SetBkColor(lvhWnd(), TranslateColor(m_clrBackColor));
        ListView_SetTextBkColor(lvhWnd(), TranslateColor(m_clrBackColor));
        //}
        my::lvSetFont(lvhWnd(), m_pFont);
        this->EnableWindow(abs(m_bEnabled));
        m_lvw.EnableWindow(abs(m_bEnabled));
        if (m_nBorderStyle == ccFixedSingle) {
            ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME);
        } else {
            ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME);
        }

        if (m_nAppearance == cc3D) {
            ModifyStyleEx(0, WS_EX_CLIENTEDGE, SWP_DRAWFRAME);
        } else {
            ModifyStyleEx(WS_EX_CLIENTEDGE, 0, SWP_DRAWFRAME);
        }

        setLabelEdit(m_labelEdit);

        setScaleUnits();

        Refresh();
    }

    void setLabelEdit(ListLabelEditConstants newMode) {
        LONG style = ::GetWindowLong(lvhWnd(), GWL_STYLE);
        if (newMode == lvwAutomatic) {
            style |= LVS_EDITLABELS;

        } else {
            style &= ~LVS_EDITLABELS;
        }

        ::SetWindowLong(lvhWnd(), GWL_STYLE, style);
        m_labelEdit = newMode;
        ASSERT(labelEditAPI() == newMode);
    }

    ListLabelEditConstants labelEditAPI() const {
        LONG style = ::GetWindowLong(lvhWnd(), GWL_STYLE);
        if ((style & LVS_EDITLABELS) == LVS_EDITLABELS) {
            return lvwAutomatic;
        } else {
            return lvwManual;
        }
    }

    long getItemCountAPI() const {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        return ListView_GetItemCount(m_lvw.m_hWnd);
    }

    HWND m_hWndVBHostSite;

    HRESULT reportError(const std::wstring& what, HRESULT hr = E_FAIL) const {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        return my::atlReportError(
            CLSID_ListControl, IID_IListControl, what, hr);
    }

    HRESULT myCreateControlWindow(
        HWND hWndParent, RECT& rcPos, DWORD style = 0) {

        AFX_MANAGE_STATE(AfxGetStaticModuleState())

        m_hWndVBHostSite = hWndParent;

        DWORD lvstyle = my::lvDefaultStyle(m_virtualMode);
        lvstyle |= style;

        // my::win32::debugShowClassName(hWndParent);

        if (m_lvw.m_hWnd) {
            // ASSERT(m_phdrSubclass);
            m_lvw.DestroyWindow();
            m_lvw.Detach();
        }
        m_viewItemCount = 0;
        if (m_lvw.m_hWnd == 0) {
            const HWND myhWnd = m_lvw.Create(m_hWnd, rcPos, _T(""), lvstyle);
            ASSERT(myhWnd);
            if (!myhWnd) {

                std::wstring ws(L"Unable to create listview itself during "
                                L"creation\nBad flags or out of memory");
                return AtlReportError(CLSID_ListControl, ws.c_str(),
                    IID_IListControl, E_OUTOFMEMORY);
            }
            SendMessage(myhWnd, LVM_SETEXTENDEDLISTVIEWSTYLE,
                LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES,
                LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
            sizeToFit();
            CComObject<CListItems> litems;
            if (m_litems) {
                m_litems->Release();
                m_litems = 0;
            }
            HRESULT hr = litems.CreateInstance(&m_litems);
            ASSERT(SUCCEEDED(hr));
            if (FAILED(hr)) {
                return AtlReportError(CLSID_ListControl,
                    L"Listview creation: unable to create internal listitems",
                    IID_IListControl, E_OUTOFMEMORY);
            }
            m_litems->AddRef();
            if (m_selItems) {
                m_selItems->Release();
                m_selItems = 0;
            }
            CComObject<CSelItemCollection> selItems;
            hr = selItems.CreateInstance(&m_selItems);
            ASSERT(SUCCEEDED(hr));
            if (FAILED(hr)) {
                return AtlReportError(CLSID_ListControl,
                    L"Listview creation: unable to create internal (selected) "
                    L"listitems",
                    IID_IListControl, E_OUTOFMEMORY);
            }
            m_selItems->AddRef();

            ASSERT(SUCCEEDED(hr));
            if (FAILED(hr)) {
                return AtlReportError(CLSID_ListControl,
                    L"Listview creation: unable to create internal listitems",
                    IID_IListControl, E_OUTOFMEMORY);
            }

            CListItems* items = m_litems;
            if (!m_virtualMode) { // not in virtualMode because SetItemCount is
                // interpreted differently: you actually "see"
                // the rows!
                const BOOL b = ListView_SetItemCount(myhWnd, 50000);
                ASSERT(b);
                if (!b) {
                    std::wstring ws(
                        L"Unable to reserve memory upon creation\n");
                    return AtlReportError(CLSID_ListControl, ws.c_str(),
                        IID_IListControl, E_OUTOFMEMORY);
                }
            }

            const HWND hh = ListView_GetHeader(myhWnd);
            ASSERT(::IsWindow(hh));
            this->setHeaderhWnd(hh);

            ASSERT(items);
            if (items) items->m_lv = this;
            (void)hr;

#define TYPICAL_ITEM_COUNT 20000

            if ((!m_virtualMode) && (viewItemCountGet() < TYPICAL_ITEM_COUNT)) {
                viewItemCountSet(TYPICAL_ITEM_COUNT);
            }

            initProperties();
        }

        return S_OK;
    }

    // where hWndParent is the VB6 site: something like
    // 0x0019f258 "ThunderFormDC" or 0x0019ec20 "ThunderUserControlDC"
    virtual HWND CreateControlWindow(HWND hWndParent, RECT& rcPos) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        HWND& hWnd = this->m_hWnd;
        if (hWnd == 0) {
            ControlBaseType::CreateControlWindow(hWndParent, rcPos);
        }
        if (!hWnd) {
            AtlReportError(CLSID_ListControl,
                L"Could not create listview window. Out of memory or bad flags",
                IID_IListControl, E_OUTOFMEMORY);
            return 0;
        }

        myCreateControlWindow(hWndParent, rcPos, LVS_SINGLESEL);

        return hWnd;
    }

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid) {
        static const IID* arr[] = {
            &IID_IListControl,
        };
        for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) { //-V1008
            if (my::InlineIsEqualGUID(*arr[i], riid)) return S_OK;
        }
        return S_FALSE;
    }

    // IViewObjectEx
    DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

    // IListControl
    public:
    STDMETHOD(get_VirtualMode)(/*[out, retval]*/ VARIANT_BOOL* pVal);
    STDMETHOD(put_VirtualMode)(/*[in]*/ VARIANT_BOOL newVal);
    STDMETHOD(ResumeLayout)();
    STDMETHOD(SuspendLayout)();
    STDMETHOD(get_LayoutSuspended)(/*[out, retval]*/ VARIANT_BOOL* pVal);
    STDMETHOD(get_RedrawEnabled)(/*[out, retval]*/ VARIANT_BOOL* pVal);
    STDMETHOD(put_RedrawEnabled)(/*[in]*/ VARIANT_BOOL newVal);
    STDMETHOD(get_ContainerhWnd)(/*[out, retval]*/ OLE_HANDLE* pVal);
    STDMETHOD(get_AxHWnd)(/*[out, retval]*/ OLE_HANDLE* pVal);
    STDMETHOD(SetRedraw)(VARIANT_BOOL ShouldRedraw);
    STDMETHOD(get_Font)(/*[out, retval]*/ IFontDisp** pVal);
    STDMETHOD(put_Font)(/*[in]*/ IFontDisp* newVal);
    STDMETHOD(putref_Font)(/*[in]*/ IFontDisp* newVal);

    STDMETHOD(get_ListItems)(/*[out, retval]*/ IListItems** pVal);
    STDMETHOD(get_ColumnHeaders)(/*[out, retval]*/ IColumnHeaders** pVal);
    STDMETHOD(get_hWnd)(/*[out, retval]*/ OLE_HANDLE* pVal);

    STDMETHOD(get_BorderStyle)(BorderStyleConstants* pVal);
    STDMETHOD(put_BorderStyle)(BorderStyleConstants newVal);

    STDMETHOD(get_Appearance)(AppearanceConstants* pVal);
    STDMETHOD(put_Appearance)(AppearanceConstants newVal);

    STDMETHOD(SetObjectRects)(LPCRECT prcPos, LPCRECT prcClip) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        IOleInPlaceObjectWindowlessImpl<CListControl>::SetObjectRects(
            prcPos, prcClip);

        sizeToFit();

        return S_OK;
    }

    SortInfo m_sortInfo;

    STDMETHOD(get_Sorted)(VARIANT_BOOL* b) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        *b = my::BoolToVB(m_sortInfo.sorted);
        return S_OK;
    }
    STDMETHOD(put_Sorted)(VARIANT_BOOL b) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_sortInfo.sorted = my::VBToBool(b);
        m_litems->sort(m_sortInfo);
        ListView_RedrawItems(lvhWnd(), 0, m_litems->size());

        return S_OK;
    }

    STDMETHOD(get_SortKey)(short* pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        *pVal = static_cast<short>(m_sortInfo.key);
        return S_OK;
    }

    STDMETHOD(put_SortKey)(short newVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_sortInfo.key = newVal;
        if (m_sortInfo.sorted) {
            m_litems->sort(m_sortInfo);
            ListView_RedrawItems(lvhWnd(), 0, m_litems->size());
        }
        return S_OK;
    }

    STDMETHOD(put_SortOrder)(ListSortOrderFlags so) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_sortInfo.order = so;
        return S_OK;
    }

    STDMETHOD(get_SortOrder)(ListSortOrderFlags* pso) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        *pso = m_sortInfo.order;
        return S_OK;
    }
    BOOL m_bMultiSelect;
    STDMETHOD(get_MultiSelect)(VARIANT_BOOL* pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        *pVal = my::BoolToVB(m_bMultiSelect);
        return S_OK;
    }

    STDMETHOD(put_MultiSelect)(VARIANT_BOOL pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        m_bMultiSelect = my::VBToBool(pVal);

        if (!m_bMultiSelect) {
            m_lvw.ModifyStyle(0, LVS_SINGLESEL);
        } else {
            m_lvw.ModifyStyle(LVS_SINGLESEL, 0);
        }

        return S_OK;
    }

    STDMETHOD(get_SelectedItems)(ISelItemCollection** pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        if (!pVal) {
            return DISP_E_TYPEMISMATCH;
        }
        if (*pVal) (*pVal)->Release();
        HRESULT hr
            = m_selItems->QueryInterface(IID_ISelItemCollection, (void**)pVal);
        if (SUCCEEDED(hr)) {
            m_selItems->setData(m_litems->m_items.getVector());
        }
        return hr;
    }

    STDMETHOD(get_SelectedItemCount)(LONG* pVal) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        *pVal = ListView_GetSelectedCount(lvhWnd());
        return S_OK;
    }

    void handleClick(const int code, const LVHITTESTINFO& data) {

        if (data.iItem < 0 && data.iSubItem < 0) {
            this->Fire_Click();
            if (code == NM_CLICK) {
                if (m_hdr->m_scaleUnits == my::win32::twipsUnits) {

                    int tppy = my::win32::twipsPerPixel(LOGPIXELSY);
                    int tppx = my::win32::twipsPerPixel(LOGPIXELSX);
                    this->Fire_ClickEx(data.pt.x * tppx, data.pt.y * tppy);
                } else {
                    this->Fire_ClickEx(data.pt.x, data.pt.y);
                }
            }

            return;
        }

        HRESULT hr = my::bounds_check(data.iItem, m_litems->m_items);
        if (FAILED(hr)) {
            // ASSERT(0);
            return;
        }

        CComPtr<CListItem> p = 0;
        hr = m_litems->m_items[data.iItem]->QueryInterface(
            IID_IListItem, (void**)&p);
        if (FAILED(hr)) {
            ASSERT(0);
            return;
        }
        ListItem* pitem = (ListItem*)p.p;
        switch (code) {
            case NM_CLICK:
                this->Fire_ItemClick(pitem);
                this->Fire_ItemClickEx(
                    pitem, static_cast<SHORT>(data.iSubItem));
                break;
            case NM_DBLCLK: break;
            case NM_RCLICK:
                this->Fire_ItemClickRight(
                    pitem, static_cast<SHORT>(data.iSubItem));
                break;
            default: ASSERT("handleClick: bad code" == 0);
        }
    }

    private: // for developer use, internally
    std::string m_sid;
    HWND m_hdrWnd;

    void setHeaderhWnd(const HWND hWnd) {
        AFX_MANAGE_STATE(AfxGetStaticModuleState())
        // ASSERT(m_phdrSubclass == 0);
        m_hdrWnd = hWnd;
        // m_phdrSubclass = new my::win32::Subclasser<CListControl>(this, hWnd);
        this->m_hdr->setup(this, lvhWnd(), m_scaleUnitsEnum);
    }

    public:
    __inline void setLvItemCount(int newCount) {
        my::lvSetItemCount(lvhWnd(), newCount);
    }

#if _MSC_VER > VC6_VERSION
#pragma warning(disable : 26819)
#endif

    LRESULT OnSubclassProc(
        HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
        switch (msg) {
            case HDM_LAYOUT:
                if (m_hdr) {
                    // LPHDLAYOUT pHL = reinterpret_cast<LPHDLAYOUT>(lParam);
                    if (m_hdr->apiSetHeight(
                            hWnd, msg, wParam, lParam, bHandled)) {
                        this->Refresh();
                        return 1;
                        break;
                    }
                }
                break;
            default: {
                break;
            }
        }

        bHandled = FALSE;
        return 0;
    }

    public:
    STDMETHOD(StartLabelEdit)();
    STDMETHOD(get_LabelEdit)(/*[out, retval]*/ ListLabelEditConstants* pVal);
    STDMETHOD(put_LabelEdit)(/*[in]*/ ListLabelEditConstants newVal);
};

#endif //__LISTCONTROL_H_
