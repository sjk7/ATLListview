// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com

#ifndef __COLUMNHEADER_H_
#define __COLUMNHEADER_H_

#include "resource.h" // main symbols
#include "my.h"
class CColumnHeaders;
class CListControl;
#define DEFAULT_WIDTH 1440
/////////////////////////////////////////////////////////////////////////////
// ColumnHeader
class ATL_NO_VTABLE ColumnHeader
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<ColumnHeader, &CLSID_ColumnHeader>,
      public ISupportErrorInfo,
      public IDispatchImpl<IColumnHeader, &IID_IColumnHeader,
          &LIBID_ATLLISTVIEWLib> {
    public:
    ColumnHeader()
        : m_VBIndex(-1)
        , m_hdrs(0)
        , m_plv(0)
        , m_width(LVSCW_AUTOSIZE_USEHEADER)
        , m_bVisible(TRUE)
        , m_resizeMode(lvwResizeHorizontalProportion)
        , m_contentType(lvText)
        , m_leftInPixels(0)
        , m_left(0)
        , m_align(lvwColumnLeft)
        , m_fixedWidth(-1) // means we never auto-resize the column. User can
                           // resize it still if he wants
        , m_VBsubItemIndex(-1)
        , m_widthInPixels(0)
        , m_lvhWnd(0)
        , m_hWnd(0)
        , m_heightInPixels(0)

    {}

    CString m_key;
    __inline void setKey(VARIANT* key) {
        HRESULT hr = my::VariantToCString(key, m_key, true);
        ASSERT(SUCCEEDED(hr));
        (void)hr;
        return;
    }
    void setListview(CListControl* p);

    __inline void setName(const CString& name) { m_name = name; }
    __inline void setVBIndex(size_t index) { m_VBIndex = index; }
    __inline void setup(
        CListControl* lv, HWND lvhWnd, CColumnHeaders* hdrs, HWND hWnd) {
        m_plv = lv;
        m_lvhWnd = lvhWnd;
        m_hdrs = hdrs;
        m_hWnd = hWnd;
    }
    CString m_name;
    int m_VBIndex;
    CColumnHeaders* m_hdrs;
    CListControl* m_plv;
    LONG m_width;
    BOOL m_bVisible;
    ColumnContentType m_contentType;
    int m_leftInPixels;
    int m_left;
    ListColumnAlignmentConstants m_align;
    int m_fixedWidth;

    long left_in_pixels();
    CString m_sKey;
    short m_VBsubItemIndex;
    CComVariant m_tag;
    // always reports the physical width of the
    // column, even if it was set to a minus value()
    // eg: LVSCW_AUTOSIZE_USEHEADER.
    int m_widthInPixels;
    HWND m_lvhWnd;
    HWND m_hWnd;
    int m_heightInPixels;

    inline HRESULT reportError(
        const std::wstring& what, HRESULT hr = E_FAIL) const {
        return my::atlReportError(
            CLSID_ColumnHeader, IID_IColumnHeader, what, hr);
    }
    // returns the new value of visible
    inline BOOL toggleVisible() {
        if (m_bVisible)
            m_bVisible = FALSE;
        else
            m_bVisible = TRUE;

        if (!m_bVisible) {
            this->put_Width(0);
        } else {
            put_Width(LVSCW_AUTOSIZE_USEHEADER);
        }

        return m_bVisible;
    }

    void applyWidth() { put_Width(m_width); }
    ListColumnResizeMode m_resizeMode;

    inline ListColumnResizeMode resizeMode() const noexcept {
        return m_resizeMode;
    }

    inline void resizeModeSet(ListColumnResizeMode newMode) {
        m_resizeMode = newMode;
    }

    inline ColumnContentType contentType() const { return m_contentType; }

    inline void contentTypeSet(ColumnContentType newval) {
        m_contentType = newval;
    }

    int apiColumnWidth(int apiIndex) const noexcept;
    BOOL apiSetText() noexcept;

    long heightInPixels() const noexcept;

    DECLARE_REGISTRY_RESOURCEID(IDR_COLUMNHEADER)

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(ColumnHeader)
    COM_INTERFACE_ENTRY(IColumnHeader)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
    END_COM_MAP()

    // ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

    // IColumnHeader
    public:
    STDMETHOD(get_Tag)(/*[out, retval]*/ VARIANT* pVal);
    STDMETHOD(put_Tag)(/*[in]*/ VARIANT newVal);
    STDMETHOD(get_SubItemIndex)(/*[out, retval]*/ short* pVal);
    STDMETHOD(get_Key)(/*[out, retval]*/ BSTR* pVal);
    STDMETHOD(get_FixedWidth)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(put_FixedWidth)(/*[in]*/ LONG newVal);
    STDMETHOD(get_Alignment)
    (/*[out, retval]*/ ListColumnAlignmentConstants* pVal);
    STDMETHOD(put_Alignment)(/*[in]*/ ListColumnAlignmentConstants newVal);
    STDMETHOD(get_LeftInPixels)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(EnsureVisible)();
    STDMETHOD(get_ColumnContentKind)(/*[out, retval]*/ ColumnContentType* pVal);
    STDMETHOD(put_ColumnContentKind)(/*[in]*/ ColumnContentType newVal);
    STDMETHOD(get_ResizeMode)(/*[out, retval]*/ ListColumnResizeMode* pVal);
    STDMETHOD(put_ResizeMode)(/*[in]*/ ListColumnResizeMode newVal);
    STDMETHOD(get_Width)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(put_Width)(/*[in]*/ LONG newVal);
    STDMETHOD(get_Index)(/*[out, retval]*/ LONG* pVal);
    STDMETHOD(get_Text)(BSTR* pVal);
    STDMETHOD(put_Text)(BSTR* newVal);

    STDMETHOD(get_WidthInPixels)(LONG* pVal);
    STDMETHOD(get_HeightInPixels)(LONG* pVal);
    STDMETHOD(put_HeightInPixels)(LONG newVal);
};

#endif //__COLUMNHEADER_H_
