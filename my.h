// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com// my.h

#pragma once

#include <vector>
#include <map>
#include <string>
#include <algorithm>

#include <sstream>

#include <commctrl.h>
#include <cassert>
#include <strsafe.h> // StringCchCopy
// #include <atlstr.h> // to help intellisense: doesn't exist in VC6
#include "ListItem.h"
class CListItems;
#define MAX_TXT_LEN 512

#if _MSC_VER > VC6_VERSION
#pragma warning(disable : 26812)
#pragma warning(disable : 4130)
#endif

#if _MSC_VER > VC6_VERSION
#pragma warning(disable : 26454)
#else
#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp) ((int)(short)LOWORD(lp))
#endif
#ifndef GET_Y_LPARAM
#define GET_Y_LPARAM(lp) ((int)(short)HIWORD(lp))
#endif
#endif

namespace my {

namespace win32 {

    struct VBPOINTF {
        float x;
        float y;
    };

    __inline vbMouseButtonConstants getVBMouseButton(
        const WPARAM* wParam = NULL) {
        int button = 0;
        if (!wParam) {
            if (GetAsyncKeyState(VK_LBUTTON)) button |= VbLeftButton;
            if (GetAsyncKeyState(VK_RBUTTON)) button |= VbRightButton;
            if (GetAsyncKeyState(VK_MBUTTON)) button |= VbMiddleButton;
        } else {

            if (*wParam & MK_LBUTTON) button |= VbLeftButton;
            if (*wParam & MK_RBUTTON) button |= VbRightButton;
            if (*wParam & MK_MBUTTON) button |= VbMiddleButton;
        }
        return static_cast<vbMouseButtonConstants>(button);
    }

    __inline static vbShiftConstants getVBKeyStates(
        short* ptrShort = 0) noexcept {
        short ivbshift = 0;
        if (GetKeyState(VK_SHIFT) < 0) {
            ivbshift |= vbShiftMask;
        }
        if (GetKeyState(VK_CONTROL) < 0) {
            ivbshift |= vbCtrlMask;
        }
        if (GetKeyState(VK_MENU) < 0) {
            ivbshift |= vbAltMask;
        }
        if (ptrShort) {
            *ptrShort = ivbshift;
        }
        return (vbShiftConstants)ivbshift;
    }

    static __inline POINT ScreenPoint() {
        POINT retval;
        const DWORD dwPos = ::GetMessagePos();
        retval.x = LOWORD(dwPos);
        retval.y = HIWORD(dwPos);
        return retval;
    }

    enum ScaleUnits { pixelUnits, twipsUnits, unknownUnits };

    static __inline ScaleUnits ScaleUnitsFromString(const CString& s) {
        if (s == "Pixel") return pixelUnits;
        if (s == "Twip") return twipsUnits;
        static bool shown_error = false;
        if (!shown_error) {
            ::MessageBox(GetDesktopWindow(),
                _T("Only ScaleModes supported are 'vbTwips' or ")
                _T("'vbPixel'.\nPlease change the properties of the ")
                _T("container\n\nProperties such as columnheader widths will ")
                _T("not work correctly unless you do."),
                _T("ListControl Message"), MB_OK | MB_ICONWARNING);

            shown_error = true;
        }
        return unknownUnits;
    }

    // direction can also be LOGPIXELSX
    static __inline int twipsPerPixel(unsigned int direction = LOGPIXELSY) {
        HWND hWnd = ::GetDesktopWindow();
        HDC hDC = ::GetDC(hWnd);
        const int logPix = ::GetDeviceCaps(hDC, direction);
        ::ReleaseDC(hWnd, hDC);
        return 1440 / logPix;
    }

    static __inline POINT VBToPt(float vbx, float vby, ScaleUnits scale) {
        POINT ret;
        ret.x = (LONG)vbx;
        ret.y = (LONG)vby;
        if (scale == twipsUnits) {

            ret.x = (LONG)((float)vbx / twipsPerPixel(LOGPIXELSX));
            ret.y = (LONG)((float)vby / twipsPerPixel(LOGPIXELSY));
        }
        return ret;
    }

    static __inline VBPOINTF ptToVB(
        const POINT& pt, const ScaleUnits units) noexcept {
        VBPOINTF vbPoint;

        if (units == twipsUnits) {
            vbPoint.x = (float)pt.x * (float)twipsPerPixel(LOGPIXELSX);
            vbPoint.y = (float)pt.y * (float)twipsPerPixel(LOGPIXELSY);
        } else {
            vbPoint.x = (float)pt.x;
            vbPoint.y = (float)pt.y;
        }
        return vbPoint;
    }

    __inline my::win32::VBPOINTF getXYFromLParam(
        const LPARAM lParam, ScaleUnits scale) {
        POINT pt;
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
        my::win32::VBPOINTF point = my::win32::ptToVB(pt, scale);
        return point;
    }

    struct InputInfo {

        InputInfo(const ScaleUnits Units, const LPARAM* lp = NULL,
            const POINT* pt = NULL, const WPARAM* wp = NULL)
            : units(Units) {

            memset(&point, 0, sizeof(point));
            button = getVBMouseButton(wp);
            shift = getVBKeyStates();
            if (pt) {
                point = ptToVB(*pt, units);
            } else if (lp) {
                point = getXYFromLParam(*lp, units);
            } else {
                ASSERT("InputInfo: either LPARAM or POINT should be set. They "
                       "are both null"
                    == 0);
            }
        }
        vbMouseButtonConstants button;
        vbShiftConstants shift;
        VBPOINTF point;
        ScaleUnits units;
    };

    static __inline CString getClassName(HWND hWnd) {
        ASSERT(IsWindow(hWnd));

        CString ret(static_cast<char>(0), 512);
        ASSERT(ret.GetLength() == 512);
        const int iret = ::GetClassName(hWnd, ret.GetBuffer(512), 512);
        ret.ReleaseBuffer();
        const int len = ret.GetLength();
        assert(len == iret);
        return ret;
    }

    __inline void debugShowClassName(HWND hWnd) {
        CString className = my::win32::getClassName(hWnd);
        TRACE(_T("ListControl: Parent's Window ClassName %p is: %s\n"), hWnd,
            className.GetBuffer(256));
    }

} // namespace win32

static __inline std::wstring build_wstring(
    const std::wstring& s1, const int i, const std::wstring& s2 = L"") {
    std::wstringstream wss;
    wss << s1 << i << s2;
    return std::wstring(wss.str());
}

static __inline BOOL VBToBool(const VARIANT_BOOL vb) {
    return (vb == VARIANT_TRUE) ? TRUE : FALSE;
}
static __inline VARIANT_BOOL BoolToVB(const BOOL b) {
    if (b) return VARIANT_TRUE;
    return VARIANT_FALSE;
}

template <typename T> static __inline T cpp_max(const T& a, const T& b) {
    return (a < b) ? b : a;
}

static bool inline InlineIsEqualGUID(const GUID a, const GUID b) { //-V801
    return !memcmp((void*)&a, (void*)&b, sizeof(GUID));
}
//////////////////////////////////////////////////////
// Collections, as VC6 does not support <atlcoll.h> //
//////////////////////////////////////////////////////

// NOTE: com_map does NOT AddRef() or Release(); it assumes you
// hang on to the pointers elsewhere.
template <typename K = CString, typename V = IUnknown> struct com_map {
    typedef std::map<K, V> map_t;
    typedef typename map_t::iterator map_it;
    typedef typename map_t::key_type key_type;
    typedef typename map_t::value_type value_type;

    map_t m_map;
    virtual ~com_map() {}

    map_it begin() { return m_map.begin(); }

    map_it end() { return m_map.end(); }

    __inline void clear() { m_map.clear(); }
    __inline bool add(const K& key, const V& value) {
        map_it it = m_map.find(key);
        if (it != m_map.end()) return false;
        m_map[key] = value;
        return true;
    }

    // removes from both map and vector storage, bool version
    __inline bool remove(const K& key) {
        map_it it = m_map.find(key);
        if (it == m_map.end()) return false;
        m_map.erase(it);
        return true;
    }

    // removes from both map and vector, returns map iterator (or map iterator
    // end() if not found)
    __inline map_it remove_it(const K& key) {
        map_it it = m_map.find(key);
        if (it == m_map.end()) return it;
        return m_map.erase(it);
    }

    __inline size_t size() const { return m_map.size(); }
    __inline int isize() const { return static_cast<int>(size()); }
};

// com_vector: AddRefs() and Releases() the interfaces it holds.
template <typename T>
struct com_vector { // NOLINT(cppcoreguidelines-special-member-functions)

    typedef std::vector<T> vec_t;
    typedef typename vec_t::iterator vec_it;

    virtual ~com_vector() {
        for (typename vec_t::iterator it = m_vec.begin(); it < m_vec.end();
             ++it) {
            (*it)->Release();
        }
    }
    __inline size_t capacity() const { return m_vec.capacity(); }
    __inline int icapacity() const { return static_cast<int>(capacity()); }

    __inline void push_back(typename vec_t::value_type c) {
        c->AddRef();
        m_vec.push_back(c);
    }

    __inline void clear() {
        for (vec_it it = m_vec.begin(); it < m_vec.end(); ++it) {
            (*it)->Release();
        }
        m_vec.clear();
    }

    bool delete_at(size_t index) {
        if (index >= m_vec.size()) return false;
        const size_t old_size = m_vec.size();

        vec_it where = m_vec.begin() + index;
        T ptr = *where;
        ptr->Release();
        vec_it it = std::remove(m_vec.begin(), m_vec.end(), ptr);
        m_vec.erase(it, m_vec.end());
        ASSERT(m_vec.size() == old_size - 1);
        return m_vec.size() == old_size - 1;
    }

    // remove item from underlying vector. Release() is called on the item.
    bool removeItem(T item) {
        vec_it it = std::find(m_vec.begin(), m_vec.end(), item);

        if (it == m_vec.end()) return false;
        size_t cnt = m_vec.size();
        (*it)->Release();
        vec_it it1 = std::remove(m_vec.begin(), m_vec.end(), *it);
        m_vec.erase(it1, m_vec.end());
        return m_vec.size() == cnt - 1;
    }

    __inline const T& operator[](size_t idx) const {
#ifdef DEBUG
        return m_vec.at(idx);
#else
        return m_vec[idx];
#endif
    }

    __inline T& operator[](size_t idx) {
#ifdef DEBUG
        return m_vec.at(idx);
#else
        return m_vec[idx];
#endif
    }

    __inline const T& at(size_t idx) const {
        return m_vec.at(idx);
    }

    __inline T& at(size_t idx) {
        return m_vec.at(idx);
    }

    __inline void reserve(size_t how_many) {
        m_vec.reserve(how_many);
    }

    __inline size_t size() const {
        return m_vec.size();
    }

    __inline int isize() const {
        return static_cast<int>(m_vec.size());
    }

    __inline typename vec_t::value_type operator[](int i) {
        return m_vec[i];
    }

    size_t vec_size() const { //-V524
        return m_vec.size();
    }

    const std::vector<T>& vec_data() const {
        return m_vec;
    }
    // std::vector<T>& vec_data() {
    //     return m_vec;
    // }
    const std::vector<T> vec_data_copy() const {
        return m_vec;
    }

    protected:
    vec_t m_vec; // not allowed to directly access as it may defeat the addref
    // and release stuff
};

template <typename T> struct interface_collection : com_vector<T*> {
    typedef com_map<CString, T*> map_type;
    typedef typename map_type::map_it map_iterator;
    typedef com_vector<T*> base_t;

    map_type m_map;

    // ReSharper disable once CppEnforceOverridingDestructorStyle
    virtual ~interface_collection() {}

    std::vector<T*>& getVector() { return this->m_vec; }

    // add with key
    __inline bool add(const CString& key, T* p) {
        const bool ret = m_map.add(key, p);
        if (ret) {
            this->push_back(p);
        }
        return ret;
    }

    // add without key
    __inline bool add(T* p) {
        this->push_back(p);
        return true;
    }

    // remove both from the keyed map and from the vector
    __inline bool remove(const CString& key) {
        size_t sz_map = m_map.size();
        size_t sz_vec = this->m_vec.size();
        map_iterator mit = m_map.remove_it(key);
        if (mit != m_map.end()) {
            this->removeItem(mit->second);
        }

        return ((m_map.size() == sz_map - 1) || (m_vec.size() == sz_vec - 1));
    }

    // removes from vector, and conditionally map if the key exists and is found
    __inline bool remove(T* value, const CString* key) {
        ASSERT(value || key);
        if (key) {
            return remove(*key);
        } else {
            if (value) {
                return base_t::removeItem(value);
            }
        }
        return false;
    }

    // ReSharper disable once CppHidingFunction
    __inline void clear() {
        m_map.clear();
        base_t::clear();
    }

    __inline const T* const operator[](size_t idx) const {
#ifdef DEBUG
        return base_t::at(idx);
#else
        return base_t::operator[](idx);
#endif
    }

    __inline T* operator[](size_t idx) {
#ifdef DEBUG
        return base_t::at(idx);
#else
        return base_t::operator[](idx);
#endif
    }

    // is it in the map?
    __inline T* find(const CString& key) {
        map_iterator it = m_map.m_map.find(key);
        if (it == m_map.m_map.end()) {
            return NULL;
        }
        return it->second;
    }

    __inline size_t size() const {
        assert(m_map.size()
            <= this->vec_size()); // there could be more in the vector if
        // things were added without keys.
        return (my::cpp_max)(this->vec_size(), m_map.size());
    }

    __inline int isize() const {
        return static_cast<int>(size());
    }
};

static inline HRESULT atlReportError(
    const CLSID& clsid, const IID& iid, const std::wstring& what, HRESULT hr) {
    if (hr == E_FAIL) {
        HRESULT mhr = HRESULT_FROM_WIN32(::GetLastError());
        const std::wstring w = my::build_wstring(what, (int)::GetLastError());
        return AtlReportError(
            clsid, w.c_str(), iid, SUCCEEDED(mhr) ? E_FAIL : mhr);
    } else {
        return AtlReportError(clsid, what.c_str(), iid, hr);
    }
}

static __inline HRESULT VariantToInt(
    VARIANT* v, int& i, bool isOptional = false) {

    if (isOptional) {
        if (!v) return S_FALSE;
        if (v->vt == VT_ERROR || v->vt == VT_EMPTY) return S_FALSE;
    }

    const long vt = v->vt;
    if (!v) {
        return DISP_E_TYPEMISMATCH;
    }

    if (v->vt != VT_ERROR) {
        if (vt & VT_I2) {
            if (vt & VT_BYREF) {
                i = *v->piVal;
                return S_OK;
            } else {
                i = v->iVal;
                return S_OK;
            }
        }

        if (vt & VT_I4) {
            if (vt & VT_BYREF) {
                i = *v->plVal;
                return S_OK;
            } else {
                i = v->lVal;
                return S_OK;
            }
        }
    }

    return DISP_E_TYPEMISMATCH;
}

static __inline HRESULT VariantToCString(
    VARIANT* Text, CString& txt, bool isOptional = false) {

    if (isOptional) {
        if (!Text) return S_FALSE;
        if (Text->vt == VT_ERROR || Text->vt == VT_EMPTY) return S_FALSE;
    }

    long vt = Text->vt;
    if (vt != VT_ERROR) {
        if (vt == (VT_BSTR | VT_BYREF)) {
            txt = *Text->pbstrVal;
        } else {
            if (vt == VT_BSTR) {
                txt = Text->bstrVal;
            } else {
                return DISP_E_TYPEMISMATCH;
            }
        }
    } else {
        return DISP_E_TYPEMISMATCH;
    }

    return S_OK;
}

template <typename INT_T, typename COLLECTION>
static __inline HRESULT bounds_check(const INT_T index, const COLLECTION& c) {

    INT_T isize = static_cast<INT_T>(c.size());
    if (index < 0 || index >= isize) return DISP_E_BADINDEX;
    return S_OK;
}

static __inline void lvHeaderHitTest(const HWND& hWndHdr, POINT& pointHdr,
    HDHITTESTINFO& hdhti, bool use_screen_coords = true) {

    if (use_screen_coords) ::ScreenToClient(hWndHdr, &pointHdr);

    memset(&hdhti, 0, sizeof(hdhti));
    hdhti.pt = pointHdr;
    ::SendMessage(hWndHdr, HDM_HITTEST, static_cast<WPARAM>(0),
        reinterpret_cast<LPARAM>((HD_HITTESTINFO FAR*)&hdhti));
}

static __inline int lvGetItemCount(HWND hWnd) {
    return ListView_GetItemCount(hWnd);
}

typedef interface_collection<IDispatch> idispatch_collection;

static __inline LRESULT lvHandleDispInfo(
    WPARAM wp, LPARAM lp, const idispatch_collection& items, NMHDR* pnmhdr) {

    (void)lp;
    (void)wp;

    NMLVDISPINFO* plvdi = (NMLVDISPINFO*)pnmhdr;

    if (-1 == plvdi->item.iItem) {
        OutputDebugString(TEXT("LVOWNER: Request for -1 item?\n"));
        ASSERT(0);
    }
    const int count = static_cast<int>(items.vec_size());
    if (plvdi->item.iItem >= count) {
        // TRACE("WARNING: Virtual listview is asking for more data than we "
        //       "have\nListItem count (in actual win32 listview): %ld\nItems
        //       (in " "vector) count: %ld\n\n",
        //    (long)lvGetItemCount(pnmhdr->hwndFrom), (long)items.vec_size());
        return 0;
    }
    const CListItem* item
        = static_cast<const CListItem*>(items[plvdi->item.iItem]);

    if (plvdi->item.mask & LVIF_STATE) {
        // Fill in the state information.
        // plvdi->item.state |= rndItem.state;
    }

    if (plvdi->item.mask & LVIF_IMAGE) {
        // Fill in the image information.
        // plvdi->item.iImage = rndItem.iIcon;
    }

    if (plvdi->item.mask & LVIF_TEXT) {
        // Fill in the text information.
        switch (plvdi->item.iSubItem) {
                // https://docs.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-lvitema
                // NOTE: iSubItem is ONE-BASED!! If its zero: if this structure
                // refers to an item rather than a subitem. (MSDN)
            case 1:
                // we don't need to copy the text here as listitems should
                // outlive the display. FIXME VC6
                plvdi->item.pszText = item->m_listItemInfo.toString();

                break;

            default:
                plvdi->item.pszText = item->m_listItemInfo.toString();
                break;
        }
    }

    return TRUE;
}

static __inline int lvGetColumnWidth(HWND hWnd, int index) {
    return ListView_GetColumnWidth(hWnd, index);
}

static __inline DWORD lvDefaultStyle(BOOL isVirtual = FALSE) {
    DWORD style = WS_VISIBLE | WS_CHILD | LVS_REPORT | WS_OVERLAPPED
        | LVS_SHOWSELALWAYS | WS_TABSTOP | LVS_SHAREIMAGELISTS | WS_VSCROLL;

    if (isVirtual) {
        style |= LVS_OWNERDATA;
    }
    return style;
}

// This for SHIFT + SELECT WITH MOUSE only in virtual mode
// You also get the states for individual items, so really can ignore this if
// you want.
// static __inline void lvHandleStateChange(LPNMLVODSTATECHANGE state) {
// TRACE(_T("state from item %d to item %d, oldstate=%d, newstate=%d.\n"),
// state->iFrom, state->iTo, 	state->uOldState, state->uNewState);
//}

/*/
https://docs.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-nmlistview
typedef struct tagNMLISTVIEW {
NMHDR  hdr;			//	NMHDR structure that contains information about this
notification message.
https://docs.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-nmhdr
int    iItem;		//	Identifies the list-view item, or -1 if not used.
int    iSubItem;	//	Identifies the subitem, or zero if none.
UINT   uNewState;	//	LVIS_ACTIVATING	Not currently supported.
LVIS_CUT The item is marked for a cut-and-paste
operation. LVIS_DROPHILITED	The item is highlighted as a drag-and-drop target.
LVIS_FOCUSED		The item has the focus, so it is
surrounded by a standard focus rectangle. Although more than one item may be
selected, only one item can have the focus. LVIS_OVERLAYMASK	The item's
overlay image index is retrieved by a mask. LVIS_SELECTED		The item is
selected. The appearance of a selected item depends on whether it has the focus
and also on the system colors used for selection. LVIS_STATEIMAGEMASK	The
item's state image index is retrieved by a mask. UINT   uOldState; UINT
uChanged;		Set of flags that indicate the item attributes that have
changed. This member is zero for notifications that do not use it. Otherwise, it
can have the same values as the mask member of the LVITEM structure.
https://docs.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-lvitema
LVIF_COLFMT			Windows Vista and later. The
piColFmt member is valid or must be set. If this flag is used, the cColumns
member is valid or must be set. LVIF_COLUMNS		The cColumns member is valid
or must be set. LVIF_DI_SETITEM		The operating system should store the
requested list item information and not ask for it again. This flag is used only
with the LVN_GETDISPINFO notification code. LVIF_GROUPID		The iGroupId
member is valid or must be set. If this flag is not set when an LVM_INSERTITEM
message is sent, the value of iGroupId is assumed to be I_GROUPIDCALLBACK.
LVIF_IMAGE			The iImage member is valid or must
be set. LVIF_INDENT			The iIndent member is valid or must be set.
LVIF_NORECOMPUTE	The control will not generate
LVN_GETDISPINFO to retrieve text information if it receives an LVM_GETITEM
message. Instead, the pszText member will contain LPSTR_TEXTCALLBACK. LVIF_PARAM
The lParam member is valid or must be set.
LVIF_STATE			The state member is valid or must be
set. LVIF_TEXT			The pszText member is valid or must be set.

  POINT  ptAction;		POINT structure that indicates the location at which the
  event occurred. This member is undefined for notification messages that do not
  use it. LPARAM lParam;			Application-defined value of the item. This
  member is undefined for notification messages that do not use it. }
NMLISTVIEW, *LPNMLISTVIEW;
/*/
template <typename T>
static __inline LRESULT lvHandleStateChange(T* pControl, HWND hWndLv,
    LPNMLISTVIEW state, const idispatch_collection& items) {

    (void)hWndLv;
    TRACE(_T("item %d, state from oldstate:%d to newstate:%d, and uChanged: ")
          _T("%d\n"),
        state->iItem, state->uOldState, state->uNewState, state->uChanged);

    int index = state->iItem;
    const int size = items.isize();

    // quirks in virtual mode: I always see this -1 instead of the correct index
    // when something is UNselected. I am hoping it is the former current item.
    // I never seem to get this if control + click is used.
    if (index == -1) {
        index = pControl->getLastSelItemIndex();
        if (index < 0 || index >= size) {
            return FALSE;
        }
    }

    if (index >= 0 && index < size) {

        if ((state->uChanged & LVIF_STATE) && (state->uOldState & LVIS_SELECTED)
            && !(state->uNewState & LVIS_SELECTED)) {
            CListItem* pitem = (CListItem*)items.at(index);
            pitem->setSelected(FALSE);
        }

        if ((state->uChanged & LVIF_STATE)
            && (state->uNewState & LVIS_SELECTED)) {
            CListItem* pitem = (CListItem*)items.at(index);

            pitem->setSelected(TRUE);
        }
    }

    return 0;
}

struct ClickData {
    ClickData(int c, int i, int s) : code(c), item(i), subitem(s) {}
    int code;
    int item;
    int subitem;
};

template <typename T>
static __inline void lvHandleClick(
    T* pControl, NMHDR* pnmhdr, const idispatch_collection&) {

    const int code = pnmhdr->code;
    NMITEMACTIVATE* itemClicked = (NMITEMACTIVATE*)pnmhdr;
    LVHITTESTINFO myinfo = {0};
    memset(&myinfo, 0, sizeof(myinfo));

    myinfo.pt = itemClicked->ptAction;
    HWND lvhWnd = pControl->lvhWnd();
    int itemNumber
        = SendMessage(lvhWnd, LVM_SUBITEMHITTEST, 0, (LPARAM)&myinfo);
    ASSERT(itemNumber == myinfo.iItem);
    pControl->handleClick(code, myinfo);
}

template <typename T>
static __inline void lvHandleMouse(T* pControl, NMHDR* phdr) {}

template <typename T>
static __inline BOOL lvHandleKeyPressFromPreTranslate(T* pControl, LPMSG pMsg) {
    // handle 'KeyPress' for VBClient:
    WPARAM wParam = pMsg->wParam;
    WORD wpHigh = HIWORD(wParam);
    WPARAM wpOrig = pMsg->wParam;
    LPARAM lParam = pMsg->lParam;
    WORD vkCode = LOWORD(wParam);
    WORD keyFlags = HIWORD(lParam);
    WORD scanCode = LOBYTE(keyFlags); // scan code
    BOOL isExtendedKey = (keyFlags & KF_EXTENDED)
        == KF_EXTENDED; // extended-key flag, 1 if scancode has 0xE0 prefix

    if (isExtendedKey) scanCode = MAKEWORD(scanCode, 0xE0);
    BYTE keystate[256];
    memset(keystate, 0, 256);
    char chars[2];
    memset(chars, 0, 2);
    int nchars = ToAscii(vkCode, 0, keystate, (WORD*)&chars, 0);
    ASSERT(nchars >= 1);
    if (nchars == 1) {
        short keyascii = *((short*)chars);
        const short origKey = keyascii;
        short virtKey = VkKeyScanA((CHAR)keyascii);
        WPARAM newWParam = MAKEWPARAM(virtKey, wpHigh);
        if (newWParam
            == wpOrig) { // check symmetry: it's not always the case, and we
                         // are only interested in ASCII characters, anyway!
            pControl->Fire_KeyPress(&keyascii);
            if (origKey != keyascii) {
                short newvirtKey = VkKeyScanA((CHAR)keyascii);
                if (keyascii == 0) {
                    newvirtKey = 0;
                }
                newWParam = MAKEWPARAM(newvirtKey, wpHigh);
                pMsg->wParam = newWParam;
                if (keyascii == 0) {
                    return TRUE;
                }
            }
        }
    }

    return FALSE;
}

template <typename T>
static __inline LRESULT lvHandleNotify(WPARAM wp, LPARAM lp, T* pControl,
    BOOL isVirtual, const idispatch_collection& items, NMHDR* pnmhdr,
    BOOL& bHandled) {

    LRESULT lrt = 0;
    bHandled = FALSE;
    switch (pnmhdr->code) {
        case LVN_GETDISPINFO: {
            if (isVirtual) return lvHandleDispInfo(wp, lp, items, pnmhdr);
            break;
        }

        case LVN_ITEMCHANGED: {
            if (GetKeyState(VK_SHIFT) < 0) {
                break; // LVN_ODSTATECHANGED will handle it.
            }
            return lvHandleStateChange(
                pControl, pnmhdr->hwndFrom, (LPNMLISTVIEW)pnmhdr, items);
            break;
        }

        case LVN_KEYDOWN: {
            LPNMLVKEYDOWN pnkd = (LPNMLVKEYDOWN)lp;
            vbShiftConstants shift = my::win32::getVBKeyStates();
            pControl->Fire_KeyDown((SHORT)pnkd->wVKey, (SHORT)shift);
            break;
        }
            /*/
            case LVN_KEYUP: {
                LPNMLVKEYDOWN pnkd = (LPNMLVKEYDOWN)lp;
                vbShiftConstants shift = my::win32::getVBKeyStates();
                // pControl->Fire_KeyUp((SHORT)pnkd->wVKey, (SHORT)shift);
                break;
            }
            /*/

        case LVN_ODSTATECHANGED: {
            LPNMLVODSTATECHANGE p = (LPNMLVODSTATECHANGE)pnmhdr;
            // range selected

            int n = p->iTo - p->iFrom;
            if (n > 0 && n < items.isize()) {

                for (int i = 0; i < n; i++) {

                    if ((p->uOldState & LVIS_SELECTED)
                        && !(p->uNewState & LVIS_SELECTED)) {
                        CListItem* pitem = (CListItem*)items.at(i);
                        pitem->setSelected(FALSE);
                    }

                    if ((p->uNewState & LVIS_SELECTED)) {
                        CListItem* pitem = (CListItem*)items.at(i);
                        pitem->setSelected(TRUE);
                    }
                }
            }
            break;
        }

        case LVN_ODCACHEHINT: {
            // NMLVCACHEHINT* pcachehint = (NMLVCACHEHINT*)pnmhdr;

            // Load the cache with the recommended range.
            // PrepCache( pcachehint->iFrom, pcachehint->iTo );
            break;
        }

        case LVN_ODFINDITEM: {
            // LPNMLVFINDITEM pnmfi = NULL;

            // pnmfi = (LPNMLVFINDITEM)pnmhdr;

            // Call a user-defined function that finds the index according to
            // LVFINDINFO (which is embedded in the LPNMLVFINDITEM structure).
            // If nothing is found, then set the return value to -1.

            break;
        }

        case NM_CLICK:
        case NM_RCLICK:
        case NM_DBLCLK: {

            lvHandleClick(pControl, pnmhdr, items);

            break;
        }

        default: break;

    } // End Switch block.

    return (lrt);
}

static __inline bool isListviewHeader(HWND hWnd) {
    static const char* HEADER_CLASSNAME = "SysHeader32";
    CString className = my::win32::getClassName(hWnd);
    ASSERT(className == HEADER_CLASSNAME);
    if (className == HEADER_CLASSNAME) return true;
    return false;
}

static __inline bool isListView(HWND hWnd) {
    static TCHAR buf[255];
    ::GetClassName(hWnd, buf, 255);
    const CString classname = buf;
    if (classname != "ATL:SysListView32") {
        assert("Not a listview window" == 0);
        return false;
    }
    return true;
}
static __inline BOOL lvSetItemCount(HWND myhWnd, size_t how_many) {
    // ASSERT(how_many >= 20000);
    ASSERT(isListView(myhWnd));
    const BOOL b = ListView_SetItemCount(myhWnd, how_many);
    return b;
}
static __inline int lvGetColumnCount(HWND hWnd) {
    const HWND hWndHdr = reinterpret_cast<const HWND>(
        ::SendMessage(hWnd, LVM_GETHEADER, 0, 0));
    const int count
        = static_cast<int>(::SendMessage(hWndHdr, HDM_GETITEMCOUNT, 0, 0L));
    return count;
}

// returns zero on success
static __inline int lvSetFont(HWND hWnd, IFontDisp* newFont) {

    assert(hWnd && newFont);

    HFONT hfont = 0;
    if (newFont) {
        CComQIPtr<IFont, &IID_IFont> pfont(newFont);
        if (pfont) {
            assert(newFont);
            HRESULT hr = pfont->get_hFont(&hfont);
            assert(SUCCEEDED(hr));
            if (SUCCEEDED(hr)) {
                long l = ::SendMessage(hWnd, WM_SETFONT, (WPARAM)hfont, 1);
                if (l != 0) {
                    assert(isListView(hWnd));
                    l = ::SendMessage(hWnd, WM_SETFONT,
                        static_cast<WPARAM>(DEFAULT_GUI_FONT), 1);
                }
                return l;
            }
        }
    }
    return -1;
}

// returns index added
static __inline int lvInsertTextColumn(HWND hWnd, const CString& txt,
    int width = 100, int fmt_flags = lvcf_text | lvcf_width, int index = -1) {

    ASSERT(isListView(hWnd));
    LVCOLUMN lvc;
    memset(&lvc, 0, sizeof(lvc));
    lvc.mask = fmt_flags;
    lvc.fmt = LVCFMT_LEFT;
    if (index == -1) index = lvGetColumnCount(hWnd);

    lvc.iSubItem = 0;
    lvc.pszText = (TCHAR*)static_cast<LPCTSTR>(txt);
    lvc.cx = width; // width of column in pixels
    int ret = ListView_InsertColumn(hWnd, index, &lvc);
    return ret;
}

/*/
Public Function ListView_SetSelectedItem(hWndLv As Long, i As Long,
iSelect As Long) As Boolean Dim Flags As LVITEM_state Dim mask
As LVITEM_state mask = LVIS_FOCUSED Or LVIS_SELECTED If iSelect
Then Flags = LVIS_FOCUSED Or LVIS_SELECTED Else Flags = 0 End If

  ListView_SetSelectedItem = ListView_SetItemState(
  hWndLv, i, Flags, mask) End Function
/*/

static __inline int lvSetSelectedItem(HWND hWnd, long index) {

    DWORD data = LVIS_SELECTED;
    DWORD mask = LVIS_SELECTED;
    ListView_SetItemState(hWnd, index, data, mask);
    return TRUE;
}

struct lv_insert_info {
    lv_insert_info(int idx = -1, BOOL bVirtual = FALSE, int itemCount = -1)
        : index(idx), isVirtual(bVirtual), viewItemCount(itemCount) {}
    int index;
    BOOL isVirtual;
    int viewItemCount;
};

// NOTE: YOU are responsible for setting the number of items
// (my::lvSetItemCount) because only you, as the caller, know how many items you
// actually want to insert.

static __inline void check_index_performance(const lv_insert_info& info) {
    if (info.index % 1000 == 0) {
        if (info.isVirtual) {
            if (info.index >= info.viewItemCount) {
                ASSERT(
                    "Internal Error: When inserting into a virtual listview."
                    "\nYou should always be inserting *inside* the view.\n"
                    "You forgot to call lvSetItemCount to the correct value, "
                    "and this spoils performance.\n"
                    == 0);
            }

        } else {
            if (info.viewItemCount <= 0) {
                ASSERT(
                    "Internal error: always send the view count when inserting "
                    "(adding) a listitem."
                    == 0);
            }
            TRACE(_T("Added %ld items so far ...\n"), info.viewItemCount);
            if (info.index
                < info.viewItemCount - 1) { // -1 because lvInsertItem is not
                // called until after we were called
                TRACE(_T(
                    "ATLListView: Did you mean to insert instead of append? \
				This will be super-slow!\n"));
            }
        }
    }
}

static __inline int lvInsertItem(
    HWND hWnd, const CString& txt, lv_insert_info& info) {
    LVITEM lvI; // using = {} chokes VC6
    memset(&lvI, 0, sizeof(lvI));
    //  Initialize LVITEM members that are common to all items.  //
    //  NOLINT(performance-no-int-to-ptr)
    if (!info.isVirtual) {
        lvI.cchTextMax = txt.GetLength();
        lvI.pszText = const_cast<TCHAR*>(static_cast<LPCTSTR>(txt));
    } else {
        lvI.pszText = LPSTR_TEXTCALLBACK;
    }
    // LPSTR_TEXTCALLBACK; // Sends an LVN_GETDISPINFO message.

    lvI.mask = LVIF_TEXT;
    lvI.stateMask = 0;
    lvI.iSubItem = 0;
    lvI.state = 0;
    if (info.index == -1) {
        TRACE(
            _T("LISTVIEW WARNING. THIS IS A SLOW WAY TO ADD LISTITEMS. IF YOU \
		KNOW THE INDEX OF THIS ITEM, YOU SHOULD SEND IT.\n\n"));
        info.index = lvGetItemCount(hWnd);
    }
    lvI.iItem = info.index;
    lvI.iImage = info.index;

    check_index_performance(info);
    ASSERT(!info.isVirtual); // DO NOT insertItem on a virtual listview; insert
    // it into your datacollection only, and use
    // lvSetItemCount()
    const int ret = ListView_InsertItem(hWnd, &lvI);
    assert(ret >= 0);
    if (ret == 0) {
        const int cnt = lvGetItemCount(hWnd);
        if (cnt == 0) {
            assert(isListView(hWnd));
        }
    }
    return ret;
}

} // namespace my
