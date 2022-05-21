// This is an independent project of an individual developer. Dear PVS-Studio,
// please check it.

// PVS-Studio Static Code Analyzer for C, C++, C#, and Java:
// http://www.viva64.com
//

// ATLListView.cpp : Implementation of DLL Exports.

// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL,
//      run nmake -f ATLListViewps.mk in the project directory.

#include "stdAfx.h"
#include "resource.h"
#include <initguid.h>
#include "ATLListView.h"

#include "ATLListView_i.c"
#include "ListControl.h"
#include "ColumnHeaders.h"
// #include "lv.h"
#include "ColumnHeader.h"
#include "ListItem.h"
#include "ListSubItem.h"
#include "ListItems.h"
#include "InterfaceCollection.h"
#include "RedrawLocker.h"
#include "RedrawLock.h"
#include "Stopwatch.h"
#include "SelItemCollection.h"
#include "ListSubItems.h"

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_ListControl, CListControl)
OBJECT_ENTRY(CLSID_ColumnHeaders, CColumnHeaders)
// OBJECT_ENTRY(CLSID_lv, Clv)
OBJECT_ENTRY(CLSID_ColumnHeader, ColumnHeader)
OBJECT_ENTRY(CLSID_ListItem, CListItem)
OBJECT_ENTRY(CLSID_ListSubItem, CListSubItem)
// OBJECT_ENTRY(CLSID_ListItems, CListItems)
OBJECT_ENTRY(CLSID_ListItems, CListItems)
OBJECT_ENTRY(CLSID_InterfaceCollection, CInterfaceCollection)
OBJECT_ENTRY(CLSID_RedrawLock, CRedrawLock)

OBJECT_ENTRY(CLSID_SelItemCollection, CSelItemCollection)
OBJECT_ENTRY(CLSID_ListSubItems, CListSubItems)
END_OBJECT_MAP()

class CATLListViewApp : public CWinApp {
    public:
    // Overrides
    // ClassWizard generated virtual function overrides
    //{{AFX_VIRTUAL(CATLListViewApp)
    public:
    virtual BOOL InitInstance();
    virtual int ExitInstance();
    virtual BOOL PreTranslateMessage(MSG* pMsg);
    //}}AFX_VIRTUAL

    //{{AFX_MSG(CATLListViewApp)
    // NOTE - the ClassWizard will add and remove member functions here.
    //    DO NOT EDIT what you see in these blocks of generated code !
    //}}AFX_MSG
    DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CATLListViewApp, CWinApp)
//{{AFX_MSG_MAP(CATLListViewApp)
// NOTE - the ClassWizard will add and remove mapping macros here.
//    DO NOT EDIT what you see in these blocks of generated code!
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CATLListViewApp theApp;

BOOL CATLListViewApp::InitInstance() {

    _Module.Init(ObjectMap, m_hInstance, &LIBID_ATLLISTVIEWLib);
    return CWinApp::InitInstance();
}

int CATLListViewApp::ExitInstance() {
    _Module.Term();
    return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void) {
    AFX_MANAGE_STATE(AfxGetStaticModuleState());
    return (AfxDllCanUnloadNow() == S_OK && _Module.GetLockCount() == 0)
        ? S_OK
        : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void) {
    // registers object, typelib and all interfaces in typelib
    return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void) {
    return _Module.UnregisterServer(TRUE);
}

BOOL CATLListViewApp::PreTranslateMessage(MSG* pMsg) {

    return CWinApp::PreTranslateMessage(pMsg);
}
