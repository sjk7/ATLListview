

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Sat May 14 12:06:40 2022
 */
/* Compiler settings for ATLListView.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __ATLListView_h__
#define __ATLListView_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IInterfaceCollection_FWD_DEFINED__
#define __IInterfaceCollection_FWD_DEFINED__
typedef interface IInterfaceCollection IInterfaceCollection;

#endif 	/* __IInterfaceCollection_FWD_DEFINED__ */


#ifndef __IListControl_FWD_DEFINED__
#define __IListControl_FWD_DEFINED__
typedef interface IListControl IListControl;

#endif 	/* __IListControl_FWD_DEFINED__ */


#ifndef ___IListControlEvents_FWD_DEFINED__
#define ___IListControlEvents_FWD_DEFINED__
typedef interface _IListControlEvents _IListControlEvents;

#endif 	/* ___IListControlEvents_FWD_DEFINED__ */


#ifndef __IColumnHeaders_FWD_DEFINED__
#define __IColumnHeaders_FWD_DEFINED__
typedef interface IColumnHeaders IColumnHeaders;

#endif 	/* __IColumnHeaders_FWD_DEFINED__ */


#ifndef __ListControl_FWD_DEFINED__
#define __ListControl_FWD_DEFINED__

#ifdef __cplusplus
typedef class ListControl ListControl;
#else
typedef struct ListControl ListControl;
#endif /* __cplusplus */

#endif 	/* __ListControl_FWD_DEFINED__ */


#ifndef __ColumnHeaders_FWD_DEFINED__
#define __ColumnHeaders_FWD_DEFINED__

#ifdef __cplusplus
typedef class ColumnHeaders ColumnHeaders;
#else
typedef struct ColumnHeaders ColumnHeaders;
#endif /* __cplusplus */

#endif 	/* __ColumnHeaders_FWD_DEFINED__ */


#ifndef __IColumnHeader_FWD_DEFINED__
#define __IColumnHeader_FWD_DEFINED__
typedef interface IColumnHeader IColumnHeader;

#endif 	/* __IColumnHeader_FWD_DEFINED__ */


#ifndef __IListSubItem_FWD_DEFINED__
#define __IListSubItem_FWD_DEFINED__
typedef interface IListSubItem IListSubItem;

#endif 	/* __IListSubItem_FWD_DEFINED__ */


#ifndef __IListItem_FWD_DEFINED__
#define __IListItem_FWD_DEFINED__
typedef interface IListItem IListItem;

#endif 	/* __IListItem_FWD_DEFINED__ */


#ifndef __IListItems_FWD_DEFINED__
#define __IListItems_FWD_DEFINED__
typedef interface IListItems IListItems;

#endif 	/* __IListItems_FWD_DEFINED__ */


#ifndef __IRedrawLock_FWD_DEFINED__
#define __IRedrawLock_FWD_DEFINED__
typedef interface IRedrawLock IRedrawLock;

#endif 	/* __IRedrawLock_FWD_DEFINED__ */


#ifndef __ISelItemCollection_FWD_DEFINED__
#define __ISelItemCollection_FWD_DEFINED__
typedef interface ISelItemCollection ISelItemCollection;

#endif 	/* __ISelItemCollection_FWD_DEFINED__ */


#ifndef __ColumnHeader_FWD_DEFINED__
#define __ColumnHeader_FWD_DEFINED__

#ifdef __cplusplus
typedef class ColumnHeader ColumnHeader;
#else
typedef struct ColumnHeader ColumnHeader;
#endif /* __cplusplus */

#endif 	/* __ColumnHeader_FWD_DEFINED__ */


#ifndef __ListItem_FWD_DEFINED__
#define __ListItem_FWD_DEFINED__

#ifdef __cplusplus
typedef class ListItem ListItem;
#else
typedef struct ListItem ListItem;
#endif /* __cplusplus */

#endif 	/* __ListItem_FWD_DEFINED__ */


#ifndef __ListSubItem_FWD_DEFINED__
#define __ListSubItem_FWD_DEFINED__

#ifdef __cplusplus
typedef class ListSubItem ListSubItem;
#else
typedef struct ListSubItem ListSubItem;
#endif /* __cplusplus */

#endif 	/* __ListSubItem_FWD_DEFINED__ */


#ifndef __ListItems_FWD_DEFINED__
#define __ListItems_FWD_DEFINED__

#ifdef __cplusplus
typedef class ListItems ListItems;
#else
typedef struct ListItems ListItems;
#endif /* __cplusplus */

#endif 	/* __ListItems_FWD_DEFINED__ */


#ifndef __InterfaceCollection_FWD_DEFINED__
#define __InterfaceCollection_FWD_DEFINED__

#ifdef __cplusplus
typedef class InterfaceCollection InterfaceCollection;
#else
typedef struct InterfaceCollection InterfaceCollection;
#endif /* __cplusplus */

#endif 	/* __InterfaceCollection_FWD_DEFINED__ */


#ifndef __RedrawLock_FWD_DEFINED__
#define __RedrawLock_FWD_DEFINED__

#ifdef __cplusplus
typedef class RedrawLock RedrawLock;
#else
typedef struct RedrawLock RedrawLock;
#endif /* __cplusplus */

#endif 	/* __RedrawLock_FWD_DEFINED__ */


#ifndef __SelItemCollection_FWD_DEFINED__
#define __SelItemCollection_FWD_DEFINED__

#ifdef __cplusplus
typedef class SelItemCollection SelItemCollection;
#else
typedef struct SelItemCollection SelItemCollection;
#endif /* __cplusplus */

#endif 	/* __SelItemCollection_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_ATLListView_0000_0000 */
/* [local] */ 









#pragma once
typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("985ECCDF-A665-489D-887B-2BF1E5C8156B") 
enum BorderStyleConstants
    {
        ccNone	= 0,
        ccFixedSingle	= 1
    } 	BorderStyleConstants;

typedef /* [uuid] */  DECLSPEC_UUID("550EC01B-FF02-4AC4-A836-4366A35F6248") 
enum ListSortOrderFlags
    {
        lvwNone	= 0,
        lvwAscending	= 1,
        lvwDescending	= 2,
        lvwNatural	= 4,
        lvwNoCase	= 8
    } 	ListSortOrderFlags;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("A81E9973-2781-4CC6-B325-42FB3E2DB3E2") 
enum ListColumnAlignmentConstants
    {
        lvwColumnLeft	= 0,
        lvwColumnRight	= 1,
        lvwColumnCenter	= 2
    } 	ListColumnAlignmentConstants;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("5407754A-1B28-45A5-8D46-B4B9B5D39B7B") 
enum ColumnContentType
    {
        lvText	= 0,
        lvNumber	= 1,
        lvDate	= 2,
        lvTimeFormat	= 4
    } 	ColumnContentType;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("6CB4DCFB-6598-4726-8A14-C14736529B98") 
enum AppearanceConstants
    {
        ccFlat	= 0,
        cc3D	= 1
    } 	AppearanceConstants;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("F690F6F2-9587-4DEE-B8D1-61F5DDA806DE") 
enum ShiftConstants
    {
        vbShiftMask	= 1,
        vbCtrlMask	= 2,
        vbAltMask	= 4
    } 	vbShiftConstants;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("2454E41D-A2CE-488B-8C1C-FCD7D05AFE4E") 
enum ListColumnResizeMode
    {
        lvwResizeHorizontalProportion	= 0,
        lvwResizeKeepWidth	= 1
    } 	ListColumnResizeMode;

typedef /* [v1_enum][uuid] */  DECLSPEC_UUID("3511FB6C-CB9B-4DE6-A95A-D4554D7885F7") 
enum MouseButtonConstants
    {
        VbLeftButton	= 1,
        VbRightButton	= 2,
        VbMiddleButton	= 4
    } 	vbMouseButtonConstants;

typedef /* [public] */ 
enum __MIDL___MIDL_itf_ATLListView_0000_0000_0001
    {
        fmt_left	= 0x1,
        fmt_right	= 0x2,
        fmt_center	= 0x4,
        fmt_justify_mask	= 0x3,
        fmt_image	= 0x800,
        fmt_bitmap_right	= 0x1000,
        fmt_has_images	= 0x8000
    } 	FormatFlags;

typedef /* [public] */ 
enum __MIDL___MIDL_itf_ATLListView_0000_0000_0002
    {
        lvcf_fmt	= 0x1,
        lvcf_width	= 0x2,
        lvcf_text	= 0x4,
        lvcf_subitem	= 0x8,
        lvcf_image	= 0x10,
        lvcf_order	= 0x20
    } 	FormatMask;

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_ATLListView_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_ATLListView_0000_0000_v0_0_s_ifspec;


#ifndef __ATLLISTVIEWLib_LIBRARY_DEFINED__
#define __ATLLISTVIEWLib_LIBRARY_DEFINED__

/* library ATLLISTVIEWLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ATLLISTVIEWLib;

#ifndef __IInterfaceCollection_INTERFACE_DEFINED__
#define __IInterfaceCollection_INTERFACE_DEFINED__

/* interface IInterfaceCollection */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IInterfaceCollection;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A59BF564-6FE7-4802-81E7-7E057338BA9F")
    IInterfaceCollection : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long *pRetVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ LONG vIndex,
            /* [retval][out] */ IDispatch **pvItem) = 0;
        
        virtual /* [helpstring][id][restricted][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppUnk) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IInterfaceCollectionVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IInterfaceCollection * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IInterfaceCollection * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IInterfaceCollection * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IInterfaceCollection * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IInterfaceCollection * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IInterfaceCollection * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IInterfaceCollection * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IInterfaceCollection * This,
            /* [retval][out] */ long *pRetVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IInterfaceCollection * This,
            /* [in] */ LONG vIndex,
            /* [retval][out] */ IDispatch **pvItem);
        
        /* [helpstring][id][restricted][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IInterfaceCollection * This,
            /* [retval][out] */ IUnknown **ppUnk);
        
        END_INTERFACE
    } IInterfaceCollectionVtbl;

    interface IInterfaceCollection
    {
        CONST_VTBL struct IInterfaceCollectionVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IInterfaceCollection_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IInterfaceCollection_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IInterfaceCollection_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IInterfaceCollection_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IInterfaceCollection_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IInterfaceCollection_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IInterfaceCollection_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IInterfaceCollection_get_Count(This,pRetVal)	\
    ( (This)->lpVtbl -> get_Count(This,pRetVal) ) 

#define IInterfaceCollection_get_Item(This,vIndex,pvItem)	\
    ( (This)->lpVtbl -> get_Item(This,vIndex,pvItem) ) 

#define IInterfaceCollection_get__NewEnum(This,ppUnk)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppUnk) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IInterfaceCollection_INTERFACE_DEFINED__ */


#ifndef __IListControl_INTERFACE_DEFINED__
#define __IListControl_INTERFACE_DEFINED__

/* interface IListControl */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IListControl;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("211C9B08-575E-4626-8E25-6A29D1DE7471")
    IListControl : public IDispatch
    {
    public:
        virtual /* [requestedit][bindable][id][propput] */ HRESULT STDMETHODCALLTYPE put_BackColor( 
            /* [in] */ OLE_COLOR clr) = 0;
        
        virtual /* [requestedit][bindable][id][propget] */ HRESULT STDMETHODCALLTYPE get_BackColor( 
            /* [retval][out] */ OLE_COLOR *pclr) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Appearance( 
            /* [in] */ AppearanceConstants appearance) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Appearance( 
            /* [retval][out] */ AppearanceConstants *pappearance) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_ForeColor( 
            /* [in] */ OLE_COLOR clr) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_ForeColor( 
            /* [retval][out] */ OLE_COLOR *pclr) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Enabled( 
            /* [in] */ VARIANT_BOOL vbool) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Enabled( 
            /* [retval][out] */ VARIANT_BOOL *pbool) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_TabStop( 
            /* [in] */ VARIANT_BOOL vbool) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_TabStop( 
            /* [retval][out] */ VARIANT_BOOL *pbool) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_hWnd( 
            /* [retval][out] */ OLE_HANDLE *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_BorderStyle( 
            /* [in] */ BorderStyleConstants style) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_BorderStyle( 
            /* [retval][out] */ BorderStyleConstants *pstyle) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Refresh( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ColumnHeaders( 
            /* [retval][out] */ IColumnHeaders **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ListItems( 
            /* [retval][out] */ IListItems **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Font( 
            /* [retval][out] */ IFontDisp **pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Font( 
            /* [in] */ IFontDisp *newVal) = 0;
        
        virtual /* [helpstring][id][propputref] */ HRESULT STDMETHODCALLTYPE putref_Font( 
            /* [in] */ IFontDisp *newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetRedraw( 
            VARIANT_BOOL ShouldRedraw) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AxHWnd( 
            /* [retval][out] */ OLE_HANDLE *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ContainerhWnd( 
            /* [retval][out] */ OLE_HANDLE *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RedrawEnabled( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RedrawEnabled( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LayoutSuspended( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SuspendLayout( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ResumeLayout( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_VirtualMode( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_VirtualMode( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Sorted( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_Sorted( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_SortKey( 
            /* [retval][out] */ short *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_SortKey( 
            /* [in] */ short newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_SortOrder( 
            /* [retval][out] */ ListSortOrderFlags *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_SortOrder( 
            /* [in] */ ListSortOrderFlags newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_MultiSelect( 
            /* [retval][out] */ VARIANT_BOOL *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_MultiSelect( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectedItems( 
            /* [retval][out] */ ISelItemCollection **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_SelectedItemCount( 
            /* [retval][out] */ LONG *pVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IListControlVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IListControl * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IListControl * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IListControl * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IListControl * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IListControl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IListControl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IListControl * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [requestedit][bindable][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_BackColor )( 
            IListControl * This,
            /* [in] */ OLE_COLOR clr);
        
        /* [requestedit][bindable][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_BackColor )( 
            IListControl * This,
            /* [retval][out] */ OLE_COLOR *pclr);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Appearance )( 
            IListControl * This,
            /* [in] */ AppearanceConstants appearance);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Appearance )( 
            IListControl * This,
            /* [retval][out] */ AppearanceConstants *pappearance);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ForeColor )( 
            IListControl * This,
            /* [in] */ OLE_COLOR clr);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ForeColor )( 
            IListControl * This,
            /* [retval][out] */ OLE_COLOR *pclr);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Enabled )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL vbool);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Enabled )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pbool);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_TabStop )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL vbool);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_TabStop )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pbool);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_hWnd )( 
            IListControl * This,
            /* [retval][out] */ OLE_HANDLE *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_BorderStyle )( 
            IListControl * This,
            /* [in] */ BorderStyleConstants style);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_BorderStyle )( 
            IListControl * This,
            /* [retval][out] */ BorderStyleConstants *pstyle);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Refresh )( 
            IListControl * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ColumnHeaders )( 
            IListControl * This,
            /* [retval][out] */ IColumnHeaders **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ListItems )( 
            IListControl * This,
            /* [retval][out] */ IListItems **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Font )( 
            IListControl * This,
            /* [retval][out] */ IFontDisp **pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Font )( 
            IListControl * This,
            /* [in] */ IFontDisp *newVal);
        
        /* [helpstring][id][propputref] */ HRESULT ( STDMETHODCALLTYPE *putref_Font )( 
            IListControl * This,
            /* [in] */ IFontDisp *newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SetRedraw )( 
            IListControl * This,
            VARIANT_BOOL ShouldRedraw);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_AxHWnd )( 
            IListControl * This,
            /* [retval][out] */ OLE_HANDLE *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ContainerhWnd )( 
            IListControl * This,
            /* [retval][out] */ OLE_HANDLE *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_RedrawEnabled )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_RedrawEnabled )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LayoutSuspended )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SuspendLayout )( 
            IListControl * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ResumeLayout )( 
            IListControl * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_VirtualMode )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_VirtualMode )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Sorted )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Sorted )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SortKey )( 
            IListControl * This,
            /* [retval][out] */ short *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_SortKey )( 
            IListControl * This,
            /* [in] */ short newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SortOrder )( 
            IListControl * This,
            /* [retval][out] */ ListSortOrderFlags *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_SortOrder )( 
            IListControl * This,
            /* [in] */ ListSortOrderFlags newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_MultiSelect )( 
            IListControl * This,
            /* [retval][out] */ VARIANT_BOOL *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_MultiSelect )( 
            IListControl * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SelectedItems )( 
            IListControl * This,
            /* [retval][out] */ ISelItemCollection **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SelectedItemCount )( 
            IListControl * This,
            /* [retval][out] */ LONG *pVal);
        
        END_INTERFACE
    } IListControlVtbl;

    interface IListControl
    {
        CONST_VTBL struct IListControlVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IListControl_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IListControl_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IListControl_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IListControl_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IListControl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IListControl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IListControl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IListControl_put_BackColor(This,clr)	\
    ( (This)->lpVtbl -> put_BackColor(This,clr) ) 

#define IListControl_get_BackColor(This,pclr)	\
    ( (This)->lpVtbl -> get_BackColor(This,pclr) ) 

#define IListControl_put_Appearance(This,appearance)	\
    ( (This)->lpVtbl -> put_Appearance(This,appearance) ) 

#define IListControl_get_Appearance(This,pappearance)	\
    ( (This)->lpVtbl -> get_Appearance(This,pappearance) ) 

#define IListControl_put_ForeColor(This,clr)	\
    ( (This)->lpVtbl -> put_ForeColor(This,clr) ) 

#define IListControl_get_ForeColor(This,pclr)	\
    ( (This)->lpVtbl -> get_ForeColor(This,pclr) ) 

#define IListControl_put_Enabled(This,vbool)	\
    ( (This)->lpVtbl -> put_Enabled(This,vbool) ) 

#define IListControl_get_Enabled(This,pbool)	\
    ( (This)->lpVtbl -> get_Enabled(This,pbool) ) 

#define IListControl_put_TabStop(This,vbool)	\
    ( (This)->lpVtbl -> put_TabStop(This,vbool) ) 

#define IListControl_get_TabStop(This,pbool)	\
    ( (This)->lpVtbl -> get_TabStop(This,pbool) ) 

#define IListControl_get_hWnd(This,pVal)	\
    ( (This)->lpVtbl -> get_hWnd(This,pVal) ) 

#define IListControl_put_BorderStyle(This,style)	\
    ( (This)->lpVtbl -> put_BorderStyle(This,style) ) 

#define IListControl_get_BorderStyle(This,pstyle)	\
    ( (This)->lpVtbl -> get_BorderStyle(This,pstyle) ) 

#define IListControl_Refresh(This)	\
    ( (This)->lpVtbl -> Refresh(This) ) 

#define IListControl_get_ColumnHeaders(This,pVal)	\
    ( (This)->lpVtbl -> get_ColumnHeaders(This,pVal) ) 

#define IListControl_get_ListItems(This,pVal)	\
    ( (This)->lpVtbl -> get_ListItems(This,pVal) ) 

#define IListControl_get_Font(This,pVal)	\
    ( (This)->lpVtbl -> get_Font(This,pVal) ) 

#define IListControl_put_Font(This,newVal)	\
    ( (This)->lpVtbl -> put_Font(This,newVal) ) 

#define IListControl_putref_Font(This,newVal)	\
    ( (This)->lpVtbl -> putref_Font(This,newVal) ) 

#define IListControl_SetRedraw(This,ShouldRedraw)	\
    ( (This)->lpVtbl -> SetRedraw(This,ShouldRedraw) ) 

#define IListControl_get_AxHWnd(This,pVal)	\
    ( (This)->lpVtbl -> get_AxHWnd(This,pVal) ) 

#define IListControl_get_ContainerhWnd(This,pVal)	\
    ( (This)->lpVtbl -> get_ContainerhWnd(This,pVal) ) 

#define IListControl_get_RedrawEnabled(This,pVal)	\
    ( (This)->lpVtbl -> get_RedrawEnabled(This,pVal) ) 

#define IListControl_put_RedrawEnabled(This,newVal)	\
    ( (This)->lpVtbl -> put_RedrawEnabled(This,newVal) ) 

#define IListControl_get_LayoutSuspended(This,pVal)	\
    ( (This)->lpVtbl -> get_LayoutSuspended(This,pVal) ) 

#define IListControl_SuspendLayout(This)	\
    ( (This)->lpVtbl -> SuspendLayout(This) ) 

#define IListControl_ResumeLayout(This)	\
    ( (This)->lpVtbl -> ResumeLayout(This) ) 

#define IListControl_get_VirtualMode(This,pVal)	\
    ( (This)->lpVtbl -> get_VirtualMode(This,pVal) ) 

#define IListControl_put_VirtualMode(This,newVal)	\
    ( (This)->lpVtbl -> put_VirtualMode(This,newVal) ) 

#define IListControl_get_Sorted(This,pVal)	\
    ( (This)->lpVtbl -> get_Sorted(This,pVal) ) 

#define IListControl_put_Sorted(This,newVal)	\
    ( (This)->lpVtbl -> put_Sorted(This,newVal) ) 

#define IListControl_get_SortKey(This,pVal)	\
    ( (This)->lpVtbl -> get_SortKey(This,pVal) ) 

#define IListControl_put_SortKey(This,newVal)	\
    ( (This)->lpVtbl -> put_SortKey(This,newVal) ) 

#define IListControl_get_SortOrder(This,pVal)	\
    ( (This)->lpVtbl -> get_SortOrder(This,pVal) ) 

#define IListControl_put_SortOrder(This,newVal)	\
    ( (This)->lpVtbl -> put_SortOrder(This,newVal) ) 

#define IListControl_get_MultiSelect(This,pVal)	\
    ( (This)->lpVtbl -> get_MultiSelect(This,pVal) ) 

#define IListControl_put_MultiSelect(This,newVal)	\
    ( (This)->lpVtbl -> put_MultiSelect(This,newVal) ) 

#define IListControl_get_SelectedItems(This,pVal)	\
    ( (This)->lpVtbl -> get_SelectedItems(This,pVal) ) 

#define IListControl_get_SelectedItemCount(This,pVal)	\
    ( (This)->lpVtbl -> get_SelectedItemCount(This,pVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IListControl_INTERFACE_DEFINED__ */


#ifndef ___IListControlEvents_DISPINTERFACE_DEFINED__
#define ___IListControlEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IListControlEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IListControlEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("A674700A-9B7B-4F42-A5D5-0DBD19A6F59D")
    _IListControlEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IListControlEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IListControlEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IListControlEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IListControlEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IListControlEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IListControlEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IListControlEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IListControlEvents * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _IListControlEventsVtbl;

    interface _IListControlEvents
    {
        CONST_VTBL struct _IListControlEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IListControlEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IListControlEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IListControlEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IListControlEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IListControlEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IListControlEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IListControlEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IListControlEvents_DISPINTERFACE_DEFINED__ */


#ifndef __IColumnHeaders_INTERFACE_DEFINED__
#define __IColumnHeaders_INTERFACE_DEFINED__

/* interface IColumnHeaders */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IColumnHeaders;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D78DC907-B60F-4E2E-AE9D-4F0ADCDBB92F")
    IColumnHeaders : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            LONG index,
            /* [retval][out] */ IColumnHeader **pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [optional][in] */ VARIANT *Index,
            /* [optional][in] */ VARIANT *Key,
            /* [optional][in] */ VARIANT *Text,
            /* [optional][in] */ VARIANT *Width,
            /* [optional][in] */ VARIANT *Alignment,
            /* [optional][in] */ VARIANT *Icon,
            /* [retval][out] */ IColumnHeader **out) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HeightInPixels( 
            /* [retval][out] */ long *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HeightInPixels( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ LONG indexToRemove) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IColumnHeadersVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IColumnHeaders * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IColumnHeaders * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IColumnHeaders * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IColumnHeaders * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IColumnHeaders * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IColumnHeaders * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IColumnHeaders * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IColumnHeaders * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IColumnHeaders * This,
            LONG index,
            /* [retval][out] */ IColumnHeader **pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Add )( 
            IColumnHeaders * This,
            /* [optional][in] */ VARIANT *Index,
            /* [optional][in] */ VARIANT *Key,
            /* [optional][in] */ VARIANT *Text,
            /* [optional][in] */ VARIANT *Width,
            /* [optional][in] */ VARIANT *Alignment,
            /* [optional][in] */ VARIANT *Icon,
            /* [retval][out] */ IColumnHeader **out);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IColumnHeaders * This,
            /* [retval][out] */ IUnknown **pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HeightInPixels )( 
            IColumnHeaders * This,
            /* [retval][out] */ long *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_HeightInPixels )( 
            IColumnHeaders * This,
            /* [in] */ long newVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Clear )( 
            IColumnHeaders * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Remove )( 
            IColumnHeaders * This,
            /* [in] */ LONG indexToRemove);
        
        END_INTERFACE
    } IColumnHeadersVtbl;

    interface IColumnHeaders
    {
        CONST_VTBL struct IColumnHeadersVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IColumnHeaders_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IColumnHeaders_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IColumnHeaders_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IColumnHeaders_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IColumnHeaders_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IColumnHeaders_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IColumnHeaders_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IColumnHeaders_get_Count(This,pVal)	\
    ( (This)->lpVtbl -> get_Count(This,pVal) ) 

#define IColumnHeaders_get_Item(This,index,pVal)	\
    ( (This)->lpVtbl -> get_Item(This,index,pVal) ) 

#define IColumnHeaders_Add(This,Index,Key,Text,Width,Alignment,Icon,out)	\
    ( (This)->lpVtbl -> Add(This,Index,Key,Text,Width,Alignment,Icon,out) ) 

#define IColumnHeaders_get__NewEnum(This,pVal)	\
    ( (This)->lpVtbl -> get__NewEnum(This,pVal) ) 

#define IColumnHeaders_get_HeightInPixels(This,pVal)	\
    ( (This)->lpVtbl -> get_HeightInPixels(This,pVal) ) 

#define IColumnHeaders_put_HeightInPixels(This,newVal)	\
    ( (This)->lpVtbl -> put_HeightInPixels(This,newVal) ) 

#define IColumnHeaders_Clear(This)	\
    ( (This)->lpVtbl -> Clear(This) ) 

#define IColumnHeaders_Remove(This,indexToRemove)	\
    ( (This)->lpVtbl -> Remove(This,indexToRemove) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IColumnHeaders_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ListControl;

#ifdef __cplusplus

class DECLSPEC_UUID("720AF342-0D90-4580-BB75-5696CA3B0F51")
ListControl;
#endif

EXTERN_C const CLSID CLSID_ColumnHeaders;

#ifdef __cplusplus

class DECLSPEC_UUID("D30B41C0-DBE6-405A-A826-02FBFF6C7C64")
ColumnHeaders;
#endif

#ifndef __IColumnHeader_INTERFACE_DEFINED__
#define __IColumnHeader_INTERFACE_DEFINED__

/* interface IColumnHeader */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IColumnHeader;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("ED03F8D1-440F-493F-8136-DF3B365190DE")
    IColumnHeader : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Text( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Text( 
            /* [in] */ BSTR *newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Index( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Width( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Width( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ResizeMode( 
            /* [retval][out] */ ListColumnResizeMode *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ResizeMode( 
            /* [in] */ ListColumnResizeMode newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ColumnContentKind( 
            /* [retval][out] */ ColumnContentType *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ColumnContentKind( 
            /* [in] */ ColumnContentType newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE EnsureVisible( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_LeftInPixels( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Alignment( 
            /* [retval][out] */ ListColumnAlignmentConstants *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Alignment( 
            /* [in] */ ListColumnAlignmentConstants newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FixedWidth( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_FixedWidth( 
            /* [in] */ LONG newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Key( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SubItemIndex( 
            /* [retval][out] */ short *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Tag( 
            /* [retval][out] */ VARIANT *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Tag( 
            /* [in] */ VARIANT newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_WidthInPixels( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_HeightInPixels( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_HeightInPixels( 
            /* [in] */ LONG newVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IColumnHeaderVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IColumnHeader * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IColumnHeader * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IColumnHeader * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IColumnHeader * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IColumnHeader * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IColumnHeader * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IColumnHeader * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Text )( 
            IColumnHeader * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Text )( 
            IColumnHeader * This,
            /* [in] */ BSTR *newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Index )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Width )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Width )( 
            IColumnHeader * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ResizeMode )( 
            IColumnHeader * This,
            /* [retval][out] */ ListColumnResizeMode *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ResizeMode )( 
            IColumnHeader * This,
            /* [in] */ ListColumnResizeMode newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ColumnContentKind )( 
            IColumnHeader * This,
            /* [retval][out] */ ColumnContentType *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_ColumnContentKind )( 
            IColumnHeader * This,
            /* [in] */ ColumnContentType newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *EnsureVisible )( 
            IColumnHeader * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_LeftInPixels )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Alignment )( 
            IColumnHeader * This,
            /* [retval][out] */ ListColumnAlignmentConstants *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Alignment )( 
            IColumnHeader * This,
            /* [in] */ ListColumnAlignmentConstants newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_FixedWidth )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_FixedWidth )( 
            IColumnHeader * This,
            /* [in] */ LONG newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Key )( 
            IColumnHeader * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_SubItemIndex )( 
            IColumnHeader * This,
            /* [retval][out] */ short *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Tag )( 
            IColumnHeader * This,
            /* [retval][out] */ VARIANT *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Tag )( 
            IColumnHeader * This,
            /* [in] */ VARIANT newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_WidthInPixels )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_HeightInPixels )( 
            IColumnHeader * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_HeightInPixels )( 
            IColumnHeader * This,
            /* [in] */ LONG newVal);
        
        END_INTERFACE
    } IColumnHeaderVtbl;

    interface IColumnHeader
    {
        CONST_VTBL struct IColumnHeaderVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IColumnHeader_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IColumnHeader_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IColumnHeader_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IColumnHeader_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IColumnHeader_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IColumnHeader_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IColumnHeader_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IColumnHeader_get_Text(This,pVal)	\
    ( (This)->lpVtbl -> get_Text(This,pVal) ) 

#define IColumnHeader_put_Text(This,newVal)	\
    ( (This)->lpVtbl -> put_Text(This,newVal) ) 

#define IColumnHeader_get_Index(This,pVal)	\
    ( (This)->lpVtbl -> get_Index(This,pVal) ) 

#define IColumnHeader_get_Width(This,pVal)	\
    ( (This)->lpVtbl -> get_Width(This,pVal) ) 

#define IColumnHeader_put_Width(This,newVal)	\
    ( (This)->lpVtbl -> put_Width(This,newVal) ) 

#define IColumnHeader_get_ResizeMode(This,pVal)	\
    ( (This)->lpVtbl -> get_ResizeMode(This,pVal) ) 

#define IColumnHeader_put_ResizeMode(This,newVal)	\
    ( (This)->lpVtbl -> put_ResizeMode(This,newVal) ) 

#define IColumnHeader_get_ColumnContentKind(This,pVal)	\
    ( (This)->lpVtbl -> get_ColumnContentKind(This,pVal) ) 

#define IColumnHeader_put_ColumnContentKind(This,newVal)	\
    ( (This)->lpVtbl -> put_ColumnContentKind(This,newVal) ) 

#define IColumnHeader_EnsureVisible(This)	\
    ( (This)->lpVtbl -> EnsureVisible(This) ) 

#define IColumnHeader_get_LeftInPixels(This,pVal)	\
    ( (This)->lpVtbl -> get_LeftInPixels(This,pVal) ) 

#define IColumnHeader_get_Alignment(This,pVal)	\
    ( (This)->lpVtbl -> get_Alignment(This,pVal) ) 

#define IColumnHeader_put_Alignment(This,newVal)	\
    ( (This)->lpVtbl -> put_Alignment(This,newVal) ) 

#define IColumnHeader_get_FixedWidth(This,pVal)	\
    ( (This)->lpVtbl -> get_FixedWidth(This,pVal) ) 

#define IColumnHeader_put_FixedWidth(This,newVal)	\
    ( (This)->lpVtbl -> put_FixedWidth(This,newVal) ) 

#define IColumnHeader_get_Key(This,pVal)	\
    ( (This)->lpVtbl -> get_Key(This,pVal) ) 

#define IColumnHeader_get_SubItemIndex(This,pVal)	\
    ( (This)->lpVtbl -> get_SubItemIndex(This,pVal) ) 

#define IColumnHeader_get_Tag(This,pVal)	\
    ( (This)->lpVtbl -> get_Tag(This,pVal) ) 

#define IColumnHeader_put_Tag(This,newVal)	\
    ( (This)->lpVtbl -> put_Tag(This,newVal) ) 

#define IColumnHeader_get_WidthInPixels(This,pVal)	\
    ( (This)->lpVtbl -> get_WidthInPixels(This,pVal) ) 

#define IColumnHeader_get_HeightInPixels(This,pVal)	\
    ( (This)->lpVtbl -> get_HeightInPixels(This,pVal) ) 

#define IColumnHeader_put_HeightInPixels(This,newVal)	\
    ( (This)->lpVtbl -> put_HeightInPixels(This,newVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IColumnHeader_INTERFACE_DEFINED__ */


#ifndef __IListSubItem_INTERFACE_DEFINED__
#define __IListSubItem_INTERFACE_DEFINED__

/* interface IListSubItem */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IListSubItem;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D32B4AFE-8D77-4FFF-959B-458867A464C3")
    IListSubItem : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IListSubItemVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IListSubItem * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IListSubItem * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IListSubItem * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IListSubItem * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IListSubItem * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IListSubItem * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IListSubItem * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IListSubItemVtbl;

    interface IListSubItem
    {
        CONST_VTBL struct IListSubItemVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IListSubItem_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IListSubItem_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IListSubItem_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IListSubItem_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IListSubItem_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IListSubItem_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IListSubItem_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IListSubItem_INTERFACE_DEFINED__ */


#ifndef __IListItem_INTERFACE_DEFINED__
#define __IListItem_INTERFACE_DEFINED__

/* interface IListItem */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IListItem;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("B0DAEBF8-61C3-4DA5-80FB-201F5D69B47B")
    IListItem : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Text( 
            /* [retval][out] */ BSTR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Text( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Index( 
            /* [retval][out] */ LONG *pVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IListItemVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IListItem * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IListItem * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IListItem * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IListItem * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IListItem * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IListItem * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IListItem * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Text )( 
            IListItem * This,
            /* [retval][out] */ BSTR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE *put_Text )( 
            IListItem * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Index )( 
            IListItem * This,
            /* [retval][out] */ LONG *pVal);
        
        END_INTERFACE
    } IListItemVtbl;

    interface IListItem
    {
        CONST_VTBL struct IListItemVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IListItem_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IListItem_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IListItem_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IListItem_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IListItem_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IListItem_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IListItem_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IListItem_get_Text(This,pVal)	\
    ( (This)->lpVtbl -> get_Text(This,pVal) ) 

#define IListItem_put_Text(This,newVal)	\
    ( (This)->lpVtbl -> put_Text(This,newVal) ) 

#define IListItem_get_Index(This,pVal)	\
    ( (This)->lpVtbl -> get_Index(This,pVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IListItem_INTERFACE_DEFINED__ */


#ifndef __IListItems_INTERFACE_DEFINED__
#define __IListItems_INTERFACE_DEFINED__

/* interface IListItems */
/* [unique][nonextensible][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IListItems;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D7697A8E-0A5F-4B36-850A-394D8FACE5C3")
    IListItems : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Add( 
            /* [optional][in] */ VARIANT *Index,
            /* [optional][in] */ VARIANT *Key,
            /* [optional][in] */ VARIANT *Text,
            /* [optional][in] */ VARIANT *Icon,
            /* [optional][in] */ VARIANT *SmallIcon,
            /* [retval][out] */ IListItem **out) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddRangeEx( 
            SAFEARRAY * *arItemTexts) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddRange( 
            /* [in] */ LONG howMany,
            /* [optional][in] */ VARIANT *Index,
            /* [retval][out] */ IInterfaceCollection **out) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ LONG index,
            /* [retval][out] */ IListItem **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppUnk) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Remove( 
            /* [in] */ LONG Index) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IListItemsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IListItems * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IListItems * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IListItems * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IListItems * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IListItems * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IListItems * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IListItems * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Add )( 
            IListItems * This,
            /* [optional][in] */ VARIANT *Index,
            /* [optional][in] */ VARIANT *Key,
            /* [optional][in] */ VARIANT *Text,
            /* [optional][in] */ VARIANT *Icon,
            /* [optional][in] */ VARIANT *SmallIcon,
            /* [retval][out] */ IListItem **out);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IListItems * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Clear )( 
            IListItems * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AddRangeEx )( 
            IListItems * This,
            SAFEARRAY * *arItemTexts);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *AddRange )( 
            IListItems * This,
            /* [in] */ LONG howMany,
            /* [optional][in] */ VARIANT *Index,
            /* [retval][out] */ IInterfaceCollection **out);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            IListItems * This,
            /* [in] */ LONG index,
            /* [retval][out] */ IListItem **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            IListItems * This,
            /* [retval][out] */ IUnknown **ppUnk);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Remove )( 
            IListItems * This,
            /* [in] */ LONG Index);
        
        END_INTERFACE
    } IListItemsVtbl;

    interface IListItems
    {
        CONST_VTBL struct IListItemsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IListItems_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IListItems_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IListItems_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IListItems_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IListItems_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IListItems_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IListItems_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IListItems_Add(This,Index,Key,Text,Icon,SmallIcon,out)	\
    ( (This)->lpVtbl -> Add(This,Index,Key,Text,Icon,SmallIcon,out) ) 

#define IListItems_get_Count(This,pVal)	\
    ( (This)->lpVtbl -> get_Count(This,pVal) ) 

#define IListItems_Clear(This)	\
    ( (This)->lpVtbl -> Clear(This) ) 

#define IListItems_AddRangeEx(This,arItemTexts)	\
    ( (This)->lpVtbl -> AddRangeEx(This,arItemTexts) ) 

#define IListItems_AddRange(This,howMany,Index,out)	\
    ( (This)->lpVtbl -> AddRange(This,howMany,Index,out) ) 

#define IListItems_get_Item(This,index,pVal)	\
    ( (This)->lpVtbl -> get_Item(This,index,pVal) ) 

#define IListItems_get__NewEnum(This,ppUnk)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppUnk) ) 

#define IListItems_Remove(This,Index)	\
    ( (This)->lpVtbl -> Remove(This,Index) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IListItems_INTERFACE_DEFINED__ */


#ifndef __IRedrawLock_INTERFACE_DEFINED__
#define __IRedrawLock_INTERFACE_DEFINED__

/* interface IRedrawLock */
/* [nonextensible][unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IRedrawLock;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1F163FF2-C8AF-4FC6-9632-DBDE28DC764F")
    IRedrawLock : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ListControlObject( 
            /* [retval][out] */ IListControl **pVal) = 0;
        
        virtual /* [helpstring][id][propputref] */ HRESULT STDMETHODCALLTYPE putref_ListControlObject( 
            /* [in] */ IListControl *newVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRedrawLockVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRedrawLock * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRedrawLock * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRedrawLock * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRedrawLock * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRedrawLock * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRedrawLock * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRedrawLock * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_ListControlObject )( 
            IRedrawLock * This,
            /* [retval][out] */ IListControl **pVal);
        
        /* [helpstring][id][propputref] */ HRESULT ( STDMETHODCALLTYPE *putref_ListControlObject )( 
            IRedrawLock * This,
            /* [in] */ IListControl *newVal);
        
        END_INTERFACE
    } IRedrawLockVtbl;

    interface IRedrawLock
    {
        CONST_VTBL struct IRedrawLockVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRedrawLock_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRedrawLock_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRedrawLock_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRedrawLock_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRedrawLock_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRedrawLock_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRedrawLock_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRedrawLock_get_ListControlObject(This,pVal)	\
    ( (This)->lpVtbl -> get_ListControlObject(This,pVal) ) 

#define IRedrawLock_putref_ListControlObject(This,newVal)	\
    ( (This)->lpVtbl -> putref_ListControlObject(This,newVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRedrawLock_INTERFACE_DEFINED__ */


#ifndef __ISelItemCollection_INTERFACE_DEFINED__
#define __ISelItemCollection_INTERFACE_DEFINED__

/* interface ISelItemCollection */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ISelItemCollection;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("8F6E31F0-063E-41E9-9C2A-F9F14BE656DA")
    ISelItemCollection : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ LONG *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Item( 
            /* [in] */ LONG index,
            /* [retval][out] */ IListItem **pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get__NewEnum( 
            /* [retval][out] */ IUnknown **ppUnk) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ISelItemCollectionVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISelItemCollection * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISelItemCollection * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISelItemCollection * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISelItemCollection * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISelItemCollection * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISelItemCollection * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISelItemCollection * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            ISelItemCollection * This,
            /* [retval][out] */ LONG *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Item )( 
            ISelItemCollection * This,
            /* [in] */ LONG index,
            /* [retval][out] */ IListItem **pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE *get__NewEnum )( 
            ISelItemCollection * This,
            /* [retval][out] */ IUnknown **ppUnk);
        
        END_INTERFACE
    } ISelItemCollectionVtbl;

    interface ISelItemCollection
    {
        CONST_VTBL struct ISelItemCollectionVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISelItemCollection_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISelItemCollection_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISelItemCollection_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISelItemCollection_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISelItemCollection_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISelItemCollection_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISelItemCollection_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ISelItemCollection_get_Count(This,pVal)	\
    ( (This)->lpVtbl -> get_Count(This,pVal) ) 

#define ISelItemCollection_get_Item(This,index,pVal)	\
    ( (This)->lpVtbl -> get_Item(This,index,pVal) ) 

#define ISelItemCollection_get__NewEnum(This,ppUnk)	\
    ( (This)->lpVtbl -> get__NewEnum(This,ppUnk) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ISelItemCollection_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_ColumnHeader;

#ifdef __cplusplus

class DECLSPEC_UUID("FFFA7381-5EA4-4BD7-AE57-09AC6A731F32")
ColumnHeader;
#endif

EXTERN_C const CLSID CLSID_ListItem;

#ifdef __cplusplus

class DECLSPEC_UUID("E43F0889-64B6-45B5-84DF-002FEAEF5F97")
ListItem;
#endif

EXTERN_C const CLSID CLSID_ListSubItem;

#ifdef __cplusplus

class DECLSPEC_UUID("E8B8BD02-42E2-4408-9EA5-102AB790B20C")
ListSubItem;
#endif

EXTERN_C const CLSID CLSID_ListItems;

#ifdef __cplusplus

class DECLSPEC_UUID("1F457B00-05C9-44FC-A3FB-DFFC776620F2")
ListItems;
#endif

EXTERN_C const CLSID CLSID_InterfaceCollection;

#ifdef __cplusplus

class DECLSPEC_UUID("91D76D82-2204-4071-B814-E95D9F05D412")
InterfaceCollection;
#endif

EXTERN_C const CLSID CLSID_RedrawLock;

#ifdef __cplusplus

class DECLSPEC_UUID("FE6B69DD-D7F0-4D48-82BD-73E347EF7FD0")
RedrawLock;
#endif

EXTERN_C const CLSID CLSID_SelItemCollection;

#ifdef __cplusplus

class DECLSPEC_UUID("03B45990-66CA-4EF9-BF2B-28544DFCDE25")
SelItemCollection;
#endif
#endif /* __ATLLISTVIEWLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


