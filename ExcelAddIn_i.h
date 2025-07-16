

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 10:14:07 2038
 */
/* Compiler settings for ExcelAddIn.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0628 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ExcelAddIn_i_h__
#define __ExcelAddIn_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __IThisAddIn_FWD_DEFINED__
#define __IThisAddIn_FWD_DEFINED__
typedef interface IThisAddIn IThisAddIn;

#endif 	/* __IThisAddIn_FWD_DEFINED__ */


#ifndef __IRibbonCallback_FWD_DEFINED__
#define __IRibbonCallback_FWD_DEFINED__
typedef interface IRibbonCallback IRibbonCallback;

#endif 	/* __IRibbonCallback_FWD_DEFINED__ */


#ifndef __ITaskPane_FWD_DEFINED__
#define __ITaskPane_FWD_DEFINED__
typedef interface ITaskPane ITaskPane;

#endif 	/* __ITaskPane_FWD_DEFINED__ */


#ifndef __IExcelCustomTaskPane_FWD_DEFINED__
#define __IExcelCustomTaskPane_FWD_DEFINED__
typedef interface IExcelCustomTaskPane IExcelCustomTaskPane;

#endif 	/* __IExcelCustomTaskPane_FWD_DEFINED__ */


#ifndef __IMyCustomTaskPane_FWD_DEFINED__
#define __IMyCustomTaskPane_FWD_DEFINED__
typedef interface IMyCustomTaskPane IMyCustomTaskPane;

#endif 	/* __IMyCustomTaskPane_FWD_DEFINED__ */


#ifndef __IMyExcelCustomTaskPane_FWD_DEFINED__
#define __IMyExcelCustomTaskPane_FWD_DEFINED__
typedef interface IMyExcelCustomTaskPane IMyExcelCustomTaskPane;

#endif 	/* __IMyExcelCustomTaskPane_FWD_DEFINED__ */


#ifndef __ThisAddIn_FWD_DEFINED__
#define __ThisAddIn_FWD_DEFINED__

#ifdef __cplusplus
typedef class ThisAddIn ThisAddIn;
#else
typedef struct ThisAddIn ThisAddIn;
#endif /* __cplusplus */

#endif 	/* __ThisAddIn_FWD_DEFINED__ */


#ifndef __TaskPane_FWD_DEFINED__
#define __TaskPane_FWD_DEFINED__

#ifdef __cplusplus
typedef class TaskPane TaskPane;
#else
typedef struct TaskPane TaskPane;
#endif /* __cplusplus */

#endif 	/* __TaskPane_FWD_DEFINED__ */


#ifndef __ExcelCustomTaskPane_FWD_DEFINED__
#define __ExcelCustomTaskPane_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExcelCustomTaskPane ExcelCustomTaskPane;
#else
typedef struct ExcelCustomTaskPane ExcelCustomTaskPane;
#endif /* __cplusplus */

#endif 	/* __ExcelCustomTaskPane_FWD_DEFINED__ */


#ifndef __MyCustomTaskPane_FWD_DEFINED__
#define __MyCustomTaskPane_FWD_DEFINED__

#ifdef __cplusplus
typedef class MyCustomTaskPane MyCustomTaskPane;
#else
typedef struct MyCustomTaskPane MyCustomTaskPane;
#endif /* __cplusplus */

#endif 	/* __MyCustomTaskPane_FWD_DEFINED__ */


#ifndef __MyExcelCustomTaskPane_FWD_DEFINED__
#define __MyExcelCustomTaskPane_FWD_DEFINED__

#ifdef __cplusplus
typedef class MyExcelCustomTaskPane MyExcelCustomTaskPane;
#else
typedef struct MyExcelCustomTaskPane MyExcelCustomTaskPane;
#endif /* __cplusplus */

#endif 	/* __MyExcelCustomTaskPane_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IThisAddIn_INTERFACE_DEFINED__
#define __IThisAddIn_INTERFACE_DEFINED__

/* interface IThisAddIn */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IThisAddIn;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("a90b8f06-8456-412c-a837-7c834af56842")
    IThisAddIn : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IThisAddInVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IThisAddIn * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IThisAddIn * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IThisAddIn * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IThisAddIn * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IThisAddIn * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IThisAddIn * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IThisAddIn * This,
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
    } IThisAddInVtbl;

    interface IThisAddIn
    {
        CONST_VTBL struct IThisAddInVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IThisAddIn_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IThisAddIn_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IThisAddIn_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IThisAddIn_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IThisAddIn_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IThisAddIn_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IThisAddIn_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IThisAddIn_INTERFACE_DEFINED__ */


#ifndef __IRibbonCallback_INTERFACE_DEFINED__
#define __IRibbonCallback_INTERFACE_DEFINED__

/* interface IRibbonCallback */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRibbonCallback;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D251AC67-BB99-4F9F-878D-1C1AC5C31C38")
    IRibbonCallback : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ButtonClicked( 
            /* [in] */ IDispatch *ribbonControl) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetImage( 
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppPicture) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRibbonCallbackVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRibbonCallback * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRibbonCallback * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRibbonCallback * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRibbonCallback * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRibbonCallback * This,
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
        
        DECLSPEC_XFGVIRT(IRibbonCallback, ButtonClicked)
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ButtonClicked )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl);
        
        DECLSPEC_XFGVIRT(IRibbonCallback, GetImage)
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetImage )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *ribbonControl,
            /* [retval][out] */ IPictureDisp **ppPicture);
        
        END_INTERFACE
    } IRibbonCallbackVtbl;

    interface IRibbonCallback
    {
        CONST_VTBL struct IRibbonCallbackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRibbonCallback_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRibbonCallback_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRibbonCallback_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRibbonCallback_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRibbonCallback_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRibbonCallback_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRibbonCallback_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRibbonCallback_ButtonClicked(This,ribbonControl)	\
    ( (This)->lpVtbl -> ButtonClicked(This,ribbonControl) ) 

#define IRibbonCallback_GetImage(This,ribbonControl,ppPicture)	\
    ( (This)->lpVtbl -> GetImage(This,ribbonControl,ppPicture) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRibbonCallback_INTERFACE_DEFINED__ */


#ifndef __ITaskPane_INTERFACE_DEFINED__
#define __ITaskPane_INTERFACE_DEFINED__

/* interface ITaskPane */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITaskPane;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0ed156c3-8913-4e25-8304-007c364aa5cf")
    ITaskPane : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ITaskPaneVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITaskPane * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITaskPane * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITaskPane * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITaskPane * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITaskPane * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITaskPane * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITaskPane * This,
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
    } ITaskPaneVtbl;

    interface ITaskPane
    {
        CONST_VTBL struct ITaskPaneVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITaskPane_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITaskPane_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITaskPane_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITaskPane_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITaskPane_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITaskPane_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITaskPane_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITaskPane_INTERFACE_DEFINED__ */


#ifndef __IExcelCustomTaskPane_INTERFACE_DEFINED__
#define __IExcelCustomTaskPane_INTERFACE_DEFINED__

/* interface IExcelCustomTaskPane */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IExcelCustomTaskPane;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fb23467d-e764-41e8-86a1-ebbf82535e18")
    IExcelCustomTaskPane : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IExcelCustomTaskPaneVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExcelCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExcelCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExcelCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExcelCustomTaskPane * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExcelCustomTaskPane * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExcelCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExcelCustomTaskPane * This,
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
    } IExcelCustomTaskPaneVtbl;

    interface IExcelCustomTaskPane
    {
        CONST_VTBL struct IExcelCustomTaskPaneVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExcelCustomTaskPane_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExcelCustomTaskPane_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExcelCustomTaskPane_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExcelCustomTaskPane_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExcelCustomTaskPane_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExcelCustomTaskPane_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExcelCustomTaskPane_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExcelCustomTaskPane_INTERFACE_DEFINED__ */


#ifndef __IMyCustomTaskPane_INTERFACE_DEFINED__
#define __IMyCustomTaskPane_INTERFACE_DEFINED__

/* interface IMyCustomTaskPane */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMyCustomTaskPane;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("fb49f83f-81f2-4c93-828d-985428a81df0")
    IMyCustomTaskPane : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IMyCustomTaskPaneVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyCustomTaskPane * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyCustomTaskPane * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyCustomTaskPane * This,
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
    } IMyCustomTaskPaneVtbl;

    interface IMyCustomTaskPane
    {
        CONST_VTBL struct IMyCustomTaskPaneVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyCustomTaskPane_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyCustomTaskPane_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyCustomTaskPane_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyCustomTaskPane_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyCustomTaskPane_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyCustomTaskPane_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyCustomTaskPane_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMyCustomTaskPane_INTERFACE_DEFINED__ */


#ifndef __IMyExcelCustomTaskPane_INTERFACE_DEFINED__
#define __IMyExcelCustomTaskPane_INTERFACE_DEFINED__

/* interface IMyExcelCustomTaskPane */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMyExcelCustomTaskPane;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("625e44fc-6158-4c4e-a754-a709c12cd5c7")
    IMyExcelCustomTaskPane : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IMyExcelCustomTaskPaneVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyExcelCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyExcelCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyExcelCustomTaskPane * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyExcelCustomTaskPane * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyExcelCustomTaskPane * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyExcelCustomTaskPane * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyExcelCustomTaskPane * This,
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
    } IMyExcelCustomTaskPaneVtbl;

    interface IMyExcelCustomTaskPane
    {
        CONST_VTBL struct IMyExcelCustomTaskPaneVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyExcelCustomTaskPane_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyExcelCustomTaskPane_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyExcelCustomTaskPane_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyExcelCustomTaskPane_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyExcelCustomTaskPane_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyExcelCustomTaskPane_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyExcelCustomTaskPane_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMyExcelCustomTaskPane_INTERFACE_DEFINED__ */



#ifndef __ExcelAddInLib_LIBRARY_DEFINED__
#define __ExcelAddInLib_LIBRARY_DEFINED__

/* library ExcelAddInLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_ExcelAddInLib;

EXTERN_C const CLSID CLSID_ThisAddIn;

#ifdef __cplusplus

class DECLSPEC_UUID("e9ff80fe-7162-4221-a1c7-dc5b41f2c335")
ThisAddIn;
#endif

EXTERN_C const CLSID CLSID_TaskPane;

#ifdef __cplusplus

class DECLSPEC_UUID("3174918f-5bf0-495d-8c68-2c7086c35d30")
TaskPane;
#endif

EXTERN_C const CLSID CLSID_ExcelCustomTaskPane;

#ifdef __cplusplus

class DECLSPEC_UUID("91612b79-adb5-4aff-938a-d2d4f99245e7")
ExcelCustomTaskPane;
#endif

EXTERN_C const CLSID CLSID_MyCustomTaskPane;

#ifdef __cplusplus

class DECLSPEC_UUID("c630b734-e4af-4b4f-a641-18e58e300c82")
MyCustomTaskPane;
#endif

EXTERN_C const CLSID CLSID_MyExcelCustomTaskPane;

#ifdef __cplusplus

class DECLSPEC_UUID("3e103957-b102-4ef4-87ba-133fde89acd1")
MyExcelCustomTaskPane;
#endif
#endif /* __ExcelAddInLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


