// ThisAddIn.h : Declaration of the CThisAddIn

#pragma once
#include "resource.h"       // main symbols
#include "ApplicationEventsSink.h"


#include "ExcelAddIn_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CThisAddIn
typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>CallbackImpl;
typedef IDispatchImpl <IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), 2, 5> RibbonImpl;
typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1, /* wMinor = */ 0> IDTImpl;

class ATL_NO_VTABLE CThisAddIn :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CThisAddIn, &CLSID_ThisAddIn>,
	public IDispatchImpl<IThisAddIn, &IID_IThisAddIn, &LIBID_ExcelAddInLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDTImpl,
	public RibbonImpl,
	public CallbackImpl,
	public IDispatchImpl<ICustomTaskPaneConsumer, &__uuidof(ICustomTaskPaneConsumer), &__uuidof(__Office), 1, 0>
{
public:
	CThisAddIn()
	{
	}

DECLARE_REGISTRY_RESOURCEID(106)


BEGIN_COM_MAP(CThisAddIn)
	COM_INTERFACE_ENTRY(IThisAddIn)
	COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(IRibbonCallback)
	COM_INTERFACE_ENTRY(ICustomTaskPaneConsumer)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}


// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(IDispatch* Application, ext_ConnectMode ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY** custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY** custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY** custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY** custom);
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR* RibbonXml);
	STDMETHOD(Invoke)(DISPID dispIdMember, const IID& riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexceptinfo, UINT* puArgErr);

	STDMETHOD(ButtonClicked)(IDispatch* ribbonControl);
	STDMETHOD(GetImage)(IDispatch* ribbonControl, IPictureDisp** ppPicture);

	STDMETHOD(CTPFactoryAvailable)(ICTPFactory* CTPFactoryInst);

private:
	Excel::_ApplicationPtr m_App;
	ApplicationEventsSink* pAppEventsSink;
	CComPtr<ICTPFactory> m_pCtpFactoryInst;
	CComPtr<_CustomTaskPane> m_pPane;
};

OBJECT_ENTRY_AUTO(__uuidof(ThisAddIn), CThisAddIn)
