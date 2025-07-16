// MyCustomTaskPane.h : Declaration of the CMyCustomTaskPane

#pragma once
#include "resource.h"       // main symbols



#include "ExcelAddIn_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CMyCustomTaskPane

class ATL_NO_VTABLE CMyCustomTaskPane :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMyCustomTaskPane, &CLSID_MyCustomTaskPane>,
	public IDispatchImpl<IMyCustomTaskPane, &IID_IMyCustomTaskPane, &LIBID_ExcelAddInLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CMyCustomTaskPane()
	{
	}

DECLARE_REGISTRY_RESOURCEID(111)


BEGIN_COM_MAP(CMyCustomTaskPane)
	COM_INTERFACE_ENTRY(IMyCustomTaskPane)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:



};

OBJECT_ENTRY_AUTO(__uuidof(MyCustomTaskPane), CMyCustomTaskPane)
