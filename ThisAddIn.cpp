// ThisAddIn.cpp : Implementation of CThisAddIn

#include "pch.h"
#include "ThisAddIn.h"


// CThisAddIn
HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes);
HRESULT GetXMLResource(int nId, BSTR* bstrXml);

STDMETHODIMP CThisAddIn::OnConnection(IDispatch* Application, ext_ConnectMode ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom) {
    if (!Application) return E_POINTER;
    m_App = Application;
    pAppEventsSink = new ApplicationEventsSink(m_App);
    if (!pAppEventsSink) return E_OUTOFMEMORY;
	return S_OK;
}

STDMETHODIMP CThisAddIn::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY** custom) {
    delete pAppEventsSink;
	return S_OK;
}

STDMETHODIMP CThisAddIn::OnAddInsUpdate(SAFEARRAY** custom) {
	return S_OK;
}

STDMETHODIMP CThisAddIn::OnStartupComplete(SAFEARRAY** custom) {
	return S_OK;
}

STDMETHODIMP CThisAddIn::OnBeginShutdown(SAFEARRAY** custom) {
	return S_OK;
}

HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes) {
    HMODULE hModule = _AtlBaseModule.GetModuleInstance();
    if (!hModule) return E_UNEXPECTED;
    HRSRC hRscr = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
    if (!hRscr) return HRESULT_FROM_WIN32(GetLastError());
    HGLOBAL hGlobal = LoadResource(hModule, hRscr);
    if (!hGlobal) return HRESULT_FROM_WIN32(GetLastError());
    *pdwSizeInBytes = SizeofResource(hModule, hRscr);
    *ppvResourceData = LockResource(hGlobal);
    return S_OK;
}

HRESULT GetXMLResource(int nId, BSTR* bstrXml) {
    LPVOID pResourceData = NULL;
    DWORD dwSizeInBytes = 0;
    HRESULT hr = HrGetResource(nId, L"XML", &pResourceData, &dwSizeInBytes);
    if (FAILED(hr)) return hr;
    char* strResourceData = (char*)pResourceData;
    int nLen = MultiByteToWideChar(CP_UTF8, 0, strResourceData, -1, NULL, 0);
    if (!nLen) return E_FAIL;
    wchar_t* wstrResourceData = (wchar_t*)CoTaskMemAlloc((size_t)nLen * sizeof(wchar_t));
    if (!wstrResourceData) return E_OUTOFMEMORY;
    if (!MultiByteToWideChar(CP_UTF8, 0, strResourceData, -1, wstrResourceData, nLen)) {
        CoTaskMemFree(wstrResourceData);
        return E_FAIL;
    }
    *bstrXml = SysAllocString(wstrResourceData);
    CoTaskMemFree(wstrResourceData);
    return (*bstrXml) ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP CThisAddIn::Invoke(DISPID dispIdMember, const IID& riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexceptinfo, UINT* puArgErr) {
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    hr = CallbackImpl::Invoke(dispIdMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);
    if (hr == DISP_E_MEMBERNOTFOUND) {
        hr = RibbonImpl::Invoke(dispIdMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);
    }
    return hr;
}

STDMETHODIMP CThisAddIn::GetCustomUI(BSTR RibbonID, BSTR* RibbonXml) {
    if (!RibbonXml) return E_POINTER;
    return GetXMLResource(IDR_XML1, RibbonXml);
}

HRESULT BitMapToIPictureDisp(HBITMAP hBitmap, IPictureDisp** ppPicture) {
    if (!hBitmap || !ppPicture) return E_INVALIDARG;
    PICTDESC pd{};
    pd.bmp.hbitmap = hBitmap;
    pd.cbSizeofstruct = sizeof(pd);
    pd.picType = PICTYPE_BITMAP;
    IPictureDisp* pPicture = NULL;
    HRESULT hr = OleCreatePictureIndirect(&pd, IID_IPictureDisp, TRUE, (LPVOID*)&pPicture);
    if (FAILED(hr)) return hr;
    *ppPicture = pPicture;
    return *ppPicture ? S_OK : E_FAIL;
}

STDMETHODIMP CThisAddIn::ButtonClicked(IDispatch* ribbonControl) {
    Excel::_WorkbookPtr wb;
    HRESULT hr = m_App->get_ActiveWorkbook(&wb);
    if (FAILED(hr)) return hr;
    Excel::SheetsPtr wss;
    hr = wb->get_Worksheets(&wss);
    if (FAILED(hr)) return hr;
    Excel::_WorksheetPtr ws;
    IDispatchPtr pDisp;
    hr = wss->get_Item(CComVariant(1), &pDisp);
    if (FAILED(hr)) return hr;
    ws = pDisp;
    Excel::RangePtr range = ws->GetRange(CComVariant("A1:A2"), vtMissing);
    SAFEARRAYBOUND sab[2];
    sab[0].cElements = 1;
    sab[0].lLbound = 1;
    sab[1].cElements = 1;
    sab[1].lLbound = 1;
    SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 2, sab);
    if (!psa) return E_OUTOFMEMORY;
    long index[2] = { 1, 1 };
    VARIANT varValue{};
    VariantInit(&varValue);
    varValue.vt = VT_R8;
    varValue.dblVal = 1;
    hr = SafeArrayPutElement(psa, index, (LPVOID)&varValue);
    VariantClear(&varValue);
    if (FAILED(hr)) {
        SafeArrayDestroy(psa);
        return hr;
    }
    _variant_t varData(psa);
    range->PutValue(vtMissing, varData);
    return S_OK;
}

STDMETHODIMP CThisAddIn::GetImage(IDispatch* ribbonControl, IPictureDisp** ppPicture) {
    if (!ppPicture) return E_POINTER;
    HBITMAP hBitmap = LoadBitmap(_AtlBaseModule.GetResourceInstance(), MAKEINTRESOURCE(IDB_BITMAP1));
    if (!hBitmap) return E_UNEXPECTED;
    HRESULT hr = BitMapToIPictureDisp(hBitmap, ppPicture);
    if (FAILED(hr)) return hr;
    return S_OK;
}

STDMETHODIMP CThisAddIn::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst) {
    if (!CTPFactoryInst) return E_POINTER;
    m_pCtpFactoryInst = CTPFactoryInst;
    HRESULT hr = m_pCtpFactoryInst->CreateCTP(CComBSTR(L"ExcelAddIn.MyExcelCustomTaskPane"), CComBSTR(L"My Task Pane"), vtMissing, &m_pPane);
    if (SUCCEEDED(hr)) m_pPane->put_Visible(VARIANT_TRUE);
    return hr;
}