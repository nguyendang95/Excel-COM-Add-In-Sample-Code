#include "pch.h"
#include "ApplicationEventsSink.h"

ApplicationEventsSink::ApplicationEventsSink(Excel::_ApplicationPtr pApp) {
	if (!pApp) return;
	m_App = pApp;
	DispEventAdvise((IUnknown*)m_App);
}

ApplicationEventsSink::~ApplicationEventsSink() {
	DispEventUnadvise((IUnknown*)m_App);
}

HRESULT ApplicationEventsSink::OnWorkbookNewSheet(Excel::Workbook* Wb, IDispatch* Sh) {
	Excel::_WorksheetPtr pSh = Sh;
	CComBSTR bstrShName = NULL;
	HRESULT hr = pSh->get_Name(&bstrShName);
	if (FAILED(hr)) return hr;
	MessageBox(NULL, bstrShName, L"", MB_OK | MB_ICONINFORMATION);
	return S_OK;
}

_ATL_FUNC_INFO ApplicationEventsSink::WorkbookNewSheetInfo = { CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_DISPATCH} };