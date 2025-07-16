// CustomTaskPane.cpp : Implementation of CCustomTaskPane

#include "pch.h"
#include "CustomTaskPane.h"


// CCustomTaskPane

STDMETHODIMP CCustomTaskPane::CTPFactoryAvailable(ICTPFactory* CTPFactoryInst) {
	if (!CTPFactoryInst) return E_POINTER;
	m_pCtpFactory = CTPFactoryInst;
	HRESULT hr = m_pCtpFactory->CreateCTP(CComBSTR(L"ExcelAddIn.CustomTaskPane"), CComBSTR(L"My Custom Task Pane"), CComVariant(), &m_pPane);
	return hr;
}