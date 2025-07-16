// MyExcelCustomTaskPane.cpp : Implementation of CMyExcelCustomTaskPane
#include "pch.h"
#include "MyExcelCustomTaskPane.h"


// CMyExcelCustomTaskPane

LRESULT CMyExcelCustomTaskPane::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
	HINSTANCE hInstance = (HINSTANCE)GetWindowLongPtr(GWLP_HINSTANCE);
	CreateWindow(L"STATIC", L"Static control", WS_VISIBLE | WS_CHILD, 100, 20, 60, 30, m_hWnd, NULL, hInstance, NULL);
	CreateWindow(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD, 100, 50, 30, 30, m_hWnd, (HMENU)1, hInstance, NULL);
	CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"Hello World!", WS_VISIBLE | WS_CHILD | ES_LEFT | ES_UPPERCASE, 100, 120, 120, 30, m_hWnd, (HMENU)2, hInstance, NULL);

	return 0;
}

LRESULT CMyExcelCustomTaskPane::OnButtonClicked(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
	switch (LOWORD(wParam)) {
	case 1:
	{
		MessageBox(L"Button clicked", L"", MB_OK | MB_ICONINFORMATION);
		break;
	}
	}
	return 0;
}