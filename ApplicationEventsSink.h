#pragma once
using namespace ATL;
class ApplicationEventsSink : public IDispEventSimpleImpl<1, ApplicationEventsSink, &__uuidof(Excel::AppEvents)>
{
public:
	ApplicationEventsSink(Excel::_ApplicationPtr pApp);
	virtual ~ApplicationEventsSink();
	static _ATL_FUNC_INFO WorkbookNewSheetInfo;
	BEGIN_SINK_MAP(ApplicationEventsSink)
		SINK_ENTRY_INFO(1, __uuidof(Excel::AppEvents), 0x00000625, OnWorkbookNewSheet, &WorkbookNewSheetInfo)
	END_SINK_MAP()

	STDMETHOD(OnWorkbookNewSheet)(Excel::Workbook* Wb, IDispatch* Sh);

private:
	Excel::_ApplicationPtr m_App = NULL;
	
};

