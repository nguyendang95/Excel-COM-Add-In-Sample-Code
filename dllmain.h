// dllmain.h : Declaration of module class.

class CExcelAddInModule : public ATL::CAtlDllModuleT< CExcelAddInModule >
{
public :
	DECLARE_LIBID(LIBID_ExcelAddInLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_EXCELADDIN, "{26df424a-02b2-440d-a7ed-a6718bfed2e4}")
};

extern class CExcelAddInModule _AtlModule;
