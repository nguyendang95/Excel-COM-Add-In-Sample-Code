

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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



#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IThisAddIn,0xa90b8f06,0x8456,0x412c,0xa8,0x37,0x7c,0x83,0x4a,0xf5,0x68,0x42);


MIDL_DEFINE_GUID(IID, IID_IRibbonCallback,0xD251AC67,0xBB99,0x4F9F,0x87,0x8D,0x1C,0x1A,0xC5,0xC3,0x1C,0x38);


MIDL_DEFINE_GUID(IID, IID_ITaskPane,0x0ed156c3,0x8913,0x4e25,0x83,0x04,0x00,0x7c,0x36,0x4a,0xa5,0xcf);


MIDL_DEFINE_GUID(IID, IID_IExcelCustomTaskPane,0xfb23467d,0xe764,0x41e8,0x86,0xa1,0xeb,0xbf,0x82,0x53,0x5e,0x18);


MIDL_DEFINE_GUID(IID, IID_IMyCustomTaskPane,0xfb49f83f,0x81f2,0x4c93,0x82,0x8d,0x98,0x54,0x28,0xa8,0x1d,0xf0);


MIDL_DEFINE_GUID(IID, IID_IMyExcelCustomTaskPane,0x625e44fc,0x6158,0x4c4e,0xa7,0x54,0xa7,0x09,0xc1,0x2c,0xd5,0xc7);


MIDL_DEFINE_GUID(IID, LIBID_ExcelAddInLib,0x26df424a,0x02b2,0x440d,0xa7,0xed,0xa6,0x71,0x8b,0xfe,0xd2,0xe4);


MIDL_DEFINE_GUID(CLSID, CLSID_ThisAddIn,0xe9ff80fe,0x7162,0x4221,0xa1,0xc7,0xdc,0x5b,0x41,0xf2,0xc3,0x35);


MIDL_DEFINE_GUID(CLSID, CLSID_TaskPane,0x3174918f,0x5bf0,0x495d,0x8c,0x68,0x2c,0x70,0x86,0xc3,0x5d,0x30);


MIDL_DEFINE_GUID(CLSID, CLSID_ExcelCustomTaskPane,0x91612b79,0xadb5,0x4aff,0x93,0x8a,0xd2,0xd4,0xf9,0x92,0x45,0xe7);


MIDL_DEFINE_GUID(CLSID, CLSID_MyCustomTaskPane,0xc630b734,0xe4af,0x4b4f,0xa6,0x41,0x18,0xe5,0x8e,0x30,0x0c,0x82);


MIDL_DEFINE_GUID(CLSID, CLSID_MyExcelCustomTaskPane,0x3e103957,0xb102,0x4ef4,0x87,0xba,0x13,0x3f,0xde,0x89,0xac,0xd1);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



