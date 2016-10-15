
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 6.00.0347 */
/* at Wed Feb 28 16:14:52 2007
 */
/* Compiler settings for Wordpro.odl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#if !defined(_M_IA64) && !defined(_M_AMD64)

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
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_WordPRO,0x6F7DEE17,0xA2AC,0x4232,0xAE,0x4F,0xC3,0x31,0x85,0xD6,0x2D,0x72);


MIDL_DEFINE_GUID(IID, IID_iifWRD_Table,0x5061A8BD,0x1850,0x49E3,0x88,0xAF,0x96,0x24,0x67,0x53,0x5C,0x62);


MIDL_DEFINE_GUID(IID, IID_iifWRD_Document,0x6B47EE8C,0xE208,0x4B4A,0x91,0xAA,0x91,0x23,0xE5,0x58,0xA2,0x7F);


MIDL_DEFINE_GUID(CLSID, CLSID_cifWRD_Document,0xDE108E32,0x4199,0x4B82,0xB3,0x47,0x97,0x5E,0x83,0x2A,0x35,0x35);


MIDL_DEFINE_GUID(CLSID, CLSID_cifWRD_Table,0x310083A8,0x4ED0,0x4D51,0xAB,0xF7,0x81,0x75,0xB6,0xF4,0xAA,0xA5);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/

