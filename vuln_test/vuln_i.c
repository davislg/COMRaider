/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Feb 23 18:26:14 2006
 */
/* Compiler settings for D:\work_data\COMRaider\vuln_test\vuln.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


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

const IID IID_Iserver = {0x426D34DC,0x2298,0x417E,{0xA8,0x40,0x6C,0xBD,0xB5,0xAE,0x93,0xE0}};


const IID LIBID_VULNLib = {0xCB85160D,0xAC62,0x4288,{0xAF,0xEE,0xE3,0x9B,0x35,0xF4,0x36,0xB7}};


const IID DIID__IserverEvents = {0x8712742F,0xB31D,0x4241,{0x82,0xD8,0x0B,0x1C,0x93,0xD6,0xF8,0xA7}};


const CLSID CLSID_server = {0x8EF2A07C,0x6E69,0x4144,{0x96,0xAA,0x22,0x47,0xD8,0x92,0xA7,0x3D}};


#ifdef __cplusplus
}
#endif

