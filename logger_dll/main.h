/*
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA
'
*/

typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;

//basically used to give us a function pointer with right prototype
//and 24 byte empty buffer inline which we assemble commands into in the
//hook proceedure. 
//#define BLOCK _asm int 3
#define ALLOC_THUNK(prototype) __declspec(naked) prototype { _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop _asm nop}	   

ALLOC_THUNK( BOOL     __stdcall Real_WriteFile(HANDLE a0,LPCVOID a1,DWORD a2,LPDWORD a3,LPOVERLAPPED a4) ); 
ALLOC_THUNK( HANDLE   __stdcall Real_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6) );
ALLOC_THUNK( BOOL	  __stdcall Real_WriteFileEx(HANDLE a0,LPCVOID a1,DWORD a2,LPOVERLAPPED a3,LPOVERLAPPED_COMPLETION_ROUTINE a4)) ;
ALLOC_THUNK( HFILE	  __stdcall Real__lcreat(LPCSTR a0,int a1));
ALLOC_THUNK( UINT	  __stdcall Real__lwrite(HFILE a0,LPCSTR a1,UINT a2));
ALLOC_THUNK( BOOL	  __stdcall Real_DeleteFileA(LPCSTR a0));
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4));
ALLOC_THUNK( int	  __stdcall Real_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5));
ALLOC_THUNK( int __stdcall Real_RegCreateKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 ) );
ALLOC_THUNK( int __stdcall Real_RegSetValueA ( HKEY a0, LPCSTR a1, DWORD a2, LPCSTR a3, DWORD a4 ) );
ALLOC_THUNK( int __stdcall Real_RegSetValueExA ( HKEY a0, LPCSTR a1, DWORD a2, DWORD a3, CONST BYTE* a4, DWORD a5 ) );
ALLOC_THUNK( int __stdcall Real_RegCreateKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, LPSTR a3, DWORD a4, REGSAM a5, LPSECURITY_ATTRIBUTES a6, PHKEY a7, LPDWORD a8 ) );

void msg(char);
void LogAPI(const char*, ...);

bool Warned=false;
HWND hServer=0;

void GetHive(HKEY hive, char* buf){

	switch((int)hive){
		case 0x80000000:
				 strcpy(buf, "HKCR\\");
				 break;
		
		case 0x80000001:
				 strcpy(buf, "HKCU\\");
				 break;

		case 0x80000002:
					 strcpy(buf, "HKLM\\");
					 break;

		case 0x80000003:
					 strcpy(buf, " HKU\\");
					 break;

		case 0x80000004 :
					 strcpy(buf, "HKPD\\");
					 break;

		case 0x80000005 :
					 strcpy(buf, "HKPD\\");
					 break;

		case 0x80000006 :
					 strcpy(buf, "HKCC\\");
					 break;
	
		default:
					 //sprintf(buf, "%x", (int)hive);
					 buf[0] = 0;

	};
}




void FindVBWindow(){
	
	const char *key = "hwnd";
	const char *path="Software\\VB and VBA Program Settings\\ComRaider\\Settings\\";
	char buf[255];
	HKEY h;
	unsigned long len;

	/*
	char *vbIDEClassName = "ThunderMDIForm" ;
	char *vbEXEClassName = "ThunderRT6MDIForm" ;
	char *vbWindowCaption = "ComRaider" ;

	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName, vbWindowCaption );
	*/

    //cheat in case they have a form maximized, caption would be changed
	if(hServer==0){ 
		RegOpenKey(HKEY_CURRENT_USER, path, &h);
		RegQueryValueEx(h, key, 0, 0, (unsigned char*)&buf, &len); 
		RegCloseKey(h);
		hServer = (HWND)atoi(buf);
	}

	if(hServer==0){
		if(!Warned){
			MessageBox(0,"Could not find msg window","",0);
			Warned=true;
		}
	}
	else{
		if(!Warned){
			//first time we are being called we could do stuff here...
			Warned=true;

		}
	}	

} 

void msg(char *Buffer){
  
  if(hServer==0) FindVBWindow();
  
  cpyData cpStructData;
  
  cpStructData.cbSize = strlen(Buffer) ;
  cpStructData.lpData = (int)Buffer;
  cpStructData.dwFlag = 4;
  
  SendMessage(hServer, WM_COPYDATA, 0,(LPARAM)&cpStructData);

} 

void LogAPI(const char *format, ...)
{
	DWORD dwErr = GetLastError();
		
	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 msg(buf);
		}
		catch(...){}
	}

	SetLastError(dwErr);
}


__declspec(naked) int CalledFrom(){ 
	
	_asm{
			 mov eax, [ebp+4]  //return address of parent function (were nekkid)
			 ret
	}
	
}

 

