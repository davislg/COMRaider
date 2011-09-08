/*
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:  David Zimmer <david@idefense.com, dzzie@yahoo.com>
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

	EnumProcessModules was sometimes not giving proper output not sure why
	never had a problem with that code before..anyway figure might as well  
	goto the horses mouth right...

    2-20-06 - these routines are still in place, but shouldnt be currently in use, 
			  decided to just extract dll names and sizes from DEBUG_EVENT_LOAD_DLL

*/

#include "windows.h"
#include "stdio.h"
#include <tlhelp32.h> 


typedef struct{
  LIST_ENTRY              InLoadOrderModuleList;
  LIST_ENTRY              InMemoryOrderModuleList;
  LIST_ENTRY              InInitializationOrderModuleList;
  PVOID                   BaseAddress;
  PVOID                   EntryPoint;
  ULONG                   SizeOfImage;
  char*			          FullDllName;
  char*			          BaseDllName;
} LDR_MODULE;

typedef struct{
  ULONG                   Length;
  BOOLEAN                 Initialized;
  PVOID                   SsHandle;
  LIST_ENTRY              InLoadOrderModuleList;
  LIST_ENTRY              InMemoryOrderModuleList;
  LIST_ENTRY              InInitializationOrderModuleList;
} PEB_LDR_DATA,*PPEB_LDR_DATA;

typedef struct PEB { //shortened
  char					  Glob1[12];
  PPEB_LDR_DATA           LoaderData;
} PEB;

typedef struct{
    PVOID  Reserved1;
    PEB* PebBaseAddress;
    PVOID  Reserved2[2];
    PVOID  UniqueProcessId;
    PVOID  Reserved3;
} PROCESS_BASIC_INFORMATION;


typedef LONG (WINAPI NTQIP)(HANDLE, WORD, PVOID, ULONG, PULONG);
NTQIP *lpfnNTQuery;

void (__stdcall *vbCallBack)(int,int,int);

char pebErr[255] = {0};
extern hProcess;


//XP SP2 randomizes PEB address so have to use this
//to support it. Default addr at bottom wont work that case.
PEB* GetPEBAddress(HANDLE hProc){

	PROCESS_BASIC_INFORMATION pbi;
	DWORD dwSize;

	if (lpfnNTQuery!=NULL){
		(*lpfnNTQuery)(hProc, 0, &pbi, sizeof(pbi), &dwSize);
		return pbi.PebBaseAddress ;
	}
	
	return (PEB*)0x7ffdf000;
}

int ScanUnicode(char* strIn, char* bufOut){
	
	int outLen=0;

	for(int i=0;i<255;i++){

		if(strIn[i] == 0 && strIn[i+1]==0) return outLen;
		
		if(strIn[i] != 0){
			bufOut[outLen] = strIn[i];
			outLen++;
		}
	}

	return outLen;
}

int BuildModuleList(void* hProc){ 
	
	PEB peb;
	LDR_MODULE mod;	
	PEB_LDR_DATA pld;

	unsigned long sz;
	char tmp[255] = {0};
	char out[255] = {0};
	char buf[255] = {0};

	memset(&mod,0,sizeof(mod));
	memset(&peb,0,sizeof(peb));
	memset(&pld,0,sizeof(pld));

	//HANDLE hProc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, 0, pid);
	void *pebAddress = (void*)GetPEBAddress(hProc);
    
	try{
			if(!ReadProcessMemory(hProc, pebAddress, (void*)&peb, sizeof(PEB), &sz ) ){
				strcpy(pebErr, "Could not extract PEB from remote process");
				return 0;
			}

			if(!ReadProcessMemory(hProc, peb.LoaderData, (void*)&pld, sizeof(PEB_LDR_DATA), &sz )){
				strcpy(pebErr, "Could not extract Loader Data from remote process\n");
				return 0;
			}

			if(!ReadProcessMemory(hProc, pld.InLoadOrderModuleList.Flink , (void*)&mod, sizeof(LDR_MODULE), &sz )){
				strcpy(pebErr, "Could not extract module Data from remote process\n");
				return 0;
			}

			while( mod.InLoadOrderModuleList.Flink != pld.InLoadOrderModuleList.Flink  ){
			
				if(!ReadProcessMemory(hProc, mod.BaseDllName, out, 254, &sz )){
					sprintf(pebErr,"Could not extract module path from remote process ptr=0x%x\n", mod.BaseDllName );
					return 0;
				}	
				
				memset(tmp,0,255);

				if(ScanUnicode(out, (char*)tmp ) == 0 ){
					strcpy(pebErr, "Error reading Scanning Unicode string");
					return 0;
				}
			
				sz = strlen(tmp);
				while(tmp[sz] != '\\' && sz > 0) sz--;
				
				if(sz>0){
					strncpy(buf, &tmp[sz+1], 254 );
				}else{
					strncpy(buf,tmp, 254);
				}

				//printf("Name: %s Base: %x  Size: %x\n", buf, mod.BaseAddress, mod.SizeOfImage  );
				//MessageBox(0,buf,"",0);
				(*vbCallBack)( (int)&buf, (int)mod.BaseAddress, mod.SizeOfImage);

				if(!ReadProcessMemory(hProc, mod.InLoadOrderModuleList.Flink , (void*)&mod, sizeof(LDR_MODULE), &sz  )){
					sprintf(pebErr,"Could not get next module from remote process ptr=0x%x\n", mod.InMemoryOrderModuleList.Flink );
					return 0;
				}

			}
	 }
	 catch(...){
		strcpy(pebErr, "Error reading Processes PEB :(");
		return 0;
	 }
	
	 return 1;

}


void __stdcall GetModuleListErrorMsg(char* buf500){ strcpy(buf500,pebErr);}

int __stdcall GetModuleList(int lpfnCallback){
		
		if(lpfnNTQuery == NULL){
			lpfnNTQuery = (NTQIP *)GetProcAddress(GetModuleHandle("ntdll.dll"), "NtQueryInformationProcess");
			if(lpfnNTQuery == NULL){
				strcpy(pebErr, "Could not GetProcAddress(NtQueryInformationProcess)");
				return 0;
			}
		}

		vbCallBack = (void (__stdcall *)(int,int,int))lpfnCallback;
		return BuildModuleList((void*)hProcess); 
}
