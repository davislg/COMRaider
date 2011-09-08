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
'         This code was ripped from sysanalyzer and sclog projects
*/


#include <windows.h>
#include <stdio.h>
#include <string.h>
//#include <stdlib.h>


void InstallHooks(void);

#include "hooker.h"
#include "main.h"   //contains a bunch of library functions in it too..



bool Installed =false;

void Closing(void){ /*msg("***** Injected Process Terminated *****");*/ }
	

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{

    if(!Installed){
		 Installed=true;
		 InstallHooks();
		 //atexit(Closing);
	}

	return TRUE;
}




//___________________________________________________hook implementations _________


HANDLE __stdcall My_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6)
{
	
	char *calledFrom=0;

	LogAPI("%x     CreateFileA(%s)", CalledFrom(), a0);

    HANDLE ret = 0;
    try{
        ret = Real_CreateFileA(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){} 
	
	return ret;

}

BOOL __stdcall My_WriteFile(HANDLE a0,LPCVOID a1,DWORD a2,LPDWORD a3,LPOVERLAPPED a4)
{
    
	LogAPI("%x     WriteFile(h=%x)", CalledFrom(), a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFile(a0, a1, a2, a3, a4);
    } 
	catch(...){	} 
    return ret;
}
 
HFILE __stdcall My__lcreat(LPCSTR a0,int a1)
{

	LogAPI("%x     _lcreat(%s,%x)", CalledFrom(), a0, a1);

    HFILE ret = 0;
    try {
        ret = Real__lcreat(a0, a1);
    } 
	catch(...){	} 
    return ret;
}


UINT __stdcall My__lwrite(HFILE a0,LPCSTR a1,UINT a2)
{
    
	LogAPI("%x     _lwrite(h=%x)", CalledFrom(), a0);

    UINT ret = 0;
    try {
        ret = Real__lwrite(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}




BOOL __stdcall My_WriteFileEx(HANDLE a0,LPCVOID a1,DWORD a2,LPOVERLAPPED a3,LPOVERLAPPED_COMPLETION_ROUTINE a4)
{
  
    LogAPI("%x     WriteFileEx(h=%x)", CalledFrom(), a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFileEx(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}


//untested
int My_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4)
{
	
	LogAPI("%x     URLDownloadToFile(%s)", CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToFileA(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}

//untested
int My_URLDownloadToCacheFile(int a0,char* a1, char* a2, DWORD a3, DWORD a4, int a5)
{
	
	LogAPI("%x     URLDownloadToCacheFile(%s)", CalledFrom(), a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToCacheFile(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}


//------------------------------------------------------------------
int __stdcall My_RegCreateKeyA ( HKEY a0, LPCSTR a1, PHKEY a2 )
{
	char h[6];
	GetHive(a0,h);
	LogAPI("%x     RegCreateKeyA (%s%s)", CalledFrom() ,h, a1 );

	
	int  ret = 0;
	try{
		ret = Real_RegCreateKeyA (a0,a1,a2);
	}
	catch(...){}

	return ret;
}

int __stdcall My_RegSetValueA ( HKEY a0, LPCSTR a1, DWORD a2, LPCSTR a3, DWORD a4 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegSetValueA (%s%s,%s)", CalledFrom(), h, a1,a3 );

	int  ret = 0;
	try{
		ret = Real_RegSetValueA (a0,a1,a2,a3,a4);
	}
	catch(...){}

	return ret;
}

int __stdcall My_RegCreateKeyExA ( HKEY a0, LPCSTR a1, DWORD a2, LPSTR a3, DWORD a4, REGSAM a5, LPSECURITY_ATTRIBUTES a6, PHKEY a7, LPDWORD a8 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegCreateKeyExA (%s%s,%s)", CalledFrom(), h, a1 , a3 );

	int  ret = 0;
	try{
		ret = Real_RegCreateKeyExA (a0,a1,a2,a3,a4,a5,a6,a7,a8);
	}
	catch(...){}

	return ret;
}

int __stdcall My_RegSetValueExA ( HKEY a0, LPCSTR a1, DWORD a2, DWORD a3, CONST BYTE* a4, DWORD a5 )
{
	char h[6];
	GetHive(a0,h);

	LogAPI("%x     RegSetValueExA (%s%s)", CalledFrom(), h, a1 );

	int  ret = 0;
	try{
		ret = Real_RegSetValueExA (a0,a1,a2,a3,a4,a5);
	}
	catch(...){}

	return ret;
}

//--------------------------------------------------------------


//_______________________________________________ install hooks fx 

void DoHook(void* real, void* hook, void* thunk, char* name){

	char err[400];

	if ( !InstallHook( real, hook, thunk) ){ //try to install the real hook here
		sprintf(err,"***** Install %s hook failed...Error: %s", name, &lastError);
		msg(err);
	} 

}


//Macro wrapper to build DoHook() call
#define ADDHOOK(name) DoHook( name, My_##name, Real_##name, #name );	


void InstallHooks(void)
{

	msg("***** Installing Hooks *****");	
 
	//ADDHOOK(DeleteFileA)
	ADDHOOK(WriteFile);
	ADDHOOK(CreateFileA);
	ADDHOOK(WriteFileEx);
	ADDHOOK(_lcreat);
	ADDHOOK(_lwrite);
	ADDHOOK(URLDownloadToFileA);
	ADDHOOK(URLDownloadToCacheFile);
	ADDHOOK(RegCreateKeyA) 
	ADDHOOK(RegSetValueA)
	ADDHOOK(RegCreateKeyExA)
	ADDHOOK(RegSetValueExA)

	 	
}




