/*
Purpose: sclog.exe

		This research application was designed to allow malcode analysts to
		quickly get an overview of an unknown shellcodes functionality by
		actually executing it within the framework of a minimal sandbox
		implemented through the use of API hooking. 

		It is not recommended to run unknown payloads outside of VMWare type
		enviroments. 

		By using this tool, you take responsibility for any results the use 
		of this tool may cause. It is NOT guaranteed to be safe.

		sclog supports the following command line arguments:

			Usage: sclog <sc_file> [/addbpx /redir /nonet /nofilt /dump /step]

			sc_file     shellcode file to execute and log
			/addbpx     Adds a breakpoint to beginning of shellcode buffer
			/redir      Changes IP specified in Connect() to localhost
			/nonet      no safety net - if set we dont block any dangerous apis
			/nofilt     no api filtering - show all hook messages
			/dump       dumps shellcode buffer to disk at first api call (self decoded)
			/step       asks the user to permit each hooked API call before executing

		Several sample shellcode payloads are provided (*.sc) 
		See the readme file for example output.

License: sclog.exe Copyright (C) 2005 David Zimmer <david@idefense.com, dzzie@yahoo.com>

		 Assembler and Disassembler engines are Copyright (C) 2001 Oleh Yuschuk
		 and used under GPL License. (disasm.h, asmserv.c, assembl.c, disasm.c)

         This program is free software; you can redistribute it and/or modify it
         under the terms of the GNU General Public License as published by the Free
         Software Foundation; either version 2 of the License, or (at your option)
         any later version.

         This program is distributed in the hope that it will be useful, but WITHOUT
         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
         more details.

         You should have received a copy of the GNU General Public License along with
         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
         Place, Suite 330, Boston, MA 02111-1307 USA

*/




#define _WIN32_WINNT 0x0401  //for IsDebuggerPresent 
#include <Winsock2.h>
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>

HANDLE STDOUT;
HANDLE STDIN;

DWORD bufsz=0;  //these are global so we can check to see if execution comes from 
char *buf;      //   this vincinity for logging ret address

int redirect=0; //cmdline option to change connect ips to 127.0.0.1
int nonet=0;    //no safety net if 1 we dont block any apis 
int nofilt=0;   //no api filters show all 
int autoDump=0; //quick autoway to get dump at first api call we detect
int stepMode=0; //call by call affirmation to allow

int infoMsgColor = 0x0E;
char sc_file[MAX_PATH];

void InstallHooks(void);

#include "hooker.h"
#include "main.h"   //contains a bunch of library functions in it too..



//___________________________________________________hook implementations _________


HANDLE __stdcall My_CreateFileA(LPCSTR a0,DWORD a1,DWORD a2,LPSECURITY_ATTRIBUTES a3,DWORD a4,DWORD a5,HANDLE a6)
{

    AddAddr( SCOffset() );	
	LogAPI("CreateFileA(%s)\n", a0);

    HANDLE ret = 0;
    try{
        ret = Real_CreateFileA(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){
	
	} 

    return ret;
}

BOOL __stdcall My_WriteFile(HANDLE a0,LPCVOID a1,DWORD a2,LPDWORD a3,LPOVERLAPPED a4)
{
    
	AddAddr( SCOffset() );	
	LogAPI("WriteFile(h=%x)\n", a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFile(a0, a1, a2, a3, a4);
    } 
	catch(...){	} 
    return ret;
}
 
HFILE __stdcall My__lcreat(LPCSTR a0,int a1)
{
    AddAddr( SCOffset() );	
	LogAPI("_lcreat(%s,%x)\n", a0, a1);

    HFILE ret = 0;
    try {
        ret = Real__lcreat(a0, a1);
    } 
	catch(...){	} 
    return ret;
}

HFILE __stdcall My__lopen(LPCSTR a0, int a1)
{
   
    AddAddr( SCOffset() );	
	LogAPI("_lopen(%s,%x)\n", a0, a1);

    HFILE ret = 0;
    try {
        ret = Real__lopen(a0, a1);
    }
	catch(...){	} 

    return ret;
}

UINT __stdcall My__lread(HFILE a0,LPVOID a1,UINT a2)
{
    AddAddr( SCOffset() );	
	LogAPI("_lread(%x,%x,%x)\n", a0, a1, a2);

    UINT ret = 0;
    try {
        ret = Real__lread(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

UINT __stdcall My__lwrite(HFILE a0,LPCSTR a1,UINT a2)
{
    
	AddAddr( SCOffset() );	
	LogAPI("_lwrite(h=%x)\n", a0);

    UINT ret = 0;
    try {
        ret = Real__lwrite(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}




BOOL __stdcall My_WriteFileEx(HANDLE a0,LPCVOID a1,DWORD a2,LPOVERLAPPED a3,LPOVERLAPPED_COMPLETION_ROUTINE a4)
{
    AddAddr( SCOffset() );	
    LogAPI("WriteFileEx(h=%x)\n", a0);

    BOOL ret = 0;
    try {
        ret = Real_WriteFileEx(a0, a1, a2, a3, a4);
    }
	catch(...){	} 

    return ret;
}

DWORD __stdcall My_WaitForSingleObject(HANDLE a0,DWORD a1)
{
   
   	if( calledFromSC() ){
		AddAddr( SCOffset() );	
		LogAPI("WaitForSingleObject(%x,%x)\n", a0, a1);
	}

    DWORD ret = 0;
    try {
        ret = Real_WaitForSingleObject(a0, a1);
    }
	catch(...){	} 

    return ret;
}


//_________ws2_32__________________________________________________________

SOCKET __stdcall My_accept(SOCKET a0,sockaddr* a1,int* a2)
{
    AddAddr( SCOffset() );	
	LogAPI("accept(%x,%x,%x)\n", a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_accept(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_bind(SOCKET a0,SOCKADDR_IN* a1, int a2)
{
    
	AddAddr( SCOffset() );	
	LogAPI("bind(%x, port=%ld)\n", a0, htons(a1->sin_port) );

    int ret = 0;
    try {
        ret = Real_bind(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_closesocket(SOCKET a0)
{
    
	AddAddr( SCOffset() );	
	LogAPI("closesocket(%x)\n", a0);

    int ret = 0;
    try {
        ret = Real_closesocket(a0);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_connect(SOCKET a0,SOCKADDR_IN* a1,int a2)
{
    
	char* ip=0;	
	ip=ipfromlng(a1);
	
	if(redirect){
		infomsg("     Connect Redirecting Enabled: %s -> 127.0.0.1\n",ip); 
		free(ip);
		a1->sin_addr.S_un.S_addr=inet_addr("127.0.0.1");
		ip=ipfromlng(a1);
	}

	AddAddr( SCOffset() );	
	LogAPI("connect( %s:%d )\n", ip, htons(a1->sin_port) );
	
	free(ip);

    int ret = 0;
    try {
        ret = Real_connect(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

hostent* __stdcall My_gethostbyaddr(char* a0,int a1,int a2)
{
    
	AddAddr( SCOffset() );	
	LogAPI("gethostbyaddr(%x)\n", a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyaddr(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

hostent* __stdcall My_gethostbyname(char* a0)
{
    AddAddr( SCOffset() );	
	LogAPI("gethostbyname(%x)\n", a0);

    hostent* ret = 0;
    try {
        ret = Real_gethostbyname(a0);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_gethostname(char* a0,int a1)
{
    AddAddr( SCOffset() );	
	LogAPI("gethostname(%x)\n", a0);

    int ret = 0;
    try {
        ret = Real_gethostname(a0, a1);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_listen(SOCKET a0,int a1)
{
    
	AddAddr( SCOffset() );	
	LogAPI("listen(h=%x )\n", a0);

    int ret = 0;
    try {
        ret = Real_listen(a0, a1);
    }
	catch(...){	} 

    return ret;
}

int __stdcall My_recv(SOCKET a0,char* a1,int a2,int a3)
{
	AddAddr( SCOffset() );	
    LogAPI("recv(h=%x)\n", a0);

    int ret = 0;
    try {
        ret = Real_recv(a0, a1, a2, a3);

		if(ret>0){
			hexdump((unsigned char*)a1,ret);
		}

    } 
	catch(...){	} 

    return ret;
}

int __stdcall My_send(SOCKET a0,char* a1,int a2,int a3)
{
    
	AddAddr( SCOffset() );	
	LogAPI("send(h=%x)\n", a0);
    int ret = 0;

    try {

		if(a2>0 && *a1 !=0)	hexdump((unsigned char*)a1,a2);
        ret = Real_send(a0, a1, a2, a3);
    
	}
	catch(...){	} 

    return ret;
}

int __stdcall My_shutdown(SOCKET a0,int a1)
{
    
	AddAddr( SCOffset() );	
	LogAPI("shutdown()\n");

    int ret = 0;
    try {
        ret = Real_shutdown(a0, a1);
    }
	catch(...){	} 

    return ret;
}

SOCKET __stdcall My_socket(int a0,int a1,int a2)
{
	
	AddAddr( SCOffset() );		
	LogAPI("socket(family=%x,type=%x,proto=%x)\n", a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_socket(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

SOCKET __stdcall My_WSASocketA(int a0,int a1,int a2,struct _WSAPROTOCOL_INFOA* a3,GROUP a4,DWORD a5)
{
    
	AddAddr( SCOffset() );	
	LogAPI("WSASocketA(fam=%x,typ=%x,proto=%x)\n", a0, a1, a2);

    SOCKET ret = 0;
    try {
        ret = Real_WSASocketA(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}



//untested
int My_URLDownloadToFileA(int a0,char* a1, char* a2, DWORD a3, int a4)
{
	
	AddAddr( SCOffset() );	
	LogAPI("URLDownloadToFile(%s)\n", a1);

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
	
	AddAddr( SCOffset() );	
	LogAPI("URLDownloadToCacheFile(%s)\n", a1);

    SOCKET ret = 0;
    try {
        ret = Real_URLDownloadToCacheFile(a0, a1, a2, a3, a4, a5);
    }
	catch(...){	} 

    return ret;
}

void __stdcall My_ExitProcess(UINT a0)
{
    
	AddAddr( SCOffset() );	
	LogAPI("ExitProcess()\n");

    try {
        Real_ExitProcess(a0);
    }
	catch(...){	} 

}

void __stdcall My_ExitThread(DWORD a0)
{
    
	AddAddr( SCOffset() );	
	LogAPI("ExitThread()\n");

    try {
        Real_ExitThread(a0);
    }
	catch(...){	} 

}

FILE* __stdcall My_fopen(const char* a0, const char* a1)
{

    AddAddr( SCOffset() );	
	LogAPI("fopen(%s)\n", a0);

	FILE* rt=0;
    try {
        rt = Real_fopen(a0,a1);
    }
	catch(...){	} 

	return rt;
}

size_t __stdcall My_fwrite(const void* a0, size_t a1, size_t a2, FILE* a3)
{

    AddAddr( SCOffset() );	
	LogAPI("fwrite(h=%x)\n", a3);

	size_t rt=0;
    try {
        rt = Real_fwrite(a0,a1,a2,a3);
    }
	catch(...){	} 

	return rt;
}

HANDLE __stdcall My_OpenProcess(DWORD a0,BOOL a1,DWORD a2)
{
    AddAddr( SCOffset() );	
	LogAPI("OpenProcess(pid=%ld)\n", a2);

    HANDLE ret = 0;
    try {
        ret = Real_OpenProcess(a0, a1, a2);
    }
	catch(...){	} 

    return ret;
}

HMODULE __stdcall My_GetModuleHandleA(LPCSTR a0)
{
    AddAddr( SCOffset() );	
	LogAPI("GetModuleHandleA(%s)\n", a0);

    HMODULE ret = 0;
    try {
        ret = Real_GetModuleHandleA(a0);
    }
	catch(...){	} 

    return ret;
}


//_________________________________________________ banned unless /nonet _______________
UINT __stdcall My_WinExec(LPCSTR a0,UINT a1)
{

	AddAddr( SCOffset() );	

    if(!nonet){
		infomsg("Skipping WinExec(%s,%x)\n", a0, a1);  
		return 0;
	}

	LogAPI("WinExec(%s,%x)\n", a0, a1);

    UINT ret = 0;
    try {
        ret = Real_WinExec(a0, a1);
    }
	catch(...){	} 

    return ret;


}

BOOL __stdcall My_DeleteFileA(LPCSTR a0)
{
	
	AddAddr( SCOffset() );	
 	infomsg("Skipping DeleteFileA(%s)\n", a0); //deleting is never cool nonet or not
	return 0;
	 

}

BOOL __stdcall My_CreateProcessA(LPCSTR a0,LPSTR a1,LPSECURITY_ATTRIBUTES a2,LPSECURITY_ATTRIBUTES a3,BOOL a4,DWORD a5,LPVOID a6,LPCSTR a7,struct _STARTUPINFOA* a8,LPPROCESS_INFORMATION a9)
{

	AddAddr( SCOffset() );	    

	if(!nonet){
		infomsg("Skipping CreateProcessA(%s,%s)\n", a0, a1);
		return 0;
	}

	LogAPI("CreateProcessA(%s,%s,%x,%s)\n", a0, a1, a6, a7);

    BOOL ret = 0;
    try {
        ret = Real_CreateProcessA(a0, a1, a2, a3, a4, a5, a6, a7, a8, a9);
    }
	catch(...){	} 

    return ret;



}

int My_system(const char* cmd)
{
    
	AddAddr( SCOffset() );	
	
	if(!nonet){
		infomsg("Skipping call to system(%s)\n", cmd);
		return 0;
	}
	
	LogAPI("system(%s)\n", cmd);

	int ret=0;
	try {
        ret = Real_system(cmd);
    }
	catch(...){	} 

    return ret;

}

HANDLE __stdcall My_CreateRemoteThread(HANDLE a0,LPSECURITY_ATTRIBUTES a1,DWORD a2,LPTHREAD_START_ROUTINE a3,LPVOID a4,DWORD a5,LPDWORD a6)
{
	
	AddAddr( SCOffset() );	

	if(!nonet){
		infomsg("Skipping CreateRemoteThread()\n");
		return 0;
	}

	LogAPI("CreateRemoteThread(h=%x, start=%x)\n", a0,a3);

    HANDLE ret = 0;
    try {
        ret = Real_CreateRemoteThread(a0, a1, a2, a3, a4, a5, a6);
    }
	catch(...){	} 

    return ret;

}

BOOL __stdcall My_WriteProcessMemory(HANDLE a0,LPVOID a1,LPVOID a2,DWORD a3,LPDWORD a4)
{

    
	AddAddr( SCOffset() );	

	if(!nonet){
		infomsg("Skipping WriteProcessMemory(h=%x,len=%x)\n", a0, a3);	
		return 0;
	}

	LogAPI("WriteProcessMemory(h=%x,len=%x)\n", a0, a3);

    BOOL ret = 0;
    try {
		
		hexdump( (unsigned char*) a2, a3 );
        ret = Real_WriteProcessMemory(a0, a1, a2, a3, a4);

    }
	catch(...){	} 

    return ret;
}

 
// ________________________________________________  monitored ________________

HMODULE __stdcall My_LoadLibraryA(char* a0)
{
    int isOK=0;
   
    int dllCnt=6,i=0;
	const char* tok = ".";
	char *p;
	char *okDlls[] = { "ws2_32","kernel32","advapi32", "urlmon", "msafd", "msvcrt" };
	HMODULE ret = 0;

	if(nonet||!*a0){
		isOK=1;
	}else{
		
		p = strtok( a0,tok);
		if(*p==0) p=a0;
		
		strlower(p);
		
		for(i=0;i<dllCnt;i++){
			if( strstr(p, okDlls[i]) > 0 ){
				 isOK=1;
				 break;
			}
		}
	}	

	if(isOK==0){	
		AddAddr( SCOffset() );
		infomsg("Halting..LoadLibrary for dll not in safe list: %s",a0);
		exit(0);
	}
		
	if( calledFromSC() ){
		AddAddr( SCOffset() );
		LogAPI("LoadLibraryA(%s)\n",  a0);
	}

	try {
		ret = Real_LoadLibraryA(a0);
	}
	catch(...){	} 



	return ret;

}


 
FARPROC __stdcall My_GetProcAddress(HMODULE a0,LPCSTR a1)
{
	
	if( calledFromSC() ){
		AddAddr( SCOffset() );	
		LogAPI("GetProcAddress(%s)\n", a1);
	}

    FARPROC ret = 0;
    try {
        ret = Real_GetProcAddress(a0, a1);
    }
	catch(...){	} 

    return ret;
}

//_________________________________________________ end of hook implementations ________

void usage(void){
	printf("           Generic Shellcode Logger v0.1 BETA\n");
	printf(" Author David Zimmer <david@idefense.com, dzzie@yahoo.com>\n");
	printf(" Uses the GPL Asm/Dsm Engines from OllyDbg (C) 2001 Oleh Yuschuk\n\n");
	SetConsoleTextAttribute(STDOUT,  0x0F); //white
	printf(" Usage: sc_log <sc_file> [/addbpx /redir /nonet /nofilt /dump /step]\n\n");
	printf("    sc_file\tshellcode file to execute and log\n");
	printf("    /addbpx\tAdds a breakpoint to beginning of shellcode buffer\n");
	printf("    /redir\tChanges IP specified in Connect() to localhost\n");
	printf("    /nonet\tno safety net - if set we dont block any dangerous apis\n");
	printf("    /nofilt\tno api filtering - show all hook messages\n");
	printf("    /dump\tdump (probably decoded) shellcode at first api call\n");
	printf("    /step\task user before each hooked api to continue\n\n");        
	SetConsoleTextAttribute(STDOUT,  0x07); //default gray
	printf(" Note that many interesting apis are logged, but not all.\n");
	printf(" Shellcode is allowed to run within a minimal sandbox..\n");
	printf(" and only known safe (hooked) dlls are allowed to load\n\n");
	printf(" It is advised to only run this in VM enviroments as not\n");
	printf(" all paths are blocked that could lead to system subversion.\n");
	printf(" As it runs, API hooks will be used to log actions skipping\n");
	printf(" many dangerous functions.\n\n");
	SetConsoleTextAttribute(STDOUT,  0x0E); //yellow
	printf(" Use at your own risk!\n");
	SetConsoleTextAttribute(STDOUT,  0x07); //default gray
	ExitProcess(0);
}

 

void main(int argc, char **argv){
	
	DWORD l;
	OFSTRUCT o;
	WSADATA WsaDat;	
	int addbpx=0;

	system("cls");
	printf("\n");

	STDOUT = GetStdHandle(STD_OUTPUT_HANDLE);
	STDIN  = GetStdHandle(STD_INPUT_HANDLE);
	SetConsoleMode(STDIN, !ENABLE_LINE_INPUT ); //turn off line input (step mode needs)

	if(argc < 2) usage();
	if(strstr(argv[1],"?") > 0 ) usage();
	if(strstr(argv[1],"-h") > 0 ) usage();

	for(int i=2; i<argc; i++){
		if(strstr(argv[i],"/addbpx") > 0 ) addbpx=1;
		if(strstr(argv[i],"/redir") > 0 )  redirect=1;
		if(strstr(argv[i],"/nonet") > 0 )  nonet=1;
		if(strstr(argv[i],"/nofilt") > 0 ) nofilt=1;
		if(strstr(argv[i],"/dump") > 0 )   autoDump=1;
		if(strstr(argv[i],"/step") > 0 )   stepMode=1; //might still have some side effects 
	}

	char* filename = argv[1];
	HANDLE h =  (HANDLE)OpenFile(filename, &o , OF_READ);
	
	if(h == INVALID_HANDLE_VALUE ){
		printf("Could not open file %s\n\n", filename);
		return;
	}

	strcpy(sc_file,argv[1]);
	bufsz = GetFileSize(h,NULL);
	
	if( bufsz == INVALID_FILE_SIZE){
		printf("Could not get filesize\n\n");
		CloseHandle(h);
		return;
	}
	
	if( bufsz > 2000){
		printf("What in the world are you loading..to big..nay i say!\n");
		CloseHandle(h);
		return;
	}

	if(addbpx){
		printf("Adding Breakpoint to beginning of shellcode buffer\n");
		bufsz++;
	}

	buf = (char*)malloc(bufsz);
	printf("Loading Shellcode into memory\n");

	if(addbpx){
		buf[0]= (unsigned char)0xCC;
		ReadFile(h, &buf[1]  , (bufsz-1) ,&l,0);
	}else{
		ReadFile(h, buf  , bufsz ,&l,0);
	}

	CloseHandle(h);

	printf("Starting up winsock\n");
	
	if ( WSAStartup(MAKEWORD(1,1), &WsaDat) !=0  ){  
		printf("Sorry WSAStartup failed exiting.."); 
		return;
	}

	printf("Installing Hooks\n" ) ;
	InstallHooks();

	msg("Executing Buffer...\n\n"); //we are hooked now only use safe display fx
	msg("_ret_____API_________________\n",0x02);

	_asm jmp buf

	//we wont ever get down here..


}





//_______________________________________________ install hooks fx 

void DoHook(void* real, void* hook, void* thunk, char* name){
	
	char err[400];

	if ( !InstallHook( real, hook, thunk) ){ //try to install the real hook here
		sprintf(err,"Install %s hook failed...Error: %s\n", name, &lastError);
		MessageBox(0,err,"",0);
		ExitProcess(0);
	}

}


//Macro wrapper to build DoHook() call
#define ADDHOOK(name) DoHook( name, My_##name, Real_##name, "##name" );	


void InstallHooks(void)
{
 
	ADDHOOK(LoadLibraryA); 
	ADDHOOK(WriteFile);
	ADDHOOK(CreateFileA);
	ADDHOOK(WriteFileEx);
	ADDHOOK(_lcreat);
	ADDHOOK(_lopen);
	ADDHOOK(_lread);
	ADDHOOK(_lwrite);
	ADDHOOK(CreateProcessA);
	ADDHOOK(WinExec);
	ADDHOOK(ExitProcess);
	ADDHOOK(ExitThread);
	ADDHOOK(GetProcAddress);
	ADDHOOK(WaitForSingleObject);
	ADDHOOK(CreateRemoteThread);
	ADDHOOK(OpenProcess);
	ADDHOOK(WriteProcessMemory);
	ADDHOOK(GetModuleHandleA);
	ADDHOOK(accept);
	ADDHOOK(bind);
	ADDHOOK(closesocket);
	ADDHOOK(connect);
	ADDHOOK(gethostbyaddr);
	ADDHOOK(gethostbyname);
	ADDHOOK(gethostname);
	ADDHOOK(listen);
	ADDHOOK(recv);
	ADDHOOK(send);
	ADDHOOK(shutdown);
	ADDHOOK(socket);
	ADDHOOK(WSASocketA);
	ADDHOOK(system);
	ADDHOOK(fopen);
	ADDHOOK(fwrite);
	ADDHOOK(URLDownloadToFileA);
	ADDHOOK(URLDownloadToCacheFile);
	 	
}


