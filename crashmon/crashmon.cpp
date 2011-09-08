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

		  portions of this file were adapted from the MS SDK Deb debugger example

*/

#include <windows.h>
#include <stdio.h>

#define DBG_COND_EXECUTE 0 
#define DBG_COND_WRITE  1 
#define DBG_COND_READWRITE  3 

int  targetHWND = 0;
int  pid = 0;
int  isDebugging = 0 ;
int  debugDecision = 0 ;
int  hProcess =0;

char proc[1000];
char errMsg[500];

DEBUG_EVENT  DebugEvent;
HANDLE		 hDebugEventThread;


typedef struct{
    int dwFlag;
    int cbSize;
    int lpData;
} cpyData;



DWORD WINAPI DebugEventThread(void*);

void __stdcall SetTargetHwnd(int hwnd){ targetHWND = hwnd;}
void __stdcall GetErrorMessage(char* buf500){ strcpy(buf500,errMsg);}
void __stdcall Continue(int decesion){ 	debugDecision = decesion /*DBG_CONTINUE*/; }
void __stdcall StopDebugger(){ 	TerminateThread(hDebugEventThread,0); isDebugging = 0;}
int  __stdcall ActivePID(){ if(isDebugging==1){return pid;}else{return 0;}}

int DbgMsg(DEBUG_EVENT *Buffer){
  
  if(targetHWND==0){
	  MessageBox(0,"SetTargetHWND First!","",0);
	  return 0;
  }
  
  cpyData cpStructData;
  
  cpStructData.cbSize = sizeof(DEBUG_EVENT);
  cpStructData.lpData = (int)Buffer;
  cpStructData.dwFlag = 3;
  
  return SendMessage((HWND)targetHWND, WM_COPYDATA, (WPARAM)targetHWND,(LPARAM)&cpStructData);

} 

int __stdcall GetTibForThread(HANDLE hThread){

	CONTEXT context; 
	LDT_ENTRY sel; 
	int tib;

	context.ContextFlags = CONTEXT_FULL; 
	GetThreadContext(hThread,&context); 
	GetThreadSelectorEntry(hThread, context.SegFs, &sel); 
	tib = sel.BaseLow|(sel.HighWord.Bytes.BaseMid<<16)|(sel.HighWord.Bytes.BaseHi<<24); 
	
	return tib;

}


int __stdcall LaunchDebugedProcess(char* filename){
	
   if(isDebugging==1){
		sprintf(errMsg,"Only one debuggee per session!");
		return 0;
   }
			
   pid=0;	
   strncpy(proc,filename,1000);
 
   //vb is single threaded, so debugger has to be in own thread
   //any communication between debugger thread and vb thread has
   //to use interprocess communication techniques.

   if( !( hDebugEventThread = CreateThread(
 							   (LPSECURITY_ATTRIBUTES) 0L,
							   (DWORD) 0,
							   (LPTHREAD_START_ROUTINE) DebugEventThread,
							   (LPVOID) 0,
							   0L,
							   (LPDWORD) 0) ) ) {

		sprintf(errMsg,"Debug thread failed to start");
		return 0;

	}
		
	  return 1;
}


BOOL LaunchProcess(LPTSTR lpszFileName){

  static STARTUPINFO           StartupInfo;
  static LPSTARTUPINFO         lpStartupInfo = &StartupInfo;
  static PROCESS_INFORMATION   ProcessInfo;
  static LPPROCESS_INFORMATION lpProcessInfo = &ProcessInfo;
	
  memset(&StartupInfo,0,sizeof(STARTUPINFO));
  memset(&ProcessInfo,0,sizeof(PROCESS_INFORMATION));

  lpStartupInfo->cb          = sizeof( STARTUPINFO );
  lpStartupInfo->lpTitle     = "SandSprite VB Debugger Extension";
  lpStartupInfo->wShowWindow = SW_SHOWDEFAULT;

  lpProcessInfo->hProcess = 0L;

  if( !CreateProcess(0L, lpszFileName,(LPSECURITY_ATTRIBUTES) 0L, 
					(LPSECURITY_ATTRIBUTES) 0L, TRUE, 
					 DEBUG_PROCESS | CREATE_NEW_CONSOLE,
					(LPVOID) 0L, (LPTSTR) 0L,lpStartupInfo,lpProcessInfo )
	 ) {

			switch( GetLastError() ) {
				  case ERROR_FILE_NOT_FOUND:
					sprintf(errMsg,"This file does not exist.");
					break;
				  case ERROR_ACCESS_DENIED:
					sprintf(errMsg,"Access denied.");
					break;
				  case ERROR_FILE_INVALID:
					sprintf(errMsg,"Invalid file." );
					break;
				  case ERROR_FILE_CORRUPT:
					sprintf(errMsg,"The file is corrupt." );
					break;
				  case ERROR_BAD_EXE_FORMAT:
					sprintf(errMsg,"The file has a bad format." );
					break;
				  default:
					break;
					//sprintf(errMsg,"CreateProcess()" );
			}
			return( FALSE );
  }
  else {
			pid = lpProcessInfo->dwProcessId; 
			hProcess = (int)lpProcessInfo->hProcess ;	
  }
  return(TRUE);
}



DWORD WINAPI DebugEventThread(void* x){


	if( !LaunchProcess( proc ) ){
		sprintf(errMsg,"DebugNewProcess Failed...");
		isDebugging = 0;
		ExitThread( FALSE );
	}
	
    isDebugging = 1;

	for(;;) {

		if( !WaitForDebugEvent( &DebugEvent, INFINITE ) ) {
		  sprintf(errMsg,"Failed to attach to Debuggee...");
		  isDebugging = 0;
		  ExitThread(FALSE);
		  break;
		}
		
		debugDecision = 0;
		DbgMsg(&DebugEvent);		 
		
		while(debugDecision==0){
			Sleep(60);
		}

		ContinueDebugEvent( DebugEvent.dwProcessId,DebugEvent.dwThreadId, debugDecision );
		
		  
	}

	return 1;
}


