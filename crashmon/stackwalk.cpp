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

		  should really add loadsymbol stuff here at some point :-\

*/

#pragma warning( disable: 4786 )

#include <windows.h>
#include <stdio.h>
//#include <stdlib.h>

#pragma pack( push, before_imagehlp, 8 )
#include <imagehlp.h>
#pragma pack( pop, before_imagehlp )

extern hProcess;
char swErr[255];
bool initilized = false;

typedef void (__stdcall *tvbCallBack)(int,int,int,int,int);

typedef BOOL (__stdcall *tStackWalk)( DWORD MachineType, HANDLE hProcess,
	HANDLE hThread, LPSTACKFRAME StackFrame, PVOID ContextRecord,
	PREAD_PROCESS_MEMORY_ROUTINE ReadMemoryRoutine,
	PFUNCTION_TABLE_ACCESS_ROUTINE FunctionTableAccessRoutine,
	PGET_MODULE_BASE_ROUTINE GetModuleBaseRoutine,
	PTRANSLATE_ADDRESS_ROUTINE TranslateAddress );

tvbCallBack vbStackWalkCallback;
tStackWalk lpfnStackWalk = NULL;


void __stdcall GetStackWalkErrMsg(char* buf500){
		strcpy(buf500,swErr);
}

int __stdcall InitStackWalk(tvbCallBack lpfnVbCallBack){

	HINSTANCE h = NULL;

	vbStackWalkCallback = lpfnVbCallBack;

	h = LoadLibrary("imagehlp.dll");

	if (h==NULL){
		strcpy(swErr, "Could not load imagehlp.dll");
		return 0;
	}

	lpfnStackWalk = (tStackWalk)GetProcAddress(h,"StackWalk");

	if ( lpfnStackWalk == NULL ){
		strcpy( swErr, "Could not get StackWalk addr");
		FreeLibrary(h);
		return 0;
	}

	initilized = true;
	return 1;

}

void __stdcall GetCallStack(HANDLE hThread, CONTEXT *c)
{
	
	DWORD imageType = IMAGE_FILE_MACHINE_I386;//should use MACHINE_INFO from ImageNtHeader()
	STACKFRAME s;
	CONTEXT cCopy;
	int frame=0;

	memcpy(&cCopy, c, sizeof(cCopy)); //XP trashs its copy. but not 2k weird

	if(!initilized) return;
	memset(&s,0,sizeof(s));

	s.AddrPC.Offset    = cCopy.Eip;
	s.AddrPC.Mode      = AddrModeFlat;
	s.AddrFrame.Offset = cCopy.Ebp;
	s.AddrFrame.Mode   = AddrModeFlat;

	while(lpfnStackWalk(imageType,(void*)hProcess,hThread,&s,&cCopy,NULL,NULL,NULL,NULL)){

		frame++;

		vbStackWalkCallback(
							 frame, 
							 s.AddrPC.Offset, 
							 s.AddrReturn.Offset,
							 s.AddrFrame.Offset, 
							 s.AddrStack.Offset
						   );

	} 

}
