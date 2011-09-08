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

	thanks to Felix Kasza for example code.

*/

#pragma warning( disable: 4786 )

#include <windows.h>
#include <stdio.h>

#pragma pack( push, before_imagehlp, 8 )
#include <imagehlp.h>
#pragma pack( pop, before_imagehlp )
 

char symSearchPath[500] = {0};
HANDLE hProcess=0;
DWORD symOptions;


int __stdcall InitSearchPath(char* path, HANDLE m_hProcess){

	hProcess = m_hProcess;

	symOptions = SymGetOptions();
	symOptions |= SYMOPT_LOAD_LINES;
	symOptions &= ~SYMOPT_UNDNAME;
	SymSetOptions( symOptions );

	if(strlen(path) > 500 ) path[499]=0;
	strncpy(symSearchPath,path,499);

	return SymInitialize( hProcess, symSearchPath, false );

}

void __stdcall Cleanup(){
	if( hProcess != 0) SymCleanup(hProcess);
	hProcess = 0 ;
}

int __stdcall LoadModuleSymbols(char* imgName, int dllBase, int dllSize){

	char tmp[256] = {0};

	for(int i=0;i<255;i++){ //copy base name of dll to tmp (assumes not full path just name)
		if( imgName[i] == 0) break;
		if( imgName[i] == '.') break;
		tmp[i] = imgName[i];
	}
	
	return SymLoadModule( hProcess, 0, imgName, tmp, dllBase, dllSize ); 

}



int __stdcall GetSymbol(int m_offset, char* name, char* decl, int bufLen){

	DWORD ret;

	IMAGEHLP_SYMBOL *pSym = (IMAGEHLP_SYMBOL *) malloc( sizeof(IMAGEHLP_SYMBOL) + bufLen );
	
	memset( pSym, '\0', sizeof(IMAGEHLP_SYMBOL) + bufLen );
	pSym->SizeOfStruct = sizeof(IMAGEHLP_SYMBOL);
	pSym->MaxNameLength = bufLen;

	if ( SymGetSymFromAddr( hProcess, m_offset, &ret, pSym ) ){		
		UnDecorateSymbolName( pSym->Name, name, bufLen, UNDNAME_NAME_ONLY );
		UnDecorateSymbolName( pSym->Name, decl, bufLen, UNDNAME_COMPLETE );
	    ret = 1;
	}
	else{ 
		ret = 0;
	}

	free( pSym );
	return ret;

}


/*
int main( void )
{
	HANDLE hProcess = GetCurrentProcess();
    
	const int buflen = 1024;
	char name[buflen]; 
	char decl[buflen];
	
	InitSearchPath("F:\\symbols\\XPSP1\\", hProcess);
	LoadModuleSymbols("ntdll.dll", 0x77F50000, 0x000A7000);
	GetSymbol(0x77F94153,name,decl,1024);

	printf("Addr: 0x77F94153\nName: %s\nDecl: %s\n", name,decl);

	return 0;
}
*/
