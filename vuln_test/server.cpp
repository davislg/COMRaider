// server.cpp : Implementation of Cserver
#include "stdafx.h"
#include "Vuln.h"
#include "server.h"
#include <stdio.h>

/////////////////////////////////////////////////////////////////////////////
// Cserver

void main(void){};

//these are not necessairly realistic vulns that you would find in an ATL
//or MFC project but we just need some exploitable samples for testing
//comraider against.

STDMETHODIMP Cserver::Method1(BSTR strin, long *retVal)
{
	//classic string based stack buffer overflow example
	USES_CONVERSION; 
	char sPath[200] = {0};
	char *tmp = W2A(strin);
	strcpy(sPath, tmp);
	MessageBox(0,sPath,"My Vuln Message",0);
	return S_OK;
}

STDMETHODIMP Cserver::Method4(BSTR pPath, BSTR msg, long *retVal)
{
	 
	//example scenario where local err handling could mask
	//exploitable fuzz condition. Which is why we need to
	//launch our fuzz target under debugger and examine exceptions

	USES_CONVERSION;
	char sPath[300];
	char *tmp = W2A(msg);

	//DebugBreak(); 
	strncpy(sPath, tmp, 219); 
	SysReAllocString(&msg , A2W(sPath));

	try{
		this->Method1(msg,retVal);
	}
	catch(...){
		MessageBox(0,"Caught Err\n\nLocal Error Handling Can Mask Exploitable Conditions","",0);
	}

	return S_OK;
}


STDMETHODIMP Cserver::Method2(long lin, long *retVal)
{	
	//add some unchecked signed/unsigned bug here
	return S_OK;
}

STDMETHODIMP Cserver::Method3(VARIANT vin, long *retVal)
{
	
	if(vin.vt == VT_BSTR){

		USES_CONVERSION;
		char buf[1000];
		char *tmp = W2A(vin.bstrVal);
		if(strlen(tmp) > 1000) tmp[999]=0;
		sprintf(buf, tmp); //format string bug
		MessageBox(0,buf,"",0);
	
	}

	
	return S_OK;
}


STDMETHODIMP Cserver::HeapCorruption(BSTR strIn, long bufSize, long *retVal)
{

	USES_CONVERSION;
	char *tmp = W2A(strIn);
	char *x = (char*)malloc(bufSize);

	strcpy(x,tmp);
	MessageBox(0,x,"",0);
	
	free(x);
	return S_OK;

}
