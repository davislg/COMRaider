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

		  Thanks to Shane Hird and his open source axenum for ideas on this 
		  segment of code.

*/


#include "tlbinf32.h" //import of tlbinf32.dll = conflicts with oaidl.h so had to edit tlh

#include <windows.h>
#include <objsafe.h>
#include <stdio.h>
#include <dispex.h>

using namespace TLI;

_TLIApplicationPtr pTlia(__uuidof(TLIApplication)); //one live instance for lifetime of app


void SafeRelease(IUnknown *iUnk){
	
	try{
		if(iUnk != NULL){
			iUnk->Release();
			iUnk = NULL;
		}
	}
	catch(...){
		//MessageBox(0,"Caught err in safe release","",0);
	}
}

bool gs(IObjectSafety *io, REFIID iface, DWORD *x, DWORD flags){
	
	try{
		DWORD ae=0;
		if( S_OK == io->SetInterfaceSafetyOptions(iface, flags, flags)){
			io->GetInterfaceSafetyOptions(iface, x, &ae);
			return true;
		}
	}catch(...){}

	return false;
}


int __stdcall TestClsID(LPOLESTR clsid, DWORD flags, DWORD *dispOk, DWORD *dispExOk, DWORD *perstOK, DWORD *perstStreamOK, DWORD *perstStorageOK ){
	
	//int errLine=0;
	//char buf[100];
	IUnknown *iUnk       = NULL;

	try{
		DWORD ae=0;
		CLSID target;

		IObjectSafety *io    = NULL;
		IPersist      *ip    = NULL;
		IDispatchEx   *ide   = NULL;
		IDispatch     *id    = NULL;
		IPersistStorage *ips = NULL;
		IPersistStream  *ipt = NULL;

		bool hasIP  =false;
		bool hasIDE =false;
		bool hasID  =false;
		bool hasIPS =false;
		bool hasIPT =false;

		*dispOk   = 0;
		*dispExOk = 0;
		*perstOK  = 0;
		*perstStreamOK = 0;
		*perstStorageOK =0;

		CLSIDFromString(clsid,&target);

		if(CoCreateInstance(target, NULL, flags, IID_IUnknown , (void**)&iUnk)!=S_OK){
			//MessageBox(0,"Co Create Failed","",0);
			return 0;
		}

		if(iUnk->QueryInterface(IID_IObjectSafety, (void**)&io)!=S_OK){
			//MessageBox(0,"Could not get IOSafety","",0);
			SafeRelease(iUnk);
			return 0;
		}

		if(iUnk->QueryInterface(IID_IDispatch ,  (void**)&id)==S_OK)   hasID = true;		
		if(iUnk->QueryInterface(IID_IDispatchEx ,(void**)&ide)==S_OK)  hasIDE= true; 

		if( hasID == false && hasIDE == false){
			SafeRelease(iUnk);
			return 0;
		}

		//IPersist is the base interface for three other interfaces: IPersistStorage, IPersistStream, and IPersistFile
		//You must implement the single method of IPersist in implementing any one of the other persistence interfaces
		if(iUnk->QueryInterface(IID_IPersist,        (void**)&ip) == S_OK)  hasIP =true;		
		if(iUnk->QueryInterface(IID_IPersistStream,  (void**)&ips) == S_OK) hasIPS =true;
		if(iUnk->QueryInterface(IID_IPersistStorage, (void**)&ipt) == S_OK) hasIPT =true;
		

		if(hasID)  gs(io,IID_IDispatch,      dispOk,         INTERFACESAFE_FOR_UNTRUSTED_CALLER);
		if(hasIDE) gs(io,IID_IDispatchEx,    dispExOk,       INTERFACESAFE_FOR_UNTRUSTED_CALLER);
		if(hasIP)  gs(io,IID_IPersist,       perstOK,        INTERFACESAFE_FOR_UNTRUSTED_DATA)  ;
		if(hasIPS) gs(io,IID_IPersistStream, perstStreamOK,  INTERFACESAFE_FOR_UNTRUSTED_DATA)  ;
		if(hasIPT) gs(io,IID_IPersistStorage,perstStorageOK, INTERFACESAFE_FOR_UNTRUSTED_DATA)  ;
		 
		SafeRelease(iUnk);
		return 1;

	}
	catch(...){
		
		SafeRelease(iUnk);

		/*
		if(errLine>=3){
			sprintf(buf,"Caught err at %d", errLine);
			MessageBox(0,buf,"",0);
		}*/
	}

	return 0;
	

}

int __stdcall CreateIDispObj(LPOLESTR clsid, LPVOID* outObj){
	
	CLSID target;
	int flags = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;

	CLSIDFromString(clsid,&target);

	try{
		if(CoCreateInstance(target, NULL, flags, IID_IDispatch, outObj)==S_OK) return 1;
		if(CoCreateInstance(target, NULL, flags, IID_IDispatchEx, outObj)==S_OK) return 1;
	}
	catch(...){}

	return 0;
	
}

//main class object has to be owned by vb to keep from mem leak
int __stdcall GetInterfaceInfoForClsid(LPOLESTR clsid, LPVOID* outObj, LPVOID* tmpObj){
	
	CLSID target;
	 
	int flags = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;

	CLSIDFromString(clsid,&target);

	try{
		if(CoCreateInstance(target, NULL, flags, IID_IDispatch, tmpObj)!=S_OK){
			if(CoCreateInstance(target, NULL, flags, IID_IDispatchEx, tmpObj)!=S_OK) return 0;
		}	
		if(pTlia->InterfaceInfoFromObject((IDispatch *)*tmpObj, (struct TLI::InterfaceInfo **)outObj) == S_OK )	return 1;
	}
	catch(...){}

	return 0;
	
}


