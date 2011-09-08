
vulnps.dll: dlldata.obj vuln_p.obj vuln_i.obj
	link /dll /out:vulnps.dll /def:vulnps.def /entry:DllMain dlldata.obj vuln_p.obj vuln_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del vulnps.dll
	@del vulnps.lib
	@del vulnps.exp
	@del dlldata.obj
	@del vuln_p.obj
	@del vuln_i.obj
