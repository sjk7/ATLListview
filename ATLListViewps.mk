
ATLListViewps.dll: dlldata.obj ATLListView_p.obj ATLListView_i.obj
	link /dll /out:ATLListViewps.dll /def:ATLListViewps.def /entry:DllMain dlldata.obj ATLListView_p.obj ATLListView_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del ATLListViewps.dll
	@del ATLListViewps.lib
	@del ATLListViewps.exp
	@del dlldata.obj
	@del ATLListView_p.obj
	@del ATLListView_i.obj
