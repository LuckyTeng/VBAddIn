@ECHO OFF
"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\tlbimp.exe" ^
   "C:\Program Files (x86)\Common Files\DESIGNER\MSADDNDR.DLL" ^
   /out:"MyCompany.Interop.Extensibility.dll" ^
   /keyfile:"D:\test.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 /sysarray

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\tlbimp.exe" ^
   "C:\windows\syswow64\stdole2.tlb" ^
   /out:MyCompany.Interop.Stdole.dll ^
   /keyfile:"D:\test.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\tlbimp.exe" ^
   "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\mso.dll" ^
   /out:MyCompany.Interop.Office14.dll ^
   /keyfile:"D:\test.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 ^
   /reference:MyCompany.Interop.Stdole.dll

PAUSE
CLS

"C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin\tlbimp.exe" ^
   "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\vbe6ext.olb" ^
   /out:MyCompany.Interop.VBAExtensibility.dll ^
   /keyfile:"D:\test.snk" ^
   /strictref:nopia /nologo /asmversion:1.0.0.0 ^
   /reference:MyCompany.Interop.Office14.dll

PAUSE