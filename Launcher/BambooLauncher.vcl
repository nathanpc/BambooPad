<html>
<body>
<pre>
<h1>Build Log</h1>
<h3>
--------------------Configuration: BambooLauncher - Win32 (WCE ARM) Release--------------------
</h3>
<h3>Command Lines</h3>
Creating command line "rc.exe /l 0x409 /fo"ARMRel/BambooLauncher.res" /d UNDER_CE=211 /d _WIN32_WCE=211 /d "UNICODE" /d "_UNICODE" /d "NDEBUG" /d "WIN32_PLATFORM_HPCPRO" /d "ARM" /d "_ARM_" /r "Z:\ProjectsCE\BambooPad\Launcher\BambooLauncher.rc"" 
Creating temporary file "C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP191.tmp" with contents
[
/nologo /W3 /D _WIN32_WCE=211 /D "WIN32_PLATFORM_HPCPRO" /D "ARM" /D "_ARM_" /D UNDER_CE=211 /D "UNICODE" /D "_UNICODE" /D "NDEBUG" /Fp"ARMRel/BambooLauncher.pch" /Yu"stdafx.h" /Fo"ARMRel/" /Oxs /MC /c 
"Z:\ProjectsCE\BambooPad\Launcher\BambooLauncher.cpp"
]
Creating command line "clarm.exe @C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP191.tmp" 
Creating temporary file "C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP192.tmp" with contents
[
/nologo /W3 /D _WIN32_WCE=211 /D "WIN32_PLATFORM_HPCPRO" /D "ARM" /D "_ARM_" /D UNDER_CE=211 /D "UNICODE" /D "_UNICODE" /D "NDEBUG" /Fp"ARMRel/BambooLauncher.pch" /Yc"stdafx.h" /Fo"ARMRel/" /Oxs /MC /c 
"Z:\ProjectsCE\BambooPad\Launcher\StdAfx.cpp"
]
Creating command line "clarm.exe @C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP192.tmp" 
Creating temporary file "C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP193.tmp" with contents
[
commctrl.lib coredll.lib /nologo /base:"0x00010000" /stack:0x10000,0x1000 /entry:"WinMainCRTStartup" /incremental:no /pdb:"ARMRel/BambooLauncher.pdb" /nodefaultlib:"libc.lib /nodefaultlib:libcd.lib /nodefaultlib:libcmt.lib /nodefaultlib:libcmtd.lib /nodefaultlib:msvcrt.lib /nodefaultlib:msvcrtd.lib /nodefaultlib:oldnames.lib" /out:"ARMRel/BambooLauncher.exe" /subsystem:windowsce,2.11 /align:"4096" /MACHINE:ARM 
.\ARMRel\BambooLauncher.obj
.\ARMRel\StdAfx.obj
.\ARMRel\BambooLauncher.res
]
Creating command line "link.exe @C:\DOCUME~1\NATHAN~1\LOCALS~1\Temp\RSP193.tmp"
<h3>Output Window</h3>
Compiling resources...
Compiling...
StdAfx.cpp
Compiling...
BambooLauncher.cpp
Linking...



<h3>Results</h3>
BambooLauncher.exe - 0 error(s), 0 warning(s)
</pre>
</body>
</html>
