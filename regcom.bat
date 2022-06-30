set mypath=%*
echo %mypath
echo sysPath %SystemRoot%
copy "%mypath%\MSWINSCK.OCX" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\MSWINSCK.OCX

copy "%mypath%\comctl32.ocx" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\comctl32.ocx

copy "%mypath%\MCI32.OCX" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\MCI32.OCX

copy "%mypath%\MSCOMCTL.OCX" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\MSCOMCTL.OCX

copy "%mypath%\MSINET.OCX" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\MSINET.OCX

copy "%mypath%\RICHTX32.OCX" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\RICHTX32.OCX

PAUSE