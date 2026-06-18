set mypath=%*
echo %mypath
echo sysPath %SystemRoot%
copy "%mypath%\MSWINSCK.OCX" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\MSWINSCK.OCX

copy "%mypath%\comctl32.ocx" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\comctl32.ocx

copy "%mypath%\MCI32.OCX" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\MCI32.OCX

copy "%mypath%\MSCOMCTL.OCX" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\MSCOMCTL.OCX

copy "%mypath%\MSINET.OCX" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\MSINET.OCX

copy "%mypath%\RICHTX32.OCX" %SystemRoot%\System32
REGSVR32 /s %SystemRoot%\System32\RICHTX32.OCX

PAUSE