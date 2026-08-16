@echo off
title Building 0cell Native Windows Executable...
echo [0cell] Compiling native Windows application...
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:winexe /optimize+ /win32icon:0C.ico /out:0Cell.exe /r:System.Windows.Forms.dll,System.dll ZeroCell.cs

if exist 0Cell.exe (
    echo [0cell] Build SUCCESS! Created 0Cell.exe
) else (
    echo [0cell] Compilation failed.
)
pause
