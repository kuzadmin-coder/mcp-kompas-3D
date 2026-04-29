set REG_ASM_32=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe
set REG_ASM_64=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

if not EXIST %REG_ASM_32% goto noregist

rem ѕуть к рабочей папке
set WORKDIR=%~dp0\Common
cd %WORKDIR% || goto error


%REG_ASM_32% -codebase %WORKDIR%\KAPITypes.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\Kompas6Constants.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\Kompas6Constants3D.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\Kompas6API5.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\KompasAPI7.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\Kompas6API2D5COM.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\Kompas6API3D5COM.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\CONVERTLIBINTERFACES.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\KompasLibrary.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\KGAXLib.dll || goto error
%REG_ASM_32% -codebase %WORKDIR%\VCHATCHLib.dll || goto error

if not EXIST %REG_ASM_64% goto :registend

%REG_ASM_64% -codebase %WORKDIR%\KAPITypes.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\Kompas6Constants.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\Kompas6Constants3D.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\Kompas6API5.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\KompasAPI7.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\Kompas6API2D5COM.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\Kompas6API3D5COM.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\CONVERTLIBINTERFACES.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\KompasLibrary.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\KGAXLib.dll || goto error
%REG_ASM_64% -codebase %WORKDIR%\VCHATCHLib.dll || goto error

:registend
echo Register Success!
goto end

:noregasm
echo Error: RegAsm (%REG_ASM%) was not found on this system.
goto end

:error
echo Register Error!

:end
pause