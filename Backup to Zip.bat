:: ----------------------------------- Batch Script Info -----------------------------------
:: -----------------------------------------------------------------------------------------

:: Name:            Backup to Zip
:: Description:     Specify some source files/directories and a destination and the script
::                  will generate a shortcut that when launched will output a timestamped
::                  ZIP each time. Useful for versioning things like game saves, etc.
:: Requirements:    7-Zip (CLI), Powershell (native to Windows)
:: URL:             https://github.com/chocmake/Backup-to-Zip
:: Author:          choc
:: Version:         0.1.1 (2022-07-20)

:: Note:            Keep this script in the same location, otherwise previously created
::                  shortcuts won't be able to find it.

:: Note:            If you'd like to add custom comments to the ZIP filenames add them 
::                  after the date/time and wrap the comments in square brackets.
::                  Eg: '2022-01-25 [my comment here].zip'
::                  This is so the script can detect and auto-rename ZIPs that share the
::                  same date if the 'timeinfilename' setting is disabled below.

:: Tip:             The original source/destination paths can be extracted to a text file
::                  by dragging the shortcut LNK file by itself onto the batch script.

:: -----------------------------------------------------------------------------------------

@echo off

:: --------------------------------------- Settings ----------------------------------------
:: -----------------------------------------------------------------------------------------

:: If enabled will append the current time to ZIP filenames. If disabled only dates will
:: be used and filenames auto-renamed with a counter if two or more share the same date.
set timeinfilename=yes

:: Hour format for filename timestamps
:: Valid values: 12, 24
set hourformat=12

:: If enabled will wrap the shortcut filename in square brackets so when files sorted by
:: name it will appear topmost.
set shortcutbrackets=yes

:: Script window color scheme
:: Valid values: auto, light, dark
set colorscheme=auto

:: ---------------------------------------- Script -----------------------------------------
:: -----------------------------------------------------------------------------------------

setlocal enableextensions enabledelayedexpansion

call :initformat
call :detectcolorscheme
call :detectbinaries 7z powershell
call :scriptargs

:: Check whether launched via shortcut or batch script directly
if not "!cmdcmdline!"=="!cmdcmdline:-btzdest=!" (
    call :cmdheightmanual "4"
    call :inputargs
    call :datetime

    :: Create the archive
    7z a !zippath! !input! >nul 2>&1
    if !errorlevel! neq 0 (
        set "7ziperrorlevel=!errorlevel!"
        call :7ziperror
        )
    ) else (
    :inputprompts
    call :cmdheightmanual "!cmdheight!"
    if not defined input (
        if "!cmdcmdline:~0,-2!"=="!scr!" (
            call :inputprompt input
            call :cmdheight "input" & mode con: lines=!cmdheight! & goto :inputprompts
            ) else (
            if "!cmdcmdline!"=="!cmdcmdline:-btzdest=!" (
                call :inputparse args
                call :cmdheight "input" & mode con: lines=!cmdheight! & goto :inputprompts
                )
            )
        ) else (
        call :echoprompt input
        )
    call :inputprompt destdir
    call :inputprompt destname
    call :outputprep

    :: Create the shortcut (LNK) file with custom icon
    powershell -Command "$Scr = !scr! ; $Args = !args! ; $LinkPathPS = !linkpathps! ; $W = New-Object -comObject WScript.Shell ; $S = $W.CreateShortcut($LinkPathPS) ; $S.TargetPath = $Scr ; $S.Arguments = $Args.trimstart(\""'\"").trimend(\""'\"") ; $S.IconLocation = 'shell32.dll,45' ; $S.Save()"

    :: Show the shortcut file in File Explorer upon completion
    explorer /select,!linkpath!
    )

call :prompt "Complete. Closing..." "timeout"
exit

:: ----------------------------------------- Calls -----------------------------------------
:: -----------------------------------------------------------------------------------------

:7ziperror
    if !7ziperrorlevel! equ 1 (
        set "cmdheightmanual=11"
        for %%i in (!input!) do (
            setlocal disabledelayedexpansion
            set "p=%%i"
            setlocal enabledelayedexpansion
            call :escape p
            for /f "delims=" %%p in ("!p!") do (
                endlocal & endlocal
                set "p=%%p"
                call :unescape p
                if not exist !p! (
                    set "p=ú !p!"
                    call :len p pheight & set "p="
                    call :linescalc pheight
                    set /a "cmdheightmanual=!cmdheightmanual!+(!pheight!+1)" & set "pheight="
                    )
                )
            )
        call :cmdheightmanual "!cmdheightmanual!"
        set "l=Oops. The sources below were not able to be included in the backup. They may be in use by an application or no longer exist in their original location. 7-Zip error code: !7ziperrorlevel!." & call :newlines & echo(
        set "l=The newly created zip has an '[m]' added to its filename to denote missing files." & call :newlines & echo(
        for %%i in (!input!) do (
            setlocal disabledelayedexpansion
            set "p=%%i"
            setlocal enabledelayedexpansion
            call :escape p
            for /f "delims=" %%p in ("!p!") do (
                endlocal & endlocal
                set "p=%%p"
                call :unescape p
                if not exist !p! (
                    set "l=ú !p!" & call :newlines & echo( & set "p="
                    )
                )
            )
        :: Rename newly created ZIP to denote missing files
        ren !zippath! "!destname!!suffix! [m]!ext!"
        ) else (
        call :cmdheightmanual "6"
        set "l=Oops. Backup didn't complete. 7-Zip error code: !7ziperrorlevel!." & call :newlines & echo(
        )
        call :prompt "Press any key to close..." "pause" & exit
    exit /b

:checkforlnk
    set "inputcount=0"
    for %%i in (!input!) do (
        set /a "inputcount+=1"
        )
    if !inputcount! equ 1 (
        set lnkcheck=!input:"=!
        if /i "!lnkcheck:~-3!"=="lnk" (
            call :cmdheightmanual "6"
            :: Obtain the LNK's embedded arguments
            set "input=!input:"='!"
            for /f "usebackq delims=" %%a in (`powershell -Command "$LinkPathPS = !input! ; $W = New-Object -comObject WScript.Shell ; $S = $W.CreateShortcut^($LinkPathPS^).Arguments ; echo $S"`) do (
                setlocal disabledelayedexpansion
                set "args=%%a"
                setlocal enabledelayedexpansion
                set "args=!args:"=///!"
                call :escape args
                for /f "delims=" %%s in ("!args!") do (
                    endlocal & endlocal
                    set "args=%%s"
                    call :unescape args
                    set "args=!args:///="!"
                    set "args=!args:~1,-1!"
                )
            )
            :: Write arguments to text file
            call :inputargs
            set "input=!input:""="!"
            if "!destname!"=="" (
                set "destnamefilename="
                ) else (
                set "destnamefilename=!destname! - "
                )
            for %%l in ("!lf!") do set "input=!input:" "=%%l!"
                (
                    echo Source^(s^):!lf!!input!!lf!
                    echo Destination:!lf!"!destdir:\\=!"!lf!
                    echo Name:!lf!!destname!
                ) > "!brkl!!linkprefix! - !destnamefilename!Paths!brkr!.txt"
            set "l=Extracted shortcut paths to text file." & call :newlines & echo(
            call :prompt "Press any key to close..." "pause"
            )
        )
    exit /b

:cmdheight
    if "%~1"=="input" (
        for /f "delims=" %%p in ("!input!") do (
            setlocal disabledelayedexpansion
            set "p=%%p"
            setlocal enabledelayedexpansion
            set "p=!p:"=!"
            call :len p pheight
            call :linescalc pheight
            set /a "inheightall=!inheightall!+(!pheight!+1)"
            for %%h in (!inheightall!) do endlocal & endlocal & set "inheightall=%%h"
            set /a "cmdheight=!inheightall!+6"
            )
        ) else (
        call :len %~1 %~1height
        call :linescalc %~1height
        set /a "cmdheight=!cmdheight!+!%~1height!+1"
        if "%~1"=="destname" set /a "cmdheight=!cmdheight!-2"
        )
    exit /b

:cmdheightmanual
    mode con lines=%~1
    if !cmdpadwidth! leq 3 echo(
    exit /b

:detectbinaries
    set "binarycount=0"
    for %%b in (%*) do (
        where /q %%b
        if errorlevel 1 (
            set /a "binarycount+=1"
            for %%i in (!binarycount!) do set "binarymissing[%%i]=%%b.exe"
            set "error=1" & set "errorbinariesmissing=1"
            )
        )
    if defined errorbinariesmissing (
        set "cmdheight=8" & for /l %%i in (1,1,!binarycount!) do set /a "cmdheight+=2"
        call :cmdheightmanual "!cmdheight!"
        set "l=The program(s) below couldn't be found. Please add their directory to the Windows PATH environment variable so they can be detected. Refer to the Github readme for more info." & call :newlines & echo(
        for /l %%i in (1,1,!binarycount!) do (
            set "l=ú !binarymissing[%%i]!" & call :newlines & echo(
            )
        call :prompt "Press any key to close..." "pause"
        pause
        )
    exit /b

:detectcolorscheme
    if /i "!colorscheme!"=="auto" (
        set "regquery="HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "AppsUseLightTheme""
        reg query !regquery! >nul 2>&1
        if not errorlevel 1 (
            for /f "tokens=4 delims=x " %%a in ('reg query !regquery!') do (
                if "%%a"=="0" (set "colorscheme=dark") else (set "colorscheme=light")
                )
            ) else (
            set "colorscheme=light"
            )
        )
    set "regquery=" & if "!colorscheme!"=="dark" (set "cmdcolor=0F") else (set "cmdcolor=F0") & color !cmdcolor!
    exit /b

:datetime
    :: Obtain timestamp
    setlocal
    set "YYYY="
    for /f "tokens=1-6 delims=/: " %%a in ('robocopy "|" . /njh') do if not defined YYYY (
        set "YYYY=%%a" & set "MM=%%b" & set "DD=%%c"
        set "H=%%d" & set "M=%%e" & set "S=%%f"
        if "!hourformat:~0,1!"=="1" set "TOD= AM" & if !H! gtr 12 set /a "H=!H!-12" & set "TOD= PM"
        )
    set "H=0!H!" & set "datetime=!YYYY!-!MM!-!DD!#(!H:~-2!.!M!.!S!!TOD!)"
    for /f "tokens=1,2 delims=#" %%a in ("!datetime!") do endlocal & set "date=%%a" & set "time=%%b"

    :: Filename formatting
    set "suffix=!date!" & if defined destname set "suffix= - !suffix!"
    if "!timeinfilename:~0,1!"=="y" (
        set "suffix=!suffix! !time!"
        ) else (
        if exist "!destdir!!destname!!suffix!!ext!" (
            ren "!destdir!!destname!!suffix!!ext!" "!destname!!suffix! (1)!ext!"
            ) else (
            for %%f in ("!destdir!!destname!!suffix! [*]!ext!") do (
                setlocal disabledelayedexpansion
                set "f=%%f"
                setlocal enabledelayedexpansion
                for /f "tokens=1,2 delims=[]" %%a in ("!f!") do (
                    setlocal disabledelayedexpansion
                    set "b=%%b"
                    setlocal enabledelayedexpansion
                    ren "!f!" "!destname!!suffix! (1) [!b!]!ext!"
                    endlocal & endlocal
                    )
                endlocal & endlocal
                )
            )
        set "suffixinit=!suffix!"
        call :renamecounter
        )
    set "zippath="!destdir!!destname!!suffix!!ext!""
    exit /b

:echoprompt
    set "s=!%~1!"
    set /a "inws=!cmdpadleftaddwidth!-3"
    if "!s!"=="!input!" (
        set "count=0"
        for /f "delims=" %%d in ("!input!") do (
            set /a "count+=1"
            setlocal disabledelayedexpansion
            set "d=%%d"
            setlocal enabledelayedexpansion
            if !count! equ 1 set "l=!%~1echotext! !d!" & call :newlines "alt" & echo(
            if !count! gtr 1 set "l=!ws:~-%inws%!ú  !d!" & call :newlines "alt" & echo(
            endlocal & endlocal
            )
        ) else (
        if "!s!"=="!destdir!" set "s="!s!"" & set "s=!s:""="!"
        if "!s!"=="!destname!" if "!destname!"=="[empty]" set "s="
        set "l=!%~1echotext! !s!" & call :newlines "alt" & echo(
        )
        set "s="
    exit /b

:escape
    set "s=!%~1:"=!"
    setlocal disabledelayedexpansion
    set "s=%s:!=###esc-excl###%"
    set "s=%s:^=###esc-caret###%"
    setlocal enabledelayedexpansion
    for %%a in ("!s!") do endlocal & endlocal & set "%~1=%%a"
    exit /b

:unescape
    set "s=!%~1:"=!"
    rem Checks if string contains exclamation point to later adjust number of carets for unescaping
    set "doublecaret=" & if not "!s!"=="!s:###esc-excl###=!" set "doublecaret=1"
    setlocal disabledelayedexpansion
    set "s=%s:###esc-excl###=^!%"
    if defined doublecaret (set "s=%s:###esc-caret###=^^%") else (set "s=%s:###esc-caret###=^%")
    setlocal enabledelayedexpansion
    for %%a in ("!s!") do endlocal & endlocal & set "%~1=%%a"
    exit /b

:escquote
    if "%~2"=="wrap" (
        set "%~1=\""!%~1!\"""
        ) else (
        set "%~1=!%~1:"=\""!"
        )
    exit /b

:escps
    set "%~1=@'!lf!!%~1!!lf!'@"
    exit /b

:extractarg
    set "%~2=!args:*%~1 =!"
    exit /b

:initformat
    for /f %%a in ('"prompt $H &echo on &for %%b in (1) do rem"') do set bs=%%a
    set "cmdpadtextwidth=60" & set "cmdpadwidth=3" & set "cmdpadleftaddwidth=16" & set /a "cmdwidth=!cmdpadtextwidth!+(!cmdpadwidth!*2)" & set "cmdheight=7"
    set "ws=                         " & set "ws=!ws!!ws!!ws!!ws!" & set "cmdpad=!ws:~-%cmdpadwidth%!"
    mode con: cols=!cmdwidth! lines=!cmdheight! & title Backup to Zip
    set "rencount=1"
    set "linkprefix=Backup to zip"
    set "inputprompttext=Source(s):" & set "inputechotext=Source(s)    ú "
    set "destdirprompttext=Destination:" & set "destdirechotext=Destination  ú "
    set "destnameprompttext=Backup name (optional):" & set "destnameechotext=Backup name  ú "
    if "!shortcutbrackets:~0,1!"=="y" (
        set "brkl=["
        set "brkr=]"
        )
    set lf=^


    exit /b

:inputargs
    call :extractarg "-btzname" destname
    call :trimargs "-btzname" destname
    if "!destname!"=="[empty]" set "destname="
    call :extractarg "-btzdest" destdir
    call :trimargs "-btzdest" destdir
    set "destdir=!destdir!\"
    set "ext=.zip"
    set "input=!args!" & set "args="
    exit /b

:inputparse
    set input="!%~1:"=!"
    for /f "tokens=1 delims=:" %%d in (!input!) do (
        for %%l in ("!lf!") do (
            set input=!input:%%d:\=%%l%%d:\!
            )
        )
    set "input=!input: "="!" & set "input=!input:~3!"
    call :checkforlnk
    exit /b

:inputprompt
    if not defined %~1 (
        set "l=!%~1prompttext!" & call :newlines & echo(
        set /p "%~1=!bs!!cmdpad!> "
        if "!%~1: =!"=="" set "%~1="
        if "!%~1!"==" =" set "%~1=" & rem Detect empty Enter only input
        if defined %~1 (
            if "%~1"=="input" call :inputparse input
            if "%~1"=="destdir" (
                set isfile=1&pushd "!%~1!" 2>nul&&(popd&set isfile=)||(if not exist "!%~1!" set isfile=)
                if defined isfile set "%~1=" & goto :inputprompts
                )
            :: Remove double quotes if present
            set "destdir=!destdir:"=!"
            :: Trim any trailing backslash
            if "!destdir:~-1!"=="\" (
                call :len destdir destdirlen
                set /a "destdirlen-=1"
                for %%l in (!destdirlen!) do set "destdir=!destdir:~0,%%l!"
            )
            :: Restore double quotes
            set "destdir="!destdir!""
            :: Check if directory exists, otherwise re-prompt for input
            if not exist !destdir! set "destdir=" & goto :inputprompts
            call :cmdheight "%~1" & goto :inputprompts
            ) else (
            if "%~1"=="destname" set "destname=[empty]"
            goto :inputprompts
            )
        ) else (
        call :echoprompt "%~1"
        )
    exit /b

:len
    set "s=!%~1!#"
    set "len=0"
    for %%p in (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) do (
        if "!s:~%%p,1!" neq "" ( 
            set /a "len+=%%p"
            set "s=!s:~%%p!"
            )
        )
    set "s=" & set "%~2=!len!"
    exit /b

:linescalc
    set "l=!%~1!"
    set /a "l=!l!*100"
    set /a "l=!l!/(!cmdpadtextwidth!-!cmdpadleftaddwidth!)"
    set /a "mod=!l! %% 100"
    if !l! lss 100 set "l=100"
    if !mod! gtr 0 (
        if !l! gtr 100 (
            set /a "add=100-!l:~-1!"
            set /a "l=!l!+!add!" & set "add="
            )
        )
    set /a "l=!l!/100" & set "%~1=!l!" & set "mod=" & set "l="
    exit /b

:newlines
    set l=!l:"=!
    call :len l llen
    if "%~1"=="alt" (
        set /a "lineswidth=!cmdwidth!-(!cmdpadwidth!*2)"
        set /a "lineswidth2=!cmdpadtextwidth!-!cmdpadleftaddwidth!"
        ) else (
        set "lineswidth=!cmdpadtextwidth!"
        set "lineswidth2=!cmdpadtextwidth!"
        )
    :newlinessub
    if !llen! gtr !lineswidth! (
        for %%w in (!lineswidth!) do (
            echo !cmdpad!!l:~0,%%w!
            set l=!l:~%%w!
            )
        if "%~1"=="alt" (
            if defined l goto :newlinessubalt
            :newlinessubalt
            if !llen! gtr !lineswidth! (
                for %%w in (!lineswidth2!) do (
                    echo !cmdpad!!ws:~-%cmdpadleftaddwidth%!!l:~0,%%w!
                    set l=!l:~%%w!
                    )
                if defined l goto :newlinessubalt
                )
            ) else (
            if defined l goto :newlinessub
            )
        )
    if !llen! leq !lineswidth! (
        echo !cmdpad!!l!
        )
    set "l=" & set "llen="
    exit /b

:outputprep
    if "!destname!"=="[empty]" set "destname="
    for %%l in ("!lf!") do set "input=!input:%%l=" "!"
    call :escquote input
    if defined destname set "linkprefix=!linkprefix! - "
    set "destdir=!destdir:"=!"
    set "linkpath="!destdir!\!brkl!!linkprefix!!destname!!brkr!.lnk"" & set "linkpathps=!linkpath!"
    if not defined destname set "destname=[empty]"
    call :escquote destdir "wrap"
    call :escquote destname "wrap"
    set "args='!input! -btzdest !destdir! -btzname !destname!'"
    call :escps scr
    call :escps args
    call :escps linkpathps
    exit /b

:prompt
    setlocal
    set "t=%~2" & set "p=!bs!!cmdpad!>" & set "s=!bs!"
    if defined t (
        if "!t!"=="pause" (
            pause >nul|set /p "=!p!  %~1   !s!" & exit
            )
        if "!t!"=="timeout" (
            timeout 1 >nul|set /p "=!p!  %~1   !s!" & exit
            )
        ) else (
        <nul set /p "=!p!  %~1   !s!" & exit /b
        )
    endlocal
    exit /b

:renamecounter
    :renamecountersub
    if exist "!destdir!!destname!!suffixinit! (!rencount!*!ext!" (
        set /a "rencount+=1"
        set "suffix=!suffixinit! (!rencount!)"
        goto :renamecountersub
        )
    exit /b

:scriptargs
    for %%f in ("!cmdcmdline!") do (
        setlocal disabledelayedexpansion
        set scr="%~f0"
        set scrdir="%~dp0"
        setlocal enabledelayedexpansion
        call :len scr scrlen
        call :len scrdir scrdirlen
        set "scrall=!scrlen! !scrdirlen!"
        for /f "tokens=1,2 delims= " %%a in ("!scrall!") do endlocal & endlocal & set "scrlen=%%a" & set "scrdirlen=%%b"
        )
    set "cmdcmdline=!cmdcmdline:~32!" & call set scr=%%cmdcmdline:~0,!scrlen!%%
    call set scrdir=%%cmdcmdline:~0,!scrdirlen!%%
    set "scrdir=!scrdir:~1,-1!"
    set "args=!cmdcmdline:~%scrlen%,-1!" & set "args=!args:* =!"
    exit /b

:trimargs
    set "s1=%~1" & set "s2=!%~2!"
    for %%n in (!s1!) do (
        for %%q in (!s2!) do (
            set "args=!args:%%n %%q=!" & set "args=!args:~0,-1!"
            set "%~2=!s2:"=!"
            )
        )
    set "s1=" & set "s2="
    exit /b

endlocal