:: ----------------------------------- Batch Script Info -----------------------------------
:: -----------------------------------------------------------------------------------------

:: Name:            Backup to Zip
:: Description:     Specify some source files/directories and a destination and the script
::                  will generate a shortcut that when launched will output a timestamped
::                  zip each time. Useful for versioning things like game saves, etc.
:: Requirements:    7-Zip (CLI), Powershell (native to Windows)
:: URL:             https://github.com/chocmake/Backup-to-Zip
:: Author:          choc
:: Version:         0.2 (2023-02-10)

:: Note:            Keep this script in the same location, otherwise previously created
::                  shortcuts won't be able to find it.

:: Note:            If you'd like to add custom comments to the backup filenames add them 
::                  after the date/time and wrap the comments in square brackets.
::                  Eg: '2022-01-25 [my comment here].zip'
::                  This is so the script can detect and auto-rename backups that share the
::                  same date if the 'timeinfilename' setting is disabled below.

:: Tip:             The original source/destination paths can be extracted to a text file
::                  by dragging the shortcut LNK file by itself onto the batch script.

:: -----------------------------------------------------------------------------------------

@echo off

:: --------------------------------------- Settings ----------------------------------------
:: -----------------------------------------------------------------------------------------

:: Archive type to use for output.
:: Valid values: zip, 7z
set archivetype=zip

:: If enabled will append the current time to backup filenames. If disabled only dates will
:: be used and filenames auto-renamed with a counter if two or more share the same date.
set timeinfilename=yes

:: If enabled will copy the Date Created and Date Accessed timestamp(s) of the source.
:: Normally only the Date Modified timestamp(s) are copied.
:: Note: this setting only works on newer versions of 7-Zip (tested on v21.01), make sure
:: you've updated 7-Zip before enabling otherwise you'll see an error when LNK launched.
set preservealltimestamps=no

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

call :detectcolorscheme
call :initformat
call :detectbinaries 7z powershell
call :scriptargs

:: Check whether launched via shortcut or batch script directly
if not "!cmdcmdline!"=="!cmdcmdline:-btzdest=!" (
    call :cmdheightmanual "4"
    call :inputargs
    call :filenameformat

    :: Create the archive
    !7z! !artype! !exttimestamps! a !zippath! !input! >nul 2>&1
    if !errorlevel! neq 0 call :error "7zip" "!errorlevel!"
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
    :: First checks for and removes duplicates of any manually entered paths to avoid 7-Zip error when LNK later used.
    :: Shortcut arguments are concatenated here instead of in batch to avoid Powershell replacing any consecutive spaces in paths with a single space. All double quotes from here-strings are removed prior to make this possible, then added back below.
    powershell -Command "$host.UI.RawUI.WindowTitle = '!title!' ; $Scr = !scr! ; $Input = (!input!.split('!lf!') | Select-Object -Unique) -join '!lf!' ; $DestDir = !destdir! ; $DestName = !destname! ; $LinkPathTempPS = !linkpathtempps! ; $W = New-Object -comObject WScript.Shell ; $S = $W.CreateShortcut($LinkPathTempPS) ; $S.TargetPath = $Scr ; $S.Arguments = '\""' + $Input.replace('!lf!','\"" \""') + '\"" -btzdest \""' + $DestDir + '\"" -btzname \""' + $DestName + '\""' ; $S.IconLocation = 'shell32.dll,45' ; $S.Save()"

    :: Move LNK from temp directory to destination as workaround for WScript.Shell limitations
    move "!linkpathtemp!" "!linkpathdest!" >nul

    :: Show the shortcut file in File Explorer upon completion
    explorer /select,"!linkpathdest!"
    )

call :prompt "Complete. Closing..." "timeout"
exit

:: ----------------------------------------- Calls -----------------------------------------
:: -----------------------------------------------------------------------------------------

:checkforlnk
    rem Determine if LNK for extracting arguments to text file
    set "inputcount=0"
    for %%i in (!input!) do (
        set /a "inputcount+=1"
        )
    if !inputcount! equ 1 (
        set lnkcheck=!input:"=!
        if /i "!lnkcheck:~-3!"=="lnk" (
            call :cmdheightmanual "6"
            rem Obtain the LNK's embedded arguments
            rem Can't use Powershell here-strings in for loop. Copy of LNK created to read from for scenarios where path contains problematic characters for WScript.Shell.
            set "linkpathtemp=!temp!\[Backup-to-Zip-Temp].lnk"
            copy !input! "!linkpathtemp!" >nul
            for /f "usebackq delims=" %%a in (`powershell -Command "$LinkPathTemp = '!linkpathtemp!' ; $W = New-Object -comObject WScript.Shell ; $S = $W.CreateShortcut^($LinkPathTemp^).Arguments ; echo $S"`) do (
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
            rem Delete the temp LNK copy
            del "!linkpathtemp!"
            rem Write arguments to text file
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

:datetime
    :: Obtain timestamp
    setlocal
    set "YYYY="
    for /f "tokens=1-6 delims=/: " %%a in ('robocopy "|" . /njh') do if not defined YYYY (
        set "YYYY=%%a" & set "MM=%%b" & set "DD=%%c"
        set "H=%%d" & set "M=%%e" & set "S=%%f"
        if "!hourformat:~0,1!"=="1" (
            set "TOD= AM"
            if !H! gtr 12 (
                set /a "H=!H!-12"
                set "TOD= PM"
            )
            if !H! equ 00 (
                set "H=12"
                )
            )
        )
    set "H=0!H!" & set "datetime=!YYYY!-!MM!-!DD!#(!H:~-2!.!M!.!S!!TOD!)"
    for /f "tokens=1,2 delims=#" %%a in ("!datetime!") do endlocal & set "date=%%a" & set "time=%%b"
    exit /b

:detectbinaries
    set "binarymissingcount=0"
    for %%b in (%*) do (
        where /q %%b
        if errorlevel 1 (
            if not "%%b"=="7z" call :detectbinariesmark "%%b"
            rem Program-specific default install path check, in case user hasn't configured PATH
            if "%%b"=="7z" (
                rem <path:pattern>
                where /q "C:\Program Files\7-Zip\:7z.exe"
                if errorlevel 1 (
                    call :detectbinariesmark "%%b"
                    ) else (
                        set "7z="C:\Program Files\7-Zip\7z.exe""
                    )
                )
            ) else (
            if "%%b"=="7z" set "%%b=%%b"
            )
        )
    if !binarymissingcount! gtr 0 call :error "binariesmissing"
    exit /b

:detectbinariesmark
    set /a "binarymissingcount+=1"
    for %%i in (!binarymissingcount!) do set "binarymissing[%%i]=%~1.exe"
    exit /b

:detectcolorscheme
    rem Check registry for current theme
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
    set "regquery=" & if /i "!colorscheme!"=="dark" (set "cmdcolor=0F") else (set "cmdcolor=F0") & color !cmdcolor!
    exit /b

:echoprompt
    set "e=!%~1!"
    set /a "inws=!cmdpadleftaddwidth!-3"
    if "!e!"=="!input!" (
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
        if "!e!"=="!destdir!" set "e="!e!"" & set "e=!e:""="!"
        if "!e!"=="!destname!" if "!destname!"=="[empty]" set "e="
        set "l=!%~1echotext! !e!" & call :newlines "alt" & echo(
        )
        set "e="
    exit /b

:error
    if "%~1"=="binariesmissing" (
        set "cmdheight=8" & for /l %%i in (1,1,!binarymissingcount!) do set /a "cmdheight+=2"
        call :cmdheightmanual "!cmdheight!"
        set "l=The program(s) below couldn't be found. Please add their directory to the Windows PATH environment variable so they can be detected. Refer to the Github readme for more info." & call :newlines & echo(
        for /l %%i in (1,1,!binarymissingcount!) do (
            set "l=ú !binarymissing[%%i]!" & call :newlines & echo(
            )
        call :prompt "Press any key to close..." "pause"
    )
        
    if "%~1"=="lnklimcheckreached" (
        set "cmdheight=8"
        call :cmdheightmanual "!cmdheight!"
        set "l=Oops. The length of the path(s) in the input is expected to exceed the maximum possible for the LNK shortcut field. Try again with fewer inputs." & call :newlines & echo(
        call :prompt "Press any key to close..." "pause"
    )

    if "%~1"=="7zip" (
        if "%~2"=="1" (
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
            set "l=Oops. The sources below were not able to be included in the backup. They may be in use by an application or no longer exist in their original location. 7-Zip error code: %~2." & call :newlines & echo(
            set "l=The newly created backup has an '[m]' added to its filename to denote missing files." & call :newlines & echo(
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
            set "l=Oops. Backup didn't complete. 7-Zip error code: %~2." & call :newlines & echo(
            )
        call :prompt "Press any key to close..." "pause" & exit
    )
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

:escdelims
    rem Can replace `=` and `*` characters in strings
    rem https://www.dostips.com/forum/viewtopic.php?f=3&t=1485&start=30#p50132
    setlocal disabledelayedexpansion
    call set "string=[%%%~1%%]"
    set "repl=%~3"
    set "result="
    call :len string slen
    :escdelims.loop
    for /f "delims=%~2 tokens=1*" %%s in ("%string%") do (
      set "head=%%s"
      set "tail=%%t"
    )
    set "result=%result%%head%"
    call :len head hlen
    call :len tail tlen
    set /a "n=slen-hlen-tlen"
    setlocal enabledelayedexpansion
    for /l %%n in (1,1,%n%) do set "result=!result!!repl!"
    endlocal & set "result=%result%"
    if defined tail (
      set "string=%tail%"
      set "slen=%tlen%"
      goto :escdelims.loop
    )
    endlocal & set "%~4=%result:~1,-1%"
    exit /b

:escps
    set "%~1=@'!lf!!%~1!!lf!'@"
    exit /b

:extractarg
    set "%~2=!args:*%~1 =!"
    exit /b

:filenameformat
    call :datetime
    set "suffix=!date!" & if defined destname set "suffix= - !suffix!"
    if /i "!timeinfilename:~0,1!"=="y" (
        set "suffix=!suffix! !time!"
        ) else (
        if exist "!destdir!!destname!!suffix!!ext!" (
            ren "!destdir!!destname!!suffix!!ext!" "!destname!!suffix! (1)!ext!"
            ) else (
            for %%f in ("!destdir!!destname!!suffix! [*]!ext!") do (
                setlocal disabledelayedexpansion
                set "f=%%f"
                setlocal enabledelayedexpansion
                for /f "tokens=2,3 delims=[" %%a in ("!f!") do (
                    setlocal disabledelayedexpansion
                    set "a=%%a"
                    set "b=%%b"
                    setlocal enabledelayedexpansion
                    rem Check if filename already has secondary square brackets (eg: `[m]`)
                    if exist "!destdir!!destname!!suffix! [*] [*]!ext!" (
                        ren "!f!" "!destname!!suffix! (1) [!a![!b!"
                        ) else (
                        ren "!f!" "!destname!!suffix! (1) [!a!"
                        )
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

:initformat
    for /f %%a in ('"prompt $H &echo on &for %%b in (1) do rem"') do set bs=%%a
    set "cmdpadtextwidth=60" & set "cmdpadwidth=3" & set "cmdpadleftaddwidth=16" & set /a "cmdwidth=!cmdpadtextwidth!+(!cmdpadwidth!*2)" & set "cmdheight=7"
    set "ws=                         " & set "ws=!ws!!ws!!ws!!ws!" & set "cmdpad=!ws:~-%cmdpadwidth%!"
    mode con: cols=!cmdwidth! lines=!cmdheight! & title Backup to Zip
    set "linkprefix=Backup to zip"
    set "inputprompttext=Source(s):" & set "inputechotext=Source(s)    ú "
    set "destdirprompttext=Destination:" & set "destdirechotext=Destination  ú "
    set "destnameprompttext=Backup name (optional):" & set "destnameechotext=Backup name  ú "
    call :lowercase archivetype & set "artype=-t!archivetype!"
    if /i "!preservealltimestamps:~0,1!"=="y" (
        set "exttimestamps=-mtc -mta"
        )
    if /i "!shortcutbrackets:~0,1!"=="y" (
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
    call :wscriptunesc destdir
    set "destdir=!destdir!\"
    set "ext=.!archivetype!"
    call :wscriptunesc args
    set "input=!args!" & set "args="
    exit /b

:inputparse
    set input="!%~1:"=!"
    call :escape input

    rem Split manual input paths by colon (from the drive letter). The drive letter will be on the end of the prior line, hence why the subsequent for loops extract that drive letter and append it to the start of the prior split line to concatenate the full path. Allows handling paths from different drive letters.
    set "inparsecount=0"
    for %%l in ("!lf!") do (
        for /f "delims=:" %%d in ("!input::=%%l!") do (
            set /a "inparsecount+=1"
            set "inparse[!inparsecount!]=%%d"
            )
        )
    set "input="

    rem Subtract for subsequent operations since last line lacks a drive letter at its end
    set /a "inparsecount-=1"

    for /l %%p in (1,1,!inparsecount!) do (
        set "inparse[%%p]=!inparse[%%p]:"=!"
        call :len inparse[%%p] inparse[%%p]len
        set /a "inparse[%%p]lentrim=!inparse[%%p]len! - 1"
        
        for %%t in (!inparse[%%p]lentrim!) do (
            for %%l in (!inparse[%%p]len!) do (
                rem Extract the drive letter from the end of the path line
                set "inparse[%%p]dr=!inparse[%%p]:~%%t,%%l!:"
                rem Trim the original path
                set "inparse[%%p]=!inparse[%%p]:~0,%%t!"
                )
            )
        )

    rem Loop again to pair the drive letters with the now trimmed paths
    for /l %%p in (1,1,!inparsecount!) do (
        rem Define subsequent line's count
        set /a "inparsecountnext=%%p + 1"
        for %%t in (!inparse[%%p]lentrim!) do (
            for %%l in (!inparse[%%p]len!) do (
            for %%n in (!inparsecountnext!) do (
                    set "inparse[%%n]=!inparse[%%n]:"=!"
                    call :trimspace "inparse[%%n]" "noesc"
                    call :unescape inparse[%%n]
                    rem Concatenate and wrap each line in double quotes
                    set "input=!input!!lf!"!inparse[%%p]dr!!inparse[%%n]:"=!""
                    set "inparse[%%p]=" & set "inparse[%%p]len=" & set "inparse[%%p]lentrim=" & set "inparse[%%p]dr="
                    )
                )
            )
        )

    rem Trim leading newline
    set "input=!input:~1!"

    rem Check if length after formatting approaches LNK field length limit
    for %%l in ("!lf!") do set "lnklimcheckin=!input:%%l=" "!"
    call :len lnklimcheckin lnklimcheckinlen
    rem Estimate final length based on additional arguments that will get added
    rem Script path + inputs + destination path (max for default Windows, and heading) + backup name (estimated, and heading)
    set /a "lnklimchecktotal=!scrlen! + !lnklimcheckinlen! + 250 + 10 + 50 + 10"
    if !lnklimchecktotal! gtr 1110 call :error "lnklimcheckreached"

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
                call :trimspace "%~1"
                set isfile=1&pushd "!%~1!" 2>nul&&(popd&set isfile=)||(if not exist "!%~1!" set isfile=)
                if defined isfile set "%~1=" & goto :inputprompts
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
                )
            if "%~1"=="destname" (
                call :triminvalid "destname"
                call :trimspace "destname"
                )
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

:lowercase
    for %%a in ("A=a" "B=b" "C=c" "D=d" "E=e" "F=f" "G=g" "H=h" "I=i"
                "J=j" "K=k" "L=l" "M=m" "N=n" "O=o" "P=p" "Q=q" "R=r"
                "S=s" "T=t" "U=u" "V=v" "W=w" "X=x" "Y=y" "Z=z") do (
        set "%~1=!%~1:%%~a!"
        )
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
    if defined destname set "linkprefix=!linkprefix! - "
    rem Save to %temp% directory initially to avoid WScript.Shell bug with certain Unicode characters
    set "destdir=!destdir:"=!"
    set "linkpathtemp=!temp!\!brkl!!linkprefix!!destname!!brkr!.lnk"
    set "linkpathdest=!destdir!\!brkl!!linkprefix!!destname!!brkr!.lnk"
    set "linkpathtempps=!linkpathtemp!" & rem Variable copied so File Explorer can later launch path, sans here-string formatting
    if not defined destname set "destname=[empty]"
    rem Escape certain characters for Windows WScript.Shell limitation workaround
    call :wscriptesc input
    call :wscriptesc destdir
    rem Remove double quotes from variables (will be added back by Powershell LNK creation command later)
    set "scr=!scr:"=!"
    set "input=!input:"=!"
    set "destname=!destname:"=!"
    rem Format as Powershell here-strings
    call :escps scr
    call :escps input
    call :escps destdir
    call :escps destname
    call :escps linkpathtempps
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
    set "rencount=1"
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

:triminvalid
    rem Remove invalid filename characters
    set "s=!%~1!"
    set "s=!s:"=!"
    set "s=!s:\=!"
    set "s=!s:/=!"
    set "s=!s::=!"
    set "s=!s:?=!"
    set "s=!s:<=!"
    set "s=!s:>=!"
    set "s=!s:|=!"
    rem If asterisk contained in string then caret/exclamation escaping will fail to be preserved but otherwise is necessary for caret/exclamation preservation with :escdelims call.
    call :escape s
    rem Use dedicated function for asterisk replacement
    call :escdelims s "*" "" s
    call :wscripttrim s
    call :unescape s
    set "%~1=!s:"=!"
    set "s="
    exit /b

:trimspace
    rem Remove leading/trailing whitespace
    set "s1=!%~1!" & set "s2=%~2"
    rem Wrap in double quotes to carry to escape call
    if not defined s2 call :escape "!s1!"
    set "x=!s1! "
    set "i=0"
    set "j="
    set "w=%x: =" & (if not defined w (if not defined j (set /a i+=1) else set /a j+=1) else set j=1) & set "w=%"
    set "x2=!x:~%i%,-%j%!"

    if not defined s2 call :unescape "!x2!"
    set "%~1=!x2!"
    set "s1=" & set "s2=" & set "x=" & set "x2="
    exit /b

:wscriptesc
    rem Escape specific Unicode characters as workaround for Windows WScript.Shell LNK limitation
    call :wscriptcp
    set "%~1=!%~1:%uni-FF1F%=###btz-esc-FF1F###!"
    set "%~1=!%~1:%uni-2215%=###btz-esc-2215###!"
    set "%~1=!%~1:%uni-003A%=###btz-esc-003A###!"
    exit /b

:wscriptunesc
    call :wscriptcp
    set "%~1=!%~1:###btz-esc-FF1F###=%uni-FF1F%!"
    set "%~1=!%~1:###btz-esc-2215###=%uni-2215%!"
    set "%~1=!%~1:###btz-esc-003A###=%uni-003A%!"
    exit /b

:wscripttrim
    call :wscriptesc "%~1"
    set "%~1=!%~1:###btz-esc-FF1F###=!"
    set "%~1=!%~1:###btz-esc-2215###=!"
    set "%~1=!%~1:###btz-esc-003A###=!"
    exit /b

:wscriptcp
    set "codepage="
    rem Detect existing codepage
    for /f "tokens=2 delims=:." %%a in ('chcp') do set "codepage=%%~a"
    if not defined codepage set "codepage=437"
    rem Change temporarily to UTF-8 codepage
    >nul chcp 65001
    rem Define characters' hex bytes based on 1252 codepage (what this script is encoded with)
    set "uni-FF1F=ï¼Ÿ" & rem Fullwidth Question Mark
    set "uni-2215=âˆ•" & rem Division Slash
    set "uni-003A=êž‰" & rem Modifier Letter Colon
    rem Switch back to original codepage
    >nul chcp !codepage!
    exit /b

endlocal