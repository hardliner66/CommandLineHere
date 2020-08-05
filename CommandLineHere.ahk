SetTitleMatchMode RegEx
return

; Stuff to do when Windows Explorer is open
;


; '!'
; This tells AutoIt to send an ALT keystroke, therefore Send("This is text!a") would send the keys "This is text" and then press "ALT+a".
;
; N.B. Some programs are very choosy about capital letters and ALT keys, i.e., "!A" is different from "!a". The first says ALT+SHIFT+A, the second is ALT+a. If in doubt, use lowercase!
;
; '+'
; This tells AutoIt to send a SHIFT keystroke; therefore, Send("Hell+o") would send the text "HellO". Send("!+a") would send "ALT+SHIFT+a".
;
; '^'
; This tells AutoIt to send a CONTROL keystroke; therefore, Send("^!a") would send "CTRL+ALT+a".
;
; N.B. Some programs are very choosy about capital letters and CTRL keys, i.e., "^A" is different from "^a". The first says CTRL+SHIFT+A, the second is CTRL+a. If in doubt, use lowercase!
;
; '#'
; The hash now sends a Windows keystroke; therefore, Send("#r") would send Win+r which launches the Run() dialog box.

#IfWinActive ahk_class ExploreWClass|CabinetWClass

    #c::
        RefreshEnvironment()
        OpenCmdInCurrent()
    return
    #+c::
        RefreshEnvironment()
        OpenCodeInCurrent()
    return
    #+r::
        RefreshEnvironment()
        OpenToolManInCurrent()
    return
    ^+m::
        RefreshEnvironment()
        OpenMergeInCurrent()
    return
#IfWinActive

GetCurrentDirectory()
{
    SetTitleMatchMode, 2
    WinGetText, out,A
    StringSplit, StringArray, out, ["`r","`n","`r`n"]
    Loop, %StringArray0%
    {
        this_string := StringArray%a_index%
        IfInString, this_string, :
        {
            this_string := Trim(StrReplace(this_string, "Adresse:", ""))
            this_string := Trim(StrReplace(this_string, "Address:", ""))
            return %this_string%
        }
    }
}

; Opens the command shell 'cmd' in the directory browsed in Explorer.
; Note: expecting to be run when the active window is Explorer.
;
OpenCmdInCurrent()
{
    ; strip to bare address
    full_path := GetCurrentDirectory()

    IfExist, %LOCALAPPDATA%\Microsoft\WindowsApps\wt.exe
      Run,  "%LOCALAPPDATA%\Microsoft\WindowsApps\wt.exe" -d "%full_path%"
   else
      Run,  cmd /K cd /D "%full_path%"
}

; Opens sublime merge in the directory browsed in Explorer.
; Note: expecting to be run when the active window is Explorer.
;
OpenMergeInCurrent()
{
    ; strip to bare address
    full_path := GetCurrentDirectory()

    IfExist, C:\Program Files\Sublime Merge\smerge.exe
        Run, "C:\Program Files\Sublime Merge\smerge.exe" "%full_path%"
}

; Opens the command shell 'code' in the directory browsed in Explorer.
; Note: expecting to be run when the active window is Explorer.
;
OpenCodeInCurrent()
{
    ; strip to bare address
    full_path := GetCurrentDirectory()

    IfExist, C:\Program Files\Microsoft VS Code\Code.exe
        Run, "C:\Program Files\Microsoft VS Code\Code.exe" "%full_path%"
    else
    	IfExist, C:\Program Files (x86)\Microsoft VS Code\Code.exe
        	Run, "C:\Program Files (x86)\Microsoft VS Code\Code.exe" "%full_path%"
        else
        	Run, "%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe" "%full_path%"
}

; Opens the command shell 'code' in the directory browsed in Explorer.
; Note: expecting to be run when the active window is Explorer.
;
OpenToolManInCurrent()
{
    ; strip to bare address
    full_path := GetCurrentDirectory()

    IfExist, c:\Tools\Toolman.exe
        Run, "c:\Tools\Toolman.exe" "%full_path%"
}

RefreshEnvironment()
{
   ; load the system-wide environment variables first in case there are user-level
   ; variables with the same name (since they override the system definitions).
   ; treat PATH and PATHEXT special - we must contactenate the user and system values.
   sysPATH := ""
   sysPATHEXT := ""
   Loop, HKLM, SYSTEM\CurrentControlSet\Control\Session Manager\Environment, 0, 0
   {
      RegRead, vEnvValue
      If (A_LoopRegType == "REG_EXPAND_SZ") {
         If (!ExpandEnvironmentStrings(vEnvValue)) {
            Return False
         }
      }
      EnvSet, %A_LoopRegName%, %vEnvValue%
      if (A_LoopRegName = "PATH") {
         sysPATH := vEnvValue
         }
      else if (A_LoopRegName = "PATHEXT") {
         sysPATHEXT := vEnvValue
         }
   }

   ; now load the user level environment variables
   Loop, HKCU, Environment, 0, 0
   {
      RegRead, vEnvValue
      If (A_LoopRegType == "REG_EXPAND_SZ") {
         If (!ExpandEnvironmentStrings(vEnvValue)) {
            Return False
         }
      }
      envVal := vEnvValue
      if (A_LoopRegName = "PATH") {
         envVal := envVal . ";" . sysPATH
        }
      else if (A_LoopRegName = "PATHEXT") {
         envVal := envVal . ";" . sysPATHEXT
        }
      EnvSet, %A_LoopRegName%, %envVal%
   }

   ; return success.
   Return True
}

ExpandEnvironmentStrings(ByRef vInputString)
{
   ; get the required size for the expanded string
   vSizeNeeded := DllCall("ExpandEnvironmentStrings", "Str", vInputString, "Int", 0, "Int", 0)
   If (vSizeNeeded == "" || vSizeNeeded <= 0)
      return False ; unable to get the size for the expanded string for some reason

   vByteSize := vSizeNeeded + 1
   If (A_PtrSize == 8) { ; Only 64-Bit builds of AHK_L will return 8, all others will be 4 or blank
      vByteSize *= 2 ; need to expand to wide character sizes
   }
   VarSetCapacity(vTempValue, vByteSize, 0)

   ; attempt to expand the environment string
   If (!DllCall("ExpandEnvironmentStrings", "Str", vInputString, "Str", vTempValue, "Int", vSizeNeeded))
      return False ; unable to expand the environment string
   vInputString := vTempValue

   ; return success
   Return True
}
