; <COMPILER: v1.1.33.10>
global hBitmap
TrayTip, CaseLauncher v1.7, Started running in the System Tray., 20, 17
Sleep 3000
#NoEnv
#SingleInstance force
OpenOutlookAppointment(caseid)
{
olApp := ComObjCreate("Outlook.Application")
olAppoint := olApp.CreateItem(1)
olAppoint.Subject := "Meeting TrackingID#"+caseid
olAppoint.Duration := 45
olAppoint.Display
}
OpenOutlookSearch(caseID)
{
static olSearchScopeAllFolders := 1
ClipWait, 1, 1
if ErrorLevel
return
olApp := ComObjActive("Outlook.Application")
olApp.ActiveExplorer.Search("""" caseID "*""", olSearchScopeAllFolders)
WinActivate, ahk_class rctrl_renwnd32
ControlSend, RICHEDIT60W1, {Tab}, ahk_class rctrl_renwnd32
}
#SingleInstance force
#Persistent
global caseID
OnClipboardChange("ClipChanged")
return
ClipChanged()
{
checkemail()
if (StrLen(clipboard) <= 10 )
{
nString := SubStr(Clipboard, 1, 9)
if (RegExMatch(nString, "/^$|^\d{9}$") )
{
IcmGUI("x1550 y750", "Case Launcher ", nString)
} else if(StrLen(clipboard) <= 9)
{
nString := SubStr(Clipboard, 1, 8)
if (RegExMatch(nString, "/^$|^\d{8}$") )
{
RaveGUI("x1550 y750", "Case Launcher ", nString)
}
}
}
else if (StrLen(clipboard) <= 19 )
{
nString := SubStr(Clipboard, 1, 16)
If (RegExMatch(nString, "/^$|^\d{16}$") )
{
LauncherGui("x1550 y750", "Case Launcher ", nString)
}
}
}
LauncherGui(position, title, nString)
{
if !A_IsCompiled
hBitmap := LoadPicture("C:\Users\dikadali\Documents\temp\Calender.bmp")
else {
hModule := DllCall("GetModuleHandle", "Str", A_ScriptFullPath, "Ptr")
hBitmap := DllCall("LoadImage", "Ptr", hModule, "Str", "CAL1", "UInt", 0, "Int", 0, "Int", 0, "UInt", 0, "Ptr")
}
caseID :=nString
static hexaColor, rgbColor
gui, LauncherGui:new
gui, Default
gui, +AlwaysOnTop -MaximizeBox -MinimizeBox ToolWindow -SysMenu
Gui, Font, s10 cBlue, Arial
Gui, Add, Text, cBlue,Click below to open SR# %nString%
Gui, Font, s12 cBlue, Arial
Gui, +AlwaysOnTop +Resize
Gui, Add, Picture, w30 h-1 section xs gLaunchDfmAscCb Icon1
, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Gui, Add, Button, ys gLaunchDfM, DfM
Gui, Add, Button, ys gLaunchCaseBuddy, CaseBuddy
Gui, Add, Button, ys gLaunchServiceDesk, SD
Gui, Add, Picture, w30 h-1 section xs gLaunchAscDtm Icon1
, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Gui, Add, Button, ys gLaunchASC, ASC
Gui, Add, Button, ys gLaunchDTM, DTM
Gui, Add, Picture, w30 h-1 section xs gLaunchEdge Icon1
, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Gui, Add, Button, ys gLaunchCaseParserStudio, Case Parser Studio
Gui, Add, Picture, w30 h-1 section xs gLaunchAppointment Icon1
, HBITMAP: %hBitmap%
Gui, Add, Picture, w30 h-1 section ys gLaunchOutlookSearch Icon1
, C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
Gui, Add, Picture, w30 h-1 section ys gLaunchTeamsContact Icon1
, C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\app.ico
Gui, Add, Picture, w30 h-1 section ys gLaunchOnenote Icon1
, C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE
Gui, Font, s8 cBlue, Arial
Gui, Add, Link,, Love this? <a href="mailto:CaseLauncherTalk@microsoft.com?subject=CaseLauncher Feedback">Send feedback</a>
gui, show, AutoSize NA %position%, %title%
SetTimer, CloseGuiTimer, -7000
return winexist()
LaunchEdge:
Run C:\Windows\explorer.exe shell:Appsfolder\Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge
WinClose
Return
LaunchDfM:
Run https://onesupport.crm.dynamics.com/main.aspx?appid=101acb62-8d00-eb11-a813-000d3a8b3117&pagetype=search&searchText=%Clipboard%
WinClose
Return
LaunchCaseBuddy:
Run mscb:case?%caseID%
WinClose
Return
LaunchServiceDesk:
Run https://servicedesk.microsoft.com/#/customer/cases?caseNumber=%caseID%
WinClose
Return
LaunchASC:
Run https://azuresupportcenter.msftcloudes.com/ticket?srId=%caseID%
WinClose
Return
LaunchDTM:
Run https://client.dtmnebula.microsoft.com/Home?srNumber=%caseID%
WinClose
Return
LaunchCaseParserStudio:
Run C:\Users\%A_UserName%\AppData\Local\Programs\case-parser-studio\Case-Parser-Studio.exe %Clipboard%
WinClose
Return
LaunchDfmAscCb:
Run https://onesupport.crm.dynamics.com/main.aspx?appid=101acb62-8d00-eb11-a813-000d3a8b3117&pagetype=search&searchText=%Clipboard%
Run https://azuresupportcenter.msftcloudes.com/ticket?srId=%caseID%
Run mscb:case?nString
WinClose
Return
LaunchAscDtm:
Run https://azuresupportcenter.msftcloudes.com/ticket?srId=%caseID%
Run https://client.dtmnebula.microsoft.com/Home?srNumber=%caseID%
WinClose
Return
LaunchAppointment:
Run mscb:meetnow?%caseID%
WinClose
Return
LaunchOutlookSearch:
OpenOutlookSearch(caseID)
WinClose
Return
LaunchTeamsContact:
Run mscb:chat?%caseID%
WinClose
Return
LaunchOneNote:
Run mscb:onenote?%caseID%
WinClose
Return
}
LauncherGuiGuiEscape:
{
Gui, LauncherGui:Destroy
return
}
LauncherGuiGuiCancel:
{
Gui, LauncherGui:Destroy
return
}
OnOk:
{
Gui, LauncherGui:Destroy
return
}
CloseGuiTimer:
{
Gui, LauncherGui:Destroy
return
}
IcmGUI(position, title, nstring)
{
static hexaColor, rgbColor
gui, IcmGUI:new
gui, Default
gui, +AlwaysOnTop -MaximizeBox -MinimizeBox ToolWindow -SysMenu
Gui, Add, Text, cBlue, Open ICM# %clipboard%
Gui, Font, s13 cBlue, Arial
Gui, +AlwaysOnTop +Resize
Gui, Add, Picture, w28 h-1 section xs gLaunchICM Icon1
, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Gui, Add, Button, ys gLaunchICM, ICM Portal
Gui, Add, Picture, w28 h-1 section xs gLaunchICM Icon1
, C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Gui, Add, Button, ys gLaunchIRIDIAS, Iridias Portal
Gui, Font, s8 cBlue, Arial
Gui, Add, Link,, Love this? <a href="mailto:dikadali@microsoft.com?subject=CaseLauncher Feedback">Send feedback</a>
gui, show, AutoSize NA %position%, %title%
SetTimer, CloseGuiTimer1, -7000
return winexist()
LaunchICM:
Run https://portal.microsofticm.com/imp/v3/incidents/details/%clipboard%/home
WinClose
Return
LaunchIRIDIAS:
Run https://iridias.microsoft.com/incidentcentral?id=%clipboard%
WinClose
Return
}
IcmGUIGuiEscape:
{
Gui, IcmGUI:Destroy
return
}
IcmGUIGuiCancel:
{
Gui, IcmGUI:Destroy
return
}
OnOk1:
{
Gui, IcmGUI:Destroy
return
}
CloseGuiTimer1:
{
Gui, IcmGUI:Destroy
return
}
RaveGUI(position, title, nstring)
{
static hexaColor, rgbColor
gui, RaveGUI:new
gui, Default
gui, +AlwaysOnTop -MaximizeBox -MinimizeBox ToolWindow -SysMenu
Gui, Add, Text, cBlue, Open Rave# %clipboard%
Gui, Font, s15 cBlue, Arial
Gui, +AlwaysOnTop +Resize
Gui, Add, Picture, w30 h-1 section gLaunchRave Icon1
, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
Gui, Add, Button, ys gLaunchRave, Rave#%clipboard%
Gui, Add, Picture, w30 h-1 section xs gLaunchASC Icon1
, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
Gui, Add, Button, ys gLaunchDTMRave, DTM
Gui, Font, s8 cBlue, Arial
Gui, Add, Link,, Love this? <a href="mailto:dikadali@microsoft.com?subject=CaseLauncher Feedback">Send feedback</a>
gui, show, AutoSize NA %position%, %title%
SetTimer, CloseGuiTimer2, -7000
return winexist()
LaunchRave:
Run https://rave.microsoft.com/search?query=%Clipboard%
WinClose
Return
LaunchDTMRave:
Run https://client.dtmnebula.microsoft.com/Home?srNumber=%clipboard%
WinClose
Return
}
RaveGUIGuiEscape:
{
Gui, RaveGUI:Destroy
return
}
RaveGUIGuiCancel:
{
Gui, RaveGUI:Destroy
return
}
OnOk2:
{
Gui, RaveGUI:Destroy
return
}
CloseGuiTimer2:
{
Gui, RaveGUI:Destroy
return
}
#s::
::.sd::
{
Run https://servicedesk.microsoft.com/#/customer/cases?caseNumber=%clipboard%
Return
}
::.asc::
{
Run https://azuresupportcenter.msftcloudes.com/ticket?srId=%clipboard%
Return
}
::.dtm::
{
Run https://client.dtmnebula.microsoft.com/Home?srNumber=%clipboard%
Return
}
::.icm::
{
Run https://portal.microsofticm.com/imp/v3/incidents/details/%clipboard%/home
Return
}
::.utc::
::!utc::
{
run tzutil /s "UTC"
Return
}
::.ist::
::!ist::
{
run tzutil /s "India Standard Time"
Return
}
::.ist::
::!ist::
{
run tzutil /s "Pacific Standard Time"
Return
}
::.meet::
{
OpenOutlookAppointment(%clipboard%)
Return
}
Checkemail()
{
nString := Clipboard
if (StrLen(clipboard) >= 5 )
{
if (RegExMatch(nString,"^[a-zA-Z0-9+_.-]+@[a-zA-Z0-9.-]+$") )
{
CommsGUI("x1550 y750", "Case Launcher ", nString)
}
}
}
CommsGUI(position, title, nstring)
{
static hexaColor, rgbColor
gui, CommsGUI:new
gui, Default
gui, +AlwaysOnTop -MaximizeBox -MinimizeBox ToolWindow -SysMenu
Gui, Add, Text, cBlue, Chat With: %clipboard%
Gui, Font, s12 cBlue, Arial
Gui, +AlwaysOnTop +Resize
Gui, Add, Picture, w25 h-1 section xs gLaunchTeamsChat Icon1
, C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\app.ico
Gui, Add, Button, ys gLaunchTeamsChat, Teams Chat
Gui, Add, Picture, w25 h-1 section xs gLaunchOutlookEmail Icon1
, C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
Gui, Add, Button, ys gLaunchOutlookEmail, Outlook Email
hrefstring := "mailto:" . nString . "?subject=TrackingID#"
Gui, Font, s8 cBlue, Arial
Gui, Add, Link,, Love this? <a href="mailto:dikadali@microsoft.com?subject=CaseLauncher Feedback">Send feedback</a>
gui, show, AutoSize NA %position%, %title%
SetTimer, CloseGuiTimer3, -7000
return winexist()
LaunchTeamsChat:
Run sip:%clipboard%
WinClose
Return
LaunchOutlookEmail:
Run mailto:%clipboard%?subject=TrackingID#
WinClose
Return
}
CommsGUIGuiEscape:
{
Gui, CommsGUI:Destroy
return
}
CommsGUIGuiCancel:
{
Gui, CommsGUI:Destroy
return
}
OnOk3:
{
Gui, CommsGUI:Destroy
return
}
CloseGuiTimer3:
{
Gui, CommsGUI:Destroy
return
}