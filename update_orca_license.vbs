Option Explicit

Dim licenseKeyPath, fso, licenseKey, file, WshShell, started, t0, windowFound, clipboard, i, j, kundenReferenz

Set fso = CreateObject("Scripting.FileSystemObject")

licenseKeyPath = "\\141.30.148.247\adminsh\ORCA.txt"

kundenReferenz = "17297"

' WScript.Echo licenseKeyPath
' WScript.Echo "Pruefe Pfad: " & licenseKeyPath
' WScript.Echo "Existiert Datei? " & fso.FileExists(licenseKeyPath)

' If Not fso.FileExists(licenseKeyPath) Then
' 	WScript.Echo "FEHLER: Datei nicht gefunden: " & licenseKeyPath
' 	WScript.Quit 1
' End If

Set file = fso.OpenTextFile(licenseKeyPath, 1, False)
licenseKey = Trim(file.ReadAll)
file.close
' WScript.Echo "Lizenzschluessel: " & licenseKey

Set clipboard = CreateObject("htmlfile")
clipboard.ParentWindow.ClipboardData.SetData "text", licenseKey

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run """C:\ProgramData\Microsoft\Windows\StartM~1\Programs\ORCASo~1\ORCASo~1.lnk"""
' WScript.Echo "Starte ORCA..."
WScript.Sleep 2000

started = False
t0 = Timer

Do
    windowFound = False

    If WshShell.AppActivate("Orca Manager (Preview)") Then
        windowFound = True
        Exit Do
    End If

    If WshShell.AppActivate("ORCA") Then
        ' Nur schließen, wenn es NICHT das Preview-Fenster ist
        If Not WshShell.AppActivate("Orca Manager (Preview)") Then
            WScript.Sleep 300
            WshShell.SendKeys "%{F4}"  ' Alt+F4
            WScript.Sleep 800
        End If
    Else
        ' Kein Fenster gefunden – kurz warten
        WScript.Sleep 500
    End If

Loop While (Not windowFound) And (Timer - t0 < 60)

WshShell.SendKeys "{ENTER}"
WScript.Sleep 1000

If WshShell.AppActivate("Orca Infocenter") Then
    WScript.Sleep 300
    WshShell.SendKeys "%{F4}"  ' Alt + F4 = Fenster schließen
    WScript.Sleep 1000
End If

WshShell.AppActivate "ORCA Manager (Preview)"

For i = 1 To 16
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 100
Next

WshShell.SendKeys " "   ' Leertaste = Klick

WScript.Sleep 1000

For j = 1 to 4
	WshShell.SendKeys "{TAB}"
    WScript.Sleep 100
Next

WshShell.SendKeys kundenReferenz
WScript.Sleep 250
WshShell.SendKeys "{TAB}"
WScript.Sleep 200
WshShell.SendKeys licenseKey
WScript.Sleep 300
WshShell.SendKeys "{ENTER}"