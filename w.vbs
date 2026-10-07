Dim shotFld
shotFld = "D:\App\this-aint-rpa-main\screenshot"
' GENERATED ENGINE v1.5.0 Multi-Sequence & CHECK PIXEL Edition
Dim shl, ax, fso, currentLoop, maxLoop, runForever, ts, lastPhysX, lastPhysY, capExe, RunSeqShotIdx
RunSeqShotIdx = 0
Set shl = CreateObject("WScript.Shell")
Set ax = CreateObject("Excel.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
capExe = fso.GetParentFolderName(WScript.ScriptFullName) & "\ScreenCap.exe"
WScript.Sleep 1500
maxLoop = 999999: runForever = True
Sub UpdateLastCursorPos()
   On Error Resume Next
   Dim posHex, fullPos
   fullPos = ax.ExecuteExcel4Macro("CALL(""user32"",""GetMessagePos"",""J"")")
   posHex = Hex(fullPos)
   Do While Len(posHex) < 8 : posHex = "0" & posHex : Loop
   lastPhysX = CLng("&H" & Right(posHex, 4))
   lastPhysY = CLng("&H" & Left(posHex, 4))
   If lastPhysX > 32767 Then lastPhysX = lastPhysX - 65536
   If lastPhysY > 32767 Then lastPhysY = lastPhysY - 65536
End Sub
Function IsMouseMoved()
   On Error Resume Next
   IsMouseMoved = False
   Dim posHex, fullPos, curX, curY
   fullPos = ax.ExecuteExcel4Macro("CALL(""user32"",""GetMessagePos"",""J"")")
   posHex = Hex(fullPos)
   Do While Len(posHex) < 8 : posHex = "0" & posHex : Loop
   curX = CLng("&H" & Right(posHex, 4))
   curY = CLng("&H" & Left(posHex, 4))
   If curX > 32767 Then curX = curX - 65536
   If curY > 32767 Then curY = curY - 65536
   If Abs(curX - lastPhysX) > 8 Or Abs(curY - lastPhysY) > 8 Then
       IsMouseMoved = True
   End If
End Function
Sub LogHaltedStatus(remTimeSec)
   On Error Resume Next
   Set ts = fso.CreateTextFile("D:\\App\\this-aint-rpa-main\\wstatus.tmp", True)
   ts.WriteLine "PHYSICAL_HOLD;" & remTimeSec
   ts.Close
End Sub
Sub CheckPhysicalHold()
   On Error Resume Next
   If IsMouseMoved() Then
       Dim idleMs, remSec
       idleMs = 0
       Do
           WScript.Sleep 100
           Call UpdateLastCursorPos()
           If IsMouseMoved() Then
               idleMs = 0
           Else
               idleMs = idleMs + 100
           End If
           remSec = FormatNumber((3000 - idleMs) / 1000, 1)
           If remSec < 0 Then remSec = 0
           Call LogHaltedStatus(remSec)
       Loop While idleMs < 3000
   End If
End Sub
Sub LogStatus(stepIdx, totalSteps, stepName)
   On Error Resume Next
   Set ts = fso.CreateTextFile("D:\\App\\this-aint-rpa-main\\wstatus.tmp", True)
   If runForever Then
       ts.WriteLine currentLoop & "/INF;" & stepIdx & "/" & totalSteps & ";" & stepName
   Else
       ts.WriteLine currentLoop & "/" & maxLoop & ";" & stepIdx & "/" & totalSteps & ";" & stepName
   End If
   ts.Close
End Sub
Function WaitDynamicHoldColor(targetX, targetY, targetHexColor)
   On Error Resume Next
   If targetHexColor = "" Or UCase(targetHexColor) = "STATIC WAIT" Then Exit Function
   Dim hDC, rgbVal, curR, curG, curB, curHex
   Do
       Call CheckPhysicalHold()
       hDC = ax.ExecuteExcel4Macro("CALL(""user32"",""GetDC"",""JJ"",0)")
       rgbVal = ax.ExecuteExcel4Macro("CALL(""gdi32"",""GetPixel"",""JJJJ""," & hDC & "," & targetX & "," & targetY & ")")
       ax.ExecuteExcel4Macro "CALL(""user32"",""ReleaseDC"",""JJJ"",0," & hDC & ")"
       If rgbVal >= 0 Then
           curR = rgbVal Mod 256
           curG = (rgbVal \ 256) Mod 256
           curB = (rgbVal \ 65536) Mod 256
           curHex = "#" & Right("0" & Hex(curR), 2) & Right("0" & Hex(curG), 2) & Right("0" & Hex(curB), 2)
           If UCase(curHex) = UCase(targetHexColor) Then Exit Do
       End If
       WScript.Sleep 50
   Loop
End Function
Function ReadPixelHex(targetX, targetY)
   On Error Resume Next
   Dim hDC, rgbVal, curR, curG, curB
   ReadPixelHex = "#000000"
   hDC = ax.ExecuteExcel4Macro("CALL(""user32"",""GetDC"",""JJ"",0)")
   rgbVal = ax.ExecuteExcel4Macro("CALL(""gdi32"",""GetPixel"",""JJJJ""," & hDC & "," & targetX & "," & targetY & ")")
   ax.ExecuteExcel4Macro "CALL(""user32"",""ReleaseDC"",""JJJ"",0," & hDC & ")"
   If rgbVal >= 0 Then
       curR = rgbVal Mod 256
       curG = (rgbVal \ 256) Mod 256
       curB = (rgbVal \ 65536) Mod 256
       ReadPixelHex = "#" & Right("0" & Hex(curR), 2) & Right("0" & Hex(curG), 2) & Right("0" & Hex(curB), 2)
   End If
End Function
Sub AutoSwitchActiveWindow(targetTitle)
   Dim cleanTitle, wmi, processes, exeName, isRunning, i
   On Error Resume Next
   cleanTitle = targetTitle
   If InStr(cleanTitle, ".exe") > 0 Then
       exeName = cleanTitle: cleanTitle = Split(cleanTitle, ".exe")(0)
   Else
       If LCase(cleanTitle) = "notepad" Then exeName = "notepad.exe"
       If LCase(cleanTitle) = "word" Then exeName = "winword.exe"
       If LCase(cleanTitle) = "excel" Then exeName = "excel.exe"
       If exeName = "" Then exeName = cleanTitle & ".exe"
   End If
   Do
       isRunning = False
       Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
       Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & exeName & "'")
       If processes.Count > 0 Then isRunning = True
       If isRunning Then
           For i = 1 To 2: shl.AppActivate(targetTitle): shl.AppActivate(cleanTitle): WScript.Sleep 150: Next
           Exit Do
       Else
           If MsgBox("Target '" & exeName & "' belum aktif. Buka lalu klik OK.", 49, "Error Focus") = 2 Then WScript.Quit
       End If
   Loop
End Sub
Sub RunScreenshot(outFile, rx, ry, rw, rh)
   On Error Resume Next
   If Not fso.FileExists(capExe) Then Exit Sub
   If rx >= 0 And rw > 0 Then
    shl.Run Chr(34) & capExe & Chr(34) & " region " & Chr(34) & outFile & Chr(34) & " " & rx & " " & ry & " " & rw & " " & rh, 0, True
   Else
    shl.Run Chr(34) & capExe & Chr(34) & " full " & Chr(34) & outFile & Chr(34), 0, True
   End If
   WScript.Sleep 200
End Sub
Sub RunSequence_MAIN()
    Call CheckPhysicalHold()
    Call LogStatus(1, 3, "MAIN #1: BRANCH [CHECK PIXEL]")
    If UCase(ReadPixelHex(413, 521)) = UCase("#266183") Then Call RunSequence_SSequence
    WScript.Sleep 600
    Call CheckPhysicalHold()
    Call LogStatus(2, 3, "MAIN #2: BRANCH [CHECK PIXEL]")
    If UCase(ReadPixelHex(789, 501)) = UCase("#5EA61E") Then Call RunSequence_SSequence_2
    WScript.Sleep 600
    Call CheckPhysicalHold()
    Call LogStatus(3, 3, "MAIN #3: BRANCH [CHECK PIXEL]")
    If UCase(ReadPixelHex(638, 593)) = UCase("#F20000") Then Call RunSequence_SSequence_3
    WScript.Sleep 600
End Sub
Sub RunSequence_SSequence()
    Call CheckPhysicalHold()
    Call LogStatus(1, 2, "Sequence #1: Langkah 5 [CLICK]")
    Call WaitDynamicHoldColor(396, 597, "")
    ax.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",396,597)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",2,0,0,0,0)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",4,0,0,0,0)"
    Call UpdateLastCursorPos()
    WScript.Sleep 600
    Call CheckPhysicalHold()
    Call LogStatus(2, 2, "Sequence #2: Langkah 6 [WRITE]")
    shl.SendKeys "Sequence1"
    WScript.Sleep 600
End Sub
Sub RunSequence_SSequence_2()
    Call CheckPhysicalHold()
    Call LogStatus(1, 2, "Sequence 2 #1: Langkah 7 [CLICK]")
    Call WaitDynamicHoldColor(754, 547, "")
    ax.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",754,547)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",2,0,0,0,0)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",4,0,0,0,0)"
    Call UpdateLastCursorPos()
    WScript.Sleep 600
    Call CheckPhysicalHold()
    Call LogStatus(2, 2, "Sequence 2 #2: Langkah 8 [WRITE]")
    shl.SendKeys "Sequence2"
    WScript.Sleep 600
End Sub
Sub RunSequence_SSequence_3()
    Call CheckPhysicalHold()
    Call LogStatus(1, 2, "Sequence 3 #1: Langkah 9 [CLICK]")
    Call WaitDynamicHoldColor(592, 692, "")
    ax.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ"",592,692)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",2,0,0,0,0)"
    ax.ExecuteExcel4Macro "CALL(""user32"",""mouse_event"",""JJJJJJ"",4,0,0,0,0)"
    Call UpdateLastCursorPos()
    WScript.Sleep 600
    Call CheckPhysicalHold()
    Call LogStatus(2, 2, "Sequence 3 #2: Langkah 10 [WRITE]")
    shl.SendKeys "Sequence3"
    WScript.Sleep 600
End Sub
currentLoop = 0
Do While (currentLoop < maxLoop) Or runForever
    currentLoop = currentLoop + 1
    Call RunSequence_MAIN()
    If Not runForever And currentLoop >= maxLoop Then Exit Do
    WScript.Sleep 300
Loop
ax.Quit: Set ax = Nothing
On Error Resume Next
fso.DeleteFile("D:\\App\\this-aint-rpa-main\\wstatus.tmp")
MsgBox "Kelar Deh Kerjaan Lu.", 64, "Kelar Cuy"
