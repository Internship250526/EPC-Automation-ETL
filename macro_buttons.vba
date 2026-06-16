' =============================================================
'  EPC AUTOMATION — Excel Macro Buttons
'  Reliance Retail | EPC Department
' =============================================================
'
'  HOW TO ADD THIS TO YOUR EXCEL FILE:
'  1. Open your master .xlsx file
'  2. Save it as .xlsm  (File → Save As → Excel Macro-Enabled Workbook)
'  3. Go to Developer tab → Visual Basic  (or press Alt + F11)
'  4. In the editor: Insert → Module
'  5. Paste this entire file into the module
'  6. Close the editor
'  7. Add buttons to your sheet:
'       Developer tab → Insert → Button (Form Control)
'       Draw the button → Assign Macro → pick RunSplit or RunMerge
'       Right-click button → Edit Text → rename it
'
'  UPDATE THE PATH BELOW before running:
' =============================================================

' ← CHANGE THIS to the actual folder path on your PC / shared drive
Const EPC_FOLDER As String = "C:\EPC_Automation\"

' -------------------------------------------------------------
'  Button 1 — Generate PM Files (Split)
' -------------------------------------------------------------
Sub RunSplit()

    Dim answer As Integer
    answer = MsgBox( _
        "This will generate individual Excel files for all 6 PMs." & vbCrLf & vbCrLf & _
        "Each PM will get a file with only their sites." & vbCrLf & _
        "Files will be saved in the PM_Files folder." & vbCrLf & vbCrLf & _
        "Make sure the master file is saved before continuing." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion, "EPC Automation — Generate PM Files")

    If answer = vbNo Then Exit Sub

    ' Save this workbook first
    ThisWorkbook.Save

    ' Run split.py
    Dim cmd As String
    cmd = "cmd /c cd /d """ & EPC_FOLDER & """ && python split.py"
    Shell cmd, vbNormalFocus

    MsgBox _
        "Split started!" & vbCrLf & vbCrLf & _
        "A terminal window will show progress." & vbCrLf & _
        "PM files will appear in the PM_Files folder when done." & vbCrLf & vbCrLf & _
        "Share the PM_Files folder with the respective PMs.", _
        vbInformation, "EPC Automation — Split Running"

End Sub

' -------------------------------------------------------------
'  Button 2 — Merge PM Files back into Master
' -------------------------------------------------------------
Sub RunMerge()

    Dim answer As Integer
    answer = MsgBox( _
        "This will merge all PM files back into the Master file." & vbCrLf & vbCrLf & _
        "Before continuing, make sure:" & vbCrLf & _
        "  • All 6 PMs have saved their files in PM_Files folder" & vbCrLf & _
        "  • The master file is closed (this window will close it)" & vbCrLf & vbCrLf & _
        "A backup of the current master will be saved automatically." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion, "EPC Automation — Merge PM Files")

    If answer = vbNo Then Exit Sub

    ' Save and remember the master file path
    Dim masterPath As String
    masterPath = ThisWorkbook.FullName
    ThisWorkbook.Save

    ' Close master so Python can write to it
    ThisWorkbook.Close SaveChanges:=False

    ' Run merge.py
    Dim cmd As String
    cmd = "cmd /c cd /d """ & EPC_FOLDER & """ && python merge.py"
    Shell cmd, vbNormalFocus

    ' Wait for merge to finish then reopen master
    Application.Wait Now + TimeValue("00:00:08")
    Workbooks.Open masterPath

    MsgBox _
        "Merge complete!" & vbCrLf & vbCrLf & _
        "Changes in the master file are highlighted:" & vbCrLf & _
        "  YELLOW  =  cleanly updated cell" & vbCrLf & _
        "  RED     =  conflict (two PMs updated same cell — check Backups folder)" & vbCrLf & vbCrLf & _
        "Audit log saved in the Backups folder.", _
        vbInformation, "EPC Automation — Merge Complete"

End Sub

' =============================================================
'  OPTIONAL: Status check — lists which PM files are present
' =============================================================
Sub CheckPMFiles()

    Dim pmFolder As String
    pmFolder = EPC_FOLDER & "PM_Files\"

    Dim msg As String
    msg = "PM Files in folder:" & vbCrLf & vbCrLf

    Dim pms(5) As String
    pms(0) = "Dyaneshwar_Wagh"
    pms(1) = "Jitendra_Talreja"
    pms(2) = "Purushottam_M"
    pms(3) = "Santanu_Halder"
    pms(4) = "Saurav_Sagar"
    pms(5) = "Shanmuga_Rajan"

    Dim i As Integer
    For i = 0 To 5
        Dim found As Boolean
        found = False
        Dim fname As String
        fname = Dir(pmFolder & pms(i) & "*.xlsx")
        If fname <> "" Then
            msg = msg & "  ✓  " & pms(i) & vbCrLf
        Else
            msg = msg & "  ✗  " & pms(i) & "  (not yet submitted)" & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "EPC Automation — PM File Status"

End Sub
