Option Compare Database
Option Explicit
'Public Function InjectSQL(ParamArray sql_string() As Variant) As String
'Public Function IsAdjoiningRoom(ByVal room_guid As String) As Boolean
'Public Function RoomIDToName(ByVal room_id As String) As String
'Public Function RoomNameToID(ByVal room_name As String) As String
'Public Function GetAdjoinRoomID(ByVal room_id As String) As String
'Public Function GetAdjoinRoomName(ByVal room_name As String) As String
'Public Function TurnOverRoom(inv_rec_id As String, room_name As String) As String
'Public Function ChangeRoom(transaction_id As String, room_name As String, Optional remarks As Variant = "") As String
'Private Function ManualIngest(fqfn As String, LoggingFunct As String) As String
'Public Function StandardCheckout(guest_id As String, Optional remarks As String = "Check out") As String
'Public Function StandardCheckin(guest_id As String, room_guid As String, Optional remarks As String = "Check in") As String
'Public Function ReadyToCheckin(ByVal guest_id As String) As String
'Public Function ReadyToCheckout(ByVal guest_id As String) As String
'Public Function FindReservation(ByVal guest_id As String) As String
'Public Function FindAnyone(ByVal guest_id As String) As Boolean
'Public Function FindGuest(ByVal guest_id As String) As Boolean
'Public Function FindRelatedPerson(ByVal guest_id As String) As Boolean
'Public Function IsAvailableRoomID(ByVal room_id As String) As Boolean
'Public Function IsAvailableRoomName(ByVal room_name As String) As Boolean
'Public Function IsReservedRoomID(ByVal room_id As String) As Boolean
'Public Function IsReservedRoomName(ByVal room_name As String) As Boolean
'Public Function CountCheckout() As Integer
'Public Function RoomsUnavailable() As Integer
'Public Function CountCheckin() As Integer
'Public Function CountReservationNoRoom() As Integer
'Public Function RoomsAvailable() As Integer
'Public Sub HideNavPane()
'Public Sub ShowNavPane()


Private dtmNext As Date
Private Type POINTAPI
    X As Long
    Y As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function SetThreadExecutionState Lib "Kernel32.dll" (ByVal esFlags As Long) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (Point As POINTAPI) As Long
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Function SetThreadExecutionState Lib "Kernel32.dll" (ByVal esFlags As Long) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (Point As POINTAPI) As Long
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long)
#End If

Private Const micVerboseSummary = 1
Private Const micVerboseListAll = 2

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

Public Property Let PreventSleepMode(ByVal bPrevent As Boolean)
    'https://www.mrexcel.com/board/threads/how-to-prevent-my-pc-from-going-to-stand-by-sleep-mode.1192484/
    Const ES_SYSTEM_REQUIRED As Long = &H1
    Const ES_DISPLAY_REQUIRED As Long = &H2
    Const ES_AWAYMODE_REQUIRED = &H40
    Const ES_CONTINUOUS As Long = &H80000000
   
    If bPrevent Then
        Call SetThreadExecutionState(ES_CONTINUOUS Or ES_DISPLAY_REQUIRED Or ES_SYSTEM_REQUIRED Or ES_AWAYMODE_REQUIRED)
    Else
        Call SetThreadExecutionState(ES_CONTINUOUS)
    End If

End Property

Sub dontsleep()
    PreventSleepMode = True
End Sub

Sub gosleep()
    PreventSleepMode = False
End Sub

'Purpose:   Count the number of lines of code in your database.
'Author:    Allen Browne (allen@allenbrowne.com)
'Release:   26 November 2007
'Copyright: None. You may use this and modify it for any database you write.
'           All we ask is that you acknowledge the source (leave these comments in your code.)
'Documentation: http://allenbrowne.com/vba-CountLines.html

Public Function CountLines(Optional iVerboseLevel As Integer = 3) As Long
On Error GoTo Err_Handler
    'Purpose:   Count the number of lines of code in modules of current database.
    'Requires:  Access 2000 or later.
    'Argument:  This number is a bit field, indicating what should print to the Immediate Window:
    '               0 displays nothing
    '               1 displays a summary for the module type (form, report, stand-alone.)
    '               2 list the lines in each module
    '               3 displays the summary and the list of modules.
    'Notes:     Code will error if dirty (i.e. the project is not compiled and saved.)
    '           Just click Ok if a form/report is assigned to a non-existent printer.
    '           Side effect: all modules behind forms and reports will be closed.
    '           Code window will flash, since modules cannot be opened hidden.
    Dim accobj As AccessObject  'Each module/form/report.
    Dim strDoc As String        'Name of each form/report
    Dim lngObjectCount As Long  'Number of modules/forms/reports
    Dim lngObjectTotal As Long  'Total number of objects.
    Dim lngLineCount As Long    'Number of lines for this object type.
    Dim lngLineTotal As Long    'Total number of lines for all object types.
    Dim bWasOpen As Boolean     'Flag to leave form/report open if it was open.
    Dim counter As Integer
    
    counter = 0
    
    DoCmd.Hourglass True
    Application.Echo False
    'Stand-alone modules.
    lngObjectCount = 0&
    lngLineCount = 0&
    
    For Each accobj In CurrentProject.AllModules
        'OPTIONAL: TO EXCLUDE THE CODE IN THIS MODULE FROM THE COUNT:
        '  a) Uncomment the If ... and End If lines (3 lines later), by removing the single-quote.
        '  b) Replace MODULE_NAME with the name of the module you saved this in (e.g. "Module1")
        '  c) Check that the code compiles after your changes (Compile on Debug menu.)
        DoEvents
        counter = counter + 1
        If counter Mod 10 = 0 Then
            DoEvents
        End If
        If accobj.Name <> "mod_cbk_dntsleep" Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines(accobj.Name, True, iVerboseLevel)
        End If

    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print Format$(lngLineCount, "@@@@@@") & " line(s) in " & lngObjectCount & " stand-alone module(s)"
        'Debug.Print lngLineCount & " line(s) in " & lngObjectCount & " stand-alone module(s)"
        Debug.Print
    End If
    
    'Modules behind forms.
    lngObjectCount = 0&
    lngLineCount = 0&
    For Each accobj In CurrentProject.AllForms
        DoEvents
        counter = counter + 1
        If counter Mod 10 = 0 Then
            DoEvents
        End If
        
        strDoc = accobj.Name
        bWasOpen = accobj.IsLoaded
        If Not bWasOpen Then
            DoCmd.OpenForm strDoc, acDesign, WindowMode:=acHidden
        End If
        If Forms(strDoc).HasModule Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines("Form_" & strDoc, False, iVerboseLevel)
        End If
        If Not bWasOpen Then
            DoCmd.Close acForm, strDoc, acSaveNo
        End If
    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print Format$(lngLineCount, "@@@@@@") & " line(s) in " & lngObjectCount & " module(s) behind forms"
        'Debug.Print lngLineCount & " line(s) in " & lngObjectCount & " module(s) behind forms"
        Debug.Print
    End If
    
    'Modules behind reports.
    lngObjectCount = 0&
    lngLineCount = 0&
    For Each accobj In CurrentProject.AllReports
        DoEvents
        counter = counter + 1
        If counter Mod 10 = 0 Then
            DoEvents
        End If
        
        strDoc = accobj.Name
        bWasOpen = accobj.IsLoaded
        If Not bWasOpen Then
            'In Access 2000, remove the ", WindowMode:=acHidden" from the next line.
            DoCmd.OpenReport strDoc, acDesign, WindowMode:=acHidden
        End If
        If Reports(strDoc).HasModule Then
            lngObjectCount = lngObjectCount + 1&
            lngLineCount = lngLineCount + GetModuleLines("Report_" & strDoc, False, iVerboseLevel)
        End If
        If Not bWasOpen Then
            DoCmd.Close acReport, strDoc, acSaveNo
        End If
    Next
    lngLineTotal = lngLineTotal + lngLineCount
    lngObjectTotal = lngObjectTotal + lngObjectCount
    If (iVerboseLevel And micVerboseSummary) <> 0 Then
        Debug.Print Format$(lngLineCount, "@@@@@@") & " line(s) in " & lngObjectCount & " module(s) behind reports"
        Debug.Print Format$(lngLineTotal, "@@@@@@") & " line(s) in " & lngObjectTotal & " module(s)"
    End If

    CountLines = lngLineTotal
    
Exit_Handler:
    DoCmd.Hourglass False
    Application.Echo True
    Exit Function
    
Err_Handler:
    Select Case Err.Number
    Case 29068&     'This error actually occurs in GetModuleLines()
        MsgBox "Cannot complete operation." & vbCrLf & "Make sure code is compiled and saved."
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Handler

End Function

Private Function GetModuleLines(strModule As String, bIsStandAlone As Boolean, iVerboseLevel As Integer) As Long
    'Usage:     Called by CountLines().
    'Note:      Do not use error handling: must pass error back to parent routine.
    Dim bWasOpen As Boolean     'Flag applies to standalone modules only.
    
    If bIsStandAlone Then
        bWasOpen = CurrentProject.AllModules(strModule).IsLoaded
    End If
    If Not bWasOpen Then
        DoCmd.OpenModule strModule
    End If
    If (iVerboseLevel And micVerboseListAll) <> 0 Then
        Debug.Print Format$(Modules(strModule).CountOfLines, "@@@@@@"), strModule
    End If
    GetModuleLines = Modules(strModule).CountOfLines
    If Not bWasOpen Then
        DoCmd.Close acModule, strModule, acSaveYes
    End If
End Function

Private Sub SingleClick()
    SetCursorPos 100, 100 'x and y position
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub DoubleClick()
    'Double click as a quick series of two clicks
    SetCursorPos 100, 100 'x and y position
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub RightClick()
    'Right click
    SetCursorPos 200, 200 'x and y position
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Sub Move_Cursor()
    Dim Hold As POINTAPI
    GetCursorPos Hold
    'SetCursorPos Hold.x + 30, Hold.y
    SetCursorPos Hold.X + 300, Hold.Y + 300
    dtmNext = DateAdd("n", 30, Now)
    'Application.OnTime dtmNext, "Move_Cursor"
End Sub

Private Sub ProcessFile()
'https://stackoverflow.com/questions/54699532/read-text-file-in-access-vba
    Dim fd As Office.FileDialog, strfilepath As String, strSql As String
    Dim fname As String, foldername As String, bname As String, ext As String
    Dim fieldname As String
    Dim i As Integer, max As Integer
    Dim db As Database, rst As Recordset, fld As Field, fsoObj As Object
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt", 1
        .Title = "Select the file list"
        .AllowMultiSelect = False
        If .Show = True Then
            strfilepath = .SelectedItems(1)
        End If
    End With
       
    Set fsoObj = CreateObject("Scripting.FileSystemObject")
    fname = fsoObj.GetFileName(strfilepath)
    foldername = fsoObj.GetParentFolderName(strfilepath)
    bname = fsoObj.GetBaseName(strfilepath)
    ext = fsoObj.GetExtensionName(strfilepath)
    strSql = "Select * FROM " & strfilepath
    
    Set db = OpenDatabase(foldername, False, False, "Text; Format=Delimited(|);HDR=Yes;CharacterSet=437")
    Set rst = db.OpenRecordset(strSql)
    Debug.Print rst.Fields(0).Name, rst.Fields(0)

    With rst
        .MoveLast
        max = .RecordCount
        .MoveFirst
        For i = 1 To max
            If Not IsNull(rst.Fields(0).Value) Then
                Debug.Print i, FileExists(rst.Fields(0).Value), rst.Fields(0).Value
            End If
            .MoveNext
        Next
'        For Each fld In rst.Fields
'            Debug.Print fld.Name, fld.Value
'        Next
    End With
    
    rst.Close
    db.Close
    Set rst = Nothing
    Set db = Nothing

End Sub

Public Sub test3()
    Dim s As String
    Dim tdf As TableDef, pty As Property
    
    For Each tdf In CurrentDb.TableDefs
        If Not StringLimitedCompare(tdf.Name, "tdf") Then
            Debug.Print tdf.Name
            For Each pty In tdf.Properties
                If s = "" Then
                    s = pty.Name
                Else
                    s = s & ", " & pty.Name
                End If
            Next pty
            Debug.Print s
        End If
    Next tdf

End Sub

Public Sub ID_Generate_string(ByVal str As String)
    Dim i As Integer
    Dim vstr() As String, tmp As String, strFile As String
    Dim fd As Office.FileDialog
    Dim fso As Object, oFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Show
    End With
    strFile = fd.SelectedItems(1)
    
    If Len(strFile) = 0 Then
        MsgBox "No filename given", vbInformation, "Data Import"
        Exit Sub
    End If
    
    Set oFile = fso.CreateTextFile(strFile)
    
    ReDim vstr(Len(str))
    For i = 1 To Len(str)
        tmp = UAT_ID_Generator(Mid(str, i, 1))
        Do While WhereInArray(vstr, tmp) <> -1
            tmp = UAT_ID_Generator(Mid(str, i, 1))
        Loop
        oFile.WriteLine tmp
        vstr(i) = tmp
    Next i
    Erase vstr
    'oFile.WriteLine "test"
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
    MsgBox strFile & " created"

End Sub

Public Function UAT_ID_Generator(Optional ch As String) As String
    Dim num1 As Integer, total As Integer, upp_range As Long, low_range As Long, xtra As Integer, rem_ch_len As Integer
    Dim prefix As String, body As String, suffix As String, rem_ch As String
    
    Randomize
    If ch = "" Or InStr(1, "mfgst", left(LCase(ch), 1)) = 0 Then
        prefix = Mid("mfgst", RandBetween(1, 5), 1)
    Else
        prefix = left(LCase(ch), 1)
    End If
    'Debug.Print "Prefix: [" & prefix & "]"
    If Len(ch) > 1 Then
        rem_ch_len = Len(ch) - 1
        rem_ch = Right(ch, rem_ch_len)
        If rem_ch_len > 7 Then
            rem_ch = left(rem_ch, 7)
        End If
        If IsNumeric(rem_ch) Then
            If Len(rem_ch) < 7 Then
                upp_range = CLng(String(7 - rem_ch_len, "9"))
                low_range = 1
                body = rem_ch & Format(RandBetween(low_range, upp_range), String(7 - rem_ch_len, "0"))
            Else
                body = rem_ch
            End If
        End If
    End If
    If body = "" Then
        If InStr(1, "s", prefix, vbTextCompare) > 0 Then
            upp_range = 5000000
            low_range = 3000000
        ElseIf InStr(1, "tg", prefix, vbTextCompare) > 0 Then
            upp_range = 9999999
            low_range = 3000000
        ElseIf InStr(1, "fm", prefix, vbTextCompare) > 0 Then
            upp_range = 2000000
            low_range = 1
        End If
        body = Format(RandBetween(low_range, upp_range), "0000000")
    End If
    If InStr(1, "tg", prefix) > 0 Then
        xtra = 4
    ElseIf InStr(1, "m", prefix) > 0 Then
        xtra = 3
    Else
        xtra = 0
    End If
    total = (CInt(Mid(body, 1, 1)) * 2) + _
            (CInt(Mid(body, 2, 1)) * 7) + _
            (CInt(Mid(body, 3, 1)) * 6) + _
            (CInt(Mid(body, 4, 1)) * 5) + _
            (CInt(Mid(body, 5, 1)) * 4) + _
            (CInt(Mid(body, 6, 1)) * 3) + _
            (CInt(Mid(body, 7, 1)) * 2) + xtra
    
    Select Case total Mod 11
        Case 0
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "x"
            Else
                suffix = "j"
            End If
        Case 1
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "w"
            Else
                suffix = "z"
            End If
        Case 2
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "u"
            Else
                suffix = "i"
            End If
        Case 3
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "t"
            Else
                suffix = "h"
            End If
        Case 4
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "r"
            Else
                suffix = "g"
            End If
        Case 5
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "q"
            Else
                suffix = "f"
            End If
        Case 6
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "p"
            Else
                suffix = "e"
            End If
        Case 7
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "n"
            Else
                suffix = "d"
            End If
        Case 8
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "m"
            Else
                suffix = "c"
            End If
        Case 9
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "l"
            Else
                suffix = "b"
            End If
        Case 10
            If InStr(1, "gfm", prefix) > 0 Then
                suffix = "k"
            Else
                suffix = "a"
            End If
    End Select
    UAT_ID_Generator = UCase(prefix & body & suffix)

End Function

Public Function IsInArray(arr As Variant, stringToBeFound As String) As Integer
    IsInArray = UBound(Filter(arr, stringToBeFound))
End Function

Private Sub TestProc()
    Dim dummy_i As Long, dummy_j As Long, dummy_s As String
    
    dummy_i = 1000000
    dummy_j = 777
    dummy_s = CStr(dummy_i / dummy_j)
End Sub

Private Sub SpeedTest()
    Dim cntr_i As Long, i As Long, j As Long, dummy_i As Long, dummy_j As Long, dummy_s As String
    Dim starttime As Double, duration As Double, cycle_limit As Double
    Dim rs As Recordset
    
    cycle_limit = 100000000
''''''''''''''''''''''''
    
    starttime = Timer
    'SELECT Count([room_id]) AS [Total Beds] FROM tbl_cbk_room_inventory;
    Set rs = CurrentDb.OpenRecordset("SELECT Count([room_id]) AS [Total Beds] FROM tbl_cbk_room_inventory;", dbOpenSnapshot, dbReadOnly)
    Debug.Print rs![Total Beds]
    
'    cntr_i = 0
'    Do While cntr_i < cycle_limit
'        dummy_i = 99999
'        dummy_j = 999
'        j = dummy_i / dummy_j
'        cntr_i = cntr_i + 1
'    Loop
    
    'duration = Round((Timer - starttime) / IIf(cycle_limit > 100000, 1000, 1), 4)
    duration = Round((Timer - starttime), 4)
    Debug.Print Format("SpeedTest_DirectQuery: ", "!" & String(50, "@")), duration & " seconds"
''''''''''''''''''''''''
''''''''''''''''''''''''''
    starttime = Timer
    
    Set rs = CurrentDb.OpenRecordset("qry_cbk_db_total_beds", dbOpenSnapshot, dbReadOnly)
    Debug.Print rs![Total Beds]
'    For cntr_i = 0 To cycle_limit
'        dummy_i = 99999
'        dummy_j = 999
'        j = dummy_i / dummy_j
'    Next
    
    'duration = Round((Timer - starttime) / IIf(cycle_limit > 100000, 1000, 1), 4)
    duration = Round((Timer - starttime), 4)
    Debug.Print Format("SpeedTest_SavedQuery: ", "!" & String(50, "@")), duration & " seconds"
    
    
'    starttime = Timer
'
'    cntr_i = 0
'    While cntr_i < cycle_limit
'        dummy_i = 99999
'        dummy_j = 999
'        j = dummy_i / dummy_j
'        cntr_i = cntr_i + 1
'    Wend
'
'    'duration = Round((Timer - starttime) / IIf(cycle_limit > 100000, 1000, 1), 4)
'    duration = Round((Timer - starttime), 4)
'    Debug.Print Format("SpeedTest_While: ", "!" & String(50, "@")), duration & " seconds"

End Sub


Public Function FF_ListFilesInDir(sPath As String, Optional sFilter As String = "*") As Variant
    Dim aFiles()              As String
    Dim sFile                 As String
    Dim i                     As Long
 
    'On Error GoTo Error_Handler
 
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sFile = Dir(sPath & "*." & sFilter)
    Do While sFile <> vbNullString
        If sFile <> "." And sFile <> ".." Then
            ReDim Preserve aFiles(i)
            aFiles(i) = sFile
            MyPrint sFile
            i = i + 1
        End If
        sFile = Dir     'Loop through the next file that was found
    Loop
    FF_ListFilesInDir = aFiles
 
Exit_Handler:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: FF_ListFilesInDir" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Exit_Handler
End Function

Sub test2()
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fileexists-method
    
    Dim fs As New FileSystemObject
    
    'If fs.FileExists("GMS_schedule.accdb") Then
    'End If
    Debug.Print fs.FileExists("GMS_schedule.accdb")
End Sub

' VBA Notes
' https://support.microsoft.com/en-us/office/show-trust-by-adding-a-digital-signature-5f4ebff3-360d-4b61-b2f8-ce0dfb53adf6#bm2
' TempVars can only handle singular, TempVars cannot handle objects
' IIf will pre-process the conditional branch, so the conditional branch that will throw Error will be triggered
'   e.g.
'   iif(isnull(a), 0, cint(a))
'   will trigger error if a is null, even though the logic of the statement is intended to avoid the error.
'
' Order of events for database objects
' https://support.microsoft.com/en-us/office/order-of-events-for-database-objects-e76fbbfe-6180-4a52-8787-ce86553682f9#bm1

' regEx Reference
' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ms974570(v=msdn.10)?redirectedfrom=MSDN

'Timer Reference
'https://www.thespreadsheetguru.com/vba/2015/1/28/vba-calculate-macro-run-time

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/raise-method
' https://trumpexcel.com/vba-error-handling/#Err-Clear-Method
' https://stackoverflow.com/questions/39378307/saving-control-values-to-class-object-property
' https://excelhelphq.com/how-to-move-and-click-the-mouse-in-vba/
' https://support.microsoft.com/en-us/office/build-an-access-database-to-share-on-the-web-cca08e35-8e51-45ce-9269-8942b0deab26
' https://stackoverflow.com/questions/31815681/can-an-excel-vba-dictionary-be-used-to-call-a-function
